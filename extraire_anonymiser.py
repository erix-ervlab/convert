#!/usr/bin/env python3
"""
Extrait le texte de documents Word (et PDF) en Markdown structuré, avec OCR
des images, et pseudonymise une liste de termes, pour analyse par IA.

Table de pseudonymisation (pseudonymisation.csv, séparateur ;) :
    terme;remplacement;mode;categorie;commentaire
    ACME;ORG-01;prefixe;ORG;          -> ACME, ACMEBOX, acme.example, ACME's
    CLIENTX;;mot;CLIENT;              -> code auto CLIENT-01, enregistré dans la table

Usage :
    # essayer la table sur une phrase, avec le rendu Markdown
    python extraire_anonymiser.py --table pseudonymisation.csv --essai "Le ACMEBOX de ACME"

    # essai sur 5 fichiers, avec le détail de chaque remplacement (contextes.csv)
    python extraire_anonymiser.py src/ out/ --table pseudonymisation.csv --limite 5 --contexte

    # traitement complet
    python extraire_anonymiser.py src/ out/ --table pseudonymisation.csv --workers 8
    python extraire_anonymiser.py src/ out/ --table pseudonymisation.csv --sans-ocr

Prérequis :
    pandoc                                  -> .docx
    pip install pymupdf4llm                 -> .pdf
    Tesseract + pack français (fra)         -> OCR
    pip install pytesseract pillow
    LibreOffice (optionnel)                 -> OCR des images .emf/.wmf

Sorties : un .md par document (même arborescence), rapport.csv (par fichier),
occurrences.csv (par fichier et par terme), contextes.csv (option --contexte).
"""
import argparse
import csv
import hashlib
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile
from collections import Counter
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

# ---------------------------------------------------------------- pandoc ----
# Met les tableaux au format Markdown, remplace chaque image par un marqueur ⟦IMG⁞chemin⁞légende⟧ (traité ensuite
# en Python) et supprime la table des matières.
FILTRE_LUA = r'''
-- Tableaux : chaque cellule tient sur une ligne (paragraphes, listes et
-- sauts de ligne -> <br>, notes de bas de page renvoyées sous le tableau),
-- pour obtenir un tableau Markdown au lieu d'un tableau HTML.
-- Aplatit chaque cellule en une seule ligne (paragraphes et listes séparés par <br>)
local function aplatir(blocks)
  local inl = {}
  local function ajouter(xs)
    if #inl > 0 then table.insert(inl, pandoc.RawInline("html", "<br>")) end
    for _, x in ipairs(xs) do table.insert(inl, x) end
  end
  local function parcourir(bs, puce)
    for _, b in ipairs(bs) do
      if b.t == "Para" or b.t == "Plain" then
        local xs = pandoc.List(b.content)
        if puce then xs:insert(1, pandoc.Str(puce .. " ")) end
        ajouter(xs)
      elseif b.t == "BulletList" or b.t == "OrderedList" then
        for _, item in ipairs(b.content) do parcourir(item, "•") end
      elseif b.t == "Header" then ajouter(b.content)
      elseif b.t == "Div" or b.t == "BlockQuote" then parcourir(b.content, puce)
      elseif b.t == "LineBlock" then for _, l in ipairs(b.content) do ajouter(l) end
      elseif b.t == "Table" then ajouter({pandoc.Str(pandoc.utils.stringify(b))})
      end
    end
  end
  parcourir(blocks, nil)
  return {pandoc.Plain(inl)}
end
function Table(t)
  local notes = {}
  local function nettoyer_cellule(bs)
    return pandoc.walk_block(pandoc.Div(bs), {
      LineBreak = function() return pandoc.RawInline("html", "<br>") end,
      SoftBreak = function() return pandoc.Space() end,
      Note = function(n)
        table.insert(notes, pandoc.utils.stringify(n.content))
        return pandoc.Str("(" .. #notes .. ")")
      end,
    }).content
  end
  local function faire(rows) for _, r in ipairs(rows) do for _, c in ipairs(r.cells) do c.contents = aplatir(nettoyer_cellule(c.contents)) end end end
  faire(t.head.rows)
  for _, b in ipairs(t.bodies) do faire(b.head); faire(b.body) end
  faire(t.foot.rows)
  for i, spec in ipairs(t.colspecs) do t.colspecs[i] = {spec[1], pandoc.ColWidthDefault} end
  local sortie = {t}
  for i, n in ipairs(notes) do table.insert(sortie, pandoc.Para({pandoc.Str("(" .. i .. ") " .. n)})) end
  return sortie
end

function Image(img)
  local alt = pandoc.utils.stringify(img.caption):gsub("⟧", " ")
  return pandoc.Str("⟦IMG⁞" .. img.src .. "⁞" .. alt .. "⟧")
end

function Pandoc(doc)
  local out, dans_toc = {}, false
  for _, b in ipairs(doc.blocks) do
    if b.t == "Header" then
      local t = pandoc.utils.stringify(b):lower()
      dans_toc = t:match("table des mati") or t:match("^contents") or t:match("^sommaire")
                 or t:match("table of contents")
    end
    if not dans_toc then table.insert(out, b) end
  end
  return pandoc.Pandoc(out, doc.meta)
end
'''
MARQUEUR = re.compile(r"⟦IMG⁞([^⁞⟧]*)⁞([^⟧]*)⟧")


def legende_utile(alt):
    alt = alt.strip()
    return alt and not re.match(r"^([A-Za-z]:\\|.*\.(png|jpe?g|emf|wmf)$|Une image contenant|"
                                r"A picture containing|Graphique|Image|Picture \d)", alt, re.I)


# ------------------------------------------------------------------- OCR ----
class OCR:
    def __init__(self, langue, confiance, taille_min):
        import pytesseract
        from PIL import Image
        self.tess, self.Image = pytesseract, Image
        dispo = set(pytesseract.get_languages())
        voulues = [l for l in langue.split("+") if l in dispo]
        if not voulues:
            raise RuntimeError(f"Aucune langue Tesseract parmi {langue} (dispo : {sorted(dispo)})")
        self.langue, self.conf, self.taille_min = "+".join(voulues), confiance, taille_min
        self.cache = {}

    def lire(self, octets):
        cle = hashlib.md5(octets).hexdigest()
        if cle not in self.cache:
            self.cache[cle] = self._lire(octets)
        return self.cache[cle]

    def _lire(self, octets):
        try:
            img = self.Image.open(io.BytesIO(octets))
            img.load()
        except Exception:
            return ""
        if min(img.size) < self.taille_min:          # icônes, puces, logos minuscules
            return ""
        img = img.convert("L")
        if img.width < 1500:                          # agrandir aide beaucoup Tesseract
            f = 1500 / img.width
            img = img.resize((int(img.width * f), int(img.height * f)))
        d = self.tess.image_to_data(img, lang=self.langue, config="--psm 3",
                                    output_type=self.tess.Output.DICT)
        lignes = {}
        for i, mot in enumerate(d["text"]):
            if not mot.strip():
                continue
            cle = (d["block_num"][i], d["par_num"][i], d["line_num"][i])
            lignes.setdefault(cle, []).append((mot, float(d["conf"][i])))
        gardees, total = [], 0
        for cle in sorted(lignes):
            mots = lignes[cle]
            bons = [m for m, c in mots if c >= self.conf and re.search(r"\w", m)]
            # une ligne n'est gardée que si elle est majoritairement lisible
            if bons and len(bons) >= 0.6 * len(mots) and any(len(re.sub(r"\W", "", m)) >= 3 for m in bons):
                gardees.append(" ".join(bons))
                total += len(bons)
        return "\n".join(gardees) if total >= 3 else ""


def emf_vers_png(chemins, dossier):
    """Convertit .emf/.wmf en .png via LibreOffice (si installé)."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice or not chemins:
        return {}
    profil = Path(dossier) / "lo_profil"
    subprocess.run([soffice, f"-env:UserInstallation=file://{profil}", "--headless",
                    "--convert-to", "png", "--outdir", str(dossier), *map(str, chemins)],
                   capture_output=True, timeout=300)
    return {c: Path(dossier) / (Path(c).stem + ".png") for c in chemins
            if (Path(dossier) / (Path(c).stem + ".png")).exists()}


def bloc_ocr(texte, en_ligne):
    if en_ligne:   # dans un tableau : tout sur une ligne
        return "[Texte d'image : " + " / ".join(texte.splitlines()) + "]"
    return ("\n\n> *Texte extrait d'une image (OCR)*\n"
            + "\n".join("> " + l for l in texte.splitlines()) + "\n\n")


# ------------------------------------------------------------------ DOCX ----
def docx_vers_md(chemin, filtre, ocr):
    with tempfile.TemporaryDirectory() as tmp:
        r = subprocess.run(
            ["pandoc", str(chemin), "-t", "gfm", "--wrap=none", "--track-changes=accept",
             "--extract-media", tmp, "--lua-filter", filtre],
            capture_output=True, text=True, encoding="utf-8")
        if r.returncode:
            raise RuntimeError(r.stderr.strip()[:300])
        md = r.stdout
        nb_images = 0

        def fichier(src):   # selon la version de pandoc, chemin absolu ou relatif
            src = re.sub(r"\\(.)", r"\1", src)   # échappements Markdown
            p = Path(src)
            return p if p.exists() else Path(tmp) / src

        vecteurs = [m for m in set(MARQUEUR.findall(md))
                    if m[0].lower().endswith((".emf", ".wmf")) and fichier(m[0]).exists()]
        convertis = emf_vers_png([fichier(m[0]) for m in vecteurs], tmp) if ocr else {}

        def remplacer(m, ligne):
            nonlocal nb_images
            src, alt = m.group(1), m.group(2)
            texte = ""
            if ocr:
                p = convertis.get(fichier(src), fichier(src))
                if p.exists():
                    texte = ocr.lire(p.read_bytes())
            if texte:
                nb_images += 1
                en_ligne = ligne.lstrip().startswith("|") or "<td" in ligne or "<th" in ligne
                return bloc_ocr(texte, en_ligne)
            return f"[Image : {alt.strip()}]" if legende_utile(alt) else ""

        sortie = []
        for ligne in md.splitlines():
            sortie.append(MARQUEUR.sub(lambda m: remplacer(m, ligne), ligne))
        return "\n".join(sortie), nb_images


# ------------------------------------------------------------------- PDF ----
def pdf_vers_md(chemin, ocr):
    import pymupdf
    import pymupdf4llm
    doc = pymupdf.open(chemin)
    pages = pymupdf4llm.to_markdown(doc, page_chunks=True, ignore_images=True,
                                    ignore_graphics=True, show_progress=False,
                                    use_ocr=False, header=False, footer=False)
    # images présentes sur la moitié des pages ou plus = logo / décor
    freq = Counter(x[0] for p in doc for x in p.get_images(full=True))
    decor = {x for x, n in freq.items() if n >= max(2, len(doc) / 2)}
    morceaux, nb_images, vus = [], 0, set()
    for num, chunk in enumerate(pages):
        morceaux.append(chunk["text"])
        if not ocr:
            continue
        for x in doc[num].get_images(full=True):
            xref = x[0]
            if xref in decor or xref in vus:
                continue
            vus.add(xref)
            try:
                octets = doc.extract_image(xref)["image"]
            except Exception:
                continue
            texte = ocr.lire(octets)
            if texte:
                nb_images += 1
                morceaux.append(bloc_ocr(texte, False))
    return "\n\n".join(morceaux), nb_images


# --------------------------------------------------------------- commun ----
def retirer_repetitions(md, seuil=3):
    """PDF : retire les lignes répétées (pieds de page, bandeaux)."""
    lignes = md.splitlines()
    c = Counter(l.strip() for l in lignes if len(l.strip()) > 25 and not l.startswith(("|", ">")))
    return "\n".join(l for l in lignes if c.get(l.strip(), 0) < seuil)


def nettoyer_tableaux(md):
    """Supprime les tableaux vides (grilles de mise en page) et compacte les espaces."""
    sortie, bloc = [], []

    def vider():
        # en-tête vide (tableau Word sans ligne d'en-tête) : la 1re ligne devient l'en-tête
        if len(bloc) > 2 and not re.search(r"\w", bloc[0]) and re.match(r"^\|[-:| ]+\|$", bloc[1]):
            bloc[:] = [bloc[2], bloc[1]] + bloc[3:]
        if bloc and any(re.search(r"\w", l) for l in bloc if not re.match(r"^\|[-:| ]+\|$", l)):
            sortie.extend(re.sub(r"-{3,}", "---", re.sub(r" {2,}", " ", l)) for l in bloc)
        bloc.clear()

    for ligne in md.splitlines():
        if ligne.startswith("|"):
            bloc.append(ligne)
        else:
            vider()
            sortie.append(ligne)
    vider()
    return "\n".join(sortie)


def nettoyer(md):
    md = nettoyer_tableaux(md)
    md = re.sub(r"<img [^>]*>", "", md)
    md = re.sub(r"<!-- Start of picture text -->.*?<!-- End of picture text -->", "", md, flags=re.S)
    md = re.sub(r"\n{3,}", "\n\n", md)
    return md.strip() + "\n"


# ------------------------------------------------------ pseudonymisation ----
MODES = ("prefixe", "mot", "partout", "regex")
CARS_MARKDOWN = re.compile(r"[\\`*_\[\]<>#|~]")   # { } sans effet en Markdown : autorisées
BLOC_OCR = re.compile(r"> \*Texte extrait d'une image \(OCR\)\*\n(?:> .*\n?)+|\[Texte d'image : [^\]]*\]")
AVANT = r"(?<![^\W_])"          # pas précédé d'une lettre ou d'un chiffre
APRES = r"(?![^\W_])"           # pas suivi d'une lettre ou d'un chiffre


def lire_table(chemin):
    """Lit la table CSV (séparateur ;). Les lignes commençant par # sont ignorées.
    Colonnes : terme ; remplacement ; mode ; categorie ; commentaire"""
    with open(chemin, encoding="utf-8-sig", newline="") as f:
        lignes = [l for l in f if l.strip() and not l.lstrip().startswith("#")]
    entrees = list(csv.DictReader(lignes, delimiter=";"))
    erreurs = []
    for i, e in enumerate(entrees, 2):
        e["terme"] = (e.get("terme") or "").strip()
        e["remplacement"] = (e.get("remplacement") or "").strip()
        e["mode"] = (e.get("mode") or "prefixe").strip().lower()
        e["categorie"] = (e.get("categorie") or "").strip()
        if not e["terme"]:
            erreurs.append(f"ligne {i} : terme vide")
        if e["mode"] not in MODES:
            erreurs.append(f"ligne {i} : mode '{e['mode']}' inconnu ({', '.join(MODES)})")
        if e["mode"] == "regex":
            try:
                re.compile(e["terme"])
            except re.error as x:
                erreurs.append(f"ligne {i} : regex invalide ({x})")
    if erreurs:
        sys.exit("Table de pseudonymisation invalide :\n  " + "\n  ".join(erreurs))
    return entrees


def completer_codes(entrees, chemin):
    """Attribue un code stable (CATEGORIE-01, -02…) aux termes sans remplacement,
    et réécrit la table pour que les codes restent identiques aux prochains passages."""
    pris = {e["remplacement"] for e in entrees if e["remplacement"]}
    ajout = False
    for e in entrees:
        if e["remplacement"]:
            continue
        prefixe = (e["categorie"] or "TERME").upper()
        n = 1
        while f"{prefixe}-{n:02d}" in pris:
            n += 1
        e["remplacement"] = f"{prefixe}-{n:02d}"
        pris.add(e["remplacement"])
        ajout = True
    if ajout:
        with open(chemin, encoding="utf-8-sig") as f:
            commentaires = [l.rstrip("\n") for l in f if l.lstrip().startswith("#")]
        with open(chemin, "w", encoding="utf-8-sig", newline="") as f:
            for c in commentaires:
                f.write(c + "\n")
            w = csv.writer(f, delimiter=";")
            w.writerow(["terme", "remplacement", "mode", "categorie", "commentaire"])
            for e in entrees:
                w.writerow([e["terme"], e["remplacement"], e["mode"], e["categorie"], e.get("commentaire") or ""])
    return ajout


class Pseudonymiseur:
    def __init__(self, entrees, echapper=False, forme="{}"):
        # termes les plus longs d'abord : ACMEBOX passe avant ACME s'ils sont tous deux dans la table
        self.entrees = sorted(entrees, key=lambda e: (e["mode"] == "regex", -len(e["terme"])))
        self.echapper = echapper
        self.forme = forme
        self.rx = self._compiler(ocr=False)
        self.rx_ocr = self._compiler(ocr=True)
        # contrôle des restes : chaque terme cherché n'importe où
        self.rx_restes = [(e, re.compile(self._motif(e, "partout"), re.I)) for e in self.entrees]

    @staticmethod
    def _motif(e, mode=None):
        mode = mode or e["mode"]
        if e["mode"] == "regex":
            return e["terme"]
        t = r"\s+".join(re.escape(x) for x in e["terme"].split())   # plusieurs mots : espaces souples
        return {"prefixe": AVANT + t, "mot": AVANT + t + APRES, "partout": t}[mode]

    def _compiler(self, ocr):
        # un seul passage avec une alternance : un remplacement n'est jamais re-remplacé
        parties = [f"(?P<t{i}>{self._motif(e, 'partout' if ocr and e['mode'] != 'regex' else None)})"
                   for i, e in enumerate(self.entrees)]
        return re.compile("|".join(parties), re.I)

    def jeton(self, e):
        r = self.forme.replace("{}", e["remplacement"])      # ex. _{}_ -> _ORG-01_
        return CARS_MARKDOWN.sub(lambda m: "\\" + m.group(0), r) if self.echapper else r

    def appliquer(self, texte, ocr=False, journal=None, decalage=0):
        compte = Counter()

        def remplacer(m):
            e = self.entrees[int(m.lastgroup[1:])]
            compte[e["terme"]] += 1
            if journal is not None:
                journal.append((decalage + m.start(), decalage + m.end(), m.group(0), e, ocr))
            return self.jeton(e)

        return (self.rx_ocr if ocr else self.rx).sub(remplacer, texte), compte

    def document(self, md, journal=None):
        """Texte normal : selon le mode de chaque terme. Texte OCR : recherche partout
        (l'OCR colle souvent des caractères parasites aux mots)."""
        total, morceaux, pos = Counter(), [], 0
        for m in BLOC_OCR.finditer(md):
            for debut, fin, ocr in ((pos, m.start(), False), (m.start(), m.end(), True)):
                t, c = self.appliquer(md[debut:fin], ocr, journal, debut)
                morceaux.append(t)
                total += c
            pos = m.end()
        t, c = self.appliquer(md[pos:], False, journal, pos)
        morceaux.append(t)
        return "".join(morceaux), total + c

    def restes(self, texte):
        """Termes encore présents n'importe où (hors remplacements eux-mêmes)."""
        for e in self.entrees:
            for r in {e["remplacement"], self.jeton(e)}:
                texte = re.sub(re.escape(r), " ", texte, flags=re.I)
        return Counter({e["terme"]: len(rx.findall(texte)) for e, rx in self.rx_restes
                        if rx.findall(texte)})


def avertir_markdown(entrees, echapper, forme):
    risques = [forme.replace("{}", e["remplacement"]) for e in entrees]
    risques = [r for r in risques if CARS_MARKDOWN.search(r)]
    if risques and not echapper:
        print("ATTENTION : ces remplacements contiennent des caractères Markdown "
              f"({', '.join(sorted(set(risques))[:5])}).\n"
              "  Dans un visualiseur Markdown ils peuvent s'afficher autrement (ex. _AN_ -> AN en italique).\n"
              "  Ajoutez --echapper, ou utilisez une forme sans _ * [ ] # | (ex. --forme '{{}}').\n")


def note_pseudonymisation(forme, entrees):
    exemple = forme.replace("{}", entrees[0]["remplacement"])
    return ("> **Note : document pseudonymisé.** Certains noms (organisations, clients, "
            f"équipements…) ont été remplacés par des codes de la forme `{exemple}`. "
            "Un même code désigne toujours la même entité, dans ce document et dans les autres.\n\n")


def rendu_markdown(texte):
    """Ce qu'affiche un visualiseur Markdown (via pandoc), pour les essais."""
    if not shutil.which("pandoc"):
        return "(pandoc absent : rendu non disponible)"
    r = subprocess.run(["pandoc", "-f", "gfm", "-t", "plain", "--wrap=none"],
                       input=texte, capture_output=True, text=True, encoding="utf-8")
    return r.stdout.strip()


# ------------------------------------------------------------ traitement ----
def extrait(texte, debut, fin, marge=40):
    return texte[max(0, debut - marge):fin + marge].replace("\n", " ⏎ ")


def traiter(chemin, src, dst, a, filtre, entrees):
    rel = chemin.relative_to(src)
    try:
        pseudo = Pseudonymiseur(entrees, a.echapper, a.forme)
        ocr = None if a.sans_ocr else OCR(a.ocr_langue, a.ocr_confiance, a.ocr_taille_min)
        if chemin.suffix.lower() == ".docx":
            md, nb_img = docx_vers_md(chemin, filtre, ocr)
        else:
            md, nb_img = pdf_vers_md(chemin, ocr)
            md = retirer_repetitions(md)
        md = nettoyer(md)
        journal = [] if a.contexte else None
        resultat, compte = pseudo.document(md, journal)
        if a.note:
            resultat = note_pseudonymisation(a.forme, entrees) + resultat
        restes = pseudo.restes(resultat)
        contextes = []
        for debut, fin, trouve, e, ocr in journal or []:
            avant = extrait(md, debut, fin)
            contextes.append([str(rel), md.count("\n", 0, debut) + 1, "image (OCR)" if ocr else "texte",
                              e["terme"], trouve, e["remplacement"], avant,
                              pseudo.appliquer(avant, ocr)[0]])

        # noms de fichiers : code brut, sans habillage ni échappement (\ interdit sous Windows)
        nom = Pseudonymiseur(entrees).appliquer(str(rel.with_suffix(".md")), ocr=True)[0]
        sortie = dst / nom
        sortie.parent.mkdir(parents=True, exist_ok=True)
        sortie.write_text(resultat, encoding="utf-8")
        return {"ligne": [str(rel), "ok", sum(compte.values()), sum(restes.values()), nb_img, len(resultat), ""],
                "compte": compte, "restes": restes, "contextes": contextes}
    except Exception as e:
        return {"ligne": [str(rel), "erreur", 0, "", 0, 0, str(e)[:300]],
                "compte": Counter(), "restes": Counter(), "contextes": []}


def mode_essai(texte, entrees, echapper, forme):
    pseudo = Pseudonymiseur(entrees, echapper, forme)
    resultat, compte = pseudo.appliquer(texte)
    print("Entrée      :", texte)
    print("Sortie brute:", resultat)
    print("Affichage   :", rendu_markdown(resultat))
    print("Remplacements :", dict(compte) or "aucun")
    restes = pseudo.restes(resultat)
    if restes:
        print("Termes encore présents :", dict(restes))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("source", type=Path, nargs="?")
    ap.add_argument("sortie", type=Path, nargs="?")
    ap.add_argument("--table", type=Path, help="table de pseudonymisation (CSV ;)")
    ap.add_argument("--mot", help="sans --table : terme unique à remplacer")
    ap.add_argument("--jeton", default="ORG-01", help="sans --table : remplacement")
    ap.add_argument("--forme", default="{}",
                    help="habillage des codes pour signaler un remplacement, ex. '_{}_' ou '[{}]'")
    ap.add_argument("--note", action="store_true",
                    help="ajoute en tête de chaque fichier une note expliquant la pseudonymisation")
    ap.add_argument("--echapper", action="store_true",
                    help="échappe les caractères Markdown des remplacements (\\_AN\\_)")
    ap.add_argument("--essai", metavar="TEXTE", help="teste la table sur un texte et affiche le résultat")
    ap.add_argument("--limite", type=int, help="ne traite que les N premiers fichiers (essais)")
    ap.add_argument("--contexte", action="store_true",
                    help="écrit contextes.csv : chaque occurrence avec son extrait, pour contrôle")
    ap.add_argument("--workers", type=int, default=max(1, (os.cpu_count() or 2) // 2))
    ap.add_argument("--sans-ocr", action="store_true", help="ne pas lire le texte des images")
    ap.add_argument("--ocr-langue", default="fra+eng")
    ap.add_argument("--ocr-confiance", type=float, default=70,
                    help="confiance mini par mot (0-100) ; plus haut = moins de bruit")
    ap.add_argument("--ocr-taille-min", type=int, default=80,
                    help="ignore les images dont un côté fait moins de N pixels")
    a = ap.parse_args()

    if a.table:
        entrees = lire_table(a.table)
        if completer_codes(entrees, a.table):
            print(f"Codes attribués et enregistrés dans {a.table}")
    elif a.mot:
        entrees = [{"terme": a.mot, "remplacement": a.jeton, "mode": "prefixe", "categorie": ""}]
    else:
        ap.error("indiquez une table (--table) ou un terme (--mot)")
    if not entrees:
        sys.exit("La table de pseudonymisation est vide.")
    if "{}" not in a.forme:
        ap.error("--forme doit contenir {} (emplacement du code), ex. '_{}_'")
    avertir_markdown(entrees, a.echapper, a.forme)

    if a.essai is not None:
        mode_essai(a.essai, entrees, a.echapper, a.forme)
        return
    if not a.source or not a.sortie:
        ap.error("indiquez le dossier source et le dossier de sortie (ou --essai)")

    fichiers = sorted(p for p in a.source.rglob("*")
                      if p.suffix.lower() in (".docx", ".pdf") and not p.name.startswith("~$"))
    if a.limite:
        fichiers = fichiers[:a.limite]
    if not fichiers:
        sys.exit("Aucun .docx ou .pdf trouvé.")
    a.sortie.mkdir(parents=True, exist_ok=True)

    with tempfile.NamedTemporaryFile("w", suffix=".lua", delete=False, encoding="utf-8") as f:
        f.write(FILTRE_LUA)
        filtre = f.name

    resultats = []
    with ProcessPoolExecutor(a.workers) as ex:
        taches = [ex.submit(traiter, p, a.source, a.sortie, a, filtre, entrees) for p in fichiers]
        for i, t in enumerate(as_completed(taches), 1):
            r = t.result()
            resultats.append(r)
            l = r["ligne"]
            alerte = f"  /!\\ {l[3]} restant(s)" if l[3] else ""
            print(f"[{i}/{len(fichiers)}] {l[1]:6} {l[2]:5} rempl. {l[4]:3} img OCR  {l[0]}{alerte}  {l[6]}")

    with open(a.sortie / "rapport.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["fichier", "statut", "remplacements", "restants", "images_ocr", "caracteres", "erreur"])
        w.writerows(sorted(r["ligne"] for r in resultats))

    # détail par terme et par fichier
    with open(a.sortie / "occurrences.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["fichier", "terme", "remplacement", "remplacees", "restantes"])
        for r in sorted(resultats, key=lambda r: r["ligne"][0]):
            for e in entrees:
                n, k = r["compte"].get(e["terme"], 0), r["restes"].get(e["terme"], 0)
                if n or k:
                    w.writerow([r["ligne"][0], e["terme"], e["remplacement"], n, k])

    if a.contexte:
        with open(a.sortie / "contextes.csv", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["fichier", "ligne", "origine", "terme", "trouve", "remplacement", "avant", "apres"])
            for r in sorted(resultats, key=lambda r: r["ligne"][0]):
                w.writerows(r["contextes"])

    ok = [r["ligne"] for r in resultats if r["ligne"][1] == "ok"]
    total = Counter()
    for r in resultats:
        total += r["compte"]
    print(f"\n{len(ok)}/{len(resultats)} fichiers traités, {sum(l[2] for l in ok)} remplacements, "
          f"{sum(l[4] for l in ok)} images lues.")
    for e in entrees:
        print(f"  {e['terme']:25} -> {e['remplacement']:15} {total.get(e['terme'], 0):6}")
    print(f"Rapports : {a.sortie / 'rapport.csv'}, occurrences.csv" + (", contextes.csv" if a.contexte else ""))


if __name__ == "__main__":
    main()
