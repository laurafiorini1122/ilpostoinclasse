"""
Genera il sito "Il mio posto c'è sempre" partendo dall'Excel dei soggetti
e dai file di testo delle cartelle testi/ e idee/.

Come si usa: doppio clic su aggiorna_sito.bat (oppure: python genera.py).
Si può anche indicare un Excel preciso: python genera.py "percorso\\file.xlsx"

Lo script:
- legge il foglio "Soggetti" (solo le colonne elencate in COLONNE qui sotto:
  le note, lo stato e i prompt non escono mai);
- legge i testi delle pagine fisse da testi/ e le idee da idee/;
- crea in idee/ i file che mancano (uno per soggetto), ma non tocca mai quelli che ci sono;
- ricrea da zero la cartella sito/ (tranne i disegni, che copia da disegni/);
- non modifica mai l'Excel.

Usa solo Python standard: non serve installare niente.
"""

import html
import json
import re
import shutil
import sys
import time
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

# ---------------------------------------------------------------------------
# IMPOSTAZIONI
# ---------------------------------------------------------------------------

# Cartella su Google Drive con l'Excel. Lo script prende da sola la versione
# più recente (Quaderno_classe_soggetti_v7, v8, v9...).
CARTELLA_EXCEL = Path(r"G:\Il mio Drive\0 Spazio flessibile\8 SiO - Scuola")
NOME_EXCEL = "Quaderno_classe_soggetti"
FOGLIO = "Soggetti"

QUI = Path(__file__).resolve().parent
SITO = QUI / "sito"
MODELLO = QUI / "modello"
DISEGNI = QUI / "disegni"
TESTI = QUI / "testi"
IDEE = QUI / "idee"

# Intestazioni delle colonne da leggere (basta l'inizio del testo).
# Così, se in futuro inserisci una colonna nell'Excel, lo script la ritrova.
COLONNE = {
    "numero": "N.",
    "titolo": "Soggetto",
    "tipo": "Tipo",
    "stimolo": "Frase stimolo",
    "classe": "Che cosa fa la classe",
    "titolo_quaderno": "Titolo della pagina da incollare",
    "indirizzo": "Indirizzo sul sito",
    "modo": "Come lavora la classe",
    "arriva": "Che cosa arriva al bambino",
}

# ---------------------------------------------------------------------------
# TESTI FISSI DEL SITO (si possono modificare qui)
# I testi lunghi delle pagine stanno invece nella cartella testi/.
# ---------------------------------------------------------------------------

NOME_SITO = "Il mio posto c'è sempre"
SOTTOTITOLO = "Un quaderno da colorare a staffetta con la mia classe"
AUTRICE = "Laura Fiorini"

MENU = [
    # (voce del menu, cartella, chiave)
    ("Il quaderno", "", "home"),
    ("Come funziona", "come-funziona/", "come-funziona"),
    ("I soggetti", "classe/", "elenco"),
    ("Per cominciare", "per-cominciare/", "per-cominciare"),
    ("Idee e consigli", "idee-e-consigli/", "idee-e-consigli"),
    ("Come averlo", "come-averlo/", "come-averlo"),
]

INGRESSO_SOGGETTO = "Cerco un soggetto"
INGRESSO_INIZIO = "Accompagno un bambino: da dove comincio"

INTRO_ELENCO = ("Cerca il numero e il titolo che ti hanno comunicato la famiglia "
                "o il docente della scuola in ospedale e apri la pagina del soggetto.")

IDEE_TITOLO = "Idee dalle classi"
IDEE_TESTO = (
    "Qui raccogliamo le varianti di questo soggetto suggerite da chi l'ha già proposto "
    "in classe: come adattarlo all'età dei bambini, a una classe numerosa o ai materiali "
    "che si hanno. Hai fatto questo soggetto in un modo diverso? Raccontalo: le istruzioni "
    "sono nella pagina {link}."
)
IDEE_VUOTO_SOGGETTO = "Non ci sono ancora idee per questo soggetto."
IDEE_VUOTO_GENERALE = "Le prime idee arriveranno presto."

FOGLIO_IN_PREPARAZIONE = "Foglio in preparazione"
RIGHE_FOGLIO_NOMI = 20

PIEDE_PDF = "Presentazione completa in PDF (in preparazione)"
PIEDE_STAMPA = "Le pagine si possono stampare per uso scolastico."
PIEDE_DATI = "Il sito non usa cookie e non raccoglie dati."

# ---------------------------------------------------------------------------
# LETTURA DELL'EXCEL (senza librerie esterne: un .xlsx è un archivio zip)
# ---------------------------------------------------------------------------

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
M = "{%s}" % NS["m"]


class ErroreExcel(Exception):
    pass


def trova_excel():
    if len(sys.argv) > 1:
        return Path(sys.argv[1])
    candidati = []
    for f in CARTELLA_EXCEL.glob(NOME_EXCEL + "*.xlsx"):
        if f.name.startswith("~$"):
            continue  # file temporaneo di Excel aperto
        m = re.search(r"_v(\d+)\.xlsx$", f.name)
        candidati.append((int(m.group(1)) if m else 0, f))
    if not candidati:
        raise ErroreExcel(f"Non trovo nessun file {NOME_EXCEL}*.xlsx in {CARTELLA_EXCEL}")
    return max(candidati)[1]


def testo_cella(c, condivisi):
    tipo = c.get("t")
    v = c.find("m:v", NS)
    if tipo == "s" and v is not None:
        return condivisi[int(v.text)]
    if tipo == "inlineStr":
        return "".join(t.text or "" for t in c.iter(M + "t"))
    return v.text if v is not None and v.text else ""


def leggi_righe(percorso):
    with zipfile.ZipFile(percorso) as z:
        condivisi = []
        if "xl/sharedStrings.xml" in z.namelist():
            radice = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in radice.findall("m:si", NS):
                # Testo normale e testo con formattazione; esclusa la fonetica (rPh).
                parti = [t.text or "" for t in si.findall("m:t", NS)]
                for r in si.findall("m:r", NS):
                    parti += [t.text or "" for t in r.findall("m:t", NS)]
                condivisi.append("".join(parti))

        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rel = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        destinazioni = {r.get("Id"): r.get("Target") for r in rel}
        percorso_foglio = None
        for s in wb.find("m:sheets", NS):
            if s.get("name") == FOGLIO:
                t = destinazioni[s.get("{%s}id" % NS["r"])].lstrip("/")
                percorso_foglio = t if t.startswith("xl/") else "xl/" + t
        if not percorso_foglio:
            raise ErroreExcel(f"Nell'Excel non c'è il foglio «{FOGLIO}».")

        foglio = ET.fromstring(z.read(percorso_foglio))
        righe = []
        for riga in foglio.iter(M + "row"):
            celle = {}
            for c in riga.findall("m:c", NS):
                colonna = re.match(r"[A-Z]+", c.get("r")).group()
                celle[colonna] = testo_cella(c, condivisi).strip()
            righe.append((int(riga.get("r")), celle))
        return righe


def righe_di(testo):
    """Divide una cella in righe non vuote, con gli spazi sistemati."""
    righe = [re.sub(r"\s+", " ", r).strip() for r in testo.replace("\r", "").split("\n")]
    return [r for r in righe if r]


def passi_di(testo):
    """La colonna «Che cosa fa la classe»: righe «1. ...» diventano passi numerati,
    le altre righe (per esempio «Primo tempo: la caccia») diventano sottotitoli."""
    parti = []
    for r in righe_di(testo):
        m = re.match(r"(\d+)[.)]\s*(.+)", r)
        if m:
            parti.append(("passo", int(m.group(1)), m.group(2)))
        else:
            parti.append(("titolo", 0, r))
    return parti


def leggi_soggetti(percorso):
    righe = leggi_righe(percorso)
    if not righe:
        raise ErroreExcel("Il foglio Soggetti è vuoto.")
    _, intestazioni = righe[0]
    posizioni = {}
    for chiave, inizio in COLONNE.items():
        for col, testo in intestazioni.items():
            if testo.startswith(inizio) and not (chiave == "titolo" and testo != "Soggetto"):
                posizioni[chiave] = col
                break
        else:
            raise ErroreExcel(f"Non trovo la colonna «{inizio}» nella prima riga del foglio.")

    soggetti, errori, avvisi = [], [], []
    visti = {}
    for n_riga, celle in righe[1:]:
        numero = celle.get(posizioni["numero"], "")
        if not re.fullmatch(r"\d+(\.0)?", numero):
            continue  # righe di riepilogo o vuote
        dato = {k: celle.get(col, "") for k, col in posizioni.items()}
        m = re.search(r"classe/(\d{2})/?\s*$", dato["indirizzo"])
        if not m:
            errori.append(f"riga {n_riga}: l'indirizzo «{dato['indirizzo']}» non finisce con classe/ e due cifre")
            continue
        codice = m.group(1)
        if codice in visti:
            errori.append(f"riga {n_riga}: l'indirizzo classe/{codice} c'è già alla riga {visti[codice]}")
            continue
        visti[codice] = n_riga
        if int(float(numero)) != int(codice):
            avvisi.append(f"riga {n_riga}: il numero {numero} e l'indirizzo classe/{codice} non coincidono (sul sito vale l'indirizzo)")
        for campo in ("titolo", "classe"):
            if not dato[campo]:
                errori.append(f"riga {n_riga}: manca la colonna «{COLONNE[campo]}»")
        for campo in ("stimolo", "modo", "arriva", "titolo_quaderno"):
            if not dato[campo]:
                avvisi.append(f"riga {n_riga}: la colonna «{COLONNE[campo]}» è vuota")
        for campo in ("titolo", "stimolo", "classe", "titolo_quaderno", "modo", "arriva"):
            if re.search(r"[\u2014\u2013]", dato[campo]):
                avvisi.append(f"riga {n_riga}: nella colonna «{COLONNE[campo]}» c'è un trattino lungo")
        soggetti.append({
            "codice": codice,
            "titolo": dato["titolo"],
            "tipo": dato["tipo"],
            "stimolo": righe_di(dato["stimolo"]),
            "passi": passi_di(dato["classe"]),
            "titolo_quaderno": dato["titolo_quaderno"],
            "modo": dato["modo"],
            "arriva": righe_di(dato["arriva"]),
        })
    if errori:
        raise ErroreExcel("Da correggere nell'Excel:\n  - " + "\n  - ".join(errori))
    soggetti.sort(key=lambda s: s["codice"])
    return soggetti, avvisi


# ---------------------------------------------------------------------------
# FILE DI TESTO (cartelle testi/ e idee/)
# ---------------------------------------------------------------------------

def leggi_testo(percorso):
    """Legge un file di testo e toglie le righe di commento (quelle con un solo # all'inizio)."""
    if not percorso.exists():
        return []
    righe = percorso.read_text(encoding="utf-8-sig").replace("\r", "").split("\n")
    return [r.rstrip() for r in righe if not (r.startswith("#") and not r.startswith("##"))]


def leggi_descrizioni(avvisi):
    descrizioni = {}
    for r in leggi_testo(TESTI / "descrizioni.txt"):
        m = re.match(r"\s*(\d{1,2})\s*:\s*(.+)", r)
        if m:
            descrizioni[m.group(1).zfill(2)] = m.group(2).strip()
    return descrizioni


IDEE_INTESTAZIONE = """# {chi}
# Le righe che cominciano con # non compaiono sul sito.
#
# Come aggiungere un contributo:
#   1. scrivi una riga con il ruolo di chi lo ha mandato, così:
#        Ruolo: insegnante di classe
#   2. sotto, il testo (se è lungo, separa i paragrafi con una riga vuota);
#   3. fra un contributo e l'altro scrivi una riga con tre trattini:
#        ---
#   4. doppio clic su aggiorna_sito.bat.
# Finché qui sotto non c'è niente, sul sito compare la frase «{vuoto}»
# Mai nomi di bambini, scuole, ospedali o reparti, mai informazioni sulla salute.
#
# Esempio (le righe con # non escono sul sito):
# Ruolo: insegnante di classe
# Abbiamo usato questo soggetto con una classe prima e ...
# ---
# Ruolo: genitore
# Un altro contributo.

"""


def prepara_file_idee(soggetti):
    """Crea i file delle idee che mancano. Non modifica mai quelli che esistono già."""
    IDEE.mkdir(exist_ok=True)
    creati = []
    gen = IDEE / "generale.txt"
    if not gen.exists():
        gen.write_text(IDEE_INTESTAZIONE.format(
            chi="Consigli generali: compaiono nella pagina Idee e consigli.",
            vuoto=IDEE_VUOTO_GENERALE), encoding="utf-8", newline="\n")
        creati.append(gen.name)
    for s in soggetti:
        f = IDEE / f"{s['codice']}.txt"
        if not f.exists():
            f.write_text(IDEE_INTESTAZIONE.format(
                chi=f"Idee dalle classi per il soggetto {s['codice']}, {s['titolo']}.",
                vuoto=IDEE_VUOTO_SOGGETTO), encoding="utf-8", newline="\n")
            creati.append(f.name)
    return creati


def leggi_idee(percorso, avvisi):
    """Restituisce una lista di contributi: (ruolo, [paragrafi])."""
    contributi = []
    blocco = []
    for r in leggi_testo(percorso) + ["---"]:
        if r.strip() == "---":
            testo = "\n".join(blocco).strip()
            blocco = []
            if not testo:
                continue
            righe = testo.split("\n")
            ruolo = ""
            m = re.match(r"\s*ruolo\s*:\s*(.*)", righe[0], re.I)
            if m:
                ruolo = m.group(1).strip()
                righe = righe[1:]
            else:
                avvisi.append(f"idee/{percorso.name}: un contributo non comincia con «Ruolo:»")
            paragrafi = [re.sub(r"\s+", " ", p).strip() for p in re.split(r"\n\s*\n", "\n".join(righe))]
            paragrafi = [p for p in paragrafi if p]
            if paragrafi:
                contributi.append((ruolo, paragrafi))
        else:
            blocco.append(r)
    return contributi


# ---------------------------------------------------------------------------
# DAL TESTO SEMPLICE ALL'HTML
# ---------------------------------------------------------------------------

def e(testo):
    return html.escape(testo, quote=True)


def in_linea(testo):
    """Grassetto **così** e segnaposto [[così]] dentro una riga."""
    t = e(testo)
    t = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", t)
    t = re.sub(r"\[\[(.+?)\]\]", r'<span class="segnaposto-breve">\1</span>', t)
    return t


def ancora(testo):
    s = unicodedata.normalize("NFKD", testo).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "-", s).strip("-")


def separa_id(testo):
    """«Titolo {#ancora}» → ('Titolo', 'ancora')."""
    m = re.search(r"\s*\{#([\w-]+)\}\s*", testo)
    if not m:
        return testo.strip(), ""
    return (testo[:m.start()] + " " + testo[m.end():]).strip(), m.group(1)


def blocchi_html(righe, segnaposti=None):
    """Trasforma le righe di un file di testi/ in HTML."""
    segnaposti = segnaposti or {}
    out = []
    gruppi, attuale = [], []
    for r in righe + [""]:
        if r.strip():
            attuale.append(r.strip())
        elif attuale:
            gruppi.append(attuale)
            attuale = []
    for g in gruppi:
        primo = g[0]
        if len(g) == 1 and primo in segnaposti:
            out.append(segnaposti[primo])
        elif primo.startswith("### "):
            titolo, id_ = separa_id(primo[4:])
            out.append(f'<h3 id="{id_ or ancora(titolo)}">{in_linea(titolo)}</h3>')
        elif primo.startswith("## "):
            titolo, id_ = separa_id(primo[3:])
            out.append(f'<h2 id="{id_ or ancora(titolo)}">{in_linea(titolo)}</h2>')
        elif all(x.startswith("- ") for x in g):
            voci = []
            for x in g:
                testo, id_ = separa_id(x[2:])
                attr = f' id="{id_}"' if id_ else ""
                voci.append(f"<li{attr}>{in_linea(testo)}</li>")
            out.append('<ul class="voci">\n' + "\n".join(voci) + "\n</ul>")
        elif all(re.match(r"\d+\.\s", x) for x in g):
            inizio = int(re.match(r"(\d+)", primo).group(1))
            voci = "\n".join("<li>" + in_linea(re.sub(r"^\d+\.\s*", "", x)) + "</li>" for x in g)
            attr = f' start="{inizio}" style="counter-reset: passo {inizio - 1}"' if inizio != 1 else ""
            out.append(f'<ol class="passi"{attr}>\n{voci}\n</ol>')
        elif all(x.startswith("|") for x in g):
            celle = [[c.strip() for c in x.strip("|").split("|")] for x in g]
            testa = "".join(f'<th scope="col">{in_linea(c)}</th>' for c in celle[0])
            corpo = "\n".join(
                "<tr>" + "".join(
                    (f'<th scope="row">{in_linea(c)}</th>' if i == 0 else f"<td>{in_linea(c)}</td>")
                    for i, c in enumerate(riga)) + "</tr>"
                for riga in celle[1:])
            out.append(f'<div class="tabella"><table>\n<thead><tr>{testa}</tr></thead>\n<tbody>\n{corpo}\n</tbody>\n</table></div>')
        elif len(g) == 1 and re.fullmatch(r"\[\[(.+)\]\]", primo):
            out.append(f'<div class="segnaposto"><p>{e(primo[2:-2])}</p></div>')
        else:
            out.append(f"<p>{in_linea(' '.join(g))}</p>")
    return out


def pagina_a_riquadri(righe, segnaposti=None):
    """Il testo prima del primo titolo diventa l'introduzione; ogni «## » apre un riquadro."""
    pezzi = blocchi_html(righe, segnaposti)
    intro, sezioni = [], []
    for p in pezzi:
        if p.startswith("<h2"):
            sezioni.append([p])
        elif sezioni:
            sezioni[-1].append(p)
        else:
            intro.append(p)
    html_ = ""
    if intro and not sezioni:
        # pagina senza titoli interni: tutto il testo in un solo riquadro
        html_ += '<div class="riquadro">\n' + "\n".join(intro) + "\n</div>\n"
    elif intro:
        html_ += '<div class="intro">\n' + "\n".join(intro) + "\n</div>\n"
    for s in sezioni:
        id_ = re.search(r'id="([^"]+)"', s[0]).group(1)
        html_ += (f'<section class="riquadro" aria-labelledby="{id_}">\n'
                  + "\n".join(s) + "\n</section>\n")
    return html_


def contributi_html(contributi, vuoto):
    if not contributi:
        return f'<p class="vuoto">{e(vuoto)}</p>'
    voci = []
    for ruolo, paragrafi in contributi:
        testo = "\n".join(f"<p>{e(p)}</p>" for p in paragrafi)
        chi = f'\n<figcaption>{e(ruolo[:1].upper() + ruolo[1:])}</figcaption>' if ruolo else ""
        voci.append(f'<li><figure class="contributo"><blockquote>\n{testo}\n</blockquote>{chi}</figure></li>')
    return '<ul class="contributi">\n' + "\n".join(voci) + "\n</ul>"


# ---------------------------------------------------------------------------
# COSTRUZIONE DELLE PAGINE
# ---------------------------------------------------------------------------

def tema(tipo):
    t = tipo.lower()
    if "legame" in t:
        return "legame"
    if "vita" in t:
        return "vita"
    return "libero"


def ancora_modo(modo):
    t = modo.lower()
    if "cartellone" in t:
        return "cartellone"
    if "gruppi" in t:
        return "piccoli-gruppi"
    if "tutta" in t:
        return "tutta-la-classe"
    return "come-lavora"


def usa_foglio_nomi(s):
    return "cartellone" in s["modo"].lower()


def nome_download(s):
    base = unicodedata.normalize("NFKD", s["titolo"]).encode("ascii", "ignore").decode()
    base = re.sub(r"[^A-Za-z0-9]+", "-", base).strip("-")
    return f"Il-mio-posto-{s['codice']}-{base}.pdf"


def peso(byte):
    if byte < 1024 * 1024:
        return f"{max(1, round(byte / 1024))} KB"
    return f"{byte / 1024 / 1024:.1f} MB".replace(".", ",")


def pagina(*, radice, titolo, descrizione, corpo, classe_body, attiva="", sezione="", extra_testa="", extra_piede=""):
    """radice: percorso relativo verso la cartella principale del sito ('', '../', '../../').
    attiva: la voce del menu che corrisponde a questa pagina; sezione: la voce a cui la pagina appartiene."""
    home = radice or "./"

    def voce(nome, cartella, chiave):
        href = radice + cartella if cartella else home
        if attiva == chiave:
            stato = ' aria-current="page"'
        elif sezione == chiave:
            stato = ' class="sezione"'
        else:
            stato = ""
        return f'<li><a href="{href}"{stato}>{e(nome)}</a></li>'

    menu = "\n        ".join(voce(*v) for v in MENU)
    return f"""<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{e(titolo)}</title>
<meta name="description" content="{e(descrizione)}">
<meta name="theme-color" content="#fff4c7">
{extra_testa}<link rel="icon" href="{radice}favicon.svg" type="image/svg+xml">
<link rel="stylesheet" href="{radice}stile.css">
</head>
<body class="tema-{classe_body}">
<a class="salta" href="#contenuto">Vai al contenuto</a>
<header class="testata">
  <div class="testata-interno">
    <a class="marchio" href="{home}">
      <svg class="marchio-icona" viewBox="0 0 40 40" aria-hidden="true" focusable="false"><rect x="6" y="9" width="28" height="7" rx="2.5" fill="#f2b632"/><rect x="9" y="16" width="4" height="17" rx="2" fill="#d9573b"/><rect x="27" y="16" width="4" height="17" rx="2" fill="#d9573b"/><circle cx="20" cy="5" r="3.5" fill="#4e9bd8"/></svg>
      <span>{e(NOME_SITO)}</span>
    </a>
    <nav class="menu" aria-label="Menu principale">
      <ul>
        {menu}
      </ul>
    </nav>
  </div>
</header>
<main id="contenuto" tabindex="-1">
{corpo}
</main>
<footer class="piede">
  <ul>
    <li><span class="segnaposto-breve">{e(PIEDE_PDF)}</span></li>
    <li>{e(PIEDE_STAMPA)}</li>
    <li>{e(PIEDE_DATI)}</li>
  </ul>
</footer>
{extra_piede}</body>
</html>
"""


def corpo_home():
    blocchi = pagina_a_riquadri(leggi_testo(TESTI / "home.txt"))
    return f"""<section class="apertura">
  <h1>{e(NOME_SITO)}</h1>
  <p class="sottotitolo">{e(SOTTOTITOLO)}</p>
</section>

{blocchi}
<nav class="ingressi" aria-label="Da dove cominciare">
  <a class="ingresso ingresso-soggetti" href="classe/">
    <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><circle cx="10.5" cy="10.5" r="6.5" fill="none" stroke="currentColor" stroke-width="2.5"/><path d="m15.5 15.5 5 5" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"/></svg>
    <span>{e(INGRESSO_SOGGETTO)}</span>
  </a>
  <a class="ingresso ingresso-inizio" href="per-cominciare/">
    <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M4 12h15m0 0-6-6m6 6-6 6" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/></svg>
    <span>{e(INGRESSO_INIZIO)}</span>
  </a>
</nav>
"""


def corpo_elenco(soggetti, descrizioni):
    voci = []
    for s in soggetti:
        descr = descrizioni.get(s["codice"]) or next(
            (t for k, _, t in s["passi"] if k == "passo"), "")
        voci.append(
            f'  <li class="scheda tema-{tema(s["tipo"])}"><a href="{s["codice"]}/">'
            f'<span class="scheda-numero" aria-hidden="true">{s["codice"]}</span>'
            f'<span class="scheda-testo"><span class="scheda-titolo">'
            f'<span class="visually-hidden">Soggetto {s["codice"]}: </span>{e(s["titolo"])}</span>'
            f'<span class="scheda-descrizione">{e(descr)}</span></span></a></li>')
    return f"""<section class="apertura">
  <h1>I {len(soggetti)} soggetti</h1>
  <p class="intro">{e(INTRO_ELENCO)}</p>
</section>
<ol class="elenco">
{chr(10).join(voci)}
</ol>
"""


def passi_html(passi):
    out, voci, inizio = [], [], None

    def chiudi():
        nonlocal voci, inizio
        if voci:
            attr = f' start="{inizio}" style="counter-reset: passo {inizio - 1}"' if inizio and inizio != 1 else ""
            out.append(f'<ol class="passi"{attr}>\n' + "\n".join(voci) + "\n</ol>")
        voci, inizio = [], None

    for tipo, n, testo in passi:
        if tipo == "passo":
            if inizio is None:
                inizio = n
            voci.append(f"<li>{e(testo)}</li>")
        else:
            chiudi()
            out.append(f"<h3>{e(testo)}</h3>")
    chiudi()
    return "\n".join(out)


def corpo_soggetto(s):
    c = s["codice"]
    stimolo = "\n".join(f"<p>{e(p)}</p>" for p in s["stimolo"])
    arriva = "\n".join(f"<p>{e(p)}</p>" for p in s["arriva"])
    pagina4 = (f'<p class="pagina4">Nel suo quaderno, la pagina 4 di questo soggetto si intitola '
               f'<strong>{e(s["titolo_quaderno"])}</strong>.</p>' if s["titolo_quaderno"] else "")

    pdf = DISEGNI / f"{c}.pdf"
    if pdf.exists():
        foglio = f"""<p><a class="pulsante pulsante-grande" href="../../disegni/{c}.pdf" download="{e(nome_download(s))}">
    <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M12 3v12m0 0-5-5m5 5 5-5M4 20h16" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/></svg>
    <span>Scarica il foglio «{e(s['titolo'])}»<span class="pulsante-dettaglio">PDF, A4, {peso(pdf.stat().st_size)}</span></span>
  </a></p>"""
    else:
        foglio = f"""<div class="segnaposto">
  <p><strong>{e(FOGLIO_IN_PREPARAZIONE)}</strong></p>
</div>"""
    if usa_foglio_nomi(s):
        foglio += f"""
<h3>Il foglio dei nomi</h3>
<p>Un foglio A4 con il titolo del soggetto e le righe per i nomi.</p>
<p><a class="pulsante" href="../../foglio-dei-nomi/?n={c}">Apri il foglio dei nomi per «{e(s['titolo'])}»</a></p>"""

    idee = contributi_html(leggi_idee(IDEE / f"{c}.txt", AVVISI), IDEE_VUOTO_SOGGETTO)
    link_idee = '<a href="../../idee-e-consigli/">Idee e consigli</a>'
    testo_idee = in_linea(IDEE_TESTO).replace("{link}", link_idee)
    modo = (f'<p class="modo"><strong>{e(s["modo"])}</strong></p>\n'
            f'<p><a href="../../come-funziona/#{ancora_modo(s["modo"])}">Come funziona questo modo di lavorare</a></p>'
            if s["modo"] else "")

    return f"""<aside class="prima-volta" aria-label="Se è la prima volta">
  <p>È la prima volta? <a href="../../per-cominciare/">Leggi Per cominciare</a></p>
</aside>

<section class="apertura apertura-soggetto">
  <p class="numero"><span class="visually-hidden">Soggetto numero </span><span aria-hidden="true">Soggetto</span> <span class="numero-cifre">{c}</span></p>
  <h1>{e(s['titolo'])}</h1>
</section>

<section class="riquadro" aria-labelledby="bambino">
  <h2 id="bambino">Che cosa ha fatto il bambino</h2>
  <p class="etichetta">Nella sua pagina ha trovato questa domanda:</p>
  <blockquote class="stimolo">
{stimolo}
  </blockquote>
</section>

<section class="riquadro" aria-labelledby="modo">
  <h2 id="modo">Come lavora la classe</h2>
  {modo}
</section>

<section class="riquadro" aria-labelledby="passi">
  <h2 id="passi">Che cosa fa la classe, passo per passo</h2>
{passi_html(s["passi"])}
</section>

<section class="riquadro" aria-labelledby="arriva">
  <h2 id="arriva">Che cosa arriva al bambino</h2>
{arriva}
{pagina4}
</section>

<section class="riquadro" aria-labelledby="fogli">
  <h2 id="fogli">Fogli da stampare</h2>
{foglio}
</section>

<section class="riquadro riquadro-idee" aria-labelledby="idee">
  <h2 id="idee">{e(IDEE_TITOLO)}</h2>
  <p>{testo_idee}</p>
{idee}
</section>

<nav class="fondo-pagina" aria-label="Altre pagine">
  <a href="../../come-funziona/#attenzioni">Alcune attenzioni</a>
  <a href="../">Torna all'elenco dei soggetti</a>
</nav>
"""


def corpo_testo(titolo, file, segnaposti=None):
    return (f'<section class="apertura">\n  <h1>{e(titolo)}</h1>\n</section>\n'
            + pagina_a_riquadri(leggi_testo(TESTI / file), segnaposti))


def corpo_foglio_nomi(soggetti):
    righe = "\n".join(
        '<li><span class="nome">Nome</span><span class="linea"></span>'
        '<span class="due-punti">:</span><span class="linea linea-lunga"></span></li>'
        for _ in range(RIGHE_FOGLIO_NOMI))
    return f"""<div class="no-stampa">
  <p class="intro">Il foglio dei nomi serve nei soggetti a cartellone. Si stampa su un foglio A4.</p>
  <p><button type="button" class="pulsante" id="stampa">Stampa il foglio</button></p>
  <p id="ritorno" hidden><a href="../classe/">Torna al soggetto</a></p>
</div>

<article class="foglio" aria-labelledby="titolo-foglio">
  <p class="foglio-sito">{e(NOME_SITO)}</p>
  <h1 class="foglio-titolo" id="titolo-foglio">Soggetto: ____________________</h1>
  <ol class="foglio-righe" aria-label="Righe per i nomi">
{righe}
  </ol>
</article>
"""


def script_foglio_nomi(soggetti):
    # Unico script del sito: mette il titolo del soggetto in base all'indirizzo (?n=02).
    # Senza script il foglio funziona lo stesso, con la riga «Soggetto: ____».
    titoli = json.dumps({s["codice"]: f"{s['codice']}. {s['titolo']}" for s in soggetti},
                        ensure_ascii=False)
    return f"""<script>
(function () {{
  var titoli = {titoli};
  var n = new URLSearchParams(location.search).get('n') || '';
  if (/^\\d$/.test(n)) n = '0' + n;
  if (titoli[n]) {{
    document.getElementById('titolo-foglio').textContent = titoli[n];
    document.title = 'Foglio dei nomi, ' + titoli[n];
    var r = document.getElementById('ritorno');
    r.querySelector('a').href = '../classe/' + n + '/';
    r.hidden = false;
  }}
  document.getElementById('stampa').addEventListener('click', function () {{ window.print(); }});
}})();
</script>
"""


def pagina_rinvio(destinazione, codice):
    """Pagina /classe/1 che porta a /classe/01, per chi scrive il numero senza zero."""
    return f"""<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex">
<meta http-equiv="refresh" content="0; url={destinazione}">
<title>Soggetto {codice}</title>
</head>
<body>
<p><a href="{destinazione}">Vai alla pagina del soggetto {codice}</a></p>
</body>
</html>
"""


def pagina_404():
    # La pagina 404 può comparire a qualsiasi indirizzo, quindi i link relativi non bastano.
    # Il piccolo script trova la cartella principale sia sull'indirizzo provvisorio
    # (nomeutente.github.io/ilpostoinclasse/) sia sul dominio definitivo.
    testa = """<base href="/">
<script>
(function () {
  if (location.hostname.slice(-10) === '.github.io') {
    var parte = location.pathname.split('/')[1];
    if (parte) document.querySelector('base').href = '/' + parte + '/';
  }
})();
</script>
"""
    corpo = """<section class="apertura">
  <h1>Questa pagina non c'è</h1>
  <p class="intro">Forse l'indirizzo è scritto in modo diverso. Le pagine dei soggetti hanno sempre due cifre: per esempio <span class="indirizzo">ilpostoinclasse.it/classe/01</span>.</p>
</section>
<p class="azione"><a class="pulsante" href="classe/">Vai ai soggetti</a></p>
"""
    testo = pagina(radice="", titolo=f"Pagina non trovata · {NOME_SITO}",
                   descrizione="La pagina cercata non esiste.", corpo=corpo,
                   classe_body="casa", extra_testa=testa)
    # In questa pagina i link partono dalla cartella principale indicata da <base>.
    return testo.replace('href="./"', 'href=""')


def scrivi(percorso, testo):
    percorso.parent.mkdir(parents=True, exist_ok=True)
    percorso.write_text(testo, encoding="utf-8", newline="\n")


def svuota_sito():
    """Ricrea sito/ da zero. Per sicurezza tocca la cartella solo se l'ha creata questo script.
    Cancella i file uno per uno e lascia le cartelle che Windows tiene bloccate per un attimo."""
    marcatore = SITO / ".generato"
    if SITO.exists() and any(SITO.iterdir()) and not marcatore.exists():
        raise ErroreExcel(f"La cartella {SITO} non è stata creata da questo script: non la tocco.")
    SITO.mkdir(exist_ok=True)
    marcatore.write_text("Cartella creata da genera.py: non modificarla a mano.\n", encoding="utf-8")
    for f in sorted(SITO.rglob("*"), key=lambda p: len(p.parts), reverse=True):
        if f == marcatore:
            continue
        for tentativo in range(5):
            try:
                f.unlink() if f.is_file() else f.rmdir()
                break
            except OSError:
                if f.is_dir():
                    break  # cartella bloccata o non vuota: si riusa
                time.sleep(0.3)
        else:
            raise PermissionError(f"non riesco a cancellare {f}")


AVVISI = []


def costruisci(soggetti):
    svuota_sito()
    (SITO / ".nojekyll").write_text("", encoding="utf-8")

    shutil.copy(MODELLO / "stile.css", SITO / "stile.css")
    shutil.copy(MODELLO / "favicon.svg", SITO / "favicon.svg")
    shutil.copytree(MODELLO / "caratteri", SITO / "caratteri", dirs_exist_ok=True)

    descrizioni = leggi_descrizioni(AVVISI)
    for s in soggetti:
        if s["codice"] not in descrizioni:
            AVVISI.append(f"in testi/descrizioni.txt manca la riga del soggetto {s['codice']}: "
                          "nell'elenco uso il primo passo di «Che cosa fa la classe»")

    scrivi(SITO / "index.html", pagina(
        radice="", titolo=f"{NOME_SITO} · {SOTTOTITOLO}",
        descrizione="Un quaderno da colorare a staffetta tra un bambino e la sua classe. Qui l'insegnante trova che cosa fa la classe e i fogli da stampare.",
        corpo=corpo_home(), classe_body="casa", attiva="home"))

    scrivi(SITO / "classe" / "index.html", pagina(
        radice="../", titolo=f"I soggetti · {NOME_SITO}",
        descrizione="L'elenco dei soggetti del quaderno, ognuno con la pagina per la classe.",
        corpo=corpo_elenco(soggetti, descrizioni), classe_body="elenco", attiva="elenco"))

    for s in soggetti:
        cartella = SITO / "classe" / s["codice"]
        scrivi(cartella / "index.html", pagina(
            radice="../../", titolo=f"{s['codice']}. {s['titolo']} · {NOME_SITO}",
            descrizione=f"Soggetto {s['codice']}, {s['titolo']}: che cosa fa la classe, passo per passo, e i fogli da stampare.",
            corpo=corpo_soggetto(s), classe_body=tema(s["tipo"]), sezione="elenco"))
        senza_zero = str(int(s["codice"]))
        if senza_zero != s["codice"]:
            scrivi(SITO / "classe" / senza_zero / "index.html",
                   pagina_rinvio(f"../{s['codice']}/", s["codice"]))

    consigli = contributi_html(leggi_idee(IDEE / "generale.txt", AVVISI), IDEE_VUOTO_GENERALE)
    fisse = [
        ("per-cominciare", "Per cominciare", "per-cominciare.txt", None,
         "Da dove cominciare: chi accompagna il bambino e il primo passo della staffetta."),
        ("come-funziona", "Come funziona", "come-funziona.txt", None,
         "Come funziona la staffetta, come lavora la classe, il quaderno gemello e alcune attenzioni."),
        ("idee-e-consigli", "I consigli di chi lo usa", "idee-e-consigli.txt", {"{{consigli}}": consigli},
         "I consigli di insegnanti, docenti e famiglie che usano il quaderno."),
        ("come-averlo", "Come averlo", "come-averlo.txt", None,
         "Come avere il quaderno: dal sito, gratis, oppure stampato e rilegato."),
    ]
    for cartella, titolo, file, segnaposti, descr in fisse:
        if not (TESTI / file).exists():
            AVVISI.append(f"manca il file testi/{file}: la pagina «{titolo}» è vuota")
        scrivi(SITO / cartella / "index.html", pagina(
            radice="../", titolo=f"{titolo} · {NOME_SITO}", descrizione=descr,
            corpo=corpo_testo(titolo, file, segnaposti), classe_body="casa", attiva=cartella))

    scrivi(SITO / "foglio-dei-nomi" / "index.html", pagina(
        radice="../", titolo=f"Foglio dei nomi · {NOME_SITO}",
        descrizione="Il foglio dei nomi da stampare in A4 per i soggetti a cartellone.",
        corpo=corpo_foglio_nomi(soggetti), classe_body="casa foglio-nomi", sezione="elenco",
        extra_piede=script_foglio_nomi(soggetti)))

    scrivi(SITO / "404.html", pagina_404())

    # Disegni: copiati per l'anteprima sul computer (su GitHub li copia la pubblicazione).
    (SITO / "disegni").mkdir(exist_ok=True)
    for pdf in DISEGNI.glob("*.pdf"):
        shutil.copy(pdf, SITO / "disegni" / pdf.name)


def main():
    try:
        excel = trova_excel()
        print(f"Leggo: {excel}")
        soggetti, avvisi = leggi_soggetti(excel)
        AVVISI.extend(avvisi)
        creati = prepara_file_idee(soggetti)
        costruisci(soggetti)
    except ErroreExcel as err:
        print("\nATTENZIONE. Il sito NON è stato aggiornato.\n" + str(err))
        return 1
    except PermissionError as err:
        print(f"\nATTENZIONE. Non riesco ad aprire un file: {err}\nChiudi i file del sito aperti in altri programmi e riprova tra qualche secondo.")
        return 1

    codici = {s["codice"] for s in soggetti}
    presenti = sorted(p.stem for p in DISEGNI.glob("*.pdf"))
    estranei = [p for p in presenti if p not in codici]
    mancanti = [c for c in sorted(codici) if c not in presenti]

    print(f"\nFatto. Pagine create: {len(soggetti)} soggetti, più home, elenco, quattro pagine di testo e il foglio dei nomi.")
    print(f"Fogli da stampare presenti: {len(presenti) - len(estranei)} su {len(soggetti)}.")
    if mancanti:
        print("Fogli che mancano ancora (al loro posto c'è «Foglio in preparazione»): " + ", ".join(mancanti))
    if creati:
        print("File delle idee creati vuoti in idee/: " + ", ".join(creati))
    for p in estranei:
        print(f"Avviso: il file disegni/{p}.pdf non corrisponde a nessun soggetto. Il nome deve essere 01.pdf, 02.pdf...")
    for a in AVVISI:
        print("Avviso: " + a)
    print("\nPer vedere il sito sul computer: doppio clic su anteprima.bat")
    return 0


if __name__ == "__main__":
    sys.exit(main())
