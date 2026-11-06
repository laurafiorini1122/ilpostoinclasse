"""
Genera il sito "Il mio posto c'è sempre" partendo dall'Excel dei soggetti.

Come si usa: doppio clic su aggiorna_sito.bat (oppure: python genera.py).
Si può anche indicare un Excel preciso: python genera.py "percorso\\file.xlsx"

Lo script:
- legge il foglio "Soggetti" (solo le colonne N., Soggetto, Che cosa fa la classe,
  Titolo della pagina da incollare, Indirizzo sul sito: le note e i prompt non escono mai);
- ricrea da zero la cartella sito/ (tranne i disegni, che copia da disegni/);
- non modifica mai l'Excel.

Usa solo Python standard: non serve installare niente.
"""

import html
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
# più recente (Quaderno_classe_soggetti_v5, v6, v7...).
CARTELLA_EXCEL = Path(r"G:\Il mio Drive\0 Spazio flessibile\8 SiO - Scuola")
NOME_EXCEL = "Quaderno_classe_soggetti"
FOGLIO = "Soggetti"

QUI = Path(__file__).resolve().parent
SITO = QUI / "sito"
MODELLO = QUI / "modello"
DISEGNI = QUI / "disegni"

# Intestazioni delle colonne da leggere (basta l'inizio del testo).
# Così, se in futuro inserisci una colonna nell'Excel, lo script la ritrova.
COLONNE = {
    "numero": "N.",
    "titolo": "Soggetto",
    "classe": "Che cosa fa la classe",
    "titolo_quaderno": "Titolo della pagina da incollare",
    "indirizzo": "Indirizzo sul sito",
    "tipo": "Tipo",
}

# ---------------------------------------------------------------------------
# TESTI DEL SITO (si possono modificare qui)
# ---------------------------------------------------------------------------

NOME_SITO = "Il mio posto c'è sempre"
SOTTOTITOLO = "Un quaderno da colorare a staffetta con la mia classe"
AUTRICE = "Laura Fiorini"
ANNO = "2026"

CONSIGLIO_STAMPA = (
    "Stampa il disegno su un foglio A4 a dimensione reale (100%) e non con "
    "«Adatta alla pagina». Basta un foglio solo per tutta la classe. "
    "Se usate i pennarelli, una carta un po' più spessa regge meglio."
)
PROMEMORIA = (
    "Prima di spedire il foglio scrivete in un angolo la data e i nomi di tutti: "
    "per il bambino è bello sapere chi ha partecipato e quando."
)

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


def pulisci(testo):
    testo = testo.replace("\r", "")
    paragrafi = [re.sub(r"\s+", " ", p).strip() for p in re.split(r"\n\s*\n", testo)]
    return [p for p in paragrafi if p]


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
        for campo in ("titolo", "classe", "titolo_quaderno"):
            if re.search(r"[\u2014\u2013]", dato[campo]):
                avvisi.append(f"riga {n_riga}: nella colonna «{COLONNE[campo]}» c'è un trattino lungo")
        soggetti.append({
            "codice": codice,
            "titolo": dato["titolo"],
            "titolo_quaderno": dato["titolo_quaderno"],
            "classe": pulisci(dato["classe"]),
            "tipo": dato["tipo"],
        })
    if errori:
        raise ErroreExcel("Da correggere nell'Excel:\n  - " + "\n  - ".join(errori))
    soggetti.sort(key=lambda s: s["codice"])
    return soggetti, avvisi


# ---------------------------------------------------------------------------
# COSTRUZIONE DELLE PAGINE
# ---------------------------------------------------------------------------

def e(testo):
    return html.escape(testo, quote=True)


def tema(tipo):
    t = tipo.lower()
    if "legame" in t:
        return "legame"
    if "vita" in t:
        return "vita"
    return "libero"


def nome_download(s):
    base = unicodedata.normalize("NFKD", s["titolo"]).encode("ascii", "ignore").decode()
    base = re.sub(r"[^A-Za-z0-9]+", "-", base).strip("-")
    return f"Il-mio-posto-{s['codice']}-{base}.pdf"


def peso(byte):
    if byte < 1024 * 1024:
        return f"{max(1, round(byte / 1024))} KB"
    return f"{byte / 1024 / 1024:.1f} MB".replace(".", ",")


def pagina(*, radice, titolo, descrizione, corpo, classe_body, attiva="", extra_testa=""):
    """radice: percorso relativo verso la cartella principale del sito ('' , '../', '../../')."""
    voce = lambda nome, href, chiave: (
        f'<li><a href="{href}"{" aria-current=\"page\"" if attiva == chiave else ""}>{nome}</a></li>'
    )
    home = radice or "./"
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
    <nav aria-label="Menu principale">
      <ul>
        {voce("Home", home, "home")}
        {voce("I soggetti", radice + "classe/", "elenco")}
      </ul>
    </nav>
  </div>
</header>
<main id="contenuto" tabindex="-1">
{corpo}
</main>
<footer class="piede">
  <p><strong>{e(NOME_SITO)}</strong>. {e(SOTTOTITOLO)}.<br>Un quaderno di {e(AUTRICE)}. © {ANNO} {e(AUTRICE)}. Le pagine per la classe si possono stampare per uso scolastico.</p>
  <p>Questo sito non usa cookie, non raccoglie dati e non contiene statistiche.</p>
</footer>
</body>
</html>
"""


def corpo_home():
    return f"""<section class="apertura">
  <h1>{e(NOME_SITO)}</h1>
  <p class="sottotitolo">{e(SOTTOTITOLO)}</p>
  <p class="intro">Questo quaderno accompagna un bambino o una bambina che per un periodo non può venire a scuola. Dentro ci sono quindici disegni da colorare. Per ogni disegno anche la classe ne completa uno, così il bambino e i compagni continuano a fare le cose insieme.</p>
</section>

<section class="riquadro" aria-labelledby="come">
  <h2 id="come">Come funziona per te, insegnante</h2>
  <ol class="passi">
    <li><span class="passo-titolo">Ricevi un numero.</span> La famiglia ti dice quale soggetto ha colorato il bambino, per esempio il numero 01.</li>
    <li><span class="passo-titolo">Apri la pagina del soggetto.</span> All'indirizzo <span class="indirizzo">ilpostoinclasse.it/classe/01</span> trovi che cosa fa la classe e il disegno da stampare.</li>
    <li><span class="passo-titolo">La classe completa un foglio solo.</span> Ognuno aggiunge un pezzetto e il suo nome. Se qualcuno fa anche un disegno suo, va bene: il bambino lo conserva nella busta in fondo al quaderno.</li>
    <li><span class="passo-titolo">Il foglio arriva al bambino.</span> Lo consegni alla famiglia e il bambino lo incolla nel quaderno, accanto al suo disegno.</li>
  </ol>
  <p class="nota">Non c'è un ordine da seguire e non c'è fretta. Le cose disegnate non devono essere per forza vere: un po' di fantasia crea un mondo che conoscono solo il bambino e la sua classe.</p>
</section>

<p class="azione"><a class="pulsante" href="classe/">Vai ai 15 soggetti</a></p>
"""


def corpo_elenco(soggetti):
    voci = "\n".join(
        f'  <li class="scheda tema-{tema(s["tipo"])}"><a href="{s["codice"]}/">'
        f'<span class="scheda-numero" aria-hidden="true">{s["codice"]}</span>'
        f'<span class="scheda-titolo"><span class="visually-hidden">Soggetto {s["codice"]}: </span>{e(s["titolo"])}</span></a></li>'
        for s in soggetti
    )
    return f"""<section class="apertura">
  <h1>I {len(soggetti)} soggetti</h1>
  <p class="intro">Cerca il numero che ti ha comunicato la famiglia e apri la pagina: lì trovi che cosa fa la classe e il disegno da stampare.</p>
</section>
<ol class="elenco">
{voci}
</ol>
"""


def corpo_soggetto(s, prec, succ):
    pdf = DISEGNI / f"{s['codice']}.pdf"
    if pdf.exists():
        disegno = f"""<p><a class="pulsante pulsante-grande" href="../../disegni/{s['codice']}.pdf" download="{e(nome_download(s))}">
    <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M12 3v12m0 0-5-5m5 5 5-5M4 20h16" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/></svg>
    <span>Scarica il disegno «{e(s['titolo'])}»<span class="pulsante-dettaglio">PDF, A4, {peso(pdf.stat().st_size)}</span></span>
  </a></p>"""
    else:
        disegno = """<div class="segnaposto">
    <p><strong>Il disegno sarà disponibile a breve.</strong></p>
    <p>Stiamo preparando il file da stampare. Puoi tornare su questa pagina allo stesso indirizzo tra qualche giorno.</p>
  </div>"""

    classe = "\n  ".join(f"<p>{e(p)}</p>" for p in s["classe"])
    nel_quaderno = (
        f'<p class="nel-quaderno">Nel quaderno del bambino questa pagina si chiama <em>{e(s["titolo_quaderno"])}</em>.</p>'
        if s["titolo_quaderno"] else ""
    )
    vicini = []
    if prec:
        vicini.append(f'<a class="vicino vicino-prec" href="../{prec["codice"]}/"><span class="vicino-etichetta">Soggetto precedente</span> {prec["codice"]}. {e(prec["titolo"])}</a>')
    if succ:
        vicini.append(f'<a class="vicino vicino-succ" href="../{succ["codice"]}/"><span class="vicino-etichetta">Soggetto successivo</span> {succ["codice"]}. {e(succ["titolo"])}</a>')

    return f"""<nav class="briciole" aria-label="Sei qui">
  <ol>
    <li><a href="../../">Home</a></li>
    <li><a href="../">I soggetti</a></li>
    <li><a href="./" aria-current="page">Soggetto {s['codice']}</a></li>
  </ol>
</nav>

<section class="apertura apertura-soggetto">
  <p class="numero"><span class="visually-hidden">Soggetto numero </span><span aria-hidden="true">Soggetto</span> <span class="numero-cifre">{s['codice']}</span></p>
  <h1>{e(s['titolo'])}</h1>
  {nel_quaderno}
</section>

<section class="riquadro" aria-labelledby="fa">
  <h2 id="fa">Che cosa fa la classe</h2>
  {classe}
</section>

<section class="riquadro" aria-labelledby="disegno">
  <h2 id="disegno">Il disegno da stampare</h2>
  {disegno}
  <h3>Come stamparlo</h3>
  <p>{e(CONSIGLIO_STAMPA)}</p>
</section>

<aside class="promemoria" aria-labelledby="ricorda">
  <h2 id="ricorda">Ricordati</h2>
  <p>{e(PROMEMORIA)}</p>
</aside>

<nav class="vicini" aria-label="Altri soggetti">
  {"".join(vicini)}
</nav>
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


def costruisci(soggetti):
    svuota_sito()
    (SITO / ".nojekyll").write_text("", encoding="utf-8")

    shutil.copy(MODELLO / "stile.css", SITO / "stile.css")
    shutil.copy(MODELLO / "favicon.svg", SITO / "favicon.svg")
    shutil.copytree(MODELLO / "caratteri", SITO / "caratteri", dirs_exist_ok=True)

    scrivi(SITO / "index.html", pagina(
        radice="", titolo=f"{NOME_SITO} · Il quaderno a staffetta",
        descrizione="Un quaderno da colorare a staffetta tra un bambino e la sua classe. Qui l'insegnante trova le pagine da stampare.",
        corpo=corpo_home(), classe_body="casa", attiva="home"))

    scrivi(SITO / "classe" / "index.html", pagina(
        radice="../", titolo=f"I soggetti · {NOME_SITO}",
        descrizione="L'elenco dei soggetti del quaderno, ognuno con la sua pagina da stampare per la classe.",
        corpo=corpo_elenco(soggetti), classe_body="elenco", attiva="elenco"))

    for i, s in enumerate(soggetti):
        prec = soggetti[i - 1] if i > 0 else None
        succ = soggetti[i + 1] if i + 1 < len(soggetti) else None
        cartella = SITO / "classe" / s["codice"]
        scrivi(cartella / "index.html", pagina(
            radice="../../", titolo=f"{s['codice']}. {s['titolo']} · {NOME_SITO}",
            descrizione=f"Soggetto {s['codice']}, {s['titolo']}: che cosa fa la classe e il disegno da stampare in A4.",
            corpo=corpo_soggetto(s, prec, succ), classe_body=tema(s["tipo"])))
        senza_zero = str(int(s["codice"]))
        if senza_zero != s["codice"]:
            scrivi(SITO / "classe" / senza_zero / "index.html",
                   pagina_rinvio(f"../{s['codice']}/", s["codice"]))

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

    print(f"\nFatto. Pagine create: {len(soggetti)} soggetti, più home ed elenco.")
    print(f"Disegni presenti: {len(presenti) - len(estranei)} su {len(soggetti)}.")
    if mancanti:
        print("Disegni che mancano ancora (al loro posto c'è l'avviso «disponibile a breve»): " + ", ".join(mancanti))
    for p in estranei:
        print(f"Avviso: il file disegni/{p}.pdf non corrisponde a nessun soggetto. Il nome deve essere 01.pdf, 02.pdf...")
    for a in avvisi:
        print("Avviso: " + a)
    print("\nPer vedere il sito sul computer: doppio clic su anteprima.bat")
    return 0


if __name__ == "__main__":
    sys.exit(main())
