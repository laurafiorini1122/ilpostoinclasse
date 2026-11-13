# CLAUDE.md

Sito statico del quaderno «Il mio posto c'è sempre. Un quaderno da colorare a staffetta
con la mia classe» di Laura Fiorini. Il bambino colora un soggetto, la sua classe scarica
da qui la pagina collegata e la completa a scuola. I lettori sono insegnanti di primaria,
spesso da telefono.

## Come lavorare con Laura

- Laura non è una programmatrice: si spiega ogni passaggio in italiano semplice.
- Prima di cambiare la struttura si propone e si aspetta il via.
- Commit con messaggio in italiano; **niente push senza che Laura lo chieda**.
- Ogni testo scritto da Claude è una proposta.

## Regole fisse

- Gli indirizzi `/classe/01` … `/classe/15` non cambiano mai (colonna «Indirizzo sul sito»).
- Nessun dominio nel codice: solo link relativi. Unica eccezione la 404, che calcola la radice.
- Repository pubblico: l'Excel (su Google Drive, `G:\Il mio Drive\0 Spazio flessibile\8 SiO - Scuola`)
  non entra mai. Lo script ne legge solo le colonne N., Soggetto, Tipo, Frase stimolo (D),
  Che cosa fa la classe (I), Titolo della pagina da incollare (L), Indirizzo sul sito (M),
  Come lavora la classe (R), Che cosa arriva al bambino (S). Mai G, H, J, K, Stato, Note, P, T, U.
- Nessun cookie, statistica, modulo o risorsa esterna. Caratteri AHN in `modello/caratteri` (OFL).
- WCAG 2.1 AA almeno: corpo 20 px, contrasto ≥ 4,5:1, pulsanti ≥ 48 px, tastiera, zoom 200%.
- Mai nominare la malattia nei testi scritti da Claude. I testi della presentazione
  (che citano la scuola in ospedale) si riportano come sono: li ha voluti Laura (13/11/2026).
- Testi: niente trattini lunghi; niente virgola prima di «e» e «o» se non serve; si dà del tu.
- Nessun legame con gli altri siti di Laura.

## Come funziona

- `genera.py` (solo libreria standard) legge l'Excel più recente e ricrea `sito/`.
  Frasi brevi fisse in alto nello script; testi lunghi delle pagine in `testi/*.txt`
  (presi dalla presentazione, non riscriverli); descrizioni dell'elenco in `testi/descrizioni.txt`;
  grafica in `modello/stile.css`.
- Menu: Il quaderno · I soggetti · Per cominciare · Come funziona · Idee e consigli · Come averlo.
- `idee/NN.txt` e `idee/generale.txt`: contributi scritti da Laura («Ruolo:» + testo, separati da `---`).
  Lo script crea i file mancanti ma non modifica mai quelli esistenti.
- Foglio dei nomi unico in `/foglio-dei-nomi/?n=NN` (titolo messo da un piccolo script inline),
  linkato dai soggetti con «Cartellone a tessere» nella colonna R.
- Segnaposto sempre tratteggiati (`.segnaposto`, `.segnaposto-breve`, `[[testo]]` nei file di testi/).
- `disegni/NN.pdf` li mette Laura. In `sito/disegni` c'è solo una copia per l'anteprima
  (ignorata da git); la Action `.github/workflows/pubblica.yml` copia `disegni/` in `sito/` e pubblica.
- Fondo per tipo: Legame giallo, Vita di classe verde, Libero azzurro; home e pagine di testo pesca, elenco lilla.
- `/classe/1` … `/classe/9` rimandano a `/classe/01` … `/classe/09`.
- Anteprima: `anteprima.bat`, oppure `preview_start` con la configurazione `ilpostoinclasse`.

## Pubblicazione

GitHub Pages con GitHub Actions dall'account `laurafiorini1122`, archivio `ilpostoinclasse`.
Email git del progetto: `284369532+laurafiorini1122@users.noreply.github.com`.
Dominio futuro ilpostoinclasse.it: passi in `GUIDA.md`.
