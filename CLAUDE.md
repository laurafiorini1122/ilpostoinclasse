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
  non entra mai. Lo script ne legge solo le colonne N., Soggetto, Tipo, Che cosa fa la classe,
  Titolo della pagina da incollare, Indirizzo sul sito.
- Nessun cookie, statistica, modulo o risorsa esterna. Caratteri AHN in `modello/caratteri` (OFL).
- WCAG 2.1 AA almeno: corpo 20 px, contrasto ≥ 4,5:1, pulsanti ≥ 48 px, tastiera, zoom 200%.
- Nessun riferimento all'ospedale o alla malattia.
- Testi: niente trattini lunghi; niente virgola prima di «e» e «o» se non serve; si dà del tu.
- Nessun legame con gli altri siti di Laura.

## Come funziona

- `genera.py` (solo libreria standard) legge l'Excel più recente e ricrea `sito/`.
  I testi fissi stanno in alto nello script, la grafica in `modello/stile.css`.
- `disegni/NN.pdf` li mette Laura. In `sito/disegni` c'è solo una copia per l'anteprima
  (ignorata da git); la Action `.github/workflows/pubblica.yml` copia `disegni/` in `sito/` e pubblica.
- Fondo per tipo: Legame giallo, Vita di classe verde, Libero azzurro; home pesca, elenco lilla.
- `/classe/1` … `/classe/9` rimandano a `/classe/01` … `/classe/09`.
- Anteprima: `anteprima.bat`, oppure `preview_start` con la configurazione `ilpostoinclasse`.

## Pubblicazione

GitHub Pages con GitHub Actions dall'account `laurafiorini1122`, archivio `ilpostoinclasse`.
Email git del progetto: `284369532+laurafiorini1122@users.noreply.github.com`.
Dominio futuro ilpostoinclasse.it: passi in `GUIDA.md`.
