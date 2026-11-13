# Il mio posto c'è sempre: il sito

Il sito del quaderno «Il mio posto c'è sempre. Un quaderno da colorare a staffetta
con la mia classe». Qui sotto trovi le cose che servono di più, in parole semplici.
Per pubblicare il sito e per il dominio leggi `GUIDA.md`.

## Dove sta ogni cosa

| Che cosa | Dove |
|---|---|
| I soggetti (titolo, domanda, passi della classe...) | Nell'Excel su Google Drive, foglio «Soggetti» |
| La riga di descrizione di ogni soggetto nell'elenco | `testi/descrizioni.txt` |
| I testi delle pagine Il quaderno, Per cominciare, Come funziona, Idee e consigli, Come averlo | `testi/` (un file per pagina) |
| Le idee dalle classi e i consigli | `idee/` (un file per soggetto e `generale.txt`) |
| I fogli da stampare per la classe | `disegni/` (`01.pdf`, `02.pdf` ...) |
| Colori e grafica | `modello/stile.css` |
| Il sito pronto | `sito/`: lo crea lo script, non si modifica a mano |

## Aggiornare il sito dopo aver cambiato l'Excel

1. Salva l'Excel. Se lo salvi con un numero nuovo (per esempio `_v8`), lo script lo trova da solo.
2. Doppio clic su `aggiorna_sito.bat`.
3. Leggi la finestra nera. Se c'è scritto **ATTENZIONE**, il sito non è cambiato
   e la finestra dice quale riga dell'Excel correggere. Le righe **Avviso** invece
   segnalano cose da controllare, ma il sito è stato aggiornato.
4. Doppio clic su `anteprima.bat` per vederlo nel browser.

Lo script rifà tutte le pagine dei soggetti, ma **non tocca mai i file delle idee**.

Dall'Excel il sito prende solo queste colonne: N., Soggetto, Tipo, Frase stimolo,
Che cosa fa la classe, Titolo della pagina da incollare, Indirizzo sul sito,
Come lavora la classe, Che cosa arriva al bambino. Le note, lo stato, i prompt
e le risposte di Claude non escono mai.

Nella colonna «Che cosa fa la classe» ogni passo va su una riga sua e comincia
con il numero e il punto (`1. `, `2. `...). Una riga senza numero, come
«Primo tempo: la caccia», diventa un piccolo titolo.

## Aggiungere un'idea o un consiglio

Le idee arrivano per email. Dopo averle sistemate:

1. Apri la cartella `idee/`.
   - Un'idea su un soggetto va nel file con il suo numero: per il soggetto 3, `03.txt`.
   - Un consiglio generale va in `generale.txt`.
2. Apri il file con il Blocco note (clic destro, Apri con, Blocco note).
3. In fondo, sotto le righe che cominciano con `#`, scrivi:

   ```
   Ruolo: insegnante di classe
   Il testo dell'idea, così come vuoi che compaia sul sito.
   ```

   Se nel file c'è già un'idea, prima della nuova scrivi una riga con tre trattini:

   ```
   ---
   Ruolo: docente della scuola in ospedale
   Il testo della nuova idea.
   ```

4. Salva e fai doppio clic su `aggiorna_sito.bat`.

Le righe che cominciano con `#` sono istruzioni e non compaiono sul sito.
Quando un file non ha ancora contributi, sul sito compare
«Non ci sono ancora idee per questo soggetto.» (oppure, nella pagina Idee e consigli,
«Le prime idee arriveranno presto.»).

Prima di pubblicare un contributo controlla che non ci siano nomi di bambini,
della scuola, dell'ospedale o del reparto, né informazioni sulla salute.

## Cambiare i testi delle pagine

I file in `testi/` si aprono con il Blocco note. Le regole sono scritte in cima a ogni file:

- una riga vuota separa i paragrafi;
- `## ` all'inizio della riga fa un titolo, `### ` un titolo più piccolo;
- `- ` all'inizio fa una voce di elenco, `1. ` un passo numerato;
- `**parole**` le mette in grassetto;
- `[[testo]]` fa un segnaposto in un riquadro tratteggiato: serve per le parti ancora da scrivere.

Dopo una modifica: doppio clic su `aggiorna_sito.bat`.

## I segnaposto ancora da completare

- La traccia del messaggio alla classe: `testi/per-cominciare.txt`.
- L'indirizzo email per le idee: `testi/idee-e-consigli.txt`, dove c'è `[[indirizzo da definire]]`.
- La presentazione in PDF nel piè di pagina: in alto in `genera.py` (`PIEDE_PDF`).
- I fogli da stampare: quando un PDF è pronto, mettilo in `disegni/` con il numero del soggetto.
