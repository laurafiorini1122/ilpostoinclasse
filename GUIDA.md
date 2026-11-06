# Il mio posto c'è sempre: guida al sito

Questa guida spiega come aggiornare il sito, come pubblicarlo la prima volta
su GitHub Pages e come collegare in futuro il dominio ilpostoinclasse.it.

## Com'è fatta la cartella

| Che cosa | A che cosa serve |
|---|---|
| `aggiorna_sito.bat` | Doppio clic: legge l'Excel e ricrea il sito. |
| `anteprima.bat` | Doppio clic: apre il sito nel browser, solo sul tuo computer. |
| `disegni/` | Qui metti i PDF dei disegni: `01.pdf`, `02.pdf` … `15.pdf`. |
| `sito/` | Il sito pronto. Lo crea lo script: non modificarlo a mano. |
| `genera.py` | Lo script. In alto ci sono i testi fissi (consiglio di stampa, promemoria). |
| `modello/` | Colori, grafica e caratteri. |

L'Excel resta su Google Drive e non entra mai nel sito: lo script ne legge solo
le colonne N., Soggetto, Che cosa fa la classe, Titolo della pagina da incollare,
Indirizzo sul sito e Tipo. Le note e i prompt restano privati.
Se salvi una versione nuova (per esempio `_v6`), lo script la trova da solo.

Il colore di fondo di ogni pagina dipende dalla colonna Tipo:
giallo per Legame, verde per Vita di classe, azzurro per Libero.

## Aggiornare il sito

1. Modifica l'Excel o aggiungi un PDF in `disegni/`.
2. Doppio clic su `aggiorna_sito.bat`. Leggi il messaggio nella finestra nera:
   se c'è scritto **ATTENZIONE**, il sito non è cambiato e la finestra ti dice
   quale riga dell'Excel correggere.
3. Se vuoi, doppio clic su `anteprima.bat` per vedere il risultato.
4. Pubblica: chiedi a Claude «pubblica le modifiche del sito» oppure usa
   GitHub Desktop (Commit, poi Push). Dopo uno o due minuti il sito online è aggiornato.

## Pubblicare la prima volta su GitHub Pages

GitHub è il servizio dove sta già La scatola di latta. GitHub Pages pubblica gratis
i siti fatti solo di pagine, come questo. Usi lo stesso account `laurafiorini1122`.

**Attenzione:** un sito gratuito su GitHub Pages sta in un archivio pubblico.
Chiunque può vedere i file della cartella. Per questo l'Excel non c'è.

### Passo 1. Creare l'archivio (lo fai tu, sul sito di GitHub)

1. Vai su github.com ed entra con il tuo account.
2. In alto a destra clicca **+** e poi **New repository**.
3. In **Repository name** scrivi `ilpostoinclasse`.
4. Scegli **Public**. Non spuntare nessuna delle caselle sotto (README, .gitignore, license).
5. Clicca **Create repository**. Si apre una pagina con delle istruzioni: non servono.

### Passo 2. Mandare i file su GitHub (lo fa Claude)

Scrivi a Claude: «ho creato l'archivio, carica il sito». Claude prepara i file,
li salva con una descrizione delle modifiche (il *commit*) e li invia a GitHub (il *push*).
La prima volta Windows potrebbe aprire una finestra per accedere a GitHub: accedi tu.

### Passo 3. Accendere GitHub Pages (lo fai tu, sul sito di GitHub)

1. Apri l'archivio `ilpostoinclasse` su github.com.
2. Clicca **Settings** (l'ingranaggio in alto), poi **Pages** nella colonna a sinistra.
3. Sotto **Build and deployment**, in **Source** scegli **GitHub Actions**.
4. Clicca la scheda **Actions** in alto: vedrai «Pubblica il sito». Se non è partito,
   cliccalo e poi **Run workflow**. Quando compare il segno verde il sito è online.

### Passo 4. Controllare

Il sito è all'indirizzo provvisorio
`https://laurafiorini1122.github.io/ilpostoinclasse/`
e la pagina del soggetto 1 è
`https://laurafiorini1122.github.io/ilpostoinclasse/classe/01`.

Aprilo anche dal telefono.

## Collegare in futuro il dominio ilpostoinclasse.it

Gli indirizzi delle pagine non cambiano: `/classe/01` resta `/classe/01`,
cambia solo la parte prima. Tutti i link del sito sono relativi,
quindi non c'è niente da modificare nel sito. Anche i link provvisori
già mandati continueranno a funzionare: GitHub porta da solo al dominio nuovo.

Sono gli stessi passi fatti per lascatoladilatta.it.

1. **Compra il dominio** ilpostoinclasse.it (per esempio su Register.it, come l'altro).
2. **Verifica il dominio su GitHub** (protegge il dominio da usi altrui):
   sulla tua foto in alto a destra, **Settings**, poi **Pages** nella colonna a sinistra,
   **Add a domain**. GitHub ti dà un record TXT da copiare.
3. **Nel pannello DNS del dominio** (su Register.it: Dominio, Gestione DNS) aggiungi:
   - il record TXT del punto 2;
   - quattro record **A** per il dominio principale (`@`), uno per ognuno di questi valori:
     `185.199.108.153`, `185.199.109.153`, `185.199.110.153`, `185.199.111.153`;
   - un record **CNAME** per `www` con valore `laurafiorini1122.github.io`.

   Non toccare gli altri record già presenti (per esempio i record MX della posta).
4. **Nell'archivio su GitHub**: Settings, Pages, **Custom domain**: scrivi
   `ilpostoinclasse.it` e clicca Save. Aspetta il segno verde (da pochi minuti a qualche ora).
5. Spunta **Enforce HTTPS**, così il sito si apre con il lucchetto.

Quando hai fatto, avvisa Claude: aggiornerà il promemoria del progetto con il dominio.

## Riservatezza

Il sito non usa cookie, statistiche, moduli né risorse da altri siti:
i caratteri sono dentro la cartella. GitHub, come ogni servizio che ospita siti,
registra per motivi di sicurezza gli indirizzi IP di chi visita.
Non sono statistiche e non usano cookie.
