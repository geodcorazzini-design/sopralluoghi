# App Sopralluoghi – installazione

App per Android che funziona anche senza rete. Compila la check-list di sopralluogo ASPI, verifica il personale sul file ITP e produce il verbale in Word (modulo ASPI con logo) e in Excel. Quando torna la rete carica tutto sul tuo Google Drive.

I dati restano sul telefono e sul tuo Drive. Il codice pubblicato online non contiene né dati né il modello ASPI: il modello e i file Excel li carichi dal telefono.

---

## 1. Pubblica l'app (una volta, circa 10 minuti)

Serve un indirizzo https per installarla sul telefono. GitHub Pages è gratuito.

1. Crea un account su **github.com** (gratuito).
2. In alto a destra **+ → New repository**. Nome: `sopralluoghi`, visibilità **Public**, poi **Create repository**.
3. Nella pagina del repository: **uploading an existing file**. Trascina **tutto il contenuto** di questa cartella: `index.html`, `app.js`, `app.css`, `sw.js`, `manifest.webmanifest` e le cartelle `lib`, `vendor`, `icons`. Poi **Commit changes**.
4. **Settings → Pages**. In *Build and deployment* scegli **Deploy from a branch**, ramo **main**, cartella **/(root)**, poi **Save**.
5. Dopo 1–2 minuti l'app è su `https://TUO-UTENTE.github.io/sopralluoghi/`.

## 2. Collega Google Drive (una volta, circa 15 minuti)

1. Apri **console.cloud.google.com** con il tuo account Gmail e crea un progetto chiamato `Sopralluoghi`.
2. **API e servizi → Libreria**: cerca **Google Drive API** e premi **Abilita**.
3. **Google Auth Platform** (o *Schermata consenso OAuth*):
   - tipo di utenti **Esterno**, nome app `Sopralluoghi` e la tua email;
   - in **Pubblico / Utenti di test** aggiungi il tuo indirizzo Gmail;
   - in **Accesso ai dati / Ambiti** aggiungi `.../auth/drive.file`. Con questo permesso l'app vede solo i file che crea lei, non il resto del tuo Drive.
4. **Client → Crea client**: tipo **Applicazione web**. In *Origini JavaScript autorizzate* inserisci `https://TUO-UTENTE.github.io` (senza la barra finale). Premi **Crea** e copia l'**ID client** (finisce con `.apps.googleusercontent.com`).
5. Al primo accesso Google può avvisare che l'app non è verificata: è la tua app, premi *Continua*.

## 3. Installa sul telefono

1. Sul telefono apri Chrome all'indirizzo del punto 1.5, poi **menu ⋮ → Installa app** (o *Aggiungi a schermata Home*).
2. Apri l'app e vai su **Impostazioni**:
   - **Modello Word ASPI**: carica `Modello_sopralluogo_ASPI_app.docx`. Dal selettore file di Android puoi prenderlo direttamente da Drive;
   - **Verifiche ITP**: carica il tuo xlsx operatori e contratti, e ricaricalo ogni volta che lo aggiorni;
   - **Programma cantieri** (facoltativo): carica il programma settimanale ASPI per precompilare dati e imprese;
   - incolla l'**ID client** e premi **Collega e verifica**. Su Drive viene creata la cartella `VERBALE SOPRALLUOGHI ASPI`.

## Uso in cantiere

- **+ Nuovo sopralluogo**: scegli il cantiere dal programma, oppure parti da un verbale vuoto.
- Inserisci imprese, preposto e operatori: i nomi si completano dal file ITP. Sotto ogni nome compare l'esito della verifica (autorizzazione a manovra, contratto, formazione, UNILAV, idoneità sanitaria).
- Scrivi nel campo **Contratti del cantiere** i C.A./AQ: l'app segnala le autorizzazioni legate a un contratto diverso.
- **Rilievi dal file ITP**: con *Applica* il rilievo diventa NC su A.2 o una riga delle osservazioni.
- Compila C / NC / N/A, aggiungi le foto (la voce associata è facoltativa) e premi **Concludi**.
- **Sincronizza**, quando c'è rete: crea `VERBALE SOPRALLUOGHI ASPI/AAAA-MM/<nome verbale>/` con il Word, l'Excel e le foto, e aggiorna `Registro_sopralluoghi.xlsx`.
- Per far vedere i verbali ad altri, condividi la cartella da Google Drive.

## Da sapere

- Finché non sincronizzi, i verbali esistono solo sul telefono. Se disinstalli l'app o cancelli i dati di Chrome li perdi: sincronizza a fine turno.
- Le foto vengono ridotte a 1600 px per contenere le dimensioni del Word.
- La firma non è compresa: si firma il documento stampato, oppure il PDF con firma digitale.
- Scarico BAW: si aggiunge appena arriva il file.
- Per aggiornare l'app basta ricaricare i file su GitHub; il telefono prende la nuova versione alla riapertura successiva.
