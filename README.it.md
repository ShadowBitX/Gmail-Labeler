
# Gmail Labeler

[English](README.md) | [Italiano](README.it.md)

Gmail Labeler è un'applicazione per Windows che consente di applicare un'etichetta Gmail a grandi quantità di email, selezionandole in base alle date.

L'applicazione permette di organizzare i messaggi senza doverli selezionare manualmente uno per uno.

## Funzionalità

- **Etichettatura per intervallo di date:** applica un'etichetta alle email comprese tra due date, estremi inclusi.
- **Etichettatura delle email precedenti a una data:** considera i messaggi antecedenti al giorno selezionato, escludendo quest'ultimo.
- **Etichettatura delle email successive a una data:** considera i messaggi dal giorno selezionato in poi, includendolo.
- **Creazione automatica delle etichette:** se l'etichetta indicata non esiste, viene creata automaticamente; se esiste, viene riutilizzata.
- **Etichettatura in blocco:** consente di applicare una singola etichetta a molti messaggi in un'unica operazione.
- **Conferma preventiva:** prima di modificare i messaggi, mostra il numero di email individuate e chiede conferma.
- **Interruzione dell'elaborazione:** permette di interrompere l'operazione in corso.
- **Supporto per email lette e non lette:** include sia i messaggi letti sia quelli non letti.
- **Supporto per email archiviate:** comprende anche i messaggi archiviati che corrispondono ai criteri di ricerca.
- **Interfaccia multilingua:** italiano e inglese.
- **Temi grafici:** chiaro, scuro, oceano, foresta, tramonto e viola.

Gmail Labeler aggiunge un'etichetta ai messaggi: non li cancella, non li sposta e non ne modifica il contenuto.

La ricerca utilizza le funzionalità standard di Gmail e normalmente non include i messaggi presenti in Spam e Cestino.

## Struttura del progetto

GmailLabeler/
- src/
  - Setup-GmailLabeler.ps1
  - Build-GmailLabeler.ps1
- assets/
  - icon.ico
- Compilator.bat
- LICENSE
- README.it.md
- README.md

### Descrizione dei file

- `Setup-GmailLabeler.ps1`: contiene il codice dell'applicazione e la procedura guidata di installazione. Il codice Python dell'applicazione è incorporato direttamente nello script.
- `Build-GmailLabeler.ps1`: gestisce la compilazione dello script PowerShell in un eseguibile Windows tramite PS2EXE.
- `icon.ico`: icona utilizzata dall'applicazione e dal programma di installazione.
- `Compilator.bat`: avvia la procedura di compilazione.

## Requisiti

### Per chi utilizza l'applicazione

- Windows a 64 bit.
- Connessione Internet durante l'installazione.
- Permessi amministrativi se richiesti dalla cartella di installazione scelta.

Il programma di installazione verifica la presenza di Python e delle dipendenze necessarie. Se non trova una versione compatibile di Python, può installarla automaticamente.

### Per chi sviluppa e compila il progetto

- Windows.
- Windows PowerShell.
- PS2EXE, installabile automaticamente durante la compilazione.
- Connessione Internet per scaricare PS2EXE, se necessario.

## Configurazione Google Cloud

Per consentire a Gmail Labeler di accedere alla casella di posta è necessario configurare un progetto Google Cloud e abilitare Gmail API.

1. Accedere a [Google Cloud Console](https://console.cloud.google.com/).
2. Creare un nuovo progetto oppure selezionarne uno esistente.
3. Abilitare **Gmail API**.
4. Configurare la schermata di consenso OAuth.
5. Creare un ID client OAuth scegliendo il tipo di applicazione **App desktop**.
6. Scaricare il file JSON delle credenziali OAuth.

Il file scaricato dovrà essere selezionato durante il primo avvio dell'applicazione, quando richiesto.

## Autorizzazioni e sicurezza

Gmail Labeler utilizza l'autorizzazione OAuth `gmail.modify`, necessaria per applicare etichette ai messaggi.

Le credenziali vengono conservate sul computer dell'utente:

- `credentials.json`: file delle credenziali OAuth.
- `token.json`: token di autenticazione generato durante l'autorizzazione.

I file vengono archiviati nella cartella:

%APPDATA%\Gmail Labeler\credentials

Le preferenze dell'applicazione, come lingua e tema, vengono salvate in:

%APPDATA%\Gmail Labeler\user_data

Le credenziali e i token non devono essere pubblicati.

## Installazione

1. Avviare `Setup-GmailLabeler.exe`.
2. Seguire le istruzioni della procedura guidata.
3. Scegliere la cartella di destinazione.
4. Attendere il completamento dell'installazione.

Il programma di installazione prepara automaticamente l'ambiente necessario:

- verifica la presenza di Python;
- installa Python, se necessario;
- verifica e installa i pacchetti richiesti;
- genera l'applicazione tramite PyInstaller;
- crea il programma di disinstallazione;
- copia i file nella cartella scelta;
- registra l'applicazione tra quelle installate in Windows;
- offre la possibilità di creare un collegamento sul Desktop.

L'installazione richiede una connessione Internet per gli eventuali download di Python e dei pacchetti necessari.

## Utilizzo

Al primo avvio viene richiesta l'autorizzazione ad accedere all'account Google.

Dopo l'autenticazione è possibile scegliere una delle tre modalità di ricerca:

### 1. Intervallo di date

Seleziona una data iniziale e una data finale.

Vengono considerate entrambe le date.

### 2. Email precedenti a una data

Seleziona una data.

Vengono considerate le email precedenti al giorno selezionato, senza includere quel giorno.

### 3. Email successive a una data

Seleziona una data.

Vengono considerate le email dal giorno selezionato in poi, includendo quel giorno.

Dopo aver scelto il criterio, inserisci il nome dell'etichetta da applicare.

L'applicazione individua i messaggi corrispondenti e mostra una richiesta di conferma prima di procedere.

Se l'etichetta non esiste, viene creata automaticamente.

## Elaborazione dei messaggi

Per gestire grandi quantità di email, l'applicazione utilizza una ricerca paginata e operazioni in blocco.

- La ricerca recupera fino a 500 risultati per pagina.
- Le operazioni di modifica possono elaborare fino a 1.000 messaggi per blocco.
- È possibile interrompere l'elaborazione.

## Esportazione delle email

Gmail Labeler può essere utilizzato anche per preparare le email all'esportazione tramite Google Takeout.

È possibile applicare un'etichetta ai messaggi desiderati e successivamente utilizzare Google Takeout per esportare i messaggi associati a quell'etichetta in formato MBOX.

## Interfaccia e preferenze

L'applicazione supporta italiano e inglese e dispone di sei temi grafici:

- Chiaro
- Scuro
- Oceano
- Foresta
- Tramonto
- Viola

La lingua viene inizialmente rilevata in base alla lingua di Windows. Le preferenze di tema e lingua possono essere conservate tra gli avvii dell'applicazione.

## Compilazione del programma di installazione

Il repository contiene gli script necessari per generare `Setup-GmailLabeler.exe`.

Per avviare la compilazione:

1. Aprire la cartella principale del progetto.
2. Avviare `Compilator.bat`.
3. Se PS2EXE non è installato, lo script chiederà se si desidera installarlo.
4. Confermare l'installazione, se richiesta.
5. Attendere il completamento della compilazione.

Al termine verrà generato il file `Setup-GmailLabeler.exe` nella cartella principale del progetto.

La compilazione utilizza PS2EXE con le seguenti caratteristiche:

- eseguibile senza console;
- modalità STA;
- richiesta di privilegi amministrativi;
- architettura x64;
- icona personalizzata;
- informazioni di versione, titolo, descrizione, azienda e prodotto.

**Nota:** la compilazione tramite PS2EXE genera il programma di installazione. È poi il programma di installazione, una volta avviato sul computer dell'utente, a preparare l'ambiente necessario e generare l'applicazione Gmail Labeler tramite PyInstaller.

## Disinstallazione

L'applicazione include un programma di disinstallazione.

Durante la rimozione viene offerta la possibilità di eliminare anche le credenziali Google e i dati utente, comprese le preferenze salvate.

Se si sceglie di conservarli, questi dati rimangono sul computer.

## GitHub e distribuzione

GitHub viene utilizzato per ospitare il codice sorgente e distribuire GmailLabeler.

Il repository include:

- Script PowerShell per la compilazione e l'installazione;
- Icona dell'applicazione;
- File README;
- File LICENSE.

L'installer compilato (`Setup-GmailLabeler.exe`) viene distribuito tramite GitHub Releases.

Le credenziali dell'utente e i token di autenticazione vengono salvati localmente sul computer dell'utente e non sono inclusi nel repository o nell'installer.## GitHub e distribuzione

## Licenza

Questo progetto è distribuito con licenza MIT.
Consulta il file [LICENSE](LICENSE) per maggiori informazioni.