# ============================================================
#  Gmail Labeler - Setup / Installazione guidata
#  Italiano / English (default: lingua del sistema, altrimenti English)
#  Controlla/installa i prerequisiti, compila l'app e il disinstallatore
#  e li copia nella cartella scelta. Nessun riavvio necessario.
#  Richiede i permessi di amministratore (installa in Programmi).
#  Contiene già incorporati: gmail_labeler_gui.py, uninstall.cs, icon.ico
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ErrorActionPreference = 'Stop'

# ------------------------------------------------------------
# File incorporati
# ------------------------------------------------------------
$PY_SOURCE = @'
"""
Gmail Labeler - applicazione grafica / graphical application
Etichetta le email prima di una data, da una data in poi o in un intervallo di date,
per poi esportarle con Google Takeout.
Labels emails before a date, from a date onwards or within a date range,
so they can be exported with Google Takeout.
Non elimina, sposta o modifica il contenuto dei messaggi.
"""

import os
import queue
import random
import shutil
import sys
import threading
import time
import tkinter as tk
import webbrowser
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ============================================================
# COSTANTI
# ============================================================

SCOPES = ["https://www.googleapis.com/auth/gmail.modify"]
PAGE_SIZE = 500
BATCH_SIZE = 1000
SLEEP_BETWEEN_BATCHES = 0.5
MAX_RETRIES = 5

URL_CONSOLE = "https://console.cloud.google.com"
URL_TAKEOUT = "https://takeout.google.com"

# Con PyInstaller __file__ punta a una cartella temporanea
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent

# Cartella delle risorse dell'app: con l'exe in modalità cartella è "_internal"
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", BASE_DIR))

# Dati dell'utente in %APPDATA%\Gmail Labeler (la cartella di installazione in Programmi non è
# scrivibile da un utente normale e l'app non vi scrive mai):
#   credentials\  -> credentials.json e token.json
#   user_data\    -> theme.txt e language.txt (preferenze)
APP_DATA = Path(os.environ.get("APPDATA") or Path.home()) / "Gmail Labeler"
DATA_DIR = APP_DATA / "credentials"
USER_DATA_DIR = APP_DATA / "user_data"
for _folder in (DATA_DIR, USER_DATA_DIR):
    try:
        _folder.mkdir(parents=True, exist_ok=True)
    except OSError:
        pass

TOKEN_FILE = DATA_DIR / "token.json"
CREDENTIALS_FILE = DATA_DIR / "credentials.json"
THEME_FILE = USER_DATA_DIR / "theme.txt"
LANG_FILE = USER_DATA_DIR / "language.txt"


# ============================================================
# TESTI E LINGUA (it / en)
# ============================================================

TEXT = {
    "it": {
        "header_sub": "Etichetta le email per data, poi esportale con Google Takeout",
        "lbl_theme": "Tema",
        "lbl_language": "Lingua",
        "grp_links": "Link e file",
        "btn_console": "Apri Google Cloud Console",
        "btn_takeout": "Apri Google Takeout",
        "btn_choose_cred": "Seleziona credentials.json",
        "btn_open_cred": "Apri cartella credentials",
        "guide_closed": "▶ Guida configurazione Google Cloud Console",
        "guide_open": "▼ Guida configurazione Google Cloud Console",
        "guide": r"""SETUP UNA TANTUM (Google Cloud Console)

1. Apri console.cloud.google.com (pulsante qui sopra) e accedi.

2. Crea un progetto: menu in alto > "Nuovo progetto", nome
   per esempio "gmail-labeler".

3. Menu > API e servizi > Libreria > cerca "Gmail API" > Abilita.

4. API e servizi > Schermata consenso OAuth (o "Google Auth Platform"):
   - tipo utente: Esterno
   - inserisci nome app e la tua email
   - in "Utenti di test" aggiungi il tuo indirizzo Gmail.

5. API e servizi > Credenziali > Crea credenziali >
   ID client OAuth > tipo "App desktop".

6. Scarica il JSON, poi usa il pulsante "Seleziona credentials.json"
   in questa finestra: il file viene copiato nella cartella
   credentials dell'utente: %APPDATA%\Gmail Labeler\credentials

7. Premi "Connetti Gmail": si apre il browser, accetta i permessi.
   (Se compare "app non verificata": Avanzate > Vai a ... > Continua.)
   Dalla volta successiva non serve più il browser.

DOPO L'ETICHETTATURA
Apri Google Takeout, "Deseleziona tutto", scegli Gmail >
"Tutti i dati di Gmail inclusi" > seleziona solo l'etichetta
e esporta (formato MBOX).
""",
        "grp_settings": "Impostazioni",
        "lbl_mode": "Modalità:",
        "mode_before": "Prima di una data (esclusa)",
        "mode_after": "Da una data in poi (inclusa)",
        "mode_range": "Intervallo tra due date (incluse)",
        "date1_before": "Messaggi precedenti al:",
        "date1_after": "Messaggi dal giorno (incluso) in poi:",
        "date1_range": "Dal giorno (incluso):",
        "date2": "Al giorno (incluso):",
        "lbl_dmy": "giorno / mese / anno",
        "lbl_label_name": "Nome etichetta:",
        "label_placeholder": "Es. Backup",
        "preview": "Query Gmail: {0}",
        "grp_actions": "Operazioni",
        "btn_connect": "1. Connetti Gmail",
        "btn_search": "2. Cerca messaggi",
        "btn_apply": "3. Applica etichetta",
        "btn_stop": "Interrompi",
        "btn_disconnect": "Disconnetti account",
        "grp_log": "Log",
        "status": "credentials.json: {0}  |  Gmail: {1}  |  Messaggi trovati: {2}",
        "found": "trovato",
        "missing": "MANCANTE",
        "connected": "connesso",
        "not_connected": "non connesso",
        "err_title": "Errore",
        "invalid_date": "Data non valida (per esempio 31/02).",
        "range_inverted": "Nell'intervallo la data finale non può precedere quella iniziale.",
        "label_empty": "Il nome dell'etichetta non può essere vuoto.",
        "dlg_choose_title": "Seleziona il file JSON scaricato da Google Cloud",
        "ft_json": "File JSON",
        "ft_all": "Tutti i file",
        "period_before": "messaggi precedenti al {0}",
        "period_after": "messaggi dal {0} in poi",
        "period_range": "messaggi dal {0} al {1} (estremi inclusi)",
        "log_cred_dir": "Cartella credenziali: {0}",
        "log_cred_copied": "credentials.json copiato in {0}",
        "log_err_copy": "ERRORE copia credentials.json: {0}",
        "log_err_open": "ERRORE apertura cartella: {0}",
        "log_connecting": "Connessione a Gmail in corso...",
        "log_connected": "Connesso come {0} ({1} messaggi totali)",
        "log_token_deleted": "token.json eliminato: alla prossima connessione servirà una nuova autorizzazione.",
        "log_no_token": "Nessun token da eliminare.",
        "log_period": "Periodo: {0}",
        "log_query": "Query: {0}",
        "log_total": "Totale messaggi trovati: {0}",
        "log_none": "Nessun messaggio trovato.",
        "log_interrupted": "Operazione interrotta dall'utente.",
        "log_err_http": "ERRORE Google API HTTP {0}: {1}",
        "log_err": "ERRORE: {0}",
        "log_applying": "Applicazione dell'etichetta in corso...",
        "log_cancelled": "Operazione annullata.",
        "log_done": "COMPLETATO: {0} messaggi elaborati con l'etichetta '{1}'.",
        "log_done2": "Ora puoi verificare l'etichetta in Gmail e usare Google Takeout.",
        "dlg_repeat_title": "Ricerca da ripetere",
        "dlg_repeat_msg": "Modalità o date sono cambiate dopo l'ultima ricerca. Premi di nuovo 'Cerca messaggi'.",
        "dlg_confirm_title": "Conferma",
        "dlg_confirm_msg": "Applicare l'etichetta «{0}» a {1} messaggi?\n\nPeriodo: {2}\nQuery Gmail: {3}\n\n"
                           "Verrà aggiunta solo l'etichetta: nessun messaggio sarà eliminato o spostato.",
        "retry": "Errore temporaneo HTTP {0}. Nuovo tentativo tra {1}s...",
        "cred_missing": "credentials.json non trovato in:\n{0}\nUsa il pulsante 'Seleziona credentials.json'.",
        "token_unreadable": "Token esistente non leggibile: serve una nuova autorizzazione.",
        "token_refresh_fail": "Impossibile rinnovare il token: serve una nuova autorizzazione.",
        "opening_browser": "Apertura del browser per l'autorizzazione...",
        "label_exists": "Etichetta già esistente: {0} (ID: {1})",
        "label_created": "Etichetta creata: {0} (ID: {1})",
        "msgs_collected": "Messaggi raccolti: {0}",
        "msgs_processed": "Messaggi elaborati: {0}/{1}",
    },
    "en": {
        "header_sub": "Label emails by date, then export them with Google Takeout",
        "lbl_theme": "Theme",
        "lbl_language": "Language",
        "grp_links": "Links and files",
        "btn_console": "Open Google Cloud Console",
        "btn_takeout": "Open Google Takeout",
        "btn_choose_cred": "Select credentials.json",
        "btn_open_cred": "Open credentials folder",
        "guide_closed": "▶ Google Cloud Console setup guide",
        "guide_open": "▼ Google Cloud Console setup guide",
        "guide": r"""ONE-TIME SETUP (Google Cloud Console)

1. Open console.cloud.google.com (button above) and sign in.

2. Create a project: top menu > "New project", name it
   for example "gmail-labeler".

3. Menu > APIs & Services > Library > search "Gmail API" > Enable.

4. APIs & Services > OAuth consent screen (or "Google Auth Platform"):
   - user type: External
   - enter the app name and your email
   - under "Test users" add your Gmail address.

5. APIs & Services > Credentials > Create credentials >
   OAuth client ID > type "Desktop app".

6. Download the JSON, then use the "Select credentials.json" button
   in this window: the file is copied to the user's
   credentials folder: %APPDATA%\Gmail Labeler\credentials

7. Press "Connect Gmail": the browser opens, accept the permissions.
   (If "unverified app" appears: Advanced > Go to ... > Continue.)
   From the next time the browser is no longer needed.

AFTER LABELLING
Open Google Takeout, "Deselect all", choose Gmail >
"All Mail data included" > select only the label
and export (MBOX format).
""",
        "grp_settings": "Settings",
        "lbl_mode": "Mode:",
        "mode_before": "Before a date (excluded)",
        "mode_after": "From a date onwards (included)",
        "mode_range": "Date range (both dates included)",
        "date1_before": "Messages before:",
        "date1_after": "Messages from this day (included) onwards:",
        "date1_range": "From day (included):",
        "date2": "To day (included):",
        "lbl_dmy": "day / month / year",
        "lbl_label_name": "Label name:",
        "label_placeholder": "Ex. Backup",
        "preview": "Gmail query: {0}",
        "grp_actions": "Actions",
        "btn_connect": "1. Connect Gmail",
        "btn_search": "2. Search messages",
        "btn_apply": "3. Apply label",
        "btn_stop": "Stop",
        "btn_disconnect": "Disconnect account",
        "grp_log": "Log",
        "status": "credentials.json: {0}  |  Gmail: {1}  |  Messages found: {2}",
        "found": "found",
        "missing": "MISSING",
        "connected": "connected",
        "not_connected": "not connected",
        "err_title": "Error",
        "invalid_date": "Invalid date (for example 31/02).",
        "range_inverted": "In a range the end date cannot be earlier than the start date.",
        "label_empty": "The label name cannot be empty.",
        "dlg_choose_title": "Select the JSON file downloaded from Google Cloud",
        "ft_json": "JSON files",
        "ft_all": "All files",
        "period_before": "messages before {0}",
        "period_after": "messages from {0} onwards",
        "period_range": "messages from {0} to {1} (both dates included)",
        "log_cred_dir": "Credentials folder: {0}",
        "log_cred_copied": "credentials.json copied to {0}",
        "log_err_copy": "ERROR copying credentials.json: {0}",
        "log_err_open": "ERROR opening folder: {0}",
        "log_connecting": "Connecting to Gmail...",
        "log_connected": "Connected as {0} ({1} messages in total)",
        "log_token_deleted": "token.json deleted: a new authorization will be required at the next connection.",
        "log_no_token": "No token to delete.",
        "log_period": "Period: {0}",
        "log_query": "Query: {0}",
        "log_total": "Total messages found: {0}",
        "log_none": "No messages found.",
        "log_interrupted": "Operation stopped by the user.",
        "log_err_http": "ERROR Google API HTTP {0}: {1}",
        "log_err": "ERROR: {0}",
        "log_applying": "Applying the label...",
        "log_cancelled": "Operation cancelled.",
        "log_done": "COMPLETED: {0} messages processed with label '{1}'.",
        "log_done2": "You can now check the label in Gmail and use Google Takeout.",
        "dlg_repeat_title": "Search must be repeated",
        "dlg_repeat_msg": "Mode or dates changed after the last search. Press 'Search messages' again.",
        "dlg_confirm_title": "Confirm",
        "dlg_confirm_msg": "Apply the label «{0}» to {1} messages?\n\nPeriod: {2}\nGmail query: {3}\n\n"
                           "Only the label will be added: no message will be deleted or moved.",
        "retry": "Temporary HTTP error {0}. Retrying in {1}s...",
        "cred_missing": "credentials.json not found in:\n{0}\nUse the 'Select credentials.json' button.",
        "token_unreadable": "Existing token unreadable: a new authorization is required.",
        "token_refresh_fail": "Unable to refresh the token: a new authorization is required.",
        "opening_browser": "Opening the browser for authorization...",
        "label_exists": "Label already exists: {0} (ID: {1})",
        "label_created": "Label created: {0} (ID: {1})",
        "msgs_collected": "Messages collected: {0}",
        "msgs_processed": "Messages processed: {0}/{1}",
    },
}

LANGUAGE_LABELS = {"it": "Italiano", "en": "English"}
LANG = "it"


def tr(key, *args):
    """Testo nella lingua corrente; i segnaposto {0}, {1}... vengono sostituiti da args."""
    text = TEXT[LANG].get(key) or TEXT["en"][key]
    return text.format(*args) if args else text


def set_language(lang):
    global LANG
    LANG = lang if lang in TEXT else "en"


def _windows_ui_language_id():
    """LANGID della lingua di visualizzazione di Windows (None se non disponibile)."""
    if sys.platform != "win32":
        return None
    try:
        import ctypes
        return int(ctypes.windll.kernel32.GetUserDefaultUILanguage()) or None
    except (AttributeError, OSError, ValueError):
        return None


def detect_system_language(root):
    """Lingua di sistema: italiano se Windows è in italiano, altrimenti inglese."""
    langid = _windows_ui_language_id()
    if langid:
        return "it" if (langid & 0x3FF) == 0x10 else "en"     # 0x10 = lingua italiana
    # ripieghi: impostazioni internazionali di Tcl e variabili d'ambiente
    locale_name = ""
    try:
        locale_name = str(root.tk.call("::msgcat::mclocale")).lower()
    except tk.TclError:
        pass
    if not locale_name:
        locale_name = (os.environ.get("LC_ALL") or os.environ.get("LANG") or "").lower()
    return "it" if locale_name.startswith("it") else "en"


def read_saved_language():
    """Lingua salvata (it / en) oppure None se l'utente non l'ha ancora fissata."""
    try:
        value = LANG_FILE.read_text(encoding="utf-8").strip().lower()
        if value in TEXT:
            return value
    except OSError:
        pass
    return None


def load_language(root):
    """Lingua salvata; al primo avvio quella di sistema."""
    return read_saved_language() or detect_system_language(root)


def save_language(lang):
    try:
        LANG_FILE.parent.mkdir(parents=True, exist_ok=True)
        LANG_FILE.write_text(lang, encoding="utf-8")
    except OSError:
        pass


# ============================================================
# LOGICA GMAIL
# ============================================================

class Cancelled(Exception):
    pass


def execute_with_retry(request, log):
    for attempt in range(MAX_RETRIES + 1):
        try:
            return request.execute()
        except HttpError as exc:
            status = exc.resp.status
            retryable = status in (429, 500, 502, 503, 504)
            if not retryable or attempt >= MAX_RETRIES:
                raise
            delay = min(2 ** attempt, 30) + random.uniform(0, 1)
            log(tr("retry", status, f"{delay:.1f}"))
            time.sleep(delay)


def get_service(log):
    if not CREDENTIALS_FILE.is_file():
        raise FileNotFoundError(tr("cred_missing", DATA_DIR))

    creds = None
    if TOKEN_FILE.is_file():
        try:
            creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
        except (ValueError, OSError):
            log(tr("token_unreadable"))

    if creds and not creds.valid:
        if creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except RefreshError:
                log(tr("token_refresh_fail"))
                creds = None
        else:
            creds = None

    if not creds or not creds.valid:
        log(tr("opening_browser"))
        flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
        creds = flow.run_local_server(port=0)

    TOKEN_FILE.write_text(creds.to_json(), encoding="utf-8")
    return build("gmail", "v1", credentials=creds, cache_discovery=False)


def get_or_create_label(service, name, log):
    response = execute_with_retry(service.users().labels().list(userId="me"), log)
    for label in response.get("labels", []):
        if label["name"] == name:
            log(tr("label_exists", name, label["id"]))
            return label["id"]

    response = execute_with_retry(
        service.users().labels().create(
            userId="me",
            body={
                "name": name,
                "labelListVisibility": "labelShow",
                "messageListVisibility": "show",
            },
        ),
        log,
    )
    log(tr("label_created", name, response["id"]))
    return response["id"]


def list_all_messages(service, query, log, cancel):
    all_ids, seen, page_token = [], set(), None
    while True:
        if cancel.is_set():
            raise Cancelled()
        params = {"userId": "me", "q": query, "maxResults": PAGE_SIZE}
        if page_token:
            params["pageToken"] = page_token
        response = execute_with_retry(service.users().messages().list(**params), log)
        for message in response.get("messages", []):
            if message["id"] not in seen:
                seen.add(message["id"])
                all_ids.append(message["id"])
        log(tr("msgs_collected", len(all_ids)))
        page_token = response.get("nextPageToken")
        if not page_token:
            break
    return all_ids


def batch_label(service, ids, label_id, log, cancel):
    total, processed = len(ids), 0
    for start in range(0, total, BATCH_SIZE):
        if cancel.is_set():
            raise Cancelled()
        batch = ids[start:start + BATCH_SIZE]
        execute_with_retry(
            service.users().messages().batchModify(
                userId="me", body={"ids": batch, "addLabelIds": [label_id]}
            ),
            log,
        )
        processed += len(batch)
        log(tr("msgs_processed", processed, total))
        if processed < total:
            time.sleep(SLEEP_BETWEEN_BATCHES)
    return processed


# ============================================================
# PERIODO E QUERY GMAIL
# ============================================================

MODE_KEYS = ("before", "after", "range")


def build_query(mode, d1, d2=None):
    """Query Gmail: after: include il giorno indicato, before: lo esclude."""
    if mode == "before":
        return f"before:{d1:%Y/%m/%d}"
    if mode == "after":
        return f"after:{d1:%Y/%m/%d}"
    if d2 is None or d2 < d1:
        raise ValueError(tr("range_inverted"))
    # il giorno finale deve essere incluso: before: il giorno successivo
    return f"after:{d1:%Y/%m/%d} before:{d2 + timedelta(days=1):%Y/%m/%d}"


def describe_period(mode, d1, d2=None):
    fmt = "%d/%m/%Y" if LANG == "it" else "%Y-%m-%d"
    if mode == "before":
        return tr("period_before", d1.strftime(fmt))
    if mode == "after":
        return tr("period_after", d1.strftime(fmt))
    return tr("period_range", d1.strftime(fmt), d2.strftime(fmt))


# ============================================================
# TEMI
# ============================================================

THEMES = {
    "Chiaro": dict(bg="#F3F5F9", panel="#FFFFFF", fg="#1F2937", muted="#686f7d",
                   accent="#2563EB", accent_text="#1D4ED8", accent_fg="#FFFFFF",
                   border="#D3D9E4", field_border="#AEB8CB", entry="#FFFFFF",
                   log_bg="#FFFFFF", log_fg="#111827", ok="#15803D", err="#B91C1C"),
    "Scuro": dict(bg="#161C2C", panel="#202A42", fg="#E5E7EB", muted="#9AA4B5",
                  accent="#2563EB", accent_text="#8AB4FF", accent_fg="#FFFFFF",
                  border="#34415F", field_border="#5B6B8F", entry="#202A42",
                  log_bg="#202A42", log_fg="#E1E6F0", ok="#4ADE80", err="#F87171"),
    "Oceano": dict(bg="#E8F3FA", panel="#FFFFFF", fg="#0B3954", muted="#4d7086",
                   accent="#0369A1", accent_text="#0369A1", accent_fg="#FFFFFF",
                   border="#BBD7E8", field_border="#8FB7D1", entry="#FFFFFF",
                   log_bg="#F7FBFE", log_fg="#0B3954", ok="#0F766E", err="#B91C1C"),
    "Foresta": dict(bg="#EDF4EE", panel="#FFFFFF", fg="#1B3A2A", muted="#557261",
                    accent="#2E7D32", accent_text="#2E7D32", accent_fg="#FFFFFF",
                    border="#C5D9C8", field_border="#9CBBA2", entry="#FFFFFF",
                    log_bg="#F8FCF8", log_fg="#1B3A2A", ok="#166534", err="#B91C1C"),
    "Tramonto": dict(bg="#FFF3EA", panel="#FFFFFF", fg="#4A2C1A", muted="#8A6A55",
                     accent="#C2410C", accent_text="#C2410C", accent_fg="#FFFFFF",
                     border="#EBCDB9", field_border="#D9A98A", entry="#FFFFFF",
                     log_bg="#FFFBF8", log_fg="#4A2C1A", ok="#15803D", err="#B91C1C"),
    "Viola": dict(bg="#F4F0FB", panel="#FFFFFF", fg="#2E1A4A", muted="#75658e",
                  accent="#7C3AED", accent_text="#6D28D9", accent_fg="#FFFFFF",
                  border="#D8CCEE", field_border="#B7A4DE", entry="#FFFFFF",
                  log_bg="#FBF9FE", log_fg="#2E1A4A", ok="#15803D", err="#B91C1C"),
}
THEME_LABELS = {
    "Chiaro": {"it": "Chiaro", "en": "Light"},
    "Scuro": {"it": "Scuro", "en": "Dark"},
    "Oceano": {"it": "Oceano", "en": "Ocean"},
    "Foresta": {"it": "Foresta", "en": "Forest"},
    "Tramonto": {"it": "Tramonto", "en": "Sunset"},
    "Viola": {"it": "Viola", "en": "Violet"},
}
DEFAULT_THEME = "Chiaro"


def theme_label(key):
    return THEME_LABELS[key][LANG]


def theme_from_label(label):
    for key, names in THEME_LABELS.items():
        if label in names.values():
            return key
    return DEFAULT_THEME


def load_theme():
    try:
        name = THEME_FILE.read_text(encoding="utf-8").strip()
        if name in THEMES:
            return name
    except OSError:
        pass
    return DEFAULT_THEME


def save_theme(name):
    try:
        THEME_FILE.parent.mkdir(parents=True, exist_ok=True)
        THEME_FILE.write_text(name, encoding="utf-8")
    except OSError:
        pass


def _rgb(color):
    return tuple(int(color[i:i + 2], 16) for i in (1, 3, 5))


def shade(color, factor):
    """factor < 1 scurisce, factor > 1 schiarisce."""
    r, g, b = _rgb(color)
    if factor < 1:
        r, g, b = (int(c * factor) for c in (r, g, b))
    else:
        r, g, b = (int(c + (255 - c) * (factor - 1)) for c in (r, g, b))
    return f"#{r:02x}{g:02x}{b:02x}"


def is_dark(color):
    r, g, b = _rgb(color)
    return (0.299 * r + 0.587 * g + 0.114 * b) < 128


# ============================================================
# INTERFACCIA
# ============================================================

class DateSelector(ttk.Frame):
    """Tre menu a tendina: giorno / mese / anno."""

    def __init__(self, master, default, on_change=None):
        super().__init__(master)
        self.day = tk.StringVar(value=f"{default.day:02d}")
        self.month = tk.StringVar(value=f"{default.month:02d}")
        self.year = tk.StringVar(value=str(default.year))
        self.boxes = []
        specs = (
            (self.day, 4, [f"{d:02d}" for d in range(1, 32)]),
            (self.month, 4, [f"{m:02d}" for m in range(1, 13)]),
            (self.year, 6, [str(y) for y in range(datetime.now().year, 1999, -1)]),
        )
        for i, (var, width, values) in enumerate(specs):
            if i:
                ttk.Label(self, text="/").pack(side="left", padx=3)
            box = ttk.Combobox(self, textvariable=var, width=width, state="readonly", values=values)
            box.pack(side="left")
            if on_change:
                box.bind("<<ComboboxSelected>>", lambda _event: on_change())
            self.boxes.append(box)
        self.hint = ttk.Label(self, style="Muted.TLabel")
        self.hint.pack(side="left", padx=(12, 0))

    def get(self):
        try:
            return datetime(int(self.year.get()), int(self.month.get()), int(self.day.get()))
        except ValueError:
            raise ValueError(tr("invalid_date"))


class PlaceholderEntry(ttk.Entry):
    """Campo di testo con un suggerimento grigio quando è vuoto.

    Il suggerimento non è un valore: value() restituisce "" finché l'utente non scrive.
    """

    def __init__(self, master, placeholder="", **kwargs):
        super().__init__(master, **kwargs)
        self._placeholder = placeholder
        self._showing = False
        self.bind("<FocusIn>", self._on_focus_in)
        self.bind("<FocusOut>", self._on_focus_out)
        self._show_if_empty()

    def _show_if_empty(self):
        if not self._showing and not self.get():
            self.insert(0, self._placeholder)
            self.configure(style="Placeholder.TEntry")
            self._showing = True

    def _on_focus_in(self, _event=None):
        if self._showing:
            self.delete(0, "end")
            self.configure(style="TEntry")
            self._showing = False

    def _on_focus_out(self, _event=None):
        self._show_if_empty()

    def value(self):
        return "" if self._showing else self.get()

    def set_placeholder(self, text):
        self._placeholder = text
        if self._showing:
            self.delete(0, "end")
            self.insert(0, text)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gmail Labeler")
        self.geometry("920x800")
        self.minsize(820, 680)
        try:
            self.iconbitmap(str(RESOURCE_DIR / "icon.ico"))
        except tk.TclError:
            pass

        self.log_queue = queue.Queue()
        self.cancel = threading.Event()
        self.service = None
        self.found_ids = []
        self.found_query = None
        self.busy = False

        self.style = ttk.Style(self)
        self._texts = []       # caselle di testo (non ttk)
        self._combos = []      # menu a tendina
        self._i18n = []        # (widget, opzione, chiave) da ritradurre al cambio lingua
        self.theme_key = DEFAULT_THEME
        self.mode_key = MODE_KEYS[0]

        # La lingua di sistema vale solo al primo avvio: viene salvata subito e da lì in poi
        # resta quella scelta dall'utente con la tendina.
        saved = read_saved_language()
        set_language(saved or detect_system_language(self))
        if saved is None:
            save_language(LANG)
        self._build_ui()
        self.bind_all("<Button-1>", self._on_click_anywhere, add="+")
        self.apply_theme(load_theme(), save=False)
        self._poll_log()
        self.log(tr("log_cred_dir", DATA_DIR))
        self._on_mode_change()
        self._refresh_status()

    # ---------- costruzione UI ----------

    def _reg(self, widget, key, option="text"):
        """Imposta il testo e lo ricorda per il cambio lingua."""
        self._i18n.append((widget, option, key))
        widget.configure(**{option: tr(key)})
        return widget

    def _combo(self, parent, **kwargs):
        box = ttk.Combobox(parent, state="readonly", **kwargs)
        self._combos.append(box)
        return box

    def _make_text(self, parent, height, font):
        frame = ttk.Frame(parent)
        text = tk.Text(frame, height=height, wrap="word", state="disabled", relief="flat",
                       borderwidth=0, font=font, padx=8, pady=6, highlightthickness=2)
        bar = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=bar.set)
        bar.pack(side="right", fill="y")
        text.pack(side="left", fill="both", expand=True)
        self._texts.append(text)
        return frame, text

    def _build_ui(self):
        # Intestazione: titolo, lingua e tema
        header = ttk.Frame(self, style="Header.TFrame")
        header.pack(fill="x")
        left = ttk.Frame(header, style="Header.TFrame")
        left.pack(side="left", padx=18, pady=12)
        ttk.Label(left, text="Gmail Labeler", style="HeaderTitle.TLabel").pack(anchor="w")
        self._reg(ttk.Label(left, style="HeaderSub.TLabel"), "header_sub").pack(anchor="w")

        right = ttk.Frame(header, style="Header.TFrame")
        right.pack(side="right", padx=18)
        self._reg(ttk.Label(right, style="HeaderSub.TLabel"), "lbl_language").pack(side="left", padx=(0, 6))
        self.lang_var = tk.StringVar(value=LANGUAGE_LABELS[LANG])
        self.lang_box = self._combo(right, textvariable=self.lang_var,
                                    values=list(LANGUAGE_LABELS.values()), width=10)
        self.lang_box.pack(side="left", padx=(0, 16))
        self.lang_box.bind("<<ComboboxSelected>>", self._on_language_selected)
        self._reg(ttk.Label(right, style="HeaderSub.TLabel"), "lbl_theme").pack(side="left", padx=(0, 6))
        self.theme_var = tk.StringVar(value=theme_label(DEFAULT_THEME))
        self.theme_box = self._combo(right, textvariable=self.theme_var,
                                     values=[theme_label(k) for k in THEMES], width=11)
        self.theme_box.pack(side="left")
        self.theme_box.bind("<<ComboboxSelected>>", self._on_theme_selected)

        body = ttk.Frame(self, padding=(14, 8))
        body.pack(fill="both", expand=True)
        pad = {"pady": 5}

        # Link rapidi
        links = self._reg(ttk.LabelFrame(body), "grp_links")
        links.pack(fill="x", **pad)
        self._reg(ttk.Button(links, command=lambda: webbrowser.open(URL_CONSOLE)),
                  "btn_console").pack(side="left", padx=6, pady=8)
        self._reg(ttk.Button(links, command=lambda: webbrowser.open(URL_TAKEOUT)),
                  "btn_takeout").pack(side="left", padx=6, pady=8)
        self._reg(ttk.Button(links, command=self.choose_credentials),
                  "btn_choose_cred").pack(side="left", padx=6, pady=8)
        self._reg(ttk.Button(links, command=self.open_credentials_dir),
                  "btn_open_cred").pack(side="left", padx=6, pady=8)

        # Guida a tendina
        self.guide_open = False
        self.guide_btn = ttk.Button(body, command=self.toggle_guide)
        self.guide_btn.pack(fill="x", **pad)
        self.guide_frame, self.guide_box = self._make_text(body, 14, ("Segoe UI", 9))
        self._set_guide_button()
        self._set_guide_text()

        # Impostazioni: modalità, date, nome etichetta
        opts = self._reg(ttk.LabelFrame(body), "grp_settings")
        opts.pack(fill="x", **pad)
        opts.columnconfigure(1, weight=1)

        self._reg(ttk.Label(opts), "lbl_mode").grid(row=0, column=0, sticky="w", padx=10, pady=6)
        self.mode_var = tk.StringVar(value=tr("mode_" + self.mode_key))
        self.mode_box = self._combo(opts, textvariable=self.mode_var,
                                    values=[tr("mode_" + k) for k in MODE_KEYS], width=36)
        self.mode_box.grid(row=0, column=1, sticky="w", padx=10, pady=6)
        self.mode_box.bind("<<ComboboxSelected>>", self._on_mode_selected)

        self.lbl_date1 = ttk.Label(opts)
        self.lbl_date1.grid(row=1, column=0, sticky="w", padx=10, pady=6)
        self.date1 = DateSelector(opts, datetime(2020, 1, 1), self._update_preview)
        self.date1.grid(row=1, column=1, sticky="w", padx=10, pady=6)

        self.lbl_date2 = self._reg(ttk.Label(opts), "date2")
        self.lbl_date2.grid(row=2, column=0, sticky="w", padx=10, pady=6)
        self.date2 = DateSelector(opts, datetime(2020, 12, 31), self._update_preview)
        self.date2.grid(row=2, column=1, sticky="w", padx=10, pady=6)
        self._reg(self.date1.hint, "lbl_dmy")
        self._reg(self.date2.hint, "lbl_dmy")
        self._combos.extend(self.date1.boxes + self.date2.boxes)

        self._reg(ttk.Label(opts), "lbl_label_name").grid(row=3, column=0, sticky="w", padx=10, pady=6)
        self.label_entry = PlaceholderEntry(opts, tr("label_placeholder"), width=38)
        self.label_entry.grid(row=3, column=1, sticky="w", padx=10, pady=6)

        self.preview = tk.StringVar()
        ttk.Label(opts, textvariable=self.preview, style="Preview.TLabel").grid(
            row=4, column=0, columnspan=2, sticky="w", padx=10, pady=(2, 10))

        # Azioni
        actions = self._reg(ttk.LabelFrame(body), "grp_actions")
        actions.pack(fill="x", **pad)
        self.btn_connect = self._reg(ttk.Button(actions, style="Accent.TButton", command=self.connect),
                                     "btn_connect")
        self.btn_connect.pack(side="left", padx=6, pady=8)
        self.btn_search = self._reg(ttk.Button(actions, style="Accent.TButton", command=self.search),
                                    "btn_search")
        self.btn_search.pack(side="left", padx=6, pady=8)
        self.btn_apply = self._reg(ttk.Button(actions, style="Accent.TButton", command=self.apply_label),
                                   "btn_apply")
        self.btn_apply.pack(side="left", padx=6, pady=8)
        self.btn_stop = self._reg(ttk.Button(actions, command=self.cancel.set), "btn_stop")
        self.btn_stop.pack(side="left", padx=6, pady=8)
        self.btn_disconnect = self._reg(ttk.Button(actions, command=self.disconnect), "btn_disconnect")
        self.btn_disconnect.pack(side="right", padx=6, pady=8)

        self.status = tk.StringVar()
        ttk.Label(body, textvariable=self.status, style="Status.TLabel").pack(anchor="w", padx=4)

        # Log
        log_frame = self._reg(ttk.LabelFrame(body), "grp_log")
        log_frame.pack(fill="both", expand=True, **pad)
        box_frame, self.log_box = self._make_text(log_frame, 12, ("Consolas", 9))
        box_frame.pack(fill="both", expand=True, padx=6, pady=6)

    # ---------- temi ----------

    def apply_theme(self, key, save=True):
        if key not in THEMES:
            key = DEFAULT_THEME
        p = THEMES[key]
        self.theme_key = key
        self.theme_var.set(theme_label(key))
        hover = shade(p["panel"], 1.25 if is_dark(p["bg"]) else 0.94)
        fb = p["field_border"]          # bordo uniforme di tendine, campi di testo e caselle
        s = self.style
        s.theme_use("clam")
        self.configure(bg=p["bg"])

        s.configure(".", background=p["bg"], foreground=p["fg"], fieldbackground=p["entry"],
                    bordercolor=p["border"], lightcolor=p["border"], darkcolor=p["border"],
                    troughcolor=p["bg"], font=("Segoe UI", 9))
        s.configure("TFrame", background=p["bg"])
        s.configure("TLabel", background=p["bg"], foreground=p["fg"])
        s.configure("Muted.TLabel", foreground=p["muted"])
        s.configure("Status.TLabel", foreground=p["muted"])
        s.configure("Preview.TLabel", foreground=p["accent_text"], font=("Consolas", 9))
        s.configure("Header.TFrame", background=p["accent"])
        s.configure("HeaderTitle.TLabel", background=p["accent"], foreground=p["accent_fg"],
                    font=("Segoe UI Semibold", 16))
        s.configure("HeaderSub.TLabel", background=p["accent"], foreground=p["accent_fg"],
                    font=("Segoe UI", 9))
        s.configure("TLabelframe", background=p["bg"], bordercolor=p["border"],
                    lightcolor=p["border"], darkcolor=p["border"])
        s.configure("TLabelframe.Label", background=p["bg"], foreground=p["accent_text"],
                    font=("Segoe UI Semibold", 9))

        s.configure("TButton", background=p["panel"], foreground=p["fg"], bordercolor=p["border"],
                    lightcolor=p["panel"], darkcolor=p["panel"], padding=(12, 6))
        s.map("TButton",
              background=[("disabled", p["bg"]), ("active", hover)],
              foreground=[("disabled", p["muted"])])
        s.configure("Accent.TButton", background=p["accent"], foreground=p["accent_fg"],
                    bordercolor=p["accent"], lightcolor=p["accent"], darkcolor=p["accent"],
                    padding=(12, 6))
        s.map("Accent.TButton",
              background=[("disabled", p["border"]), ("active", shade(p["accent"], 0.88))],
              foreground=[("disabled", p["muted"])],
              bordercolor=[("disabled", p["border"])],
              lightcolor=[("disabled", p["border"])],
              darkcolor=[("disabled", p["border"])])

        s.configure("TCombobox", fieldbackground=p["entry"], background=p["panel"],
                    foreground=p["fg"], arrowcolor=p["fg"], bordercolor=fb,
                    lightcolor=fb, darkcolor=fb,
                    selectbackground=p["entry"], selectforeground=p["fg"], padding=3)
        s.map("TCombobox",
              fieldbackground=[("readonly", p["entry"]), ("disabled", p["bg"])],
              foreground=[("readonly", p["fg"]), ("disabled", p["muted"])],
              selectbackground=[("readonly", p["entry"])],
              selectforeground=[("readonly", p["fg"])],
              background=[("active", p["panel"]), ("readonly", p["panel"])],
              bordercolor=[("focus", p["accent"])],
              lightcolor=[("focus", p["accent"])],
              darkcolor=[("focus", p["accent"])])
        s.configure("TEntry", fieldbackground=p["entry"], foreground=p["fg"], insertcolor=p["fg"],
                    bordercolor=fb, lightcolor=fb, darkcolor=fb, padding=4)
        s.configure("Placeholder.TEntry", foreground=p["muted"])
        s.map("TEntry",
              bordercolor=[("focus", p["accent"])],
              lightcolor=[("focus", p["accent"])],
              darkcolor=[("focus", p["accent"])])
        s.configure("Vertical.TScrollbar", background=p["border"], troughcolor=p["bg"],
                    bordercolor=p["bg"], arrowcolor=p["fg"], lightcolor=p["border"],
                    darkcolor=p["border"])
        s.map("Vertical.TScrollbar", background=[("active", p["muted"])])
        # cornice dell'elenco che si apre dalle tendine: stesso bordo spesso e stesso colore
        s.configure("ComboboxPopdownFrame", background=p["entry"], bordercolor=fb,
                    lightcolor=fb, darkcolor=fb, borderwidth=2, relief="solid")

        # elenco a discesa dei menu e caselle di testo (widget non ttk)
        self.option_add("*TCombobox*Listbox.background", p["entry"])
        self.option_add("*TCombobox*Listbox.foreground", p["fg"])
        self.option_add("*TCombobox*Listbox.selectBackground", p["accent"])
        self.option_add("*TCombobox*Listbox.selectForeground", p["accent_fg"])
        for combo in self._combos:
            try:
                popdown = self.tk.call("ttk::combobox::PopdownWindow", str(combo))
                self.tk.call(f"{popdown}.f.l", "configure",
                             "-background", p["entry"], "-foreground", p["fg"],
                             "-selectbackground", p["accent"], "-selectforeground", p["accent_fg"])
            except tk.TclError:
                pass
        for text in self._texts:
            text.configure(background=p["log_bg"], foreground=p["log_fg"],
                           insertbackground=p["log_fg"], selectbackground=p["accent"],
                           selectforeground=p["accent_fg"], highlightthickness=2,
                           highlightbackground=fb, highlightcolor=p["accent"])
        self.log_box.tag_configure("err", foreground=p["err"])
        self.log_box.tag_configure("ok", foreground=p["ok"])

        if save:
            save_theme(key)

    def _on_click_anywhere(self, event):
        """Un clic su uno spazio vuoto (sfondo, etichette, cornici) toglie il focus ai campi."""
        widget = event.widget
        if isinstance(widget, str):          # widget creati da Tcl (es. l'elenco aperto di una tendina)
            return
        if isinstance(widget, (ttk.Entry, ttk.Combobox, ttk.Button, ttk.Scrollbar, tk.Text, tk.Listbox)):
            return
        try:
            self.focus_set()
        except tk.TclError:
            pass

    # ---------- lingua ----------

    def apply_language(self, lang, save=True):
        set_language(lang)
        self.lang_var.set(LANGUAGE_LABELS[LANG])
        for widget, option, key in self._i18n:
            widget.configure(**{option: tr(key)})
        self.theme_box.configure(values=[theme_label(k) for k in THEMES])
        self.theme_var.set(theme_label(self.theme_key))
        self.mode_box.configure(values=[tr("mode_" + k) for k in MODE_KEYS])
        self.mode_var.set(tr("mode_" + self.mode_key))
        self.label_entry.set_placeholder(tr("label_placeholder"))
        self._set_guide_text()
        self._set_guide_button()
        self._on_mode_change()
        self._refresh_status()
        if save:
            save_language(LANG)

    def _on_language_selected(self, _event=None):
        code = next((c for c, n in LANGUAGE_LABELS.items() if n == self.lang_var.get()), LANG)
        self.apply_language(code)

    def _on_theme_selected(self, _event=None):
        self.apply_theme(theme_from_label(self.theme_var.get()))

    def _on_mode_selected(self, _event=None):
        self.mode_key = next((k for k in MODE_KEYS if tr("mode_" + k) == self.mode_var.get()),
                             self.mode_key)
        self._on_mode_change()

    # ---------- utilità ----------

    def log(self, msg):
        """Thread-safe: può essere chiamata anche dai thread di lavoro."""
        self.log_queue.put(f"[{datetime.now():%H:%M:%S}] {msg}")

    @staticmethod
    def _log_tag(line):
        if "ERRORE" in line or "ERROR" in line:
            return "err"
        if "COMPLETATO" in line or "COMPLETED" in line or "Connesso come" in line or "Connected as" in line:
            return "ok"
        return ""

    def _poll_log(self):
        try:
            while True:
                line = self.log_queue.get_nowait()
                self.log_box.configure(state="normal")
                self.log_box.insert("end", line + "\n", self._log_tag(line))
                self.log_box.see("end")
                self.log_box.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(100, self._poll_log)

    def _set_guide_text(self):
        self.guide_box.configure(state="normal")
        self.guide_box.delete("1.0", "end")
        self.guide_box.insert("1.0", tr("guide"))
        self.guide_box.configure(state="disabled")

    def _set_guide_button(self):
        self.guide_btn.configure(text=tr("guide_open" if self.guide_open else "guide_closed"))

    def toggle_guide(self):
        self.guide_open = not self.guide_open
        if self.guide_open:
            self.guide_frame.pack(fill="x", pady=5, after=self.guide_btn)
        else:
            self.guide_frame.pack_forget()
        self._set_guide_button()

    def _refresh_status(self):
        cred = tr("found") if CREDENTIALS_FILE.is_file() else tr("missing")
        conn = tr("connected") if self.service else tr("not_connected")
        self.status.set(tr("status", cred, conn, len(self.found_ids)))
        busy = self.busy
        self.btn_connect.configure(state="disabled" if busy else "normal")
        self.btn_search.configure(state="normal" if (self.service and not busy) else "disabled")
        self.btn_apply.configure(
            state="normal" if (self.service and self.found_ids and not busy) else "disabled")
        self.btn_stop.configure(state="normal" if busy else "disabled")
        self.btn_disconnect.configure(state="disabled" if busy else "normal")

    def _on_mode_change(self):
        mode = self.mode_key
        self.lbl_date1.configure(text=tr("date1_" + mode))
        if mode == "range":
            self.lbl_date2.grid()
            self.date2.grid()
        else:
            self.lbl_date2.grid_remove()
            self.date2.grid_remove()
        self._update_preview()

    def _get_params(self):
        mode = self.mode_key
        d1 = self.date1.get()
        d2 = self.date2.get() if mode == "range" else None
        return mode, d1, d2

    def _update_preview(self):
        try:
            mode, d1, d2 = self._get_params()
            self.preview.set(tr("preview", build_query(mode, d1, d2)))
        except ValueError as exc:
            self.preview.set(str(exc))

    def _get_label(self):
        name = self.label_entry.value().strip()
        if not name:
            raise ValueError(tr("label_empty"))
        return name

    def run_async(self, func):
        """Esegue func in un thread, gestendo errori e stato dei pulsanti."""
        self.busy = True
        self.cancel.clear()
        self._refresh_status()

        def worker():
            try:
                func()
            except Cancelled:
                self.log(tr("log_interrupted"))
            except HttpError as exc:
                self.log(tr("log_err_http", exc.resp.status, exc))
            except Exception as exc:  # noqa: BLE001
                self.log(tr("log_err", exc))
            finally:
                self.busy = False
                self.after(0, self._refresh_status)

        threading.Thread(target=worker, daemon=True).start()

    # ---------- azioni ----------

    def choose_credentials(self):
        path = filedialog.askopenfilename(
            title=tr("dlg_choose_title"),
            filetypes=[(tr("ft_json"), "*.json"), (tr("ft_all"), "*.*")])
        if not path:
            return
        try:
            if Path(path).resolve() != CREDENTIALS_FILE.resolve():
                shutil.copyfile(path, CREDENTIALS_FILE)
            self.log(tr("log_cred_copied", CREDENTIALS_FILE))
        except OSError as exc:
            self.log(tr("log_err_copy", exc))
        self._refresh_status()

    def open_credentials_dir(self):
        try:
            DATA_DIR.mkdir(parents=True, exist_ok=True)
            os.startfile(str(DATA_DIR))
        except (OSError, AttributeError) as exc:
            self.log(tr("log_err_open", exc))

    def connect(self):
        def task():
            self.log(tr("log_connecting"))
            self.service = get_service(self.log)
            profile = self.service.users().getProfile(userId="me").execute()
            self.log(tr("log_connected", profile.get("emailAddress"), profile.get("messagesTotal")))
        self.run_async(task)

    def disconnect(self):
        self.service = None
        self.found_ids = []
        if TOKEN_FILE.is_file():
            TOKEN_FILE.unlink()
            self.log(tr("log_token_deleted"))
        else:
            self.log(tr("log_no_token"))
        self._refresh_status()

    def search(self):
        try:
            mode, d1, d2 = self._get_params()
            query = build_query(mode, d1, d2)
        except ValueError as exc:
            messagebox.showerror(tr("err_title"), str(exc))
            return
        period = describe_period(mode, d1, d2)

        def task():
            self.log(tr("log_period", period))
            self.log(tr("log_query", query))
            self.found_ids = []
            ids = list_all_messages(self.service, query, self.log, self.cancel)
            self.found_ids, self.found_query = ids, query
            self.log(tr("log_total", len(ids)))
            if not ids:
                self.log(tr("log_none"))
        self.run_async(task)

    def apply_label(self):
        try:
            mode, d1, d2 = self._get_params()
            query = build_query(mode, d1, d2)
            name = self._get_label()
        except ValueError as exc:
            messagebox.showerror(tr("err_title"), str(exc))
            return
        period = describe_period(mode, d1, d2)

        if query != self.found_query:
            messagebox.showwarning(tr("dlg_repeat_title"), tr("dlg_repeat_msg"))
            return

        count = len(self.found_ids)
        if not messagebox.askyesno(tr("dlg_confirm_title"),
                                   tr("dlg_confirm_msg", name, count, period, query)):
            self.log(tr("log_cancelled"))
            return

        def task():
            label_id = get_or_create_label(self.service, name, self.log)
            self.log(tr("log_applying"))
            processed = batch_label(self.service, self.found_ids, label_id, self.log, self.cancel)
            self.log("=" * 40)
            self.log(tr("log_done", processed, name))
            self.log(tr("log_done2"))
        self.run_async(task)


if __name__ == "__main__":
    App().mainloop()
'@

$UNINSTALL_CS = @'
// Gmail Labeler - Uninstaller / Disinstallatore
// Rimuove solo cio' che installa il setup: GmailLabeler.exe, _internal,
// (facoltativo) le credenziali e i dati utente dell'utente corrente, i collegamenti sul Desktop
// e la voce in Impostazioni > App. Non cancella altro. Non salva nulla da nessuna parte.
// Sorgente compatibile con C# 5 (csc.exe di .NET Framework 4.x).

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("Gmail Labeler Uninstaller")]
[assembly: AssemblyDescription("Gmail Labeler Uninstaller")]
[assembly: AssemblyCompany("Gmail Labeler")]
[assembly: AssemblyProduct("Gmail Labeler")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

internal static class Program
{
    private const string AppExe = "GmailLabeler.exe";
    private const string InternalDir = "_internal";
    private const string RegPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\GmailLabeler";
    private const string ShortcutName = "Gmail Labeler.lnk";

    // ------------------------------------------------------------
    // Testi
    // ------------------------------------------------------------
    private static string T(string lang, string key)
    {
        bool it = (lang == "it");
        switch (key)
        {
            case "title":
                return it ? "Disinstalla Gmail Labeler" : "Uninstall Gmail Labeler";
            case "ask":
                return it ? "Vuoi rimuovere Gmail Labeler da questo computer?\n\nCartella:\n{0}"
                          : "Do you want to remove Gmail Labeler from this computer?\n\nFolder:\n{0}";
            case "creds":
                return it ? "Elimina anche le credenziali salvate (credentials.json e token.json) dell'utente corrente"
                          : "Also delete the saved credentials (credentials.json and token.json) of the current user";
            case "userdata":
                return it ? "Elimina anche i dati utente (tema e lingua dell'app)"
                          : "Also delete the user data (app theme and language)";
            case "lang":
                return it ? "Lingua:" : "Language:";
            case "note":
                return it ? "Per revocare l'accesso all'account Google vai su myaccount.google.com/permissions."
                          : "To revoke the app's access to your Google account, visit myaccount.google.com/permissions.";
            case "ok":
                return it ? "Disinstalla" : "Uninstall";
            case "cancel":
                return it ? "Annulla" : "Cancel";
            case "invalid":
                return it ? "Il disinstallatore non si trova nella cartella di installazione di Gmail Labeler. Operazione annullata."
                          : "The uninstaller is not inside the Gmail Labeler installation folder. Operation cancelled.";
            case "startErr":
                return it ? "Impossibile avviare la disinstallazione:\n{0}"
                          : "Unable to start the uninstallation:\n{0}";
            case "running":
                return it ? "Gmail Labeler \u00e8 in esecuzione.\n\nChiudilo, poi premi Riprova per continuare la disinstallazione."
                          : "Gmail Labeler is running.\n\nClose it, then click Retry to continue the uninstallation.";
            case "done":
                return it ? "Gmail Labeler \u00e8 stato rimosso." : "Gmail Labeler has been removed.";
            case "kept":
                return it ? "Le credenziali sono state conservate in:\n{0}"
                          : "The credentials were kept in:\n{0}";
            case "keptUser":
                return it ? "I dati utente sono stati conservati in:\n{0}"
                          : "The user data was kept in:\n{0}";
            case "failed":
                return it ? "Non \u00e8 stato possibile rimuovere:\n{0}\n\nChiudi i programmi che usano questi file e riprova."
                          : "Unable to remove:\n{0}\n\nClose any program using these files and try again.";
        }
        return key;
    }

    // Lingua di sistema, rilevata a ogni avvio: non viene letta ne' salvata da nessuna parte
    private static string DetectLanguage()
    {
        return CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "it" ? "it" : "en";
    }

    // ------------------------------------------------------------
    // L'app e' aperta?
    // ------------------------------------------------------------
    private static bool IsAppRunning(string installDir)
    {
        string appExe = Path.GetFullPath(Path.Combine(installDir, AppExe));
        Process[] procs = Process.GetProcessesByName("GmailLabeler");
        try
        {
            foreach (Process proc in procs)
            {
                try
                {
                    string file = proc.MainModule.FileName;
                    if (string.Equals(Path.GetFullPath(file), appExe, StringComparison.OrdinalIgnoreCase)) return true;
                }
                catch (InvalidOperationException) { }      // il processo e' appena terminato
                catch (Exception) { return true; }         // percorso non leggibile: per sicurezza e' attivo
            }
        }
        finally
        {
            foreach (Process proc in procs) proc.Dispose();
        }
        return false;
    }

    // ------------------------------------------------------------
    // Rimozione (solo elementi noti)
    // ------------------------------------------------------------
    private static bool RemoveFile(string path)
    {
        for (int i = 0; i < 30; i++)
        {
            if (!File.Exists(path)) return true;
            try
            {
                File.SetAttributes(path, FileAttributes.Normal);
                File.Delete(path);
                return true;
            }
            catch (Exception) { Thread.Sleep(500); }
        }
        return !File.Exists(path);
    }

    private static bool RemoveDir(string path)
    {
        for (int i = 0; i < 30; i++)
        {
            if (!Directory.Exists(path)) return true;
            try { Directory.Delete(path, true); }
            catch (Exception) { Thread.Sleep(500); }
        }
        return !Directory.Exists(path);
    }

    private static void RemoveShortcuts()
    {
        Environment.SpecialFolder[] desktops = new Environment.SpecialFolder[]
        {
            Environment.SpecialFolder.CommonDesktopDirectory,
            Environment.SpecialFolder.DesktopDirectory
        };
        foreach (Environment.SpecialFolder folder in desktops)
        {
            try
            {
                string lnk = Path.Combine(Environment.GetFolderPath(folder), ShortcutName);
                if (File.Exists(lnk)) File.Delete(lnk);
            }
            catch (Exception) { }
        }
    }

    private static void RemoveRegistry()
    {
        RegistryView[] views = new RegistryView[] { RegistryView.Registry64, RegistryView.Registry32 };
        foreach (RegistryView view in views)
        {
            try
            {
                using (RegistryKey hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                {
                    hklm.DeleteSubKey(RegPath, false);
                }
            }
            catch (Exception) { }
        }
    }

    // ------------------------------------------------------------
    // Fase 2: eseguita da una copia temporanea dell'exe
    // ------------------------------------------------------------
    private static void Stage2(string installDir, string lang, bool removeCreds, bool removeUserData)
    {
        try { Directory.SetCurrentDirectory(Path.GetTempPath()); } catch (Exception) { }   // non bloccare _internal
        Thread.Sleep(1500);   // lascia terminare il processo originale

        string appExe = Path.Combine(installDir, AppExe);
        string internalDir = Path.Combine(installDir, InternalDir);
        if (!File.Exists(appExe) && !Directory.Exists(internalDir)) return;

        List<string> failed = new List<string>();
        if (!RemoveFile(appExe)) failed.Add(appExe);
        if (!RemoveDir(internalDir)) failed.Add(internalDir);

        string appDataRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Gmail Labeler");
        string credsDir = Path.Combine(appDataRoot, "credentials");
        string userDir = Path.Combine(appDataRoot, "user_data");
        if (removeCreds && !RemoveDir(credsDir)) failed.Add(credsDir);
        if (removeUserData && !RemoveDir(userDir)) failed.Add(userDir);
        if (removeCreds || removeUserData)
        {
            try { Directory.Delete(appDataRoot, false); } catch (Exception) { }   // solo se vuota
        }

        if (failed.Count == 0)
        {
            try { Directory.Delete(installDir, false); } catch (Exception) { }     // solo se vuota
            RemoveShortcuts();
            RemoveRegistry();
        }

        string message;
        MessageBoxIcon icon;
        if (failed.Count > 0)
        {
            message = string.Format(T(lang, "failed"), string.Join("\n", failed.ToArray()));
            icon = MessageBoxIcon.Error;
        }
        else
        {
            message = T(lang, "done");
            if (!removeCreds && Directory.Exists(credsDir))
            {
                message += "\n\n" + string.Format(T(lang, "kept"), credsDir);
            }
            if (!removeUserData && Directory.Exists(userDir))
            {
                message += "\n\n" + string.Format(T(lang, "keptUser"), userDir);
            }
            icon = MessageBoxIcon.Information;
        }

        using (Form owner = new Form())
        {
            owner.TopMost = true;
            owner.ShowInTaskbar = false;
            owner.FormBorderStyle = FormBorderStyle.None;
            owner.StartPosition = FormStartPosition.Manual;
            owner.Location = new Point(-2000, -2000);
            owner.Size = new Size(1, 1);
            owner.Show();
            MessageBox.Show(owner, message, T(lang, "title"), MessageBoxButtons.OK, icon);
        }

        // la piccola copia temporanea (GmailLabelerUninstall_*.exe in %TEMP%) resta in uso fino alla chiusura
        // del processo: Windows la elimina con la normale pulizia dei file temporanei.
    }

    // ------------------------------------------------------------
    // Fase 1: conferma
    // ------------------------------------------------------------
    private static string Quote(string value)
    {
        return "\"" + value + "\"";
    }

    private static void Stage1()
    {
        string exe = Application.ExecutablePath;
        string internalDir = Path.GetDirectoryName(exe);
        string installDir = Path.GetDirectoryName(internalDir);
        string lang = DetectLanguage();     // lingua di sistema a ogni avvio; non viene salvata

        bool valid = string.Equals(Path.GetFileName(internalDir), InternalDir, StringComparison.OrdinalIgnoreCase)
                     && File.Exists(Path.Combine(installDir, AppExe));
        if (!valid)
        {
            MessageBox.Show(T(lang, "invalid"), T(lang, "title"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        using (Form form = new Form())
        {
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.MaximizeBox = false;
            form.MinimizeBox = false;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.ClientSize = new Size(480, 330);
            form.Font = new Font("Segoe UI", 9f);
            try { form.Icon = Icon.ExtractAssociatedIcon(exe); } catch (Exception) { }

            Label langLabel = new Label();
            langLabel.TextAlign = ContentAlignment.MiddleRight;
            langLabel.Location = new Point(250, 12);
            langLabel.Size = new Size(96, 24);

            ComboBox langBox = new ComboBox();
            langBox.DropDownStyle = ComboBoxStyle.DropDownList;
            langBox.Items.Add("Italiano");
            langBox.Items.Add("English");
            langBox.SelectedIndex = (lang == "it") ? 0 : 1;
            langBox.Location = new Point(352, 12);
            langBox.Size = new Size(110, 24);

            Label ask = new Label();
            ask.Location = new Point(18, 48);
            ask.Size = new Size(444, 84);

            CheckBox chkCreds = new CheckBox();
            chkCreds.Location = new Point(18, 138);
            chkCreds.Size = new Size(444, 36);

            CheckBox chkUser = new CheckBox();
            chkUser.Location = new Point(18, 176);
            chkUser.Size = new Size(444, 36);

            Label note = new Label();
            note.ForeColor = Color.DimGray;
            note.Location = new Point(18, 220);
            note.Size = new Size(444, 40);

            Button ok = new Button();
            ok.Size = new Size(100, 28);
            ok.Location = new Point(262, 284);

            Button cancel = new Button();
            cancel.Size = new Size(100, 28);
            cancel.Location = new Point(368, 284);
            cancel.DialogResult = DialogResult.Cancel;
            form.CancelButton = cancel;

            Action applyLanguage = delegate
            {
                form.Text = T(lang, "title");
                langLabel.Text = T(lang, "lang");
                ask.Text = string.Format(T(lang, "ask"), installDir);
                chkCreds.Text = T(lang, "creds");
                chkUser.Text = T(lang, "userdata");
                note.Text = T(lang, "note");
                ok.Text = T(lang, "ok");
                cancel.Text = T(lang, "cancel");
            };
            applyLanguage();

            langBox.SelectedIndexChanged += delegate
            {
                lang = (langBox.SelectedIndex == 0) ? "it" : "en";
                applyLanguage();
            };

            ok.Click += delegate(object sender, EventArgs e)
            {
                while (IsAppRunning(installDir))
                {
                    DialogResult answer = MessageBox.Show(form, T(lang, "running"), T(lang, "title"),
                        MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning);
                    if (answer != DialogResult.Retry) return;
                }

                try
                {
                    string tmp = Path.Combine(Path.GetTempPath(),
                        "GmailLabelerUninstall_" + Process.GetCurrentProcess().Id.ToString() + ".exe");
                    File.Copy(exe, tmp, true);

                    ProcessStartInfo psi = new ProcessStartInfo(tmp,
                        "--stage2 " + Quote(installDir) + " " + lang + " " +
                        (chkCreds.Checked ? "1" : "0") + " " + (chkUser.Checked ? "1" : "0"));
                    psi.UseShellExecute = false;
                    psi.WorkingDirectory = Path.GetTempPath();
                    psi.CreateNoWindow = true;
                    Process.Start(psi);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format(T(lang, "startErr"), ex.Message), T(lang, "title"),
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                form.Close();
            };

            form.Controls.Add(langLabel);
            form.Controls.Add(langBox);
            form.Controls.Add(ask);
            form.Controls.Add(chkCreds);
            form.Controls.Add(chkUser);
            form.Controls.Add(note);
            form.Controls.Add(ok);
            form.Controls.Add(cancel);
            Application.Run(form);
        }
    }

    [STAThread]
    private static void Main(string[] args)
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        if (args.Length >= 4 && args[0] == "--stage2")
        {
            bool removeUserData = args.Length >= 5 && args[4] == "1";
            Stage2(args[1], args[2], args[3] == "1", removeUserData);
        }
        else
        {
            Stage1();
        }
    }
}
'@

$UNINSTALL_CS_LEGACY = @'
// Gmail Labeler - Uninstaller / Disinstallatore
// Rimuove solo cio' che installa il setup: GmailLabeler.exe, _internal,
// (facoltativo) le credenziali dell'utente corrente, i collegamenti sul Desktop
// e la voce in Impostazioni > App. Non cancella altro.
// Sorgente compatibile con C# 5 (csc.exe di .NET Framework 4.x).

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("Gmail Labeler Uninstaller")]
[assembly: AssemblyDescription("Gmail Labeler Uninstaller")]
[assembly: AssemblyCompany("Gmail Labeler")]
[assembly: AssemblyProduct("Gmail Labeler")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

internal static class Program
{
    private const string AppExe = "GmailLabeler.exe";
    private const string InternalDir = "_internal";
    private const string RegPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\GmailLabeler";
    private const string ShortcutName = "Gmail Labeler.lnk";

    // ------------------------------------------------------------
    // Testi
    // ------------------------------------------------------------
    private static string T(string lang, string key)
    {
        bool it = (lang == "it");
        switch (key)
        {
            case "title":
                return it ? "Disinstalla Gmail Labeler" : "Uninstall Gmail Labeler";
            case "ask":
                return it ? "Vuoi rimuovere Gmail Labeler da questo computer?\n\nCartella:\n{0}"
                          : "Do you want to remove Gmail Labeler from this computer?\n\nFolder:\n{0}";
            case "creds":
                return it ? "Elimina anche le credenziali salvate (credentials.json e token.json) dell'utente corrente"
                          : "Also delete the saved credentials (credentials.json and token.json) of the current user";
            case "note":
                return it ? "Per revocare l'accesso all'account Google vai su myaccount.google.com/permissions."
                          : "To revoke the app's access to your Google account, visit myaccount.google.com/permissions.";
            case "ok":
                return it ? "Disinstalla" : "Uninstall";
            case "cancel":
                return it ? "Annulla" : "Cancel";
            case "invalid":
                return it ? "Il disinstallatore non si trova nella cartella di installazione di Gmail Labeler. Operazione annullata."
                          : "The uninstaller is not inside the Gmail Labeler installation folder. Operation cancelled.";
            case "startErr":
                return it ? "Impossibile avviare la disinstallazione:\n{0}"
                          : "Unable to start the uninstallation:\n{0}";
            case "running":
                return it ? "Gmail Labeler \u00e8 in esecuzione.\n\nChiudilo, poi premi Riprova per continuare la disinstallazione."
                          : "Gmail Labeler is running.\n\nClose it, then click Retry to continue the uninstallation.";
            case "done":
                return it ? "Gmail Labeler \u00e8 stato rimosso." : "Gmail Labeler has been removed.";
            case "kept":
                return it ? "Le credenziali sono state conservate in:\n{0}"
                          : "The credentials were kept in:\n{0}";
            case "failed":
                return it ? "Non \u00e8 stato possibile rimuovere:\n{0}\n\nChiudi i programmi che usano questi file e riprova."
                          : "Unable to remove:\n{0}\n\nClose any program using these files and try again.";
        }
        return key;
    }

    private static string DetectLanguage(string internalDir)
    {
        try
        {
            string value = File.ReadAllText(Path.Combine(internalDir, "language.txt")).Trim().ToLowerInvariant();
            if (value == "it" || value == "en") return value;
        }
        catch (Exception) { }
        return CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "it" ? "it" : "en";
    }

    // ------------------------------------------------------------
    // L'app e' aperta?
    // ------------------------------------------------------------
    private static bool IsAppRunning(string installDir)
    {
        string appExe = Path.GetFullPath(Path.Combine(installDir, AppExe));
        Process[] procs = Process.GetProcessesByName("GmailLabeler");
        try
        {
            foreach (Process proc in procs)
            {
                try
                {
                    string file = proc.MainModule.FileName;
                    if (string.Equals(Path.GetFullPath(file), appExe, StringComparison.OrdinalIgnoreCase)) return true;
                }
                catch (InvalidOperationException) { }      // il processo e' appena terminato
                catch (Exception) { return true; }         // percorso non leggibile: per sicurezza e' attivo
            }
        }
        finally
        {
            foreach (Process proc in procs) proc.Dispose();
        }
        return false;
    }

    // ------------------------------------------------------------
    // Rimozione (solo elementi noti)
    // ------------------------------------------------------------
    private static bool RemoveFile(string path)
    {
        for (int i = 0; i < 30; i++)
        {
            if (!File.Exists(path)) return true;
            try
            {
                File.SetAttributes(path, FileAttributes.Normal);
                File.Delete(path);
                return true;
            }
            catch (Exception) { Thread.Sleep(500); }
        }
        return !File.Exists(path);
    }

    private static bool RemoveDir(string path)
    {
        for (int i = 0; i < 30; i++)
        {
            if (!Directory.Exists(path)) return true;
            try { Directory.Delete(path, true); }
            catch (Exception) { Thread.Sleep(500); }
        }
        return !Directory.Exists(path);
    }

    private static void RemoveShortcuts()
    {
        Environment.SpecialFolder[] desktops = new Environment.SpecialFolder[]
        {
            Environment.SpecialFolder.CommonDesktopDirectory,
            Environment.SpecialFolder.DesktopDirectory
        };
        foreach (Environment.SpecialFolder folder in desktops)
        {
            try
            {
                string lnk = Path.Combine(Environment.GetFolderPath(folder), ShortcutName);
                if (File.Exists(lnk)) File.Delete(lnk);
            }
            catch (Exception) { }
        }
    }

    private static void RemoveRegistry()
    {
        RegistryView[] views = new RegistryView[] { RegistryView.Registry64, RegistryView.Registry32 };
        foreach (RegistryView view in views)
        {
            try
            {
                using (RegistryKey hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                {
                    hklm.DeleteSubKey(RegPath, false);
                }
            }
            catch (Exception) { }
        }
    }

    // ------------------------------------------------------------
    // Fase 2: eseguita da una copia temporanea dell'exe
    // ------------------------------------------------------------
    private static void Stage2(string installDir, string lang, bool removeCreds, string selfPath)
    {
        try { Directory.SetCurrentDirectory(Path.GetTempPath()); } catch (Exception) { }   // non bloccare _internal
        Thread.Sleep(1500);   // lascia terminare il processo originale

        string appExe = Path.Combine(installDir, AppExe);
        string internalDir = Path.Combine(installDir, InternalDir);
        if (!File.Exists(appExe) && !Directory.Exists(internalDir)) return;

        List<string> failed = new List<string>();
        if (!RemoveFile(appExe)) failed.Add(appExe);
        if (!RemoveDir(internalDir)) failed.Add(internalDir);

        string appDataRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Gmail Labeler");
        string credsDir = Path.Combine(appDataRoot, "credentials");
        if (removeCreds)
        {
            if (!RemoveDir(credsDir)) failed.Add(credsDir);
            try { Directory.Delete(appDataRoot, false); } catch (Exception) { }   // solo se vuota
        }

        if (failed.Count == 0)
        {
            try { Directory.Delete(installDir, false); } catch (Exception) { }     // solo se vuota
            RemoveShortcuts();
            RemoveRegistry();
        }

        string message;
        MessageBoxIcon icon;
        if (failed.Count > 0)
        {
            message = string.Format(T(lang, "failed"), string.Join("\n", failed.ToArray()));
            icon = MessageBoxIcon.Error;
        }
        else
        {
            message = T(lang, "done");
            if (!removeCreds && Directory.Exists(credsDir))
            {
                message += "\n\n" + string.Format(T(lang, "kept"), credsDir);
            }
            icon = MessageBoxIcon.Information;
        }

        using (Form owner = new Form())
        {
            owner.TopMost = true;
            owner.ShowInTaskbar = false;
            owner.FormBorderStyle = FormBorderStyle.None;
            owner.StartPosition = FormStartPosition.Manual;
            owner.Location = new Point(-2000, -2000);
            owner.Size = new Size(1, 1);
            owner.Show();
            MessageBox.Show(owner, message, T(lang, "title"), MessageBoxButtons.OK, icon);
        }

        // la piccola copia temporanea (GmailLabelerUninstall_*.exe in %TEMP%) resta in uso fino alla chiusura
        // del processo: Windows la elimina con la normale pulizia dei file temporanei.
    }

    // ------------------------------------------------------------
    // Fase 1: conferma
    // ------------------------------------------------------------
    private static string Quote(string value)
    {
        return "\"" + value + "\"";
    }

    private static void Stage1()
    {
        string exe = Application.ExecutablePath;
        string internalDir = Path.GetDirectoryName(exe);
        string installDir = Path.GetDirectoryName(internalDir);
        string lang = DetectLanguage(internalDir);

        bool valid = string.Equals(Path.GetFileName(internalDir), InternalDir, StringComparison.OrdinalIgnoreCase)
                     && File.Exists(Path.Combine(installDir, AppExe));
        if (!valid)
        {
            MessageBox.Show(T(lang, "invalid"), T(lang, "title"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        using (Form form = new Form())
        {
            form.Text = T(lang, "title");
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.MaximizeBox = false;
            form.MinimizeBox = false;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.ClientSize = new Size(480, 270);
            form.Font = new Font("Segoe UI", 9f);
            try { form.Icon = Icon.ExtractAssociatedIcon(exe); } catch (Exception) { }

            Label ask = new Label();
            ask.Text = string.Format(T(lang, "ask"), installDir);
            ask.Location = new Point(18, 16);
            ask.Size = new Size(444, 90);

            CheckBox chk = new CheckBox();
            chk.Text = T(lang, "creds");
            chk.Location = new Point(18, 112);
            chk.Size = new Size(444, 40);

            Label note = new Label();
            note.Text = T(lang, "note");
            note.ForeColor = Color.DimGray;
            note.Location = new Point(18, 160);
            note.Size = new Size(444, 40);

            Button ok = new Button();
            ok.Text = T(lang, "ok");
            ok.Size = new Size(100, 28);
            ok.Location = new Point(262, 222);

            Button cancel = new Button();
            cancel.Text = T(lang, "cancel");
            cancel.Size = new Size(100, 28);
            cancel.Location = new Point(368, 222);
            cancel.DialogResult = DialogResult.Cancel;
            form.CancelButton = cancel;

            ok.Click += delegate(object sender, EventArgs e)
            {
                while (IsAppRunning(installDir))
                {
                    DialogResult answer = MessageBox.Show(form, T(lang, "running"), T(lang, "title"),
                        MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning);
                    if (answer != DialogResult.Retry) return;
                }

                try
                {
                    string tmp = Path.Combine(Path.GetTempPath(),
                        "GmailLabelerUninstall_" + Process.GetCurrentProcess().Id.ToString() + ".exe");
                    File.Copy(exe, tmp, true);

                    ProcessStartInfo psi = new ProcessStartInfo(tmp,
                        "--stage2 " + Quote(installDir) + " " + lang + " " + (chk.Checked ? "1" : "0"));
                    psi.UseShellExecute = false;
                    psi.WorkingDirectory = Path.GetTempPath();
                    psi.CreateNoWindow = true;
                    Process.Start(psi);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format(T(lang, "startErr"), ex.Message), T(lang, "title"),
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                form.Close();
            };

            form.Controls.Add(ask);
            form.Controls.Add(chk);
            form.Controls.Add(note);
            form.Controls.Add(ok);
            form.Controls.Add(cancel);
            Application.Run(form);
        }
    }

    [STAThread]
    private static void Main(string[] args)
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        if (args.Length >= 4 && args[0] == "--stage2")
        {
            Stage2(args[1], args[2], args[3] == "1", Application.ExecutablePath);
        }
        else
        {
            Stage1();
        }
    }
}
'@

$ICO_SMALL_B64 = @'
AAABAAQAEBAAAAEAIABoBAAARgAAABgYAAABACAAiAkAAK4EAAAgIAAAAQAgAKgQAAA2DgAAMDAAAAEAIACoJQAA3h4AACgAAAAQAAAAIAAAAAEAIAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAD/fwACAAAAAN9fHyD0Zydg7mUmaulhI2vrZCZr62Qma+thI2vpZCZr6WYoa+5pK2r3aSdh72wmIQAAAAD/AAABAAAAAOli
JFP1Zybt9mgn//ZoJ//3aCf/92gn//doJ//2aCf/+mck//ZiIf/2Xx7/8mMl//BfIu7qURRkqlVVA+5mIh7rYiX19Wcn/+pjJfvrZCX862Ml/OtjJfzqZCb8
7GQk/N9fKfyCgoX8T6G//HGMmPxlnKv/koB7//8tBifpYiVg9mgn/+ljJvzrYiT/62Ei/+thI//rYSP/6WMm//VYFP+QenT/Z5Sn/yu89P9plqj/U5+4/HuU
n/7/SABj62IkaPZoKP7rYCD862Yp/+xqLv/saS7/7Gku/+tpLv/tai3/4mQv/2qPn/9Pnbn/coeR/2iOmPyjfXH//FILbOtlKmf1Yh//7oBN/Pzq4f/97+j/
/e/p//3v6f/98Or//fDr///v5///1sb//c+///fVyf/rfU789VsZ/+lmKmvrZyxn9V0Y//Kcdvz/////////////////////////////////////////////
////////8596/PVfGf/raC1r62csZ/VeGf/ymHD8//79///7+f///////fHs//nMuf/5zLj//fHr///////++vj//v39//KadPz1Xhn/62gta+tnLGf1Xhn/
8plw/P///////////OXb//nKt//+8u7//vPv//nLt//85Nr////////////ym3T89V4Z/+toLWvrZyxn9V4Z//GZcvz//f3/+tTD//rQvv/++vj/////////
///++vj/+tG///rTwv///f3/8px1/PVeGf/raC1r62csZ/VdGP/ynHX8/OTa//zj2f/////////////////////////////////85Nr//OTa//KeePz1XRj/
62gta+tkKWj1ZCL+7XY//PnNuf/739L/+trL//rbzP/6283/+tvN//rbzP/62sv/+9/S//nNuv/td0D89mQh/uxlKGzsYyVf9mgo/+lfIPzrXh7/6l0d/+pe
Hv/qXh7/6l4e/+peHv/qXh7/6l4e/+pdHf/rXh7/6F8g/PdoKP/qYyNk7WAjHetjJPT2aCj/6mQm++pjJv3rZCb862Qm/OtkJvzrZCb862Qm/OtkJvzqYyb9
6WQm/PVoKP/rYyX2518nIAAAAADrYiZQ9mcn6vZoJ//2aCf/9mgn//ZoJ//2aCf/9mgn//ZoJ//2aCf/9mgn//ZoJ//1Zybs6WIkUwAAAAD/fwACAAAAANtg
Ix3zZiZc8GQlZutkJWbrYyVn62MlZ+tjJWfrYyVn62QlZvBkJWb0ZSZd5V0iHgAAAAD/fwACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACgAAAAYAAAAMAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/wAAAQAAAAAAAAAA/1UAA/Zm
Kh7wYSUi6GElIvBhJSLwYSUi8GElIvBhJSLwYSUi8GElIvBhJSLoYSUi6GElIvBhJSL2Zioevz8ABAAAAAAAAAAA/38AAgAAAAD/AAAB/wAAAf9/AALvZSV0
8WUm0+1kJfHrYiX162Il9etiJfXrYiX162Il9etiJfXrYiX162Il9etiJPXqZCf16mYo9epmKPXrZinx8Gkq1OtqLHhV//8DAAAAAP8AAAH/fwAC/38/BOpj
JbH/bSn/82Ym/+5kJf/uZCX/7mQl/+5kJf/uZCX/7mQl/+5kJf/uZCX/7mQl/+xmKP/zWRf/9FQS//JWFf/yVhX/91gW//9eGP/2VBG51lEcLP9/AAIAAAAA
7GQkcP9tKf/qYiT56mIl/epjJf7qYyX+6mMl/upjJf7qYyX+6mMl/upjJf7qYyX+6mYo/vRTEP6PeXT+UaC7/mWUqf5lk6f+ZpOm/VuYrviVkpT/3UwZlgAA
AAD/fz8E6mIk1vRnJv3qYiX962Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/qZSf/7lgZ/5N5cv9+hoz/Z5Wn/xPK//9rk6T/foeM/0Wrz/07tuf9
5VEY1P+qVQbrayca6mIl8O5kJv/qYyX+62Qn/+tlKP/rZSj/62Yp/+tmKf/rZin/62Yp/+tmKf/raCv/61IX/1afvf90kJr/e4qT/x/E+/9xkJ//e4qS/2yS
n/5VocH/4lcf9PZzKR/eYiAf6mMk8u5kJv/qZCb+6l8f/+pYFv/qWBb/6lgV/+pYFf/qWBX/6lgV/+pYFf/pWBb/61gU/95IDv9GncD/EsP//yi07P9UmbX/
XZau/1Khvv5QosX/4lYe9fdtJCPlZiIe6mMk8u5lJ//qXh/+7nlE//OkgP/0pYL/9KaD//Smg//0poP/9KaD//Smg//0poL/86eE//ejfv/igV7/wn9m/8SC
a/+5iHf/tntm/7NoS/7FYT3/41gg9fdwLSLlZiIe6mQl8u5kJP/rZSj+/Onh////////////////////////////////////////////////////////////
//36///9+f///fn//+XX//RiIf7zYB3/62Mm9fBhJSLlZiIe6mQl8u5hIf/scDj+/vj1///9/P///Pv///z8///9/P///Pv///z7//////////////z7///8
+//+/v7//f////3////9/////fv6/+tzPf7sYSH/62Mm9fBhJSLlZiIe6mQl8u5hIf/sbzb+/vXx/////////v7////////+/v////////////vf0//73tL/
/////////////v7////////+/v///////vf0/+xxOf7tYSD/62Mm9fBhJSLlZiIe6mQl8u5hIf/sbzb+/vby/////////v7///7+/////////v7/+c67//jB
qf/4war/+c25///9/f////////79///+/v///////vj1/+xxOf7tYSD/62Mm9fBhJSLlZiIe6mQl8u5hIf/sbzb+/vby/////////fz///////708P/4wKj/
+cq2///9/f///v7/+cy4//i/p//+8+/////////9/f///////vj1/+xxOf7tYSD/62Mm9fBhJSLlZiIe6mQl8u5hIf/sbzb+/vXx/////////////OXb//e5
n//73M7///////////////////////vd0P/3uZ7//OTa/////////////vf0/+xxOf7tYSD/62Mm9fBhJSLlZiIe6mQl8u5hIf/sbjX+/vj1///////608P/
97uh//3s5f////////7+//////////////7+///////97ef/97yi//rSwf///////vr5/+xwOP7tYSD/62Mm9fBhJSLlZiIe6mQl8u5gIP/scTn+/e3m//jE
rv/4w6z//vf0/////////Pr///38///9/P///fz///38///8+v///////vj1//jErv/4w6z//e7o/+10Pf7tYB//62Mm9fBhJSLlZiIe6mQl8u5iIv/rbjT+
+97R//zh1v/////////////////////////////////////////////////////////////////84db/+9/T/+xwN/7tYSH/62Mm9fBhJSLmYiAf6mMk8u5m
J//pXR3+8ZZs//nPu//4yrX/+Mu2//jLt//4y7f/+Mu3//jLt//4y7f/+Mu3//jLt//4y7f/+Mu2//jKtf/5z7z/8phv/+ldHf7uZSf/62Il9fBeJCPrYica
6mIl8O5kJv/qZCb+6lsa/+pZF//qWhj/6loY/+paGP/qWhj/6loY/+paGP/qWhj/6loY/+paGP/qWhj/6loY/+paGP/qWRf/6lsa/+pkJv7uZCX/62Il8+Zi
IB//Pz8E6mMl1fVnJv3qYiX962Un/+tlKP/rZSj/62Uo/+tlKP/rZSj/62Uo/+tlKP/rZSj/62Uo/+tlKP/rZSj/62Uo/+tlKP/rZSj/62Un/+piJf30Zib9
62Ik2dRVKgYAAAAA7GIlbf9tKf/qYiT56mIl/epjJf7qYyX+6mMl/upjJf7qYyX+6mMl/upjJf7qYyX+6mMl/upjJf7qYyX+6mMl/upjJf7qYyX+6mIl/epi
JPn/bSn/62MlcwAAAAD/fwAC/1UAA+tjJaz/bSn/9Gcm/+5kJv/uZCb/7mQm/+5kJv/uZCb/7mQm/+5kJv/uZCb/7mQm/+5kJv/uZCb/7mQm/+5kJv/uZCb/
82cm//9tKf/qYyWx/38/BP9/AAL/AAAB/wAAAQAAAADuZSRv8mYmzu1jJe7qYyTy6mMk8upjJPLqYyTy6mMk8upjJPLqYyTy6mMk8upjJPLqYyTy6mMk8upj
JPLtYyXu8WUm0O9kJnJ/fwACAAAAAf8AAAEAAAAAAAAAAAAAAAAAAAAAAAAAAPVrJxr2aSMd5WYiHuVmIh7lZiIe5WYiHuVmIh7lZiIe5WYiHuVmIh7lZiIe
5WYiHvZpIx31ayca/wAAAQAAAAAAAAAA/wAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAoAAAAIAAAAEAAAAABACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
/38AAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAP9/AAIAAAAAAAAAAAAAAAAAAAAAAAAAAP9/AAIAAAAAqlUqBvNlJVjwZSaf7WMlvexjJcDpYyTB62MkwetjJMHrYyTB62MkwetjJMHrYyTB
62MkwetjJMHrYyTB62MkweljJMHpYSTB6WEkwetiJcDtYyW98GQmoPBkJFu/Px8IAAAAAP8/PwQAAAAAAAAAAAAAAAD/fwACAAAAAOJgITXyZiXR/Goo//tq
KP/4aCf/+Ggn//hoJ//4aCf/+Ggn//hoJ//4aCf/+Ggn//hoJ//4aCf/+Ggn//hoJ//4aCf/+Gop//ZsK//2bCv/9mwr//ZsK//5bSz/+24s//FrK9TkdDo5
AAAAAAEAAAAAAAAAqlUAAwAAAADqYiQx8mYl8vdoJ//pYiT762Mk++pjJfzqYyX86mMl/OpjJfzqYyX86mMl/OpjJfzqYyX86mMl/OpjJfzqYyX862Ik/Opl
J/zpWx388U4O/O5QEfzvUBH88U4O/PFODfzwTg37/VUR//FREPXbTBh70FwuC/8AAAH/fwAC/38qBuliJdX3aCf/6mIk+utjJf/rYyX/62Ml/+tjJf/rYyX/
62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/qZin/7lgZ/8FbOP9sj53/eIiS/3eJk/9qkaP/aJKl/2eTpv9sjZ38foiQ/9JYLv/0PwoYAAAAAAAA
AADpYiRT/Goo/+piJPvrYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Un/+1bHP/QUyf/T6PF/za04P8WyP//
Gsb//2KZr/9sk6P/cJCd/0Ss1P8Ryv/7z2Q9+/ZTDlkAAAAAAAAAAOtjJZz8aij+6mIk++tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tj
Jf/rYyX/62Ml/+tkJv/pXyL/11Ag/zuw2/+QfHf/+jgA/0Ks1f8iwPn/bJKj/3OOmv94ipL/R6rQ/x3C//vVZDr/8VoapAAAAAAAAAAA6mEluflpJ//qYyX8
62Ml/+tjJf/rYyX/62Mm/+tjJv/rYyb/62Mm/+tjJv/rYyb/62Mm/+tjJv/rYyb/62Qn/+lgI//WUSH/RqrP/0ynx/+OfHj/Lbnv/yi99f9zjZr/foaM/3qJ
kP+DhIb/OrDe/MpmQv/wWxrAAAABAAAAAADqYiW6+Wkn/+pjJfzrYyX/62Qm/+tkJv/rYiT/62Ik/+tiJP/rYiT/62Ik/+tiJP/rYiT/62Ik/+tiJP/rYiT/
62Qm/+xcHf/XTh7/Pq/a/wTX//8Wy///Gsj//1ihv/9im7P/YJuy/16bs/8kvvr8ymZD//BbGsEAAAEAAAAAAOpiJbr5aSf/6mMl/OtkJv/rXx//6l4e/+tm
Kf/rZSj/62Yp/+tmKf/rZin/62Yp/+tmKf/rZin/62Yp/+tmKf/rZSj/6mgs/+1cH//JUij/fHh3/3x6e/97enz/b4KL/22Djf9uipf/bo+e/3qFjfzZWi3/
610fwQAAAAAAAAAA6mIluvlpJ//qZCb862Ag/+1zO//508L//Orh//zq4f/86+L//Ovi//zr4v/86+L//Ovi//zr4v/86+L//Ovi//zr4v/86eH/++zl//fL
u//7pYf/+6qN//upjP/9p4j//aaH//VlK//yTAj/8FES/PJeIf/pYSTBAAAAAAAAAADqYiW6+Wkn/+tlKPzqWBb/98Kq////////////////////////////
////////////////////////////////////////////////////////////////////////////////9sq1/+hcG//qaSz8+Woo/+tjJMEAAAAAAAAAAOpi
Jbr5aSf/6mUn/OpaGf/50Lz///38///8+////v3///79///+/f///v3///79///+/v///f3//////////////fz///79///+/f///fz//vz7//78+//+/Pv/
/vr4//78+//508L/6lsa/+tlJ/z4aCf/62MkwQAAAAAAAAAA6mIluvlpJ//qZSf86loZ//nNuf////////38//////////////////////////////7+////
///++/n//vr4/////////v7//////////////////////////////fz///////nRv//qWxr/62Un/PhoJ//rYyTBAAAAAAAAAADqYiW6+Wkn/+plJ/zqWhn/
+c66/////////fz////////////////////////+/v///////e/o//e7of/3up///e3m/////////v7////////////////////////9/P//////+dLA/+pb
Gv/rZSf8+Ggn/+tjJMEAAAAAAAAAAOpiJbr5aSf/6mUn/OpaGf/5zrr////////9/P/////////////+/v////////////vg1P/2sZP/+c26//rOvP/2sJL/
+97R//////////////7+//////////////38///////50sD/6lsa/+tlJ/z4aCf/62MkwQAAAAAAAAAA6mIluvlpJ//qZSf86loZ//nOuv////////38////
/////v3////////9/P/6z7z/9rCS//vbzf////////////vcz//2sJL/+c26///8+/////////79/////////fz///////nSwP/qWxr/62Un/PhoJ//rYyTB
AAAAAAAAAADqYiW6+Wkn/+plJ/zqWhn/+c66/////////fz///7+///////+8+//+L+m//e3m//96+P////////+/v///v7///////3s5f/3uJ3/+L2k//7y
7f////////7+///9/P//////+dLA/+pbGv/rZSf8+Ggn/+tjJMEAAAAAAAAAAOpiJbr5aSf/6mUn/OpaGf/5zrr////////9/P///////OXb//azl//4xK7/
/vf0/////////v3//////////////v3///////749v/5xrD/9rOW//zk2f////////38/////v/50sD/6lsa/+tlJ/z4aCf/62MkwQAAAAAAAAAA6mIluvlp
J//qZSf86loZ//nNuf////////79//rUxP/2r5H/+tXF//////////////7+/////////////////////////v7////////////618f/9q+R//rTwv///f3/
//////nSv//qWxr/62Un/PhoJ//rYyTBAAAAAAAAAADqYiW6+Wkn/+plJ/zqWhj/+c67//759//4wqr/9rSX//zm3P/////////////+/v//////////////
/////////////////////v7////+///////8597/97WY//jAqP/++Pb/+dPB/+pbGv/rZSf8+Ggn/+tjJMEAAAAAAAAAAOpiJbr5aSf/6mUn/OpaGf/50L3/
+ciz//e4nv/98ez////////7+v///Pv///z8///8+////Pv///z7///8+////Pv///z7///8+////Pv///v6///////+8u3/97qg//nFr//51ML/6lwb/+tk
J/z4aCf/62MkwQAAAAAAAAAA6mIluvlpJ//rZSj86lkX//Ssi///+vj/////////////////////////////////////////////////////////////////
//////////////////////////////////r4//Wxkf/qWRf/62Uo/PhoJ//rYyTBAAAAAAAAAADqYiW6+Wkn/+pjJfzrYyT/62Ej//Okf//2vKH/9rme//a7
oP/2u6D/9rug//a7oP/2u6D/9rug//a7oP/2u6D/9rug//a7oP/2u6D/9rug//a7oP/2u6D/9rqe//a7oP/zpoL/62Ik/+tiJP/qYyX8+Ggn/+tjJMEAAAAA
AAAAAOtiJbj5aSf/6mMl/OtjJf/rYyX/6lkW/+pYFf/qWBb/6lgV/+pYFf/qWBX/6lgV/+pYFf/qWBX/6lgV/+pYFf/qWBX/6lgV/+pYFf/qWBX/6lgV/+pY
Ff/qWBb/6lgV/+pYFv/rYyX/62Ml/+pjJfz4aCf/6mIlvwAAAAAAAAAA62Ilm/xqKP7qYiT762Ml/+tjJf/rZSj/62Yp/+tmKP/rZin/62Yp/+tmKf/rZin/
62Yp/+tmKf/rZin/62Yp/+tmKf/rZin/62Yp/+tmKf/rZin/62Yp/+tmKP/rZin/62Uo/+tjJf/rYyX/6mMk+/tqKP7sYyWiAAAAAAAAAADrYiZQ/Goo/+pi
JPvrYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/
62Ml/+tjJf/pYiT7/Wso/+1jJlcAAAAA/38AAv9mMwXqYiTR+Gkn/+liJPrrYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tj
Jf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/6mIk+vdoJ//qYiTX318fCH9/AAL/fwACAAAAAOhgJy3yZSbv+Gkn/+piJPvqYiT7
6mMl/OpjJfzqYyX86mMl/OpjJfzqYyX86mMl/OpjJfzqYyX86mMl/OpjJfzqYyX86mMl/OpjJfzqYyX86mMl/OpjJfzqYyX86mIk++piJPv3aCf/8mYl8upg
IzIAAAAAqlUAAwAAAAD/fwACAAAAAONhJS/yZSXL+2oo//xqKP/5aSf/+Wkn//lpJ//5aSf/+Wkn//lpJ//5aSf/+Wkn//lpJ//5aSf/+Wkn//lpJ//5aSf/
+Wkn//lpJ//5aSf/+Wkn//lpJ//8aij//Goo//JmJs/hYiI0AAAAAP9/AAIAAAAAAAAAAAAAAAD/fwACAAAAAFVVAAPyZCVR8WUml+5jJbbrYyW56mIluupi
JbrqYiW66mIluupiJbrqYiW66mIluupiJbrqYiW66mIluupiJbrqYiW66mIluupiJbrrYyW57mQltvBlJpnvZCdUvz8ABAAAAAD/fwACAAAAAAAAAAAAAAAA
AAAAAAAAAAD/fwACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAA/38AAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKAAAADAAAABgAAAAAQAgAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/VQADvz8/BAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP9V
VQP/VQADAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/wAAAf9/AAIAAAAAAAAAAP9VKgb3Zycg7mclL+lkJTDpZCUw
6WQlMOlkJTDpZCUw6WQlMOlkJTDpZCUw6WQlMOlkJTDpZCUw6WQlMOlkJTDpZCUw6WQlMOlkJTDpZCUw6WQlMOlkJTDpZCUw6WQlMOlkJTDpZCUw6WQlMO5n
JS/3ZCYh/20kBwAAAAAAAAAA/38AAv8AAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/fwACAAAAAAAAAADuYSY88mYlne5kJt3sYyX6
8GUm//BlJv/wZSb/8GUm//BlJv/wZSb/8GUm//BlJv/wZSb/8GUm//BlJv/wZSb/8GUm//BlJv/wZSb/8GUm//BlJv/wZSb/8GUm//BlJv/wZSb/8GUm//Bl
Jv/wZSb/8GUm//BlJv/sYyX77mQl3/JlJqHvYydAAAAAAAAAAAD/fwACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP9/AAIAAAAA1l0oE/BlJaPzZib/
+2on//JmJv/sZCX/6mMl/+pjJf7qYyX+6mMl/upjJf7qYyX+6mMl/upjJf7qYyX+6mMl/upjJf7qYyX+6mMl/upjJf7qYyX+6mMl/upjJf7qYyX+6mMl/upj
Jf7qYyX+6mMl/upjJf7qYyX+6mMl/upjJf/sYyX/8WUm//ppJ//1Zyb/8GQlqtRVHxgAAAAA/z8ABAAAAAAAAAAAAAAAAAAAAAAAAAAA/38AAgAAAADoYSUi
8WYl1/1rKP/qYiX96mIk++piJf3qYyX+62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tj
Jf/rYyX/6mMl/+tkJf/qZSf/6WYo/+llKP/pZSj/6WUo/+llKP/pZSj+6WUo/ellKPvoZSj9+20r//BoKt7mbTAqAAAAAP8AAAEAAAAAAAAAAAAAAAD//wAB
/38AAu9fHxDuYyXZ+Wkn/+liJPrrYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tj
Jf/rYyX/62Ml/+tjJf/rYyX/7GQl/+RfJv/jUBf/6U0S/+hOEv/oThL/6E4S/+hOEv/oThL/6E4S/+hOEv/oThL/5k0S+/RSFP/lThXp1k8dq9hMGRQAAAAA
/wAAAQAAAAD/VQADAAAAAOhiJKH+ayj/6WEk+utjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tj
Jf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/sZCX/5WEn/+FJEv+7XkD/n3Bf/6FuXv+hb1//oW9f/6NtXP+kbVz/pG1c/6RtXP+kbVz/o25e/5pxYfy7ZUj/
3k0Z/9BRIhYAAAAAfwAAAv9VAAMAAAAA62QjM/JmJv/rYyX862Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tj
Jf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+xkJf/lYyj/5EkQ/7NkSv8bxf//EMz//xzE//8bxf//GsX//w3O//8Lz///C8///wvP//8Kz///
Gcb//w/N//9dmrT87EoP+tVeLT4AAAAA/1UAA/9/PwQAAAAA6mIkmPxqKP/qYiT762Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tj
Jf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/7GQl/+djJ//lShH/vlw9/yq77/9dm7H/WJ67/yDC/f8nvvj/KL32/4iAf/+bdWn/
lndu/5Z3b/+dc2b/S6bL/xTJ/P9llKn79U8S/+NhKKQAAAAAvz8/BAAAAAD/VQAD62Ml2fNmJv/qYiX962Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tj
Jf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/6WIm/99QGv/GWDL/Lbrn/2aWqf/tQQT/3koW/02myP8bxf//
KLz2/5d3bf+talP/qG1Z/6dtWv+waE//UKPF/xHL//9llqr960oQ/+ZiJ+L/bSQHAAAAAAAAAADrYh0a6mMl9+1kJf/qYyX+62Ml/+tjJf/rYyX/62Ml/+tj
Jf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/sZCb/31ok/+FEDv9mlqr/BdP//3WMl//qQwf/
30kU/06lx/8cxP7/Jr75/w/M//8Lz///DM7//wzO//8Lz///GMb//xbI//9klqv+5UgP/+djJ/zwaSUiAAAAAAAAAADrYicn62Ml/+xjJf/rYyX+62Ml/+tj
Jf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/6WMm/99RG//IVi//
PrDZ/xnG//9plKf/Vp+9/x7D//8kv/v/J734/6VuW/++Xz3/uGJF/7lhQ/+4YkX/v107/zyw2/9dm7T/5kYN/uxlKP/vZCUwAAAAAAAAAADrYicn62Ml/+xj
Jf/rYyX+62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/
7GQl/+hkKP/kTBT/xFk1/zK36P8Syv//HsP//ya++P8lv/r/J733/3eLlP+HgYL/g4SG/4SDhf+DhIb/in9+/zG35v9gmbD/5UcO/uxlKP/vZCUwAAAAAAAA
AADrYicn62Ml/+xjJf/rYyX+62Ml/+tjJf/rYyX/62Mm/+tlKP/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/
62Ml/+tjJf/rYyX/62Ml/+xkJf/mYyj/40oR/7lhRP8uvO//Hcb//x7G//8dxv//Hcb//xDP//8N0P//DtD//w7O//8Ozf//EMz//w3P//9gmbD/5UcP/uxl
KP/vZCUwAAAAAAAAAADrYicn62Ml/+xjJf/rYyX+62Ml/+tjJf/rZCb/62Ik/+pZF//rYSL/62Ik/+tiJP/rYiT/62Ik/+tiJP/rYiT/62Ik/+tiJP/rYiT/
62Ik/+tiJP/rYiT/62Ik/+tiJP/rYiT/62Ik/+tiJP/sYyT/5WEm/+BFDf+8USz/o2JK/6hfR/+nYEj/p2BI/6peRf+qXkT/ql5F/6tnUf+raVP/rGdR/6Zr
Vv+4X0H/2kwZ/u1jJv/vZCUwAAAAAAAAAADrYicn62Ml/+xjJf/rYyX+62Ml/+tjJv/rYiP/62Ml//WzlP/74dT//OTZ//zl2v/85dv//OXb//zl2//85dv/
/OXb//zl2//85dv//OXb//zl2//85dv//OXb//zl2//85dv//OXb//zl2//75Nn//ebb//PMvf/qhGD/8IVg/++FYP/vhV//74Vf/++FX//vhV//74Nc/+lU
G//oTRD/6FAV/+lPFP/kURj/3lYh/u9kJf/vZCUwAAAAAAAAAADrYicn62Ml/+xjJf/rYyX+62Ml/+tlKP/qWhn/972j////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////bHsP/pXh7/6mcq/+pmKP/rZSf/62Ql/vBlJv/pZCUwAAAAAAAAAADrYicn62Ml/+xjJf/rYyX+62Qm/+thIv/sbDH//vby///9/P///Pz/
//79///+/f///v3///79///+/f///v3///79///+/f///v3///79///+/f///v3///79///+/f///v3///79///+/f///v3///79///9/f/+/Pz//v38//79
/P/+/fz//v38//79/P/+/Pv//v38//76+P/scDj/6mAg/+tkJv/rYyX/6mMl/vBlJv/pZCUwAAAAAAAAAADrYicn62Ml/+xjJf/rYyX+62Qm/+tgIP/tcjn/
/vf0/////////v7///////////////////////////////////////////////////////////////////7+///+/v//////////////////////////////
//////////////////////////////////////7////////7+v/td0D/6l8f/+tkJv/rYyX/6mMl/vBlJv/pZCUwAAAAAAAAAADrYicn62Ml/+xjJf/rYyX+
62Qm/+tgIP/tcTj//vf0/////////v7//////////////////////////////////////////////////////////////v7//////////////v7/////////
//////////////////////////////////////////////////////7////////7+v/tdj//6l8f/+tkJv/rYyX/6mMl/vBlJv/pZCUwAAAAAAAAAADrYicn
62Ml/+xjJf/rYyX+62Qm/+tgIP/tcTj//vf0/////////v7////////////////////////////////////////////////////////+/v///////vTv//7z
7v////////7+//////////////////////////////////////////////////////////7////////7+v/tdj//6l8f/+tkJv/rYyX/6mMl/vBlJv/pZCUw
AAAAAAAAAADrYicn62Ml/+xjJf/rYyX+62Qm/+tgIP/tcTj//vf0/////////v7//////////////////////////////////////////////v7/////////
///849j/97me//e4nf/84dX//////////////v7///////////////////////////////////////////////7////////7+v/tdj//6l8f/+tkJv/rYyX/
6mMl/vBlJv/pZCUwAAAAAAAAAADrYicn62Ml/+xjJf/rYyX+62Qm/+tgIP/tcTj//vf0/////////v7////////////////////////////////////////+
/v////////38//rUxP/2s5b/97WZ//e1mf/2spX/+tLB///8+/////////7+//////////////////////////////////////////7////////7+v/tdj//
6l8f/+tkJv/rYyX/6mMl/vBlJv/pZCUwAAAAAAAAAADrYicn62Ml/+xjJf/rYyX+62Qm/+tgIP/tcTj//vf0/////////v7/////////////////////////
//////////79///////+9PD/+cex//awkv/3uJ3//OTZ//zm3P/3uZ7/9rCS//jFr//+8+7////////+/v////////////////////////////////////7/
///////7+v/tdj//6l8f/+tkJv/rYyX/6mMl/vBlJv/pZCUwAAAAAAAAAADrYicn62Ml/+xjJf/rYyX+62Qm/+tgIP/tcTj//vf0/////////v7/////////
/////////////////v////7///////zo3//4vKL/9rCS//jBqv/97+n////////////98ev/+MOs//awkv/3u6H//OXc//////////////7+////////////
//////////////7////////7+v/tdj//6l8f/+tkJv/rYyX/6mMl/vBlJv/pZCUwAAAAAAAAAADrYicn62Ml/+xjJf/rYyX+62Qm/+tgIP/tcTj//vf0////
/////v7///////////////////7+////////////+9nL//e1mf/2sZT/+s67//769/////////79///+/f////////v5//rQvv/3spX/97SX//rXyP///v7/
///////+/v////////////////////7////////7+v/tdj//6l8f/+tkJv/rYyX/6mMl/vBlJv/pZCUwAAAAAAAAAADrYicn62Ml/+xjJf/rYyX+62Qm/+tg
IP/tcTj//vf0/////////v7//////////////v3///////749f/5y7f/9rGT//e2mv/73M///////////////v7//////////////v7////////////739L/
97ec//axk//5ybT//vbz/////////v3///////////////7////////7+v/tdj//6l8f/+tkJv/rYyX/6mMl/vBlJv/pZCUwAAAAAAAAAADrYicn62Ml/+xj
Jf/rYyX+62Qm/+tgIP/tcTj//vf0/////////v7////////+/v///////ezl//i/p//2sJL/+L6l//3r4/////////7+/////v//////////////////////
///////+/v///////e3m//i/p//2sJL/+L6l//3q4v////////7+/////v////7////////7+v/tdj//6l8f/+tkJv/rYyX/6mMl/vBlJv/pZCUwAAAAAAAA
AADrYicn62Ml/+xjJf/rYyX+62Qm/+tgIP/tcTj//vf0/////////f3////////////73tL/97eb//axk//5ybX//vbz/////////v3/////////////////
/////////////////////////////v3///////749f/5y7j/9rGT//e2mv/73M7//////////////v3////////7+v/tdj//6l8f/+tkJv/rYyX/6mMl/vBl
Jv/pZCUwAAAAAAAAAADrYicn62Ml/+xjJf/rYyX+62Qm/+tgIP/tcTj//vf0///+/v///////vv5//rQvf/2spX/97SY//vXyP///v7////////+/v//////
//////////////////////////////////////////////////7+////////////+9rM//e1mf/2sZT/+c26//759/////////7+///7+v/tdj//6l8f/+tk
Jv/rYyX/6mMl/vBlJv/pZCUwAAAAAAAAAADrYicn62Ml/+xjJf/rYyX+62Qm/+tgIP/tcTj//vf0///////98Or/+MOs//awkv/3u6H//Obd////////////
//7+/////////////////////////////////////////////////////////////////////v////7///////zo4P/4vKP/9rCS//jBqf/97uj///////77
+v/tdj//6l8f/+tkJv/rYyX/6mMl/vBlJv/pZCUwAAAAAAAAAADrYicn62Ml/+xjJf/rYyX+62Qm/+tgIP/tcTj//vf0//zj2f/3uZ7/9rCS//jFr//+8+7/
///////+/v////////////////////////////////////////////////////////////////////////////////////////79///////+9fH/+cey//aw
kv/3uJ3//OHV//76+P/tdj//6l8f/+tkJv/rYyX/6mMl/vBlJv/pZCUwAAAAAAAAAADrYicn62Ml/+xjJf/rYyX+62Qm/+tfIP/tcjr//vXx//i+pf/2r5H/
+tPC///8+/////////79///////////////////////////////////////////////////////////////////////////////////////////////////+
/v////////38//rVxf/2sZP/97uh//738//ud0H/6l4e/+tkJv/rYyX/6mMl/vBlJv/pZCUwAAAAAAAAAADrYicn62Ml/+xjJf/rYyX+62Ml/+tjJf/rYSL/
/Ong//7z7//73M7////////9/P///Pv///38///9/P///fz///38///9/P///fz///38///9/P///fz///38///9/P///fz///38///9/P///fz///38///9
/P///fz///38///9/P///Pv///38///////73dH//fDq//3v6f/rZCb/62Mk/+tjJf/rYyX/6mMl/vBlJv/pZCUwAAAAAAAAAADrYicn62Ml/+xjJf/rYyX+
62Ml/+tlJ//qXBz/74ZV//3z7v//////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////vby//CLXf/qXBv/62Un/+tjJf/rYyX/6mMl/vBlJv/pZCUwAAAAAAAAAADrYicn
62Ml/+xjJf/rYyX+62Ml/+tjJf/rZCf/6lwb/+1vNv/xk2n/8phw//KZcP/ymXD/8plw//KZcP/ymXD/8plw//KZcP/ymXD/8plw//KZcP/ymXD/8plw//KZ
cP/ymXD/8plw//KZcP/ymXD/8plw//KZcP/ymXD/8plw//KZcP/ymXD/8plw//KYcP/xlGr/7XI5/+pcG//rZCf/62Ml/+tjJf/rYyX/6mMl/vBlJv/pZCUw
AAAAAAAAAADrYicn62Ml/+xjJf/rYyX+62Ml/+tjJf/rYyX/62Un/+tgIf/qWhj/6lkX/+pZF//qWRf/6lkX/+pZF//qWRf/6lkX/+pZF//qWRf/6lkX/+pZ
F//qWRf/6lkX/+pZF//qWRf/6lkX/+pZF//qWRf/6lkX/+pZF//qWRf/6lkX/+pZF//qWRf/6lkX/+pZF//qWRj/618g/+tlJ//rYyX/62Ml/+tjJf/rYyX/
6mMl/vBlJv/pZCUwAAAAAAAAAADqZh4Z6mMk9u5kJf/qYyX+62Ml/+tjJf/rYyX/62Ml/+tkJv/rZSj/62Uo/+tlKP/rZSj/62Uo/+tlKP/rZSj/62Uo/+tl
KP/rZSj/62Uo/+tlKP/rZSj/62Uo/+tlKP/rZSj/62Uo/+tlKP/rZSj/62Uo/+tlKP/rZSj/62Uo/+tlKP/rZSj/62Uo/+tlKP/rZSj/62Qm/+tjJf/rYyX/
62Ml/+tjJf/rYyX+7GMl/+pjJPzvZCYhAAAAAAAAAAD/VQAD6mMk1/NnJv/qYiX962Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tj
Jf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/
62Ml/+tjJf/rYyX/62Ml/+tjJf/rYiX98mYm/+tjJeD/fyoGAAAAAP9/PwQAAAAA7GIllPxqKP/qYiT762Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tj
Jf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/
62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/qYiT7+2oo/utiJKAAAAAAv38/BKpVAAMAAAAA7mMmLvBlJv7sYyX862Ml/utjJf/rYyX/62Ml/+tj
Jf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/
62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX89Gcm/+hkJDgAAAAA/1UAAwAAAAD/VQADAAAAAOliJJn/ayj/6WEk+utj
Jf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/
62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+lhJPr+ayj/6WIlpQAAAAD/VQADAAAAAAAAAAD/AAAB/38AAv9q
KgzuYyXS+2on/+liJPrrYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/
62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/6WIk+vlpJ//uZCXa8GktEf9/AAJ/fwACAAAAAAAA
AAAAAAAA/38AAgAAAADjZCQc8WYmz/5rKP/qYyX+6mIk++piJf3qYyX+62Ml/utjJf7rYyX+62Ml/utjJf7rYyX+62Ml/utjJf7rYyX+62Ml/utjJf7rYyX+
62Ml/utjJf7rYyX+62Ml/utjJf7rYyX+62Ml/utjJf7rYyX+62Ml/utjJf7rYyX+62Ml/utjJf7qYyX+6mIl/epiJPvqYiX9/Wso//FmJtboYSUiAAAAAP9/
AAIAAAAAAAAAAAAAAAAAAAAAAAAAAP9/AAIAAAAA12InDe9kJJjwZSb9+2oo//NmJv/uZCX/7GMl/+xjJf/sYyX/7GMl/+xjJf/sYyX/7GMl/+xjJf/sYyX/
7GMl/+xjJf/sYyX/7GMl/+xjJf/sYyX/7GMl/+xjJf/sYyX/7GMl/+xjJf/sYyX/7GMl/+xjJf/sYyX/7GMl/+xjJf/tZCX/82Ym//tqKP/xZib/8GUmn9RV
HBIAAAAA/38AAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/fwACAAAAAAAAAADvZiMy8WUmku9lJdTtZCX062Ml/utjJf/rYyX/62Ml/+tjJf/rYyX/
62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/rYyX/62Ml/+tjJf/sYyX172Um1vFm
JZbwYyU2AAAAAAAAAAD/fwACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/wAAAapVAAMAAAAAAAAAAAAAAAD0aioY8WcpJfFkKCbrYicn
62InJ+tiJyfrYicn62InJ+tiJyfrYicn62InJ+tiJyfrYicn62InJ+tiJyfrYicn62InJ+tiJyfrYicn62InJ+tiJyfrYicn62InJ+tiJyfrYicn8WQoJvFk
KCb0ZigZ//8AAQAAAAAAAAAA/38AAv8AAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/fwAC/38/BAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAP9/PwSqVQADAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==
'@

$UNINSTALL_MANIFEST = @'
<?xml version="1.0" encoding="utf-8"?>
<assembly manifestVersion="1.0" xmlns="urn:schemas-microsoft-com:asm.v1">
  <trustInfo xmlns="urn:schemas-microsoft-com:asm.v2">
    <security>
      <requestedPrivileges xmlns="urn:schemas-microsoft-com:asm.v3">
        <requestedExecutionLevel level="requireAdministrator" uiAccess="false" />
      </requestedPrivileges>
    </security>
  </trustInfo>
  <compatibility xmlns="urn:schemas-microsoft-com:compatibility.v1">
    <application>
      <supportedOS Id="{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}" />
    </application>
  </compatibility>
</assembly>
'@

$ICO_B64 = @'
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
TkSuQmCC
'@

$VERSION_INFO = @'
VSVersionInfo(
  ffi=FixedFileInfo(filevers=(1, 0, 0, 0), prodvers=(1, 0, 0, 0), mask=0x3f, flags=0x0,
                    OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[
    StringFileInfo([StringTable('040904B0', [
      StringStruct('CompanyName', 'Gmail Labeler'),
      StringStruct('FileDescription', 'Gmail Labeler'),
      StringStruct('FileVersion', '1.0.0.0'),
      StringStruct('InternalName', 'GmailLabeler'),
      StringStruct('OriginalFilename', 'GmailLabeler.exe'),
      StringStruct('ProductName', 'Gmail Labeler'),
      StringStruct('ProductVersion', '1.0.0.0')])]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ])
'@

$PY_VERSION_TARGET = '3.14'                       # versione installata se manca Python
$PY_MIN = [version]'3.11'                         # versioni accettate se già presenti
$PY_MAX = [version]'3.14'
$PY_FALLBACKS = '3.14.8','3.14.7','3.14.6','3.14.5','3.14.4','3.14.3','3.14.2','3.14.1','3.14.0'

$script:cancel = $false
$script:installing = $false
$script:resultExe = $null

# ------------------------------------------------------------
# Lingua: default = lingua di sistema (it) altrimenti English
# ------------------------------------------------------------
$STR = @{
    it = @{
        title0 = "Benvenuto";            sub0 = "Installazione guidata di Gmail Labeler"
        title1 = "Destinazione";         sub1 = "Scegli la cartella di installazione"
        title2 = "Installazione";        sub2 = "Attendi: può richiedere alcuni minuti"
        title3 = "Fatto";                sub3 = "Operazione terminata"
        back = "< Indietro"; next = "Avanti >"; install = "Installa"; finish = "Fine"
        cancel = "Annulla"; close = "Chiudi"; browse = "Sfoglia..."
        welcome = "Questo programma prepara tutto il necessario e crea l'applicazione Gmail Labeler.\n\nVerranno eseguiti, in automatico e senza riavvii:\n\n   •  controllo di Python (se manca viene installato Python {0})\n   •  controllo delle librerie Google e di PyInstaller (installate se mancanti)\n   •  creazione di GmailLabeler.exe con la sua icona e del disinstallatore\n   •  copia dell'applicazione nella cartella che sceglierai\n\nServe una connessione a Internet. L'installazione in Programmi richiede i permessi di amministratore.\n\nPremi Avanti per continuare."
        destInfo = "Scegli dove salvare GmailLabeler.exe.\ncredentials.json e token.json verranno salvati per ogni utente in %APPDATA%\Gmail Labeler\credentials."
        shortcut = "Crea un collegamento sul Desktop"
        destEmpty = "Indica una cartella di destinazione."
        dlgDesc = "Scegli la cartella di destinazione"
        waiting = "In attesa..."; cancelling = "Annullamento in corso..."
        failed = "Installazione non riuscita."; done = "Completato."
        s1 = "1/6  Controllo Python..."
        s1b = "1/6  Installazione di Python..."
        s2 = "2/6  Controllo interfaccia grafica (Tkinter)..."
        s3 = "3/6  Controllo librerie..."
        s4 = "4/6  Creazione di GmailLabeler.exe (può richiedere qualche minuto)..."
        s5 = "5/6  Creazione del disinstallatore..."
        s6 = "6/6  Copia dei file..."
        endText = "Installazione completata.\n\nGmailLabeler.exe si trova in:\n{0}\n\nPer disinstallare: Impostazioni > App, oppure esegui _internal\uninstall.exe\n\nAl primo avvio usa il pulsante 'Seleziona credentials.json' (c'è anche la guida a tendina). Le credenziali vengono salvate in %APPDATA%\Gmail Labeler\credentials."
        openFolder = "Apri la cartella"; runApp = "Avvia Gmail Labeler"
        cancelled = "Operazione annullata."
        destInvalid = "Cartella di destinazione non valida."
        pyFound = "Python trovato: {0}"
        pyInstalling = "Python {0} non trovato: avvio l'installazione..."
        wingetFail = "winget non riuscito: {0}"
        wingetNone = "winget non ha installato Python: provo il download diretto da python.org..."
        dl = "Download di Python {0} ..."
        dlNA = "  non disponibile ({0})"
        dlFail = "Impossibile scaricare Python. Controlla la connessione a Internet."
        pyInstall = "Installazione di Python (per tutti gli utenti)..."
        pyInstCode = "Installazione di Python terminata con codice {0}."
        pyNotFound = "Python è stato installato ma non è stato possibile individuarlo."
        pyReady = "Python pronto: {0}"
        pyVer = "Versione Python: {0} - Tkinter OK"
        pkgInstall = "Installo {0} ..."
        pkgFail = "Installazione di {0} non riuscita (codice {1})."
        buildFail = "Creazione dell'exe non riuscita (codice {0}). Vedi il log."
        uninstFail = "Creazione del disinstallatore non riuscita (codice {0}). Vedi il log."
        uninstFallback = "Il disinstallatore aggiornato non compila: uso la versione precedente (vedi gli errori sopra)."
        created = "Creato: {0}"
        shortcutOk = "Collegamento creato sul Desktop."
        shortcutFail = "Collegamento non creato: {0}"
        regOk = "Disinstallazione registrata in Impostazioni > App."
        regFail = "Registrazione non riuscita: {0}"
        errorPrefix = "ERRORE: {0}"
        notAdmin = "Questo setup deve essere eseguito come amministratore (clic destro > Esegui come amministratore)."
    }
    en = @{
        title0 = "Welcome";              sub0 = "Gmail Labeler setup wizard"
        title1 = "Destination";          sub1 = "Choose the installation folder"
        title2 = "Installing";           sub2 = "Please wait: this may take a few minutes"
        title3 = "Done";                 sub3 = "Operation completed"
        back = "< Back"; next = "Next >"; install = "Install"; finish = "Finish"
        cancel = "Cancel"; close = "Close"; browse = "Browse..."
        welcome = "This program prepares everything needed and builds the Gmail Labeler application.\n\nThe following steps run automatically, with no restart:\n\n   •  check for Python (if missing, Python {0} is installed)\n   •  check for the Google libraries and PyInstaller (installed if missing)\n   •  build GmailLabeler.exe with its icon, plus the uninstaller\n   •  copy the application to the folder you choose\n\nAn Internet connection is required. Installing in Program Files requires administrator rights.\n\nClick Next to continue."
        destInfo = "Choose where to save GmailLabeler.exe.\ncredentials.json and token.json will be stored for each user in %APPDATA%\Gmail Labeler\credentials."
        shortcut = "Create a Desktop shortcut"
        destEmpty = "Please choose a destination folder."
        dlgDesc = "Choose the destination folder"
        waiting = "Waiting..."; cancelling = "Cancelling..."
        failed = "Installation failed."; done = "Completed."
        s1 = "1/6  Checking Python..."
        s1b = "1/6  Installing Python..."
        s2 = "2/6  Checking the graphical interface (Tkinter)..."
        s3 = "3/6  Checking libraries..."
        s4 = "4/6  Building GmailLabeler.exe (this may take a few minutes)..."
        s5 = "5/6  Building the uninstaller..."
        s6 = "6/6  Copying files..."
        endText = "Installation completed.\n\nGmailLabeler.exe is located in:\n{0}\n\nTo uninstall: Settings > Apps, or run _internal\uninstall.exe\n\nOn first launch use the 'Select credentials.json' button (there is also a drop-down guide). Credentials are saved in %APPDATA%\Gmail Labeler\credentials."
        openFolder = "Open the folder"; runApp = "Launch Gmail Labeler"
        cancelled = "Operation cancelled."
        destInvalid = "Invalid destination folder."
        pyFound = "Python found: {0}"
        pyInstalling = "Python {0} not found: starting installation..."
        wingetFail = "winget failed: {0}"
        wingetNone = "winget did not install Python: trying direct download from python.org..."
        dl = "Downloading Python {0} ..."
        dlNA = "  not available ({0})"
        dlFail = "Unable to download Python. Check your Internet connection."
        pyInstall = "Installing Python (for all users)..."
        pyInstCode = "The Python installer ended with code {0}."
        pyNotFound = "Python was installed but could not be located."
        pyReady = "Python ready: {0}"
        pyVer = "Python version: {0} - Tkinter OK"
        pkgInstall = "Installing {0} ..."
        pkgFail = "Installation of {0} failed (code {1})."
        buildFail = "Building the exe failed (code {0}). See the log."
        uninstFail = "Building the uninstaller failed (code {0}). See the log."
        uninstFallback = "The updated uninstaller does not compile: using the previous version (see the errors above)."
        created = "Created: {0}"
        shortcutOk = "Desktop shortcut created."
        shortcutFail = "Shortcut not created: {0}"
        regOk = "Uninstall entry registered in Settings > Apps."
        regFail = "Registration failed: {0}"
        errorPrefix = "ERROR: {0}"
        notAdmin = "This setup must be run as administrator (right-click > Run as administrator)."
    }
}

$sysLang = [System.Globalization.CultureInfo]::CurrentUICulture.TwoLetterISOLanguageName
$script:lang = if ($sysLang -eq 'it') { 'it' } else { 'en' }

function T([string]$key) {
    $v = $STR[$script:lang][$key]
    if ($null -eq $v) { $v = $STR['en'][$key] }
    $v = $v -replace '\\n', "`r`n"
    if ($args.Count -gt 0) { return ($v -f $args) }
    return $v
}

# Il setup installa in Programmi: servono i permessi di amministratore
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    [void][System.Windows.Forms.MessageBox]::Show((T 'notAdmin'), 'Gmail Labeler - Setup', 'OK', 'Warning')
    exit 1
}

# ------------------------------------------------------------
# Utilità
# ------------------------------------------------------------
function Pump { [System.Windows.Forms.Application]::DoEvents() }

function Log([string]$msg) {
    if ($null -eq $msg) { return }
    $txtLog.AppendText($msg + "`r`n")
    Pump
}

function Set-Progress([int]$v, [string]$label) {
    $bar.Value = [Math]::Min(100, [Math]::Max(0, $v))
    $lblStep.Text = $label
    Pump
}

function Refresh-Path {
    $m = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    $u = [Environment]::GetEnvironmentVariable('Path', 'User')
    $env:Path = "$m;$u"
}

# Esegue un processo mostrando l'output nel log, senza bloccare la finestra
function Invoke-Logged([string]$file, [string[]]$arguments, [string]$workDir) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $file
    $psi.Arguments = (($arguments | ForEach-Object { if ($_ -match '[\s;]') { '"' + $_ + '"' } else { $_ } }) -join ' ')
    $psi.WorkingDirectory = $workDir
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $psi.StandardOutputEncoding = [Text.Encoding]::UTF8
    $psi.StandardErrorEncoding = [Text.Encoding]::UTF8

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi
    [void]$p.Start()

    $outTask = $p.StandardOutput.ReadLineAsync()
    $errTask = $p.StandardError.ReadLineAsync()
    $outDone = $false
    $errDone = $false

    while (-not ($outDone -and $errDone)) {
        if ($script:cancel) { try { $p.Kill() } catch {}; throw (T 'cancelled') }
        $idle = $true
        if (-not $outDone -and $outTask.IsCompleted) {
            $line = $outTask.Result
            if ($null -eq $line) { $outDone = $true } else { Log $line; $outTask = $p.StandardOutput.ReadLineAsync() }
            $idle = $false
        }
        if (-not $errDone -and $errTask.IsCompleted) {
            $line = $errTask.Result
            if ($null -eq $line) { $errDone = $true } else { Log $line; $errTask = $p.StandardError.ReadLineAsync() }
            $idle = $false
        }
        Pump
        if ($idle) { Start-Sleep -Milliseconds 30 }
    }
    while (-not $p.HasExited) {
        if ($script:cancel) { try { $p.Kill() } catch {}; throw (T 'cancelled') }
        Pump; Start-Sleep -Milliseconds 30
    }
    return $p.ExitCode
}

function Download-File([string]$url, [string]$dest) {
    $wc = New-Object System.Net.WebClient
    try {
        $task = $wc.DownloadFileTaskAsync($url, $dest)
        while (-not $task.IsCompleted) {
            if ($script:cancel) { $wc.CancelAsync(); throw (T 'cancelled') }
            Pump; Start-Sleep -Milliseconds 50
        }
        if ($task.IsFaulted) { throw $task.Exception.InnerException }
    } finally { $wc.Dispose() }
}

# ------------------------------------------------------------
# Ricerca di Python (compatibile, con tkinter)
# ------------------------------------------------------------
function Test-PythonExe([string]$exe) {
    if (-not $exe -or -not (Test-Path $exe)) { return $null }
    if ($exe -match 'WindowsApps') { return $null }     # alias del Microsoft Store
    try {
        $out = & $exe -c "import sys, tkinter; print('%d.%d' % sys.version_info[:2])" 2>$null
        if ($LASTEXITCODE -ne 0) { return $null }
        $v = [version](($out | Select-Object -First 1).ToString().Trim())
        if ($v -ge $PY_MIN -and $v -le $PY_MAX) { return $v }
    } catch {}
    return $null
}

function Test-Import([string]$exe, [string]$module) {
    try {
        & $exe -c "import $module" 2>$null | Out-Null
        return ($LASTEXITCODE -eq 0)
    } catch { return $false }
}

function Find-Python {
    Refresh-Path
    $candidates = New-Object System.Collections.Generic.List[string]

    if (Get-Command py -ErrorAction SilentlyContinue) {
        try {
            foreach ($line in (& py -0p 2>$null)) {
                if ($line -match '([A-Za-z]:\\.*?python(w)?\.exe)') {
                    $path = $Matches[1] -replace 'pythonw\.exe$', 'python.exe'
                    $candidates.Add($path)
                }
            }
        } catch {}
    }
    foreach ($base in @("$env:LOCALAPPDATA\Programs\Python", $env:ProgramW6432, $env:ProgramFiles, ${env:ProgramFiles(x86)})) {
        if ($base -and (Test-Path $base)) {
            Get-ChildItem -Path $base -Directory -Filter 'Python3*' -ErrorAction SilentlyContinue |
                ForEach-Object { $candidates.Add((Join-Path $_.FullName 'python.exe')) }
        }
    }
    try {
        Get-Command python -All -ErrorAction SilentlyContinue |
            ForEach-Object { $candidates.Add($_.Source) }
    } catch {}

    $best = $null; $bestVer = [version]'0.0'
    foreach ($c in ($candidates | Select-Object -Unique)) {
        $v = Test-PythonExe $c
        if ($v -and $v -gt $bestVer) { $best = $c; $bestVer = $v }
    }
    return $best
}

function Install-Python([string]$work) {
    Log (T 'pyInstalling' $PY_VERSION_TARGET)

    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Log "> winget install Python.Python.$PY_VERSION_TARGET"
        try {
            [void](Invoke-Logged 'winget' @('install', '-e', '--id', "Python.Python.$PY_VERSION_TARGET",
                '--silent', '--accept-package-agreements', '--accept-source-agreements') $work)
        } catch { if ($script:cancel) { throw }; Log (T 'wingetFail' $_.Exception.Message) }
        $found = Find-Python
        if ($found) { return $found }
        Log (T 'wingetNone')
    }

    $installer = Join-Path $work 'python-installer.exe'
    $got = $false
    foreach ($v in $PY_FALLBACKS) {
        $url = "https://www.python.org/ftp/python/$v/python-$v-amd64.exe"
        try {
            Log (T 'dl' $v)
            Download-File $url $installer
            $got = $true; break
        } catch {
            if ($script:cancel) { throw }
            Log (T 'dlNA' $v)
        }
    }
    if (-not $got) { throw (T 'dlFail') }

    Log (T 'pyInstall')
    $code = Invoke-Logged $installer @('/quiet', 'InstallAllUsers=1', 'PrependPath=1',
        'Include_test=0', 'Include_launcher=0', 'Shortcuts=0') $work
    if ($code -ne 0) { throw (T 'pyInstCode' $code) }

    $found = Find-Python
    if (-not $found) { throw (T 'pyNotFound') }
    return $found
}

# ------------------------------------------------------------
# Procedura di installazione
# ------------------------------------------------------------
function Run-Setup {
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $dest = $txtDest.Text.Trim()
    if (-not $dest) { throw (T 'destInvalid') }

    $work = Join-Path $env:TEMP 'GmailLabelerBuild'
    if (Test-Path $work) { Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Path $work -Force | Out-Null

    # 1. Python
    Set-Progress 5 (T 's1')
    $py = Find-Python
    if ($py) {
        Log (T 'pyFound' $py)
    } else {
        Set-Progress 10 (T 's1b')
        $py = Install-Python $work
        Log (T 'pyReady' $py)
    }

    # 2. Tkinter
    Set-Progress 25 (T 's2')
    $v = Test-PythonExe $py
    Log (T 'pyVer' $v)

    # 3. Librerie
    Set-Progress 30 (T 's3')
    $pkgs = @(
        @('googleapiclient',        'google-api-python-client'),
        @('google_auth_httplib2',   'google-auth-httplib2'),
        @('google_auth_oauthlib',   'google-auth-oauthlib'),
        @('PyInstaller',            'pyinstaller'),
        @('PIL',                    'pillow')
    )
    [void](Invoke-Logged $py @('-m', 'ensurepip', '--upgrade') $work)
    $i = 0
    foreach ($p in $pkgs) {
        $i++
        if (Test-Import $py $p[0]) {
            Log "OK: $($p[1])"
        } else {
            Log (T 'pkgInstall' $p[1])
            $code = Invoke-Logged $py @('-m', 'pip', 'install', '--upgrade', '--disable-pip-version-check', $p[1]) $work
            if ($code -ne 0) { throw (T 'pkgFail' $p[1] $code) }
        }
        Set-Progress (30 + 4 * $i) (T 's3')
    }

    # 4. App (cartella)
    Set-Progress 55 (T 's4')
    $pyFile  = Join-Path $work 'gmail_labeler_gui.py'
    $icoFile = Join-Path $work 'icon.ico'
    $verFile = Join-Path $work 'version_info.txt'
    [IO.File]::WriteAllText($pyFile, $PY_SOURCE, $utf8)
    [IO.File]::WriteAllBytes($icoFile, [Convert]::FromBase64String(($ICO_B64 -replace '\s', '')))
    [IO.File]::WriteAllText($verFile, $VERSION_INFO, $utf8)

    $code = Invoke-Logged $py @('-m', 'PyInstaller', '--noconfirm', '--clean', '--onedir', '--windowed', '--noupx',
        '--name', 'GmailLabeler', '--icon', 'icon.ico', '--add-data', 'icon.ico;.', '--version-file', 'version_info.txt',
        '--collect-all', 'googleapiclient', 'gmail_labeler_gui.py') $work
    $builtDir = Join-Path $work 'dist\GmailLabeler'
    $built = Join-Path $builtDir 'GmailLabeler.exe'
    if ($code -ne 0 -or -not (Test-Path $built)) { throw (T 'buildFail' $code) }

    # 5. Disinstallatore: mini exe C# compilato con csc.exe (incluso in Windows)
    #    Icona leggera a frame BMP: un'icona PNG da 256 px renderebbe l'exe quasi tutto "dati ad alta entropia"
    Set-Progress 78 (T 's5')
    [IO.File]::WriteAllText((Join-Path $work 'uninstall.cs'), $UNINSTALL_CS, $utf8)
    [IO.File]::WriteAllText((Join-Path $work 'uninstall.manifest'), $UNINSTALL_MANIFEST, $utf8)
    [IO.File]::WriteAllBytes((Join-Path $work 'icon_small.ico'), [Convert]::FromBase64String(($ICO_SMALL_B64 -replace '\s', '')))

    $csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
    if (-not (Test-Path $csc)) { $csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
    if (-not (Test-Path $csc)) { throw (T 'uninstFail' '-1') }

    $cscArgs = @('/nologo', '/target:winexe', '/platform:anycpu', '/optimize+', '/out:uninstall.exe',
        '/win32manifest:uninstall.manifest', '/r:System.Windows.Forms.dll', '/r:System.Drawing.dll')
    $unExe = Join-Path $work 'uninstall.exe'
    $code = Invoke-Logged $csc ($cscArgs + @('/win32icon:icon_small.ico', 'uninstall.cs')) $work
    if ($code -ne 0 -or -not (Test-Path $unExe)) {
        $code = Invoke-Logged $csc ($cscArgs + @('uninstall.cs')) $work     # riprova senza icona
    }
    if ($code -ne 0 -or -not (Test-Path $unExe)) {
        # sicurezza: se il sorgente aggiornato non compila, usa la versione precedente del disinstallatore
        Log (T 'uninstFallback')
        [IO.File]::WriteAllText((Join-Path $work 'uninstall.cs'), $UNINSTALL_CS_LEGACY, $utf8)
        $code = Invoke-Logged $csc ($cscArgs + @('/win32icon:icon_small.ico', 'uninstall.cs')) $work
    }
    if ($code -ne 0 -or -not (Test-Path $unExe)) { throw (T 'uninstFail' $code) }

    # 6. Copia nella destinazione (Programmi)
    Set-Progress 92 (T 's6')
    New-Item -ItemType Directory -Path $dest -Force | Out-Null
    $script:resultExe = Join-Path $dest 'GmailLabeler.exe'
    Copy-Item (Join-Path $builtDir '*') $dest -Recurse -Force

    $internal = Join-Path $dest '_internal'
    $uninstExe = Join-Path $internal 'uninstall.exe'
    Copy-Item $unExe $uninstExe -Force

    Log (T 'created' $script:resultExe)

    # Voce in Impostazioni > App (per tutti gli utenti, vista registro a 64 bit)
    try {
        $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
        $k = $hklm.CreateSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\GmailLabeler')
        $k.SetValue('DisplayName', 'Gmail Labeler')
        $k.SetValue('DisplayVersion', '1.0.0')
        $k.SetValue('Publisher', 'Gmail Labeler')
        $k.SetValue('InstallLocation', $dest)
        $k.SetValue('DisplayIcon', $script:resultExe)
        $k.SetValue('UninstallString', ('"' + $uninstExe + '"'))
        $k.SetValue('NoModify', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
        $k.SetValue('NoRepair', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
        $k.Close(); $hklm.Close()
        Log (T 'regOk')
    } catch { Log (T 'regFail' $_.Exception.Message) }

    if ($chkShortcut.Checked) {
        try {
            $desktop = [Environment]::GetFolderPath('CommonDesktopDirectory')
            $ws = New-Object -ComObject WScript.Shell
            $lnk = $ws.CreateShortcut((Join-Path $desktop 'Gmail Labeler.lnk'))
            $lnk.TargetPath = $script:resultExe
            $lnk.WorkingDirectory = $dest
            $lnk.IconLocation = $script:resultExe
            $lnk.Save()
            Log (T 'shortcutOk')
        } catch { Log (T 'shortcutFail' $_.Exception.Message) }
    }

    try { Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue } catch {}
    Set-Progress 100 (T 'done')
}

# ------------------------------------------------------------
# Interfaccia wizard
# ------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Gmail Labeler - Setup'
$form.ClientSize = New-Object System.Drawing.Size(660, 470)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
try {
    $icoBytes = [Convert]::FromBase64String(($ICO_B64 -replace '\s', ''))
    $form.Icon = New-Object System.Drawing.Icon((New-Object System.IO.MemoryStream(, $icoBytes)))
} catch {}

# Intestazione
$header = New-Object System.Windows.Forms.Panel
$header.Dock = 'Top'; $header.Height = 70; $header.BackColor = [System.Drawing.Color]::FromArgb(37, 99, 235)
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.ForeColor = 'White'; $lblTitle.AutoSize = $true
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 15)
$lblTitle.Location = New-Object System.Drawing.Point(20, 10)
$lblSub = New-Object System.Windows.Forms.Label
$lblSub.ForeColor = [System.Drawing.Color]::FromArgb(219, 234, 254); $lblSub.AutoSize = $true
$lblSub.Location = New-Object System.Drawing.Point(22, 42)

$cmbLang = New-Object System.Windows.Forms.ComboBox
$cmbLang.DropDownStyle = 'DropDownList'
$cmbLang.Size = New-Object System.Drawing.Size(100, 24)
$cmbLang.Location = New-Object System.Drawing.Point(545, 22)
[void]$cmbLang.Items.Add('Italiano')
[void]$cmbLang.Items.Add('English')
$cmbLang.SelectedIndex = $(if ($script:lang -eq 'it') { 0 } else { 1 })
$header.Controls.AddRange(@($lblTitle, $lblSub, $cmbLang))

# Pannelli pagina
function New-Page {
    $p = New-Object System.Windows.Forms.Panel
    $p.Location = New-Object System.Drawing.Point(0, 70)
    $p.Size = New-Object System.Drawing.Size(660, 340)
    $p.Visible = $false
    return $p
}
$pages = @((New-Page), (New-Page), (New-Page), (New-Page))

# --- Pagina 0: benvenuto
$lblWelcome = New-Object System.Windows.Forms.Label
$lblWelcome.Location = New-Object System.Drawing.Point(25, 20); $lblWelcome.Size = New-Object System.Drawing.Size(610, 300)
$pages[0].Controls.Add($lblWelcome)

# --- Pagina 1: destinazione
$lblDest = New-Object System.Windows.Forms.Label
$lblDest.Location = New-Object System.Drawing.Point(25, 25); $lblDest.Size = New-Object System.Drawing.Size(610, 45)
$txtDest = New-Object System.Windows.Forms.TextBox
$txtDest.Location = New-Object System.Drawing.Point(25, 85); $txtDest.Size = New-Object System.Drawing.Size(480, 24)
$pf = $(if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles })
$txtDest.Text = Join-Path $pf 'Gmail Labeler'
$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Location = New-Object System.Drawing.Point(515, 83); $btnBrowse.Size = New-Object System.Drawing.Size(110, 27)
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = (T 'dlgDesc')
    if ($dlg.ShowDialog() -eq 'OK') { $txtDest.Text = Join-Path $dlg.SelectedPath 'Gmail Labeler' }
})
$chkShortcut = New-Object System.Windows.Forms.CheckBox
$chkShortcut.Checked = $true
$chkShortcut.Location = New-Object System.Drawing.Point(25, 130); $chkShortcut.AutoSize = $true
$pages[1].Controls.AddRange(@($lblDest, $txtDest, $btnBrowse, $chkShortcut))

# --- Pagina 2: installazione
$lblStep = New-Object System.Windows.Forms.Label
$lblStep.Location = New-Object System.Drawing.Point(25, 15); $lblStep.Size = New-Object System.Drawing.Size(610, 20)
$bar = New-Object System.Windows.Forms.ProgressBar
$bar.Location = New-Object System.Drawing.Point(25, 40); $bar.Size = New-Object System.Drawing.Size(610, 22)
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true; $txtLog.ReadOnly = $true; $txtLog.ScrollBars = 'Vertical'
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 8.5)
$txtLog.BackColor = 'White'
$txtLog.Location = New-Object System.Drawing.Point(25, 75); $txtLog.Size = New-Object System.Drawing.Size(610, 250)
$pages[2].Controls.AddRange(@($lblStep, $bar, $txtLog))

# --- Pagina 3: fine
$lblEnd = New-Object System.Windows.Forms.Label
$lblEnd.Location = New-Object System.Drawing.Point(25, 15); $lblEnd.Size = New-Object System.Drawing.Size(610, 235)
$chkOpen = New-Object System.Windows.Forms.CheckBox
$chkOpen.Checked = $true
$chkOpen.Location = New-Object System.Drawing.Point(25, 260); $chkOpen.AutoSize = $true
$chkRun = New-Object System.Windows.Forms.CheckBox
$chkRun.Checked = $true
$chkRun.Location = New-Object System.Drawing.Point(25, 288); $chkRun.AutoSize = $true
$pages[3].Controls.AddRange(@($lblEnd, $chkOpen, $chkRun))

# Pulsanti
$footer = New-Object System.Windows.Forms.Panel
$footer.Dock = 'Bottom'; $footer.Height = 60
$btnBack = New-Object System.Windows.Forms.Button
$btnBack.Size = New-Object System.Drawing.Size(100, 30); $btnBack.Location = New-Object System.Drawing.Point(335, 14)
$btnNext = New-Object System.Windows.Forms.Button
$btnNext.Size = New-Object System.Drawing.Size(100, 30); $btnNext.Location = New-Object System.Drawing.Point(445, 14)
$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Size = New-Object System.Drawing.Size(100, 30); $btnCancel.Location = New-Object System.Drawing.Point(555, 14)
$footer.Controls.AddRange(@($btnBack, $btnNext, $btnCancel))

$form.Controls.AddRange([System.Windows.Forms.Control[]](@($footer, $header) + $pages))
$script:page = 0

function Show-Page([int]$n) {
    $script:page = $n
    for ($i = 0; $i -lt $pages.Count; $i++) { $pages[$i].Visible = ($i -eq $n) }
    $lblTitle.Text = (T "title$n"); $lblSub.Text = (T "sub$n")
    $btnBack.Text = (T 'back'); $btnCancel.Text = (T 'cancel')
    $btnBack.Enabled = ($n -eq 1)
    $btnNext.Enabled = $true
    $btnNext.Text = (T 'next')
    $btnCancel.Enabled = $true
    $cmbLang.Enabled = ($n -le 1)
    if ($n -eq 1) { $btnNext.Text = (T 'install') }
    if ($n -eq 2) { $btnBack.Enabled = $false; $btnNext.Enabled = $false }
    if ($n -eq 3) { $btnBack.Enabled = $false; $btnNext.Text = (T 'finish'); $btnCancel.Enabled = $false }
}

function Apply-Language {
    $lblWelcome.Text = (T 'welcome' $PY_VERSION_TARGET)
    $lblDest.Text = (T 'destInfo')
    $btnBrowse.Text = (T 'browse')
    $chkShortcut.Text = (T 'shortcut')
    $chkOpen.Text = (T 'openFolder')
    $chkRun.Text = (T 'runApp')
    $lblStep.Text = (T 'waiting')
    Show-Page $script:page
}

$cmbLang.Add_SelectedIndexChanged({
    $script:lang = $(if ($cmbLang.SelectedIndex -eq 0) { 'it' } else { 'en' })
    Apply-Language
})

function Start-Install {
    $script:cancel = $false
    $script:installing = $true
    $txtLog.Clear()
    $bar.Value = 0
    try {
        Run-Setup
        $script:installing = $false
        $lblEnd.Text = (T 'endText' $script:resultExe)
        Show-Page 3
    } catch {
        $script:installing = $false
        $msg = $_.Exception.Message
        Log ''
        Log (T 'errorPrefix' $msg)
        $lblStep.Text = (T 'failed')
        $btnBack.Enabled = $true
        $btnNext.Enabled = $false
        $btnCancel.Text = (T 'close')
        $btnCancel.Enabled = $true
    }
}

$btnBack.Add_Click({ if ($script:page -gt 0 -and -not $script:installing) { Show-Page ($script:page - 1) } })

$btnNext.Add_Click({
    switch ($script:page) {
        0 { Show-Page 1 }
        1 {
            if (-not $txtDest.Text.Trim()) {
                [System.Windows.Forms.MessageBox]::Show((T 'destEmpty'), 'Setup') | Out-Null
                return
            }
            Show-Page 2
            Start-Install
        }
        3 {
            if ($chkOpen.Checked) { Start-Process explorer.exe ('"' + (Split-Path $script:resultExe) + '"') }
            if ($chkRun.Checked)  { Start-Process explorer.exe ('"' + $script:resultExe + '"') }
            $form.Close()
        }
    }
})

$btnCancel.Add_Click({
    if ($script:installing) {
        $script:cancel = $true
        $lblStep.Text = (T 'cancelling')
    } else {
        $form.Close()
    }
})

$form.Add_FormClosing({
    if ($script:installing) {
        $script:cancel = $true
        $_.Cancel = $true
    }
})

Apply-Language
[void]$form.ShowDialog()
