Attribute VB_Name = "modLang"
Option Explicit

'Start variable support for languages
' MDI form ..
Public MDI_COMMAND_CONFIG               As String
Public MDI_COMMAND_CHAT                 As String
Public MDI_COMMAND_SENDFILE             As String
Public MDI_COMMAND_SOCIETY              As String

Public MDI_STAT_DISCONNECTED            As String
Public MDI_STAT_DCFROMSERVER            As String
Public MDI_STAT_CONNECTED               As String
Public MDI_STAT_CONNECTION_ERROR        As String
Public MDI_STAT_CONNECTING              As String

Public MDI_MSG_ERROR_FORM_LOAD          As String
Public MDI_MSG_CONFIG_NOTIFY            As String
Public MDI_MSG_NAME_TAKEN               As String
Public MDI_MSG_WRONG_ACCOUNT            As String
Public MDI_MSG_WRONG_PASSWORD           As String
Public MDI_MSG_BANNED                   As String

' Configuration form ..
Public CONFIG_COMMAND_CONNECT           As String
Public CONFIG_COMMAND_DISCONNECT        As String
Public CONFIG_COMMAND_SETTINGS          As String
Public CONFIG_COMMAND_UPDATE            As String
Public CONFIG_COMMAND_REGISTER          As String

Public CONFIG_CHECK_SAVE_PASSWORD       As String

Public CONFIG_FRAME_CONNECTION          As String

Public CONFIG_MSG_ACCOUNT               As String
Public CONFIG_MSG_PASSWORD              As String
Public CONFIG_MSG_NUMERIC               As String
Public CONFIG_MSG_PORT                  As String
Public CONFIG_MSG_NAME                  As String
Public CONFIG_MSG_IP                    As String

Public CHAT_COMMAND_SEND                As String
Public CHAT_COMMAND_CLEAR               As String

Public CHAT_TIME_TEXT                   As String

Public SF_LABEL_FILENAME                As String
Public SF_LABEL_SENDING_FILE            As String
Public SF_LABEL_SENT                    As String
Public SF_LABEL_SEND_TO                 As String

Public SF_COMMAND_BROWSE                As String
Public SF_COMMAND_SENDFILE              As String
Public SF_COMMAND_CANCEL                As String

Public SF_MSG_USER                      As String
Public SF_MSG_FILE                      As String
Public SF_MSG_INCOMMING_FILE            As String
Public SF_MSG_DECILINED                 As String

'Desp form ..
Public DESP_TEXT_NEW_MSG                As String
Public DESP_TEXT_DC_SERVER              As String

'Language form ..
Public LANG_COMMAND_ENTER               As String
Public LANG_LABEL_SELLANG               As String

Public LANG_GERMAN                      As String
Public LANG_ENGLISH                     As String
Public LANG_SPANISH                     As String
Public LANG_SWEDISH                     As String
Public LANG_ITALIAN                     As String
Public LANG_SERBIAN                     As String
Public LANG_DUTCH                       As String
Public LANG_FRENCH                      As String

Public Sub SetLangGerman()

' MDI form ..
MDI_COMMAND_CONFIG = "&Einstellungen"
MDI_COMMAND_CHAT = "Ch&at"
MDI_COMMAND_SENDFILE = "&Sende Datei"
MDI_COMMAND_SOCIETY = "&Online Liste"

MDI_STAT_DISCONNECTED = "Status: Getrennt"
MDI_STAT_DCFROMSERVER = "Status: Getrennt vom Server"
MDI_STAT_CONNECTED = "Status: Verbunden"
MDI_STAT_CONNECTION_ERROR = "Status: Getrennt aufgrund eines Verbindungsfehlers"
MDI_STAT_CONNECTING = "Status: Verbindung wird aufgebaut .."

MDI_MSG_CONFIG_NOTIFY = "Einige Konfigurationsdateien sind veraltet oder wurden beschädigt, Peach fand den Fehler und wird es mit dem nächsten Neustart korrigieren."
MDI_MSG_NAME_TAKEN = "Der Name ist bereits vergeben."
MDI_MSG_WRONG_ACCOUNT = "Der Account ist nicht vorhanden oder falsch."
MDI_MSG_WRONG_PASSWORD = "Das Passwort ist falsch."
MDI_MSG_BANNED = "Dieser Account wurde gebannt."

CONFIG_COMMAND_CONNECT = "&Verbinden"
CONFIG_COMMAND_DISCONNECT = "&Verbindung trenn."
CONFIG_COMMAND_SETTINGS = "&Einstellungen"
CONFIG_COMMAND_UPDATE = "&Aktualisieren"
CONFIG_COMMAND_REGISTER = "&Account registrieren"

CONFIG_CHECK_SAVE_PASSWORD = "&Password Speichern"

CONFIG_FRAME_CONNECTION = "Verbindungs Informationen"

LANG_GERMAN = "Deutsch"
LANG_ENGLISH = "Englisch"
LANG_SPANISH = "Spanisch"
LANG_SWEDISH = "Schwedisch"
LANG_ITALIAN = "Italienisch"
LANG_SERBIAN = "Serbisch"
LANG_DUTCH = "Niederländisch"
LANG_FRENCH = "Französisch"

CONFIG_MSG_ACCOUNT = "Du hast keinen Account eingegeben."
CONFIG_MSG_PASSWORD = "Du hast kein Passwort eingegeben."
CONFIG_MSG_NUMERIC = "Du kannst keine Ziffern in deinem Namen haben."
CONFIG_MSG_PORT = "Du hast keinen Port eingegeben."
CONFIG_MSG_NAME = "Du hast keinen Namen eingegeben."
CONFIG_MSG_IP = "Du hast keine IP eingegeben."

CHAT_COMMAND_SEND = "&Senden"
CHAT_COMMAND_CLEAR = "&Löschen"

CHAT_TIME_TEXT = " Die Zeit beträgt "

SF_LABEL_FILENAME = " Datei Name:"
SF_LABEL_SENDING_FILE = "Sende:"
SF_LABEL_SENT = "0.0% Gesendet"
SF_LABEL_SEND_TO = "Sende an:"

SF_MSG_USER = "Kein Benutzer ausgewählt."
SF_MSG_FILE = "Keine Datei ausgewählt."
SF_MSG_INCOMMING_FILE = "Sie empfangen eine Datei, möchsten sie annehmen?"
SF_MSG_DECILINED = "Der Benutzer hat die Datei abgelehnt."

SF_COMMAND_BROWSE = "&Suchen .."
SF_COMMAND_SENDFILE = "Senden"
SF_COMMAND_CANCEL = "Abbrechen .."

DESP_TEXT_NEW_MSG = "Neue Nachricht!"
DESP_TEXT_DC_SERVER = "Verbindung unterbrochen!"

LANG_COMMAND_ENTER = "&Auswählen"
LANG_LABEL_SELLANG = "Wähle deine Sprache aus:"
End Sub

Public Sub SetLangEnglish()

'MDI form
MDI_COMMAND_CONFIG = "&Configuration"
MDI_COMMAND_CHAT = "Ch&at"
MDI_COMMAND_SENDFILE = "&Send File"
MDI_COMMAND_SOCIETY = "&Online List"

MDI_STAT_DISCONNECTED = "Status: Disconnected"
MDI_STAT_DCFROMSERVER = "Status: Disconnected from Server"
MDI_STAT_CONNECTED = "Status: Connected"
MDI_STAT_CONNECTION_ERROR = "Status: Cant connect to server. (Server offline)"
MDI_STAT_CONNECTING = "Status: Connecting .."

'MDImsgbox_errorHandlerFormLoad
MDI_MSG_CONFIG_NOTIFY = "Some configuration files are outdated or got damaged, Peach found the problem and will fix it on next program launch."
MDI_MSG_NAME_TAKEN = "This name is already taken."
MDI_MSG_WRONG_ACCOUNT = "The account does not exist or is wrong."
MDI_MSG_WRONG_PASSWORD = "The password is wrong."
MDI_MSG_BANNED = "This account is banned."

' Configuration form ..
CONFIG_COMMAND_CONNECT = "&Connect"
CONFIG_COMMAND_DISCONNECT = "&Disconnect"
CONFIG_COMMAND_SETTINGS = "&Settings"
CONFIG_COMMAND_UPDATE = "&Update"
CONFIG_COMMAND_REGISTER = "&Register Account"

CONFIG_CHECK_SAVE_PASSWORD = "&Save Password"

CONFIG_FRAME_CONNECTION = "Connection Information"

LANG_GERMAN = "German"
LANG_ENGLISH = "English"
LANG_SPANISH = "Spanish"
LANG_SWEDISH = "Swedish"
LANG_ITALIAN = "Italian"
LANG_SERBIAN = "Serbian"
LANG_DUTCH = "Dutch"
LANG_FRENCH = "French"

CONFIG_MSG_ACCOUNT = "You didnt introduce an account."
CONFIG_MSG_PASSWORD = "You didnt introduce an password."
CONFIG_MSG_NUMERIC = "You cant take numeric names."
CONFIG_MSG_PORT = "You didnt introduce an port."
CONFIG_MSG_NAME = "You didnt introduce an name."
CONFIG_MSG_IP = "You didnt introduce an IP."

' Chat form ..
CHAT_COMMAND_SEND = "&Send"
CHAT_COMMAND_CLEAR = "&Clear"

CHAT_TIME_TEXT = " The time is "

' Send File form ..
SF_LABEL_FILENAME = " File Name:"
SF_LABEL_SENDING_FILE = "Sending:"
SF_LABEL_SENT = "0.0% Sent"
SF_LABEL_SEND_TO = "Send to:"

SF_MSG_USER = "No user selected."
SF_MSG_FILE = "No file selected."
SF_MSG_INCOMMING_FILE = "You are getting an incomming file, do you want to accept?"
SF_MSG_DECILINED = "File transfer was decilined."

SF_COMMAND_BROWSE = "&Search .."
SF_COMMAND_SENDFILE = "Send"
SF_COMMAND_CANCEL = "Cancel .."

DESP_TEXT_NEW_MSG = "New Message!"
DESP_TEXT_DC_SERVER = "Disconnected from Server!"

LANG_COMMAND_ENTER = "&Select"
LANG_LABEL_SELLANG = "Select your language:"
End Sub

Public Sub SetLangSpanish()

MDI_COMMAND_CONFIG = "&Configuración"
MDI_COMMAND_CHAT = "Ch&at"
MDI_COMMAND_SENDFILE = "&Enviar Archivo"
MDI_COMMAND_SOCIETY = "&Lista Online"

MDI_STAT_DISCONNECTED = "Estado: Desconectado"
MDI_STAT_DCFROMSERVER = "Estado: Desconectado del servidor"
MDI_STAT_CONNECTED = "Estado: Disponible"
MDI_STAT_CONNECTION_ERROR = "Estado: Desconectado por problemas de conexión"
MDI_STAT_CONNECTING = "Estado: Conectando .."

MDI_MSG_CONFIG_NOTIFY = "Alguna configuración de archivos estan caducados o dañados, Peach busca el problema y lo arreglará en el siguiente lanzamiento del programa."
MDI_MSG_NAME_TAKEN = "Este nombre ya esta cogido."
MDI_MSG_WRONG_ACCOUNT = "La cuenta no existe o es incorrecta."
MDI_MSG_WRONG_PASSWORD = "La contraseña es incorrecta."
MDI_MSG_BANNED = "Esta cuenta esta baneada."

CONFIG_COMMAND_CONNECT = "&Conectar"
CONFIG_COMMAND_DISCONNECT = "&Desconectar"
CONFIG_COMMAND_SETTINGS = "&Ajustes"
CONFIG_COMMAND_UPDATE = "&Actualizar"
CONFIG_COMMAND_REGISTER = "&Registrar cuenta"

CONFIG_CHECK_SAVE_PASSWORD = "&Save Password"

CONFIG_FRAME_CONNECTION = "Informaciónes de conexión"

LANG_GERMAN = "Aleman"
LANG_ENGLISH = "Inglés"
LANG_SPANISH = "Español"
LANG_SWEDISH = "Sueco"
LANG_ITALIAN = "Italiano"
LANG_DUTCH = "Holandés"
LANG_SERBIAN = "Serbio"
LANG_FRENCH = "Frances"

CONFIG_MSG_ACCOUNT = "You didnt introduce an account."
CONFIG_MSG_PASSWORD = "You didnt introduce an password."
CONFIG_MSG_NUMERIC = "No puedes coger nombres con numeros."
CONFIG_MSG_PORT = "No has introducido un puerto."
CONFIG_MSG_NAME = "No has introducido un nombre."
CONFIG_MSG_IP = "No has introducido una direccion."

CHAT_COMMAND_SEND = "&Enviar"
CHAT_COMMAND_CLEAR = "&Limpiar"

CHAT_TIME_TEXT = " El tiempo es "

SF_LABEL_FILENAME = " Nombre del archivo:"
SF_LABEL_SENDING_FILE = "Enviando:"
SF_LABEL_SENT = "0.0% Enviado"
SF_LABEL_SEND_TO = "Enviar a:"

SF_MSG_USER = "No has seleccionado a una persona."
SF_MSG_FILE = "No has seleccionado a un archivo."
SF_MSG_INCOMMING_FILE = "Estas recibiendo un archivo, quieres aceptar?"
SF_MSG_DECILINED = "El envio ha sido rechazado."

SF_COMMAND_BROWSE = "&Buscar .."
SF_COMMAND_SENDFILE = "Enviar"
SF_COMMAND_CANCEL = "Cancelar .."

DESP_TEXT_NEW_MSG = "Nuevo mensaje!"
DESP_TEXT_DC_SERVER = "Desconectado del servidor!"

LANG_COMMAND_ENTER = "&Seleccionar"
LANG_LABEL_SELLANG = "Elige tu idioma:"
End Sub

Public Sub SetLangSwedish()
' MDI form ..
MDI_COMMAND_CONFIG = "&Inställningar"
MDI_COMMAND_CHAT = "Ch&att"
MDI_COMMAND_SENDFILE = "&Sänd fil"
MDI_COMMAND_SOCIETY = "&Online Lista"

MDI_STAT_DISCONNECTED = "Status: Frånkopplad"
MDI_STAT_DCFROMSERVER = "Status: Koppla ifrån servern"
MDI_STAT_CONNECTED = "Status: Anslut"
MDI_STAT_CONNECTION_ERROR = "Status: Avkopplad på grund av anslutningsproblem"
MDI_STAT_CONNECTING = "Status: Ansluter .."

'MDImsgbox_errorHandlerFormLoad
MDI_MSG_CONFIG_NOTIFY = "Några Konfiguration filer är gamla eller skadade, Peach hittade problemet och det kommer bli reparerat nästa gång du kör programmet."
MDI_MSG_NAME_TAKEN = "Namnet är upptaget."
MDI_MSG_WRONG_ACCOUNT = "The account does not exist or is wrong."
MDI_MSG_WRONG_PASSWORD = "The password is wrong."
MDI_MSG_BANNED = "This account is banned."

' Config form
CONFIG_COMMAND_CONNECT = "&Anslut"
CONFIG_COMMAND_DISCONNECT = "&Frånkoppla"
CONFIG_COMMAND_SETTINGS = "&Settings"
CONFIG_COMMAND_UPDATE = "&Update"
CONFIG_COMMAND_REGISTER = "&Register Account"

CONFIG_CHECK_SAVE_PASSWORD = "&Save Password"

CONFIG_FRAME_CONNECTION = "Connection Information"

LANG_GERMAN = "Tyska"
LANG_ENGLISH = "Engelska"
LANG_SPANISH = "Spanska"
LANG_SWEDISH = "Svenska"
LANG_ITALIAN = "Italienska"
LANG_SERBIAN = "Serbiska"
LANG_DUTCH = "Holländska"
LANG_FRENCH = "Franska"

CONFIG_MSG_ACCOUNT = "You didnt introduce an account."
CONFIG_MSG_PASSWORD = "You didnt introduce an password."
CONFIG_MSG_NUMERIC = "Du kan inte använda siffror i namnet."
CONFIG_MSG_PORT = "Du angav inget portnummer."
CONFIG_MSG_NAME = "Du angav inte ett namn."
CONFIG_MSG_IP = "Du angav inte ett IP."

' Chat form ..
CHAT_COMMAND_SEND = "&Sänd"
CHAT_COMMAND_CLEAR = "&Rensa"

CHAT_TIME_TEXT = " Tiden är "

' Send file form ..
SF_LABEL_FILENAME = " Fil Namn:"
SF_LABEL_SENDING_FILE = "Sänder:"
SF_LABEL_SENT = "0.0% Sänt"
SF_LABEL_SEND_TO = "Send to:"

SF_MSG_USER = "No user selected."
SF_MSG_FILE = "No file selected."
SF_MSG_INCOMMING_FILE = "You are getting an incomming file, do you want to accept?"
SF_MSG_DECILINED = "File transfer was decilined."

SF_COMMAND_BROWSE = "&Sök .."
SF_COMMAND_SENDFILE = "Sänd"

DESP_TEXT_NEW_MSG = "New Message!"
DESP_TEXT_DC_SERVER = "Koppla ifrån servern!"

LANG_COMMAND_ENTER = "&Öppna"
LANG_LABEL_SELLANG = "Välj språk:"
End Sub

Public Sub SetLangItalian()
' Mdi form
MDI_COMMAND_CONFIG = "&Configurazione"
MDI_COMMAND_CHAT = "Ch&at"
MDI_COMMAND_SENDFILE = "&Invia File"
MDI_COMMAND_SOCIETY = "&Lista contatti Online"

MDI_STAT_DISCONNECTED = "Stato: Disconnesso"
MDI_STAT_DCFROMSERVER = "Stato: Disconnesso dal Server"
MDI_STAT_CONNECTED = "Stato: Connesso"
MDI_STAT_CONNECTION_ERROR = "Stato: Disconnesso a causa di problemi di connessione"
MDI_STAT_CONNECTING = "Stato: Connessione .."

'MDImsgbox_errorHandlerFormLoad
MDI_MSG_CONFIG_NOTIFY = "Alcuni file della configurazione potrebbero essere obsoleti o danneggiati, Peach ha riscontrato il problema e lo corregera' al prossimo avvio."
MDI_MSG_NAME_TAKEN = "Il nome immesso e' gia' in uso."
MDI_MSG_WRONG_ACCOUNT = "The account does not exist or is wrong."
MDI_MSG_WRONG_PASSWORD = "The password is wrong."
MDI_MSG_BANNED = "This account is banned."

' Config form ..
CONFIG_COMMAND_CONNECT = "&Connesso"
CONFIG_COMMAND_DISCONNECT = "&Disconnesso"
CONFIG_COMMAND_SETTINGS = "&Settings"
CONFIG_COMMAND_UPDATE = "&Update"
CONFIG_COMMAND_REGISTER = "&Register Account"

CONFIG_CHECK_SAVE_PASSWORD = "&Save Password"

LANG_GERMAN = "Tedesco"
LANG_ENGLISH = "Inglese"
LANG_SPANISH = "Spagnolo"
LANG_SWEDISH = "Svedese"
LANG_ITALIAN = "Italiano"
LANG_SERBIAN = "Serbo"
LANG_DUTCH = "Olandese"
LANG_FRENCH = "Francese"

CONFIG_MSG_ACCOUNT = "You didnt introduce an account."
CONFIG_MSG_PASSWORD = "You didnt introduce an password."
CONFIG_MSG_NUMERIC = "Non puoi immettere nomi composti da numeri."
CONFIG_MSG_PORT = "Non hai selezionato una porta valida."
CONFIG_MSG_NAME = "Non hai immesso un Nome utente."
CONFIG_MSG_IP = "Non hai immesso un IP."

' Chat form ..
CHAT_COMMAND_SEND = "&Invia"
CHAT_COMMAND_CLEAR = "&Clear"

CHAT_TIME_TEXT = " L'ora e' "

' Send file form ..
SF_LABEL_FILENAME = " Nome file:"
SF_LABEL_SENDING_FILE = "Inviando:"
SF_LABEL_SENT = "0.0% Inviato"
SF_LABEL_SEND_TO = "Send to:"

SF_MSG_USER = "No user selected."
SF_MSG_FILE = "No file selected."
SF_MSG_INCOMMING_FILE = "You are getting an incomming file, do you want to accept?"
SF_MSG_DECILINED = "File transfer was decilined."

SF_COMMAND_BROWSE = "&Cerca .."
SF_COMMAND_SENDFILE = "Invia"
SF_COMMAND_CANCEL = "Annulla .."

DESP_TEXT_NEW_MSG = "New Message!"
DESP_TEXT_DC_SERVER = "Disconnesso dal Server!"

LANG_COMMAND_ENTER = "&Apri"
LANG_LABEL_SELLANG = "Seleziona la tua lingua:"
End Sub

Public Sub SetLangSerbian()
' Mdi form ..
MDI_COMMAND_CONFIG = "&Konfiguracija"
MDI_COMMAND_CHAT = "Ch&at"
MDI_COMMAND_SENDFILE = "&Slanje fajla"
MDI_COMMAND_SOCIETY = "&Onlajn lista"

MDI_STAT_DISCONNECTED = "Status: Veza je prekinuta"
MDI_STAT_DCFROMSERVER = "Status: Veza sa serverom je prekinuta"
MDI_STAT_CONNECTED = "Status: Povezi"
MDI_STAT_CONNECTION_ERROR = "Status: Problem sa konekcijom veza je prekinuta "
MDI_STAT_CONNECTING = "Status: Povezi .."

'MDImsgbox_errorHandlerFormLoad
MDI_MSG_CONFIG_NOTIFY = "Datoteka konfigurac. Je zastarela ili ostecena, problem ce biti pronadjen i popravljen sledecim pokretanjem programa."
MDI_MSG_NAME_TAKEN = "Ime je vec zauzeto."
MDI_MSG_WRONG_ACCOUNT = "The account does not exist or is wrong."
MDI_MSG_WRONG_PASSWORD = "The password is wrong."
MDI_MSG_BANNED = "This account is banned."

' Config form ..
CONFIG_COMMAND_CONNECT = "&Povezi se"
CONFIG_COMMAND_DISCONNECT = "&Veza je prekinuta"
CONFIG_COMMAND_SETTINGS = "&Settings"
CONFIG_COMMAND_UPDATE = "&Update"
CONFIG_COMMAND_REGISTER = "&Register Account"

CONFIG_CHECK_SAVE_PASSWORD = "&Save Password"

LANG_GERMAN = "Nemacki"
LANG_ENGLISH = "Engleski"
LANG_SPANISH = "Spanski"
LANG_SWEDISH = "Svedski"
LANG_ITALIAN = "Italijanski"
LANG_SERBIAN = "Srpski"
LANG_DUTCH = "Holandski"
LANG_FRENCH = "Francuski"

CONFIG_MSG_ACCOUNT = "You didnt introduce an account."
CONFIG_MSG_PASSWORD = "You didnt introduce an password."
CONFIG_MSG_NUMERIC = "Ne mozete uzeti numericka imena."
CONFIG_MSG_PORT = "Niste uneli port."
CONFIG_MSG_NAME = "Niste uneli ime"
CONFIG_MSG_IP = "Niste uneli IP"

' Chat form ..
CHAT_COMMAND_SEND = "&Posalji"
CHAT_COMMAND_CLEAR = "&Obrisi"

CHAT_TIME_TEXT = " Vreme je "

' Send file form ..
SF_LABEL_FILENAME = " Ime  arhive:"
SF_LABEL_SENDING_FILE = "Slanje:"
SF_LABEL_SENT = "0.0% Poslato"
SF_LABEL_SEND_TO = "Send to:"

SF_MSG_USER = "No user selected."
SF_MSG_FILE = "No file selected."
SF_MSG_INCOMMING_FILE = "You are getting an incomming file, do you want to accept?"
SF_MSG_DECILINED = "File transfer was decilined."

SF_COMMAND_BROWSE = "Trazi .."
SF_COMMAND_SENDFILE = "Posalji"
SF_COMMAND_CANCEL = "Otkazhi .."

DESP_TEXT_NEW_MSG = "New Message!"
DESP_TEXT_DC_SERVER = "Veza sa serverom je prekinuta!"

LANG_COMMAND_ENTER = "&Otvori"
LANG_LABEL_SELLANG = "Dodaj svoj jezik:"
End Sub

Public Sub SetLangDutch()
MDI_COMMAND_CONFIG = "&Configuratie"
MDI_COMMAND_CHAT = "Ch&at"
MDI_COMMAND_SENDFILE = "&Bestand Verzenden"
MDI_COMMAND_SOCIETY = "&Online List"

MDI_STAT_DISCONNECTED = "Status: Verbinding verbroken"
MDI_STAT_DCFROMSERVER = "Status: Verbinding verbroken met de server"
MDI_STAT_CONNECTED = "Status: Verbonden"
MDI_STAT_CONNECTION_ERROR = "Status: Verbinding verbroken wegens connectie problemen"
MDI_STAT_CONNECTING = "Status: Verbinden .."

'MDImsgbox_errorHandlerFormLoad
MDI_MSG_CONFIG_NOTIFY = "Enkele bestanden zijn oud of beschadigd, Peach heeft het probleem gevonden en zal het herstellen bij de volgende herstart."
MDI_MSG_NAME_TAKEN = "Deze naam is niet beschikbaar."
MDI_MSG_WRONG_ACCOUNT = "The account does not exist or is wrong."
MDI_MSG_WRONG_PASSWORD = "The password is wrong."
MDI_MSG_BANNED = "This account is banned."

CONFIG_COMMAND_CONNECT = "&Verbind"
CONFIG_COMMAND_DISCONNECT = "&Verbreek de verbinding"
CONFIG_COMMAND_SETTINGS = "&Settings"
CONFIG_COMMAND_UPDATE = "&Update"
CONFIG_COMMAND_REGISTER = "&Register Account"

CONFIG_CHECK_SAVE_PASSWORD = "&Save Password"

LANG_GERMAN = "Duits"
LANG_ENGLISH = "Engels"
LANG_SPANISH = "Spaans"
LANG_SWEDISH = "Zweeds"
LANG_ITALIAN = "Italiaans"
LANG_SERBIAN = "Serbisch"
LANG_DUTCH = "Nederlands"
LANG_FRENCH = "Frans"

CONFIG_MSG_ACCOUNT = "You didnt introduce an account."
CONFIG_MSG_PASSWORD = "You didnt introduce an password."
CONFIG_MSG_NUMERIC = "U kan geen naam nemen dat nummers bevat."
CONFIG_MSG_PORT = "U hebt geen poort ingesteld."
CONFIG_MSG_NAME = "U hebt geen naam gegoven."
CONFIG_MSG_IP = "U hebt geen IP gegoven."

CHAT_COMMAND_SEND = "&Zend"
CHAT_COMMAND_CLEAR = "&Leegmaken"

CHAT_TIME_TEXT = " De Tijd is: "

SF_LABEL_FILENAME = " Bestandsnaam:"
SF_LABEL_SENDING_FILE = "verZenden:"
SF_LABEL_SENT = "0.0% verzonden"
SF_LABEL_SEND_TO = "Send to:"

SF_MSG_USER = "No user selected."
SF_MSG_FILE = "No file selected."
SF_MSG_INCOMMING_FILE = "You are getting an incomming file, do you want to accept?"
SF_MSG_DECILINED = "File transfer was decilined."

SF_COMMAND_BROWSE = "&Zoeken .."
SF_COMMAND_SENDFILE = "&Stuur"
SF_COMMAND_CANCEL = "&Annuleren .."

DESP_TEXT_NEW_MSG = "New Message!"
DESP_TEXT_DC_SERVER = "Verbinding verbroken met de server"

LANG_COMMAND_ENTER = "&Openen"
LANG_LABEL_SELLANG = "Selecteer jou taal:"
End Sub

Public Sub SetLangFrench()
MDI_COMMAND_CONFIG = "&Configuration"
MDI_COMMAND_CHAT = "Ch&at"
MDI_COMMAND_SENDFILE = "&Envoi File"
MDI_COMMAND_SOCIETY = "&Liste contact Online"

MDI_STAT_DISCONNECTED = "Etat: Deconnecté"
MDI_STAT_DCFROMSERVER = "Etat: Deconnecté du Server"
MDI_STAT_CONNECTED = "Etat: Connecté"
MDI_STAT_CONNECTION_ERROR = "Etat: Deconnecté à cause de problèmes do connection"
MDI_STAT_CONNECTING = "Etat: Connection .."

MDI_MSG_CONFIG_NOTIFY = "Quelques files de la configuration pourrait etre daumagés ou obsolète , Peach a trouvé le problème et le corrigerà au prochain envoi."
MDI_MSG_NAME_TAKEN = "Le nom inséré est déjà utilizé."
MDI_MSG_WRONG_ACCOUNT = "The account does not exist or is wrong."
MDI_MSG_WRONG_PASSWORD = "The password is wrong."
MDI_MSG_BANNED = "This account is banned."

CONFIG_COMMAND_CONNECT = "&Connecté"
CONFIG_COMMAND_DISCONNECT = "&Deconnecté"
CONFIG_COMMAND_SETTINGS = "&Settings"
CONFIG_COMMAND_UPDATE = "&Update"
CONFIG_COMMAND_REGISTER = "&Register Account"

CONFIG_CHECK_SAVE_PASSWORD = "&Save Password"

LANG_GERMAN = "Alleman"
LANG_ENGLISH = "Anglais"
LANG_SPANISH = "Espagnol"
LANG_SWEDISH = "Suédois"
LANG_ITALIAN = "Italien"
LANG_SERBIAN = "Serbois"
LANG_DUTCH = "Hollandais"
LANG_FRENCH = "Français"

CONFIG_MSG_ACCOUNT = "You didnt introduce an account."
CONFIG_MSG_PASSWORD = "You didnt introduce an password."
CONFIG_MSG_NUMERIC = "Tu ne peut pas insérer noms composé de numeros."
CONFIG_MSG_PORT = "Tu n'as pas selectionner une porte valide."
CONFIG_MSG_NAME = "Tu n'as pas innecté un Nom utilizateur."
CONFIG_MSG_IP = "Tu n'as pas innecté un IP."

CHAT_COMMAND_SEND = "&Envoi"
CHAT_COMMAND_CLEAR = "&Clear"

CHAT_TIME_TEXT = " L'heure est "

SF_LABEL_FILENAME = " Nom file:"
SF_LABEL_SENDING_FILE = "Envoyant:"
SF_LABEL_SENT = "0.0% Envoyé"
SF_LABEL_SEND_TO = "Send to:"

SF_MSG_USER = "No user selected."
SF_MSG_FILE = "No file selected."
SF_MSG_INCOMMING_FILE = "You are getting an incomming file, do you want to accept?"
SF_MSG_DECILINED = "File transfer was decilined."

SF_COMMAND_BROWSE = "&Cherche .."
SF_COMMAND_SENDFILE = "Envoi"
SF_COMMAND_CANCEL = "Annuler .."

DESP_TEXT_NEW_MSG = "New Message!"
DESP_TEXT_DC_SERVER = "Deconnecté du Server"

LANG_COMMAND_ENTER = "&Ouvrir"
LANG_LABEL_SELLANG = "Choisissez votre langue:"
End Sub
