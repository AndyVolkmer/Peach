Attribute VB_Name = "modLang"
Option Explicit

Public CURRENT_LANG                     As Long
'Start variable support for languages
' MDI form ..
Public MDI_COMMAND_CHAT                 As String
Public MDI_COMMAND_SENDFILE             As String
Public MDI_COMMAND_SOCIETY              As String

Public MDI_STAT_DISCONNECTED            As String
Public MDI_STAT_DISCONNECT              As String
Public MDI_STAT_CONNECTED               As String
Public MDI_STAT_CONNECTION_ERROR        As String
Public MDI_STAT_CONNECTING              As String

Public MDI_MSG_ERROR_FORM_LOAD          As String
Public MDI_MSG_NAME_TAKEN               As String
Public MDI_MSG_WRONG_ACCOUNT            As String
Public MDI_MSG_WRONG_PASSWORD           As String
Public MDI_MSG_BANNED                   As String
Public MDI_MSG_UNLOAD                   As String

' Configuration form ..
Public CONFIG_COMMAND_CONNECT           As String
Public CONFIG_COMMAND_DISCONNECT        As String
Public CONFIG_COMMAND_SETTINGS          As String
Public CONFIG_COMMAND_UPDATE            As String
Public CONFIG_COMMAND_REGISTER          As String
Public CONFIG_COMMAND_FORGOT_PASSWORD   As String

Public CONFIG_CHECK_SAVE_PASSWORD       As String

Public CONFIG_FRAME_CONNECTION          As String

Public CONFIG_MSG_ACCOUNT               As String
Public CONFIG_MSG_PASSWORD              As String
Public CONFIG_MSG_NUMERIC               As String
Public CONFIG_MSG_PORT                  As String
Public CONFIG_MSG_IP                    As String
Public CONFIG_MSG_NAME                  As String
Public CONFIG_MSG_NAME_SHORT            As String
Public CONFIG_MSG_NAME_INVALID          As String
Public CONFIG_MSG_UPDATE_FILE           As String

Public CHAT_COMMAND_SEND                As String
Public CHAT_COMMAND_CLEAR               As String

Public SF_LABEL_FILENAME                As String
Public SF_LABEL_SENDING_FILE            As String
Public SF_LABEL_SENT                    As String
Public SF_LABEL_SEND_TO                 As String

Public SF_COMMAND_BROWSE                As String
Public SF_COMMAND_SENDFILE              As String
Public SF_COMMAND_CANCEL                As String

Public SF_MSG_USER                      As String
Public SF_MSG_FILE                      As String
Public SF_MSG_INCOMMING_FILE_1          As String
Public SF_MSG_INCOMMING_FILE_2          As String
Public SF_MSG_INCOMMING_FILE_3          As String
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

Public SOC_FRIEND_LIST                  As String
Public SOC_ONLINE_LIST

Public SOC_COMMAND_ADD                  As String
Public SOC_COMMAND_REMOVE               As String

Public SOC_ASK_DEL_1                    As String
Public SOC_ASK_DEL_2                    As String

'Register account form
Public REG_CAPTION                      As String

Public REG_FRAME_DETAIL                 As String

Public REG_LABEL_ACCOUNT_NAME           As String
Public REG_LABEL_PASSWORD               As String
Public REG_LABEL_PASSWORD_CONFIRM       As String
Public REG_LABEL_PASSWORD_WEAK          As String
Public REG_LABEL_PASSWORD_NORMAL        As String
Public REG_LABEL_PASSWORD_STRONG        As String
Public REG_LABEL_ERROR                  As String
Public REG_LABEL_SECRET_QUESTION        As String
Public REG_LABEL_SECRET_ANSWER          As String

Public REG_CHECK_PASSWORD_SHOW          As String

Public REG_COMMAND_SUBMIT               As String
Public REG_COMMAND_CLOSE                As String

Public REG_MSG_ACCOUNT_EXIST            As String
Public REG_MSG_ACCOUNT_INVALID          As String
Public REG_MSG_ACCOUNT_NUMERIC          As String
Public REG_MSG_ACCOUNT_EMPTY            As String
Public REG_MSG_ACCOUNT_SHORT            As String

Public REG_MSG_SUCCESSFULLY             As String
Public REG_MSG_ERROR                    As String
Public REG_MSG_ERROR_OCCURED            As String
Public REG_MSG_LOADING                  As String
Public REG_MSG_CONNECTION_BROKEN        As String
Public REG_MSG_PASSWORD_MATCH           As String
Public REG_MSG_PASSWORD_SHORT           As String
Public REG_MSG_PASSWORD_EMPTY           As String
Public REG_MSG_SECRET_ANSWER_EMPTY      As String

Public REG_CMB_SECRET_QUESTION_0        As String
Public REG_CMB_SECRET_QUESTION_1        As String
Public REG_CMB_SECRET_QUESTION_2        As String
Public REG_CMB_SECRET_QUESTION_3        As String
Public REG_CMB_SECRET_QUESTION_4        As String
Public REG_CMB_SECRET_QUESTION_5        As String

'Settings form
Public SET_LABEL_COLOR                  As String

Public SET_FRAME_OPTIONS                As String
Public SET_FRAME_CONNECTION             As String

Public SET_CHECK_SAVE_ACCOUNT           As String
Public SET_CHECK_SAVE_PASSWORD          As String
Public SET_CHECK_ASK_CLOSING            As String
Public SET_CHECK_MINIMIZE               As String

Public SET_COMMAND_LANGUAGE             As String
Public SET_COMMAND_SAVE                 As String

Public SF2_COMMAND_OPEN_FILE            As String

Public FP_FRAME_FORGOT_PASSWORD         As String
Public FP_LABEL_ACCOUNT                 As String
Public FP_LABEL_SECRET_QUESTION         As String
Public FP_LABEL_SECRET_ANSWER           As String
Public FP_COMMAND_REQUEST               As String
Public FP_CAPTION                       As String

Public FP_MSG_SUCCESSFULL               As String
Public FP_MSG_WRONG_ANSWER              As String

Public Sub SET_LANG_GERMAN()
CURRENT_LANG = 0

' MDI form ..
MDI_COMMAND_CHAT = "Ch&at"
MDI_COMMAND_SENDFILE = "&Sende Datei"
MDI_COMMAND_SOCIETY = "&Online Liste"

MDI_STAT_DISCONNECTED = "Status: Getrennt"
MDI_STAT_DISCONNECT = "Status: Getrennt vom Server"
MDI_STAT_CONNECTED = "Status: Verbunden"
MDI_STAT_CONNECTION_ERROR = "Status: Getrennt aufgrund eines Verbindungsfehlers"
MDI_STAT_CONNECTING = "Status: Verbindung wird aufgebaut .."

MDI_MSG_NAME_TAKEN = "Der Name ist bereits vergeben."
MDI_MSG_WRONG_ACCOUNT = "Der Account ist nicht vorhanden oder falsch."
MDI_MSG_WRONG_PASSWORD = "Das Passwort ist falsch."
MDI_MSG_BANNED = "Dieser Account wurde gebannt."
MDI_MSG_UNLOAD = "Sind Sie sicher, dass Sie Peach schliessen wollen?"

CONFIG_COMMAND_CONNECT = "&Verbinden"
CONFIG_COMMAND_DISCONNECT = "&Verbindung trenn."
CONFIG_COMMAND_SETTINGS = "&Einstellungen"
CONFIG_COMMAND_UPDATE = "&Aktualisieren"
CONFIG_COMMAND_REGISTER = "&Account registrieren"
CONFIG_COMMAND_FORGOT_PASSWORD = "&Password vergessen"

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
CONFIG_MSG_NAME_SHORT = "Du hast einen zu kurzen Namen eingegeben."
CONFIG_MSG_NAME_INVALID = "Du hast einen ungültigen Namen eingegeben."
CONFIG_MSG_UPDATE_FILE = "Sie brauchen den Peach Updater um ihr Peach zu updaten." & vbCrLf & vbCrLf & "Sie können es hier downloaden:  http://riplegion.ri.funpic.de/Peach/peachUpdater.exe"

CHAT_COMMAND_SEND = "&Senden"
CHAT_COMMAND_CLEAR = "&Löschen"

SF_LABEL_FILENAME = " Datei Name:"
SF_LABEL_SENDING_FILE = "Sende:"
SF_LABEL_SENT = "0.0% Gesendet"
SF_LABEL_SEND_TO = "Sende an:"

SF_MSG_USER = "Kein Benutzer ausgewählt."
SF_MSG_FILE = "Keine Datei ausgewählt."
SF_MSG_INCOMMING_FILE_1 = "Du empfängst gerade '"
SF_MSG_INCOMMING_FILE_2 = "' von "
SF_MSG_INCOMMING_FILE_3 = ". Willst du die Datei annehmen?"
SF_MSG_DECILINED = "Der Benutzer hat die Datei abgelehnt."

SF_COMMAND_BROWSE = "&Suchen .."
SF_COMMAND_SENDFILE = "Senden"
SF_COMMAND_CANCEL = "Abbrechen .."

DESP_TEXT_NEW_MSG = "Neue Nachricht!"
DESP_TEXT_DC_SERVER = "Verbindung unterbrochen!"

LANG_COMMAND_ENTER = "&Auswählen"
LANG_LABEL_SELLANG = "Wähle deine Sprache aus:"

SOC_FRIEND_LIST = "Freundes Liste"
SOC_ONLINE_LIST = "Online Liste"

SOC_COMMAND_ADD = "&Hinzufügen"
SOC_COMMAND_REMOVE = "&Entfernen"

SOC_ASK_DEL_1 = "Möchten Sie '"
SOC_ASK_DEL_2 = "' von ihrer Freundesliste löschen?"

REG_CAPTION = "Peach - Registration"

REG_FRAME_DETAIL = "Gebe deine Daten an"

REG_LABEL_ACCOUNT_NAME = " Benutzer Name:"
REG_LABEL_PASSWORD = " Passwort:"
REG_LABEL_PASSWORD_CONFIRM = " Passwort bestätigen:"
REG_LABEL_PASSWORD_WEAK = "Das Passwort ist schwach."
REG_LABEL_PASSWORD_NORMAL = "Das Passwort ist gut."
REG_LABEL_PASSWORD_STRONG = "Das Passwort ist stark."
REG_LABEL_SECRET_QUESTION = " Geheime Frage:"
REG_LABEL_SECRET_ANSWER = " Geheime Antwort:"

REG_COMMAND_SUBMIT = "&Registrieren"
REG_COMMAND_CLOSE = "&Schliessen"

REG_CHECK_PASSWORD_SHOW = "&Passwort anzeigen"

REG_MSG_ACCOUNT_EXIST = "Der Account Name ist bereits vergeben."
REG_MSG_ACCOUNT_INVALID = "Ungültiger Account Name."
REG_MSG_ACCOUNT_NUMERIC = "Der Account Name darf nicht aus ziffern bestehen."
REG_MSG_ACCOUNT_EMPTY = "Kein Account angegeben."
REG_MSG_ACCOUNT_SHORT = "Der Account Name ist zu kurz, muss aus wenigstens 4 Zeichen bestehen."

REG_MSG_PASSWORD_MATCH = "Die Passwörter stimmen nicht überein."
REG_MSG_PASSWORD_SHORT = "Das Passwort ist zu kurz, muss aus wenigstens 6 Zeichen bestehen."
REG_MSG_PASSWORD_EMPTY = "Kein Passwort angegeben."

REG_MSG_SECRET_ANSWER_EMPTY = "Keine geheime Antwort angegeben."

REG_MSG_SUCCESSFULLY = "Der Account wurde erfolgreich erstellt."
REG_MSG_ERROR = "Ein Fehler ist aufgetreten bitte versuchen sie es später nochmal."
REG_MSG_ERROR_OCCURED = "Fehler aufgetreten ..."
REG_MSG_LOADING = " Lädt .."
REG_MSG_CONNECTION_BROKEN = "Die Verbindung wurde unterbrochen bitte versuchen sie es später nochmal."

REG_CMB_SECRET_QUESTION_0 = "Wie heißt dein Haustier?"
REG_CMB_SECRET_QUESTION_1 = "Dein Lieblings-Buch?"
REG_CMB_SECRET_QUESTION_2 = "Dein Lieblings-Film?"
REG_CMB_SECRET_QUESTION_3 = "Dein Lieblings-Spiel?"
REG_CMB_SECRET_QUESTION_4 = "Dein Lieblings-Sänger?"
REG_CMB_SECRET_QUESTION_5 = "Geburtsort deiner mutter?"

SET_LABEL_COLOR = "Jetzige Farbe:"

SET_FRAME_OPTIONS = "Optionen"
SET_FRAME_CONNECTION = "Verbindungs Einstellungen"

SET_CHECK_SAVE_ACCOUNT = "Account speichern"
SET_CHECK_SAVE_PASSWORD = "Passwort speichern"
SET_CHECK_ASK_CLOSING = "Abfragen bevor schliessen"
SET_CHECK_MINIMIZE = "Peach-Fenster in die Taskleiste minimieren"

SET_COMMAND_LANGUAGE = "&Sprache"
SET_COMMAND_SAVE = "&Speichern"

SF2_COMMAND_OPEN_FILE = "&Datei Ordner öffnen"

FP_FRAME_FORGOT_PASSWORD = "Password vergessen"
FP_LABEL_ACCOUNT = " Gebe deinen Account ein:"
FP_LABEL_SECRET_QUESTION = " Geheime Frage:"
FP_LABEL_SECRET_ANSWER = " Geheime Antwort:"
FP_COMMAND_REQUEST = "&Abfragen"
FP_CAPTION = "Peach - Password vergessen"

FP_MSG_SUCCESSFULL = "Ihr Passwort lautet "
FP_MSG_WRONG_ANSWER = "Die Antwort ist falsch."
End Sub

Public Sub SET_LANG_ENGLISH()
CURRENT_LANG = 1

'MDI form
MDI_COMMAND_CHAT = "Ch&at"
MDI_COMMAND_SENDFILE = "&Send File"
MDI_COMMAND_SOCIETY = "&Online List"

MDI_STAT_DISCONNECTED = "Status: Disconnected"
MDI_STAT_DISCONNECT = "Status: Disconnected from Server"
MDI_STAT_CONNECTED = "Status: Connected"
MDI_STAT_CONNECTION_ERROR = "Status: Can't connect to server."
MDI_STAT_CONNECTING = "Status: Connecting .."

'MDImsgbox_errorHandlerFormLoad
MDI_MSG_NAME_TAKEN = "This name is already taken."
MDI_MSG_WRONG_ACCOUNT = "The account does not exist or is wrong."
MDI_MSG_WRONG_PASSWORD = "The password is wrong."
MDI_MSG_BANNED = "This account is banned."
MDI_MSG_UNLOAD = "Are you sure you want to close Peach?"

' Configuration form ..
CONFIG_COMMAND_CONNECT = "&Connect"
CONFIG_COMMAND_DISCONNECT = "&Disconnect"
CONFIG_COMMAND_SETTINGS = "&Settings"
CONFIG_COMMAND_UPDATE = "&Update"
CONFIG_COMMAND_REGISTER = "&Register Account"
CONFIG_COMMAND_FORGOT_PASSWORD = "&Forgot Password"

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

CONFIG_MSG_ACCOUNT = "You didn't enter an account."
CONFIG_MSG_PASSWORD = "You didn't enter a password."
CONFIG_MSG_NUMERIC = "You can't take numeric names."
CONFIG_MSG_PORT = "You didn't introduce a port."
CONFIG_MSG_NAME = "You didn't introduce a name."
CONFIG_MSG_IP = "You didn't introduce a IP."
CONFIG_MSG_NAME_SHORT = "The name you introduced is too short."
CONFIG_MSG_NAME_INVALID = "The name you introduced is invalid."
CONFIG_MSG_UPDATE_FILE = "You need the peach updater to be able to upgrade your peach." & vbCrLf & vbCrLf & "Download it here: http://riplegion.ri.funpic.de/Peach/peachUpdater.exe"

' Chat form ..
CHAT_COMMAND_SEND = "&Send"
CHAT_COMMAND_CLEAR = "&Clear"

' Send File form ..
SF_LABEL_FILENAME = " File Name:"
SF_LABEL_SENDING_FILE = "Sending:"
SF_LABEL_SENT = "0.0% Sent"
SF_LABEL_SEND_TO = "Send to:"

SF_MSG_USER = "No user selected."
SF_MSG_FILE = "No file selected."
SF_MSG_INCOMMING_FILE_1 = "You are receiving '"
SF_MSG_INCOMMING_FILE_2 = "' from "
SF_MSG_INCOMMING_FILE_3 = ". Do you want to accept?"
SF_MSG_DECILINED = "File transfer was decilined."

SF_COMMAND_BROWSE = "&Search .."
SF_COMMAND_SENDFILE = "Send"
SF_COMMAND_CANCEL = "Cancel .."

DESP_TEXT_NEW_MSG = "New Message!"
DESP_TEXT_DC_SERVER = "Disconnected from Server!"

LANG_COMMAND_ENTER = "&Select"
LANG_LABEL_SELLANG = "Select your language:"

SOC_FRIEND_LIST = "Friend List"
SOC_ONLINE_LIST = "Online List"

SOC_COMMAND_ADD = "&Add"
SOC_COMMAND_REMOVE = "&Remove"

SOC_ASK_DEL_1 = "Do you want to delete '"
SOC_ASK_DEL_2 = "' from your friendlist?"

REG_CAPTION = "Peach - Registration"

REG_FRAME_DETAIL = "Enter your details"

REG_LABEL_ACCOUNT_NAME = " Account Name:"
REG_LABEL_PASSWORD = " Password:"
REG_LABEL_PASSWORD_CONFIRM = " Confirm the Password:"
REG_LABEL_PASSWORD_WEAK = "The password is weak."
REG_LABEL_PASSWORD_NORMAL = "The password is normal."
REG_LABEL_PASSWORD_STRONG = "The password is strong."
REG_LABEL_SECRET_QUESTION = "Secret question:"
REG_LABEL_SECRET_ANSWER = "Secret answer:"

REG_COMMAND_SUBMIT = "&Submit"
REG_COMMAND_CLOSE = "&Close"

REG_CHECK_PASSWORD_SHOW = "&Show Password."

REG_MSG_ACCOUNT_EXIST = "The account name already exists."
REG_MSG_ACCOUNT_INVALID = "Invalid account name."
REG_MSG_ACCOUNT_NUMERIC = "Account can't be made of numeric characters."
REG_MSG_ACCOUNT_EMPTY = "No account entered."
REG_MSG_ACCOUNT_SHORT = "Account name to short, it requieres at least 4 characters."

REG_MSG_PASSWORD_MATCH = "The passwords dont match."
REG_MSG_PASSWORD_SHORT = "Password to short, it requieres at least 6 characters."
REG_MSG_PASSWORD_EMPTY = "No Password entered."

REG_MSG_SECRET_ANSWER_EMPTY = "No secret answered entered."

REG_MSG_SUCCESSFULLY = "The account was successfully registered."
REG_MSG_ERROR = "An error has occured please try later again."
REG_MSG_ERROR_OCCURED = "Error has occured ..."
REG_MSG_LOADING = " Loading .."
REG_MSG_CONNECTION_BROKEN = "Connection is broken please try again later."

REG_CMB_SECRET_QUESTION_0 = "What is the name of your pet?"
REG_CMB_SECRET_QUESTION_1 = "Your favorite book?"
REG_CMB_SECRET_QUESTION_2 = "Your favorite movie?"
REG_CMB_SECRET_QUESTION_3 = "Your favorite game?"
REG_CMB_SECRET_QUESTION_4 = "Your favorite singer?"
REG_CMB_SECRET_QUESTION_5 = "The place where your mother was born?"

SET_LABEL_COLOR = "Current Color:"

SET_FRAME_OPTIONS = "Options"
SET_FRAME_CONNECTION = "Connection Settings"

SET_CHECK_SAVE_ACCOUNT = "Save Account"
SET_CHECK_SAVE_PASSWORD = "Save Password"
SET_CHECK_ASK_CLOSING = "Ask before closing"
SET_CHECK_MINIMIZE = "Minimize Peach window to system tray"

SET_COMMAND_LANGUAGE = "&Language"
SET_COMMAND_SAVE = "&Save"

SF2_COMMAND_OPEN_FILE = "&Open File Folder"

FP_FRAME_FORGOT_PASSWORD = "Forgot Passwort"
FP_LABEL_ACCOUNT = " Enter your account name:"
FP_LABEL_SECRET_QUESTION = " Secret Question:"
FP_LABEL_SECRET_ANSWER = " Secret Answer:"
FP_COMMAND_REQUEST = "&Request"
FP_CAPTION = "Peach - Forgot Password"

FP_MSG_SUCCESSFULL = "Your password is "
FP_MSG_WRONG_ANSWER = "The answer is wrong."
End Sub

Public Sub SET_LANG_SPANISH()
CURRENT_LANG = 2

MDI_COMMAND_CHAT = "Ch&at"
MDI_COMMAND_SENDFILE = "&Enviar Archivo"
MDI_COMMAND_SOCIETY = "&Lista Online"

MDI_STAT_DISCONNECTED = "Estado: Desconectado"
MDI_STAT_DISCONNECT = "Estado: Desconectado del servidor"
MDI_STAT_CONNECTED = "Estado: Disponible"
MDI_STAT_CONNECTION_ERROR = "Estado: Desconectado por problemas de conexión"
MDI_STAT_CONNECTING = "Estado: Conectando .."

MDI_MSG_NAME_TAKEN = "Este nombre ya esta cogido."
MDI_MSG_WRONG_ACCOUNT = "La cuenta no existe o es incorrecta."
MDI_MSG_WRONG_PASSWORD = "La contraseña es incorrecta."
MDI_MSG_BANNED = "Esta cuenta esta baneada."
MDI_MSG_UNLOAD = "¿Esta seguro que quiere cerrar a Peach?"

CONFIG_COMMAND_CONNECT = "&Conectar"
CONFIG_COMMAND_DISCONNECT = "&Desconectar"
CONFIG_COMMAND_SETTINGS = "&Ajustes"
CONFIG_COMMAND_UPDATE = "&Actualizar"
CONFIG_COMMAND_REGISTER = "&Registrar cuenta"
CONFIG_COMMAND_FORGOT_PASSWORD = "¿Ha olvidado contraseña?"

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

CONFIG_MSG_ACCOUNT = "You did'nt introduce an account."
CONFIG_MSG_PASSWORD = "You did'nt introduce an password."
CONFIG_MSG_NUMERIC = "No puedes coger nombres con numeros."
CONFIG_MSG_PORT = "No has introducido un puerto."
CONFIG_MSG_NAME = "No has introducido un nombre."
CONFIG_MSG_IP = "No has introducido una direccion."
CONFIG_MSG_NAME_SHORT = "El nombre que has introducido es corto."
CONFIG_MSG_NAME_INVALID = "El nombre que has introducido es invalido."
CONFIG_MSG_UPDATE_FILE = "Necesitas el Peach Updater para actualizar tu Peach." & vbCrLf & vbCrLf & "Descargalo aqui http://riplegion.ri.funpic.de/Peach/peachUpdater.exe"

CHAT_COMMAND_SEND = "&Enviar"
CHAT_COMMAND_CLEAR = "&Limpiar"

SF_LABEL_FILENAME = " Nombre del archivo:"
SF_LABEL_SENDING_FILE = "Enviando:"
SF_LABEL_SENT = "0.0% Enviado"
SF_LABEL_SEND_TO = "Enviar a:"

SF_MSG_USER = "No has seleccionado a una persona."
SF_MSG_FILE = "No has seleccionado a un archivo."
SF_MSG_INCOMMING_FILE_1 = "Esta recibiendo '"
SF_MSG_INCOMMING_FILE_2 = "' de "
SF_MSG_INCOMMING_FILE_3 = ". ¿Quieres aceptar?"
SF_MSG_DECILINED = "El envio ha sido rechazado."

SF_COMMAND_BROWSE = "&Buscar .."
SF_COMMAND_SENDFILE = "Enviar"
SF_COMMAND_CANCEL = "Cancelar .."

DESP_TEXT_NEW_MSG = "Nuevo mensaje!"
DESP_TEXT_DC_SERVER = "Desconectado del servidor!"

LANG_COMMAND_ENTER = "&Seleccionar"
LANG_LABEL_SELLANG = "Elige tu idioma:"

SOC_FRIEND_LIST = "Lista de contactos"
SOC_ONLINE_LIST = "Lista de online"

SOC_COMMAND_ADD = "&Añadir"
SOC_COMMAND_REMOVE = "&Quitar"

SOC_ASK_DEL_1 = "¿Estas seguro que quieres borrar a '"
SOC_ASK_DEL_2 = "' de tu lista de amigos?"

REG_CAPTION = "Peach - Registración"

REG_FRAME_DETAIL = "Enter your details"

REG_LABEL_ACCOUNT_NAME = " Nombre de cuenta:"
REG_LABEL_PASSWORD = " Contraseña:"
REG_LABEL_PASSWORD_CONFIRM = " Confirmar contraseña:"
REG_LABEL_PASSWORD_WEAK = "La contraseña es floja."
REG_LABEL_PASSWORD_NORMAL = "La contraseña es normal."
REG_LABEL_PASSWORD_STRONG = "La contraseña es fuerte."
REG_LABEL_SECRET_QUESTION = "Pregunta secreta:"
REG_LABEL_SECRET_ANSWER = "Respuesta secreta:"

REG_COMMAND_SUBMIT = "&Registrar"
REG_COMMAND_CLOSE = "&Cerrar"

REG_CHECK_PASSWORD_SHOW = "&Ver contraseña"

REG_MSG_ACCOUNT_EXIST = "El nombre de la cuenta ya existe."
REG_MSG_ACCOUNT_INVALID = "El nombre de la cuenta es invalido."
REG_MSG_ACCOUNT_NUMERIC = "El nombre de la cuenta no puede ser numerico."
REG_MSG_ACCOUNT_EMPTY = "No ha introducido un nombre de cuenta."
REG_MSG_ACCOUNT_SHORT = "Nombre de cuenta corto, debe que tener por lo menos 4 digitos."

REG_MSG_PASSWORD_MATCH = "Las contraseñas no son las mismas."
REG_MSG_PASSWORD_SHORT = "Contraseña corta, debe que tener por lo menos 6 digitos."
REG_MSG_PASSWORD_EMPTY = "No ha introducido una contraseña."

REG_MSG_SECRET_ANSWER_EMPTY = "No ha introducido una respuesta secreta."

REG_MSG_SUCCESSFULLY = "La cuenta ha sido registrada con exito."
REG_MSG_ERROR = "Un error ha occurido intenten de nuevo despues."
REG_MSG_ERROR_OCCURED = "Error occurido ..."
REG_MSG_LOADING = " Cargando .."
REG_MSG_CONNECTION_BROKEN = "La conexión se ha roto, intenten de nuevo despues."

REG_CMB_SECRET_QUESTION_0 = "¿Cual es el nombre de tu mascota?"
REG_CMB_SECRET_QUESTION_1 = "¿Tu libro favorito?"
REG_CMB_SECRET_QUESTION_2 = "¿Tu pelicula favorita?"
REG_CMB_SECRET_QUESTION_3 = "¿Tu juego favorito?"
REG_CMB_SECRET_QUESTION_4 = "¿Tu cantante favorito?"
REG_CMB_SECRET_QUESTION_5 = "¿El lugar de nacimiento de tu madre?"

SET_LABEL_COLOR = "Color activo:"

SET_FRAME_OPTIONS = "Opciones"
SET_FRAME_CONNECTION = "Confgiuración de conexión"

SET_CHECK_SAVE_ACCOUNT = "Guardar cuenta"
SET_CHECK_SAVE_PASSWORD = "Guardar contraseña"
SET_CHECK_ASK_CLOSING = "Preguntar antes de cerrar"
SET_CHECK_MINIMIZE = "Minimizar ventana de Peach en la bandeja del sistema"

SET_COMMAND_LANGUAGE = "&Idioma"
SET_COMMAND_SAVE = "&Guardar"

SF2_COMMAND_OPEN_FILE = "&Abrir carpeta"

FP_FRAME_FORGOT_PASSWORD = "¿Ha olvidado contraseña?"
FP_LABEL_ACCOUNT = " Introduce su nombre de cuenta:"
FP_LABEL_SECRET_QUESTION = " Pregunta secreta:"
FP_LABEL_SECRET_ANSWER = " Respuesta secreta:"
FP_COMMAND_REQUEST = "&Solicitar"
FP_CAPTION = "Peach - Recuperar contraseña"

FP_MSG_SUCCESSFULL = "Tu contraseña es "
FP_MSG_WRONG_ANSWER = "La respuesta es incorrecta."
End Sub

Public Sub SET_LANG_SWEDISH()
CURRENT_LANG = 3

' MDI form ..
MDI_COMMAND_CHAT = "Ch&att"
MDI_COMMAND_SENDFILE = "&Sänd fil"
MDI_COMMAND_SOCIETY = "&Online Lista"

MDI_STAT_DISCONNECTED = "Status: Frånkopplad"
MDI_STAT_DISCONNECT = "Status: Koppla ifrån servern"
MDI_STAT_CONNECTED = "Status: Anslut"
MDI_STAT_CONNECTION_ERROR = "Status: Avkopplad på grund av anslutningsproblem"
MDI_STAT_CONNECTING = "Status: Ansluter .."

'MDImsgbox_errorHandlerFormLoad
MDI_MSG_NAME_TAKEN = "Namnet är upptaget."
MDI_MSG_WRONG_ACCOUNT = "The account does not exist or is wrong."
MDI_MSG_WRONG_PASSWORD = "The password is wrong."
MDI_MSG_BANNED = "This account is banned."
MDI_MSG_UNLOAD = "Are you sure you want to close Peach?"

' Config form
CONFIG_COMMAND_CONNECT = "&Anslut"
CONFIG_COMMAND_DISCONNECT = "&Frånkoppla"
CONFIG_COMMAND_SETTINGS = "&Inställningar"
CONFIG_COMMAND_REGISTER = "&Skapa konto"
CONFIG_COMMAND_UPDATE = "&Updatering"
CONFIG_COMMAND_FORGOT_PASSWORD = "Forgot Password"

LANG_GERMAN = "Tyska"
LANG_ENGLISH = "Engelska"
LANG_SPANISH = "Spanska"
LANG_SWEDISH = "Svenska"
LANG_ITALIAN = "Italienska"
LANG_SERBIAN = "Serbiska"
LANG_DUTCH = "Holländska"
LANG_FRENCH = "Franska"

CONFIG_MSG_ACCOUNT = "Du skrev inte in en användare."
CONFIG_MSG_PASSWORD = "Du skrev inte in ett lösenord."
CONFIG_MSG_NUMERIC = "Du kan inte använda siffror i namnet."
CONFIG_MSG_PORT = "Du angav inget portnummer."
CONFIG_MSG_IP = "Du angav inte ett IP."
CONFIG_MSG_NAME = "Du angav inte ett namn."
CONFIG_MSG_NAME_SHORT = "The name you introduced is too short."
CONFIG_MSG_NAME_INVALID = "The name you introduced is invalid."
CONFIG_MSG_UPDATE_FILE = "You need the peach updater to be able to upgrade your peach." & vbCrLf & vbCrLf & "Download it here: http://riplegion.ri.funpic.de/Peach/peachUpdater.exe"

' Chat form ..
CHAT_COMMAND_SEND = "&Sänd"
CHAT_COMMAND_CLEAR = "&Rensa"

' Send file form ..
SF_LABEL_FILENAME = " Fil Namn:"
SF_LABEL_SENDING_FILE = "Sänder:"
SF_LABEL_SENT = "0.0% Sänt"
SF_LABEL_SEND_TO = "Skicka till:"

SF_MSG_USER = "Ingen användare vald."
SF_MSG_FILE = "Ingen fil vald."
SF_MSG_INCOMMING_FILE_1 = "You are receiving '"
SF_MSG_INCOMMING_FILE_2 = "' from "
SF_MSG_INCOMMING_FILE_3 = ". Do you want to accept?"
SF_MSG_DECILINED = "Filöverföringen var nekad."

SF_COMMAND_BROWSE = "&Sök .."
SF_COMMAND_SENDFILE = "Sänd"

SF2_COMMAND_OPEN_FILE = "&Öppna fil map"

DESP_TEXT_NEW_MSG = "Nytt meddelande!"
DESP_TEXT_DC_SERVER = "Koppla ifrån servern!"

LANG_COMMAND_ENTER = "&Öppna"
LANG_LABEL_SELLANG = "Välj språk:"

SOC_FRIEND_LIST = "Kompis Lista"
SOC_ONLINE_LIST = "Online Lista"

SOC_COMMAND_ADD = "&Tillägg"
SOC_COMMAND_REMOVE = "&Ta bort"

SOC_ASK_DEL_1 = "Do you want to delete '"
SOC_ASK_DEL_2 = "' from your friendlist?"

REG_CAPTION = "Peach - Registration"

REG_FRAME_DETAIL = "Ange dina detaljer"

REG_LABEL_ACCOUNT_NAME = " Användar Namn:"
REG_LABEL_PASSWORD = " Lösenord:"
REG_LABEL_PASSWORD_CONFIRM = " Bekräfta lösenord:"
REG_LABEL_PASSWORD_WEAK = "Lösenordet är lätt."
REG_LABEL_PASSWORD_NORMAL = "Lösenordet är normalt."
REG_LABEL_PASSWORD_STRONG = "Lösenordet är svårt."
REG_LABEL_SECRET_QUESTION = "Secret question:"
REG_LABEL_SECRET_ANSWER = "Secret answer:"

REG_COMMAND_SUBMIT = "&Acceptera"
REG_COMMAND_CLOSE = "&Ständ"

REG_CHECK_PASSWORD_SHOW = "&Show Password."

REG_MSG_ACCOUNT_EXIST = "Namnet är upptaget."
REG_MSG_ACCOUNT_INVALID = "Ogiltigt namn."
REG_MSG_ACCOUNT_NUMERIC = "Namnet kan inte bestå av number."
REG_MSG_ACCOUNT_EMPTY = "Inget namn angivet."
REG_MSG_ACCOUNT_SHORT = "För kort namn, det kräver åtminstone 4 bokstäver."

REG_MSG_PASSWORD_MATCH = "Ogiltigt lösenord."
REG_MSG_PASSWORD_SHORT = "För kort lösenord, det kräver åtminstone 6 bokstäver."
REG_MSG_PASSWORD_EMPTY = "Inget lösenord angivet."

REG_MSG_SECRET_ANSWER_EMPTY = "No secret answered entered."

REG_MSG_SUCCESSFULLY = "Kontot har skapats."
REG_MSG_ERROR = "Ett fel har uppstått var snäll och försök igen."
REG_MSG_ERROR_OCCURED = "Ett fel har uppstått ..."
REG_MSG_LOADING = " Laddar .."
REG_MSG_CONNECTION_BROKEN = "Anslutnings fel, var snäll och försök igen."

REG_CMB_SECRET_QUESTION_0 = "What is the name of your pet?"
REG_CMB_SECRET_QUESTION_1 = "Your favorite book?"
REG_CMB_SECRET_QUESTION_2 = "Your favorite movie?"
REG_CMB_SECRET_QUESTION_3 = "Your favorite game?"
REG_CMB_SECRET_QUESTION_4 = "Your favorite singer?"
REG_CMB_SECRET_QUESTION_5 = "The place where your mother was born?"

SET_LABEL_COLOR = "Nuvarande färg:"

SET_FRAME_OPTIONS = "Alternativ"
SET_FRAME_CONNECTION = "Anslutnings inställningar"

SET_CHECK_SAVE_ACCOUNT = "Spara konto"
SET_CHECK_SAVE_PASSWORD = "Spara lösenord"
SET_CHECK_ASK_CLOSING = "Ask before closing"
SET_CHECK_MINIMIZE = "Minimera Peach-fönstret till Aktivitetsfältet"

SET_COMMAND_LANGUAGE = "&Språk"
SET_COMMAND_SAVE = "&Spara"

FP_FRAME_FORGOT_PASSWORD = "Forgot Passwort"
FP_LABEL_ACCOUNT = " Enter your account name:"
FP_LABEL_SECRET_QUESTION = " Secret Question:"
FP_LABEL_SECRET_ANSWER = " Secret Answer:"
FP_COMMAND_REQUEST = "&Request"
FP_CAPTION = "Peach - Forgot Password"

FP_MSG_SUCCESSFULL = "Your password is "
FP_MSG_WRONG_ANSWER = "The answer is wrong."
End Sub

Public Sub SET_LANG_ITALIAN()
CURRENT_LANG = 4

' Mdi form
MDI_COMMAND_CHAT = "Ch&at"
MDI_COMMAND_SENDFILE = "&Invia File"
MDI_COMMAND_SOCIETY = "&Lista contatti Online"

MDI_STAT_DISCONNECTED = "Stato: Disconnesso"
MDI_STAT_DISCONNECT = "Stato: Disconnesso dal Server"
MDI_STAT_CONNECTED = "Stato: Connesso"
MDI_STAT_CONNECTION_ERROR = "Stato: Disconnesso a causa di problemi di connessione"
MDI_STAT_CONNECTING = "Stato: Connessione .."

'MDImsgbox_errorHandlerFormLoad
MDI_MSG_NAME_TAKEN = "Il nome immesso e' gia' in uso."
MDI_MSG_WRONG_ACCOUNT = "The account does not exist or is wrong."
MDI_MSG_WRONG_PASSWORD = "The password is wrong."
MDI_MSG_BANNED = "This account is banned."
MDI_MSG_UNLOAD = "Are you sure you want to close Peach?"

' Config form ..
CONFIG_COMMAND_CONNECT = "&Connesso"
CONFIG_COMMAND_DISCONNECT = "&Disconnesso"
CONFIG_COMMAND_SETTINGS = "&Settings"
CONFIG_COMMAND_UPDATE = "&Update"
CONFIG_COMMAND_REGISTER = "&Register Account"
CONFIG_COMMAND_FORGOT_PASSWORD = "Forgot Password"

CONFIG_CHECK_SAVE_PASSWORD = "&Save Password"

LANG_GERMAN = "Tedesco"
LANG_ENGLISH = "Inglese"
LANG_SPANISH = "Spagnolo"
LANG_SWEDISH = "Svedese"
LANG_ITALIAN = "Italiano"
LANG_SERBIAN = "Serbo"
LANG_DUTCH = "Olandese"
LANG_FRENCH = "Francese"

CONFIG_MSG_ACCOUNT = "You did'nt introduce an account."
CONFIG_MSG_PASSWORD = "You did'nt introduce an password."
CONFIG_MSG_NUMERIC = "Non puoi immettere nomi composti da numeri."
CONFIG_MSG_PORT = "Non hai selezionato una porta valida."
CONFIG_MSG_IP = "Non hai immesso un IP."
CONFIG_MSG_NAME = "Non hai immesso un Nome utente."
CONFIG_MSG_NAME_SHORT = "The name you introduced is too short."
CONFIG_MSG_NAME_INVALID = "The name you introduced is invalid."
CONFIG_MSG_UPDATE_FILE = "You need the peach updater to be able to upgrade your peach." & vbCrLf & vbCrLf & "Download it here: http://riplegion.ri.funpic.de/Peach/peachUpdater.exe"

' Chat form ..
CHAT_COMMAND_SEND = "&Invia"
CHAT_COMMAND_CLEAR = "&Clear"

' Send file form ..
SF_LABEL_FILENAME = " Nome file:"
SF_LABEL_SENDING_FILE = "Inviando:"
SF_LABEL_SENT = "0.0% Inviato"
SF_LABEL_SEND_TO = "Send to:"

SF_MSG_USER = "No user selected."
SF_MSG_FILE = "No file selected."
SF_MSG_INCOMMING_FILE_1 = "You are receiving '"
SF_MSG_INCOMMING_FILE_2 = "' from "
SF_MSG_INCOMMING_FILE_3 = ". Do you want to accept?"
SF_MSG_DECILINED = "File transfer was decilined."

SF_COMMAND_BROWSE = "&Cerca .."
SF_COMMAND_SENDFILE = "Invia"
SF_COMMAND_CANCEL = "Annulla .."

DESP_TEXT_NEW_MSG = "New Message!"
DESP_TEXT_DC_SERVER = "Disconnesso dal Server!"

LANG_COMMAND_ENTER = "&Apri"
LANG_LABEL_SELLANG = "Seleziona la tua lingua:"

SOC_FRIEND_LIST = "Friend List"
SOC_ONLINE_LIST = "Online List"

SOC_COMMAND_ADD = "&Add"
SOC_COMMAND_REMOVE = "&Remove"

SOC_ASK_DEL_1 = "Do you want to delete '"
SOC_ASK_DEL_2 = "' from your friendlist?"

REG_CAPTION = "Peach - Registration"

REG_FRAME_DETAIL = "Enter your details"

REG_LABEL_ACCOUNT_NAME = " Account Name:"
REG_LABEL_PASSWORD = " Password:"
REG_LABEL_PASSWORD_CONFIRM = " Confirm the Password:"
REG_LABEL_PASSWORD_WEAK = "The password is weak."
REG_LABEL_PASSWORD_NORMAL = "The password is normal."
REG_LABEL_PASSWORD_STRONG = "The password is strong."
REG_LABEL_SECRET_QUESTION = "Secret question:"
REG_LABEL_SECRET_ANSWER = "Secret answer:"

REG_COMMAND_SUBMIT = "&Submit"
REG_COMMAND_CLOSE = "&Close"

REG_CHECK_PASSWORD_SHOW = "&Show Password."

REG_MSG_ACCOUNT_EXIST = "The account name already exists."
REG_MSG_ACCOUNT_INVALID = "Invalid account name."
REG_MSG_ACCOUNT_NUMERIC = "Account can't be made of numeric characters."
REG_MSG_ACCOUNT_EMPTY = "No account entered."
REG_MSG_ACCOUNT_SHORT = "Account name to short, it requieres at least 4 characters."

REG_MSG_PASSWORD_MATCH = "The passwords dont match."
REG_MSG_PASSWORD_SHORT = "Password to short, it requieres at least 6 characters."
REG_MSG_PASSWORD_EMPTY = "No Password entered."

REG_MSG_SECRET_ANSWER_EMPTY = "No secret answered entered."

REG_MSG_SUCCESSFULLY = "The account was successfully registered."
REG_MSG_ERROR = "An error has occured please try later again."
REG_MSG_ERROR_OCCURED = "Error has occured ..."
REG_MSG_LOADING = " Loading .."
REG_MSG_CONNECTION_BROKEN = "Connection is broken please try again later."

REG_CMB_SECRET_QUESTION_0 = "What is the name of your pet?"
REG_CMB_SECRET_QUESTION_1 = "Your favorite book?"
REG_CMB_SECRET_QUESTION_2 = "Your favorite movie?"
REG_CMB_SECRET_QUESTION_3 = "Your favorite game?"
REG_CMB_SECRET_QUESTION_4 = "Your favorite singer?"
REG_CMB_SECRET_QUESTION_5 = "The place where your mother was born?"

SET_LABEL_COLOR = "Current Color:"

SET_FRAME_OPTIONS = "Options"
SET_FRAME_CONNECTION = "Connection Settings"

SET_CHECK_SAVE_ACCOUNT = "Save Account"
SET_CHECK_SAVE_PASSWORD = "Save Password"
SET_CHECK_ASK_CLOSING = "Ask before closing"
SET_CHECK_MINIMIZE = "Contrai la finestra di Peach nella barra delle applicazioni"

SET_COMMAND_LANGUAGE = "&Language"
SET_COMMAND_SAVE = "&Save"

SF2_COMMAND_OPEN_FILE = "&Open File Folder"

FP_FRAME_FORGOT_PASSWORD = "Forgot Passwort"
FP_LABEL_ACCOUNT = " Enter your account name:"
FP_LABEL_SECRET_QUESTION = " Secret Question:"
FP_LABEL_SECRET_ANSWER = " Secret Answer:"
FP_COMMAND_REQUEST = "&Request"
FP_CAPTION = "Peach - Forgot Password"

FP_MSG_SUCCESSFULL = "Your password is "
FP_MSG_WRONG_ANSWER = "The answer is wrong."
End Sub

Public Sub SET_LANG_SERBIAN()
CURRENT_LANG = 6

' Mdi form ..
MDI_COMMAND_CHAT = "Ch&at"
MDI_COMMAND_SENDFILE = "&Slanje fajla"
MDI_COMMAND_SOCIETY = "&Onlajn lista"

MDI_STAT_DISCONNECTED = "Status: Veza je prekinuta"
MDI_STAT_DISCONNECT = "Status: Veza sa serverom je prekinuta"
MDI_STAT_CONNECTED = "Status: Povezi"
MDI_STAT_CONNECTION_ERROR = "Status: Problem sa konekcijom veza je prekinuta "
MDI_STAT_CONNECTING = "Status: Povezi .."

'MDImsgbox_errorHandlerFormLoad
MDI_MSG_NAME_TAKEN = "Ime je vec zauzeto."
MDI_MSG_WRONG_ACCOUNT = "The account does not exist or is wrong."
MDI_MSG_WRONG_PASSWORD = "The password is wrong."
MDI_MSG_BANNED = "This account is banned."
MDI_MSG_UNLOAD = "Are you sure you want to close Peach?"

' Config form ..
CONFIG_COMMAND_CONNECT = "&Povezi se"
CONFIG_COMMAND_DISCONNECT = "&Veza je prekinuta"
CONFIG_COMMAND_SETTINGS = "&Settings"
CONFIG_COMMAND_UPDATE = "&Update"
CONFIG_COMMAND_REGISTER = "&Register Account"
CONFIG_COMMAND_FORGOT_PASSWORD = "Forgot Password"

CONFIG_CHECK_SAVE_PASSWORD = "&Save Password"

LANG_GERMAN = "Nemacki"
LANG_ENGLISH = "Engleski"
LANG_SPANISH = "Spanski"
LANG_SWEDISH = "Svedski"
LANG_ITALIAN = "Italijanski"
LANG_SERBIAN = "Srpski"
LANG_DUTCH = "Holandski"
LANG_FRENCH = "Francuski"

CONFIG_MSG_ACCOUNT = "You did'nt introduce an account."
CONFIG_MSG_PASSWORD = "You did'nt introduce an password."
CONFIG_MSG_NUMERIC = "Ne mozete uzeti numericka imena."
CONFIG_MSG_PORT = "Niste uneli port."
CONFIG_MSG_IP = "Niste uneli IP"
CONFIG_MSG_NAME = "Niste uneli ime"
CONFIG_MSG_NAME_SHORT = "The name you introduced is too short."
CONFIG_MSG_NAME_INVALID = "The name you introduced is invalid."
CONFIG_MSG_UPDATE_FILE = "You need the peach updater to be able to upgrade your peach." & vbCrLf & vbCrLf & "Download it here: http://riplegion.ri.funpic.de/Peach/peachUpdater.exe"

'Chat form ..
CHAT_COMMAND_SEND = "&Posalji"
CHAT_COMMAND_CLEAR = "&Obrisi"

'Send file form ..
SF_LABEL_FILENAME = " Ime  arhive:"
SF_LABEL_SENDING_FILE = "Slanje:"
SF_LABEL_SENT = "0.0% Poslato"
SF_LABEL_SEND_TO = "Send to:"

SF_MSG_USER = "No user selected."
SF_MSG_FILE = "No file selected."
SF_MSG_INCOMMING_FILE_1 = "You are receiving '"
SF_MSG_INCOMMING_FILE_2 = "' from "
SF_MSG_INCOMMING_FILE_3 = ". Do you want to accept?"
SF_MSG_DECILINED = "File transfer was decilined."

SF_COMMAND_BROWSE = "Trazi .."
SF_COMMAND_SENDFILE = "Posalji"
SF_COMMAND_CANCEL = "Otkazhi .."

DESP_TEXT_NEW_MSG = "New Message!"
DESP_TEXT_DC_SERVER = "Veza sa serverom je prekinuta!"

LANG_COMMAND_ENTER = "&Otvori"
LANG_LABEL_SELLANG = "Dodaj svoj jezik:"

SOC_FRIEND_LIST = "Friend List"
SOC_ONLINE_LIST = "Online List"

SOC_COMMAND_ADD = "&Add"
SOC_COMMAND_REMOVE = "&Remove"

SOC_ASK_DEL_1 = "Do you want to delete '"
SOC_ASK_DEL_2 = "' from your friendlist?"

REG_CAPTION = "Peach - Registration"

REG_FRAME_DETAIL = "Enter your details"

REG_LABEL_ACCOUNT_NAME = " Account Name:"
REG_LABEL_PASSWORD = " Password:"
REG_LABEL_PASSWORD_CONFIRM = " Confirm the Password:"
REG_LABEL_PASSWORD_WEAK = "The password is weak."
REG_LABEL_PASSWORD_NORMAL = "The password is normal."
REG_LABEL_PASSWORD_STRONG = "The password is strong."
REG_LABEL_SECRET_QUESTION = "Secret question:"
REG_LABEL_SECRET_ANSWER = "Secret answer:"

REG_COMMAND_SUBMIT = "&Submit"
REG_COMMAND_CLOSE = "&Close"

REG_CHECK_PASSWORD_SHOW = "&Show Password."

REG_MSG_ACCOUNT_EXIST = "The account name already exists."
REG_MSG_ACCOUNT_INVALID = "Invalid account name."
REG_MSG_ACCOUNT_NUMERIC = "Account can't be made of numeric characters."
REG_MSG_ACCOUNT_EMPTY = "No account entered."
REG_MSG_ACCOUNT_SHORT = "Account name to short, it requieres at least 4 characters."

REG_MSG_PASSWORD_MATCH = "The passwords dont match."
REG_MSG_PASSWORD_SHORT = "Password to short, it requieres at least 6 characters."
REG_MSG_PASSWORD_EMPTY = "No Password entered."

REG_MSG_SECRET_ANSWER_EMPTY = "No secret answered entered."

REG_MSG_SUCCESSFULLY = "The account was successfully registered."
REG_MSG_ERROR = "An error has occured please try later again."
REG_MSG_ERROR_OCCURED = "Error has occured ..."
REG_MSG_LOADING = " Loading .."
REG_MSG_CONNECTION_BROKEN = "Connection is broken please try again later."

REG_CMB_SECRET_QUESTION_0 = "What is the name of your pet?"
REG_CMB_SECRET_QUESTION_1 = "Your favorite book?"
REG_CMB_SECRET_QUESTION_2 = "Your favorite movie?"
REG_CMB_SECRET_QUESTION_3 = "Your favorite game?"
REG_CMB_SECRET_QUESTION_4 = "Your favorite singer?"
REG_CMB_SECRET_QUESTION_5 = "The place where your mother was born?"

SET_LABEL_COLOR = "Current Color:"

SET_FRAME_OPTIONS = "Options"
SET_FRAME_CONNECTION = "Connection Settings"

SET_CHECK_SAVE_ACCOUNT = "Save Account"
SET_CHECK_SAVE_PASSWORD = "Save Password"
SET_CHECK_ASK_CLOSING = "Ask before closing"
SET_CHECK_MINIMIZE = "Minimize Peach window to system tray"

SET_COMMAND_LANGUAGE = "&Language"
SET_COMMAND_SAVE = "&Save"

SF2_COMMAND_OPEN_FILE = "&Open File Folder"

FP_FRAME_FORGOT_PASSWORD = "Forgot Passwort"
FP_LABEL_ACCOUNT = " Enter your account name:"
FP_LABEL_SECRET_QUESTION = " Secret Question:"
FP_LABEL_SECRET_ANSWER = " Secret Answer:"
FP_COMMAND_REQUEST = "&Request"
FP_CAPTION = "Peach - Forgot Password"

FP_MSG_SUCCESSFULL = "Your password is "
FP_MSG_WRONG_ANSWER = "The answer is wrong."
End Sub

Public Sub SET_LANG_DUTCH()
CURRENT_LANG = 5

MDI_COMMAND_CHAT = "Ch&at"
MDI_COMMAND_SENDFILE = "&Bestand Verzenden"
MDI_COMMAND_SOCIETY = "&Online List"

MDI_STAT_DISCONNECTED = "Status: Verbinding verbroken"
MDI_STAT_DISCONNECT = "Status: Verbinding verbroken met de server"
MDI_STAT_CONNECTED = "Status: Verbonden"
MDI_STAT_CONNECTION_ERROR = "Status: Verbinding verbroken wegens connectie problemen"
MDI_STAT_CONNECTING = "Status: Verbinden .."

'MDImsgbox_errorHandlerFormLoad
MDI_MSG_NAME_TAKEN = "Deze naam is niet beschikbaar."
MDI_MSG_WRONG_ACCOUNT = "The account does not exist or is wrong."
MDI_MSG_WRONG_PASSWORD = "The password is wrong."
MDI_MSG_BANNED = "This account is banned."
MDI_MSG_UNLOAD = "Are you sure you want to close Peach?"

CONFIG_COMMAND_CONNECT = "&Verbind"
CONFIG_COMMAND_DISCONNECT = "&Verbinding verbreken"
CONFIG_COMMAND_SETTINGS = "&Instellingen"
CONFIG_COMMAND_REGISTER = "&Account aanmaken"
CONFIG_COMMAND_UPDATE = "&Update"
CONFIG_COMMAND_FORGOT_PASSWORD = "Forgot Password"

CONFIG_CHECK_SAVE_PASSWORD = "&Save Password"

LANG_GERMAN = "Duits"
LANG_ENGLISH = "Engels"
LANG_SPANISH = "Spaans"
LANG_SWEDISH = "Zweeds"
LANG_ITALIAN = "Italiaans"
LANG_SERBIAN = "Serbisch"
LANG_DUTCH = "Nederlands"
LANG_FRENCH = "Frans"

CONFIG_MSG_ACCOUNT = "Je hebt geen gebruikersnaam ingevuld."
CONFIG_MSG_PASSWORD = "Je hebt geen wachtwoord ingevuld."
CONFIG_MSG_NUMERIC = "U kan geen naam nemen dat nummers bevat."
CONFIG_MSG_PORT = "U hebt geen poort ingesteld."
CONFIG_MSG_IP = "U hebt geen IP gegoven."
CONFIG_MSG_NAME = "U hebt geen naam gegoven."
CONFIG_MSG_NAME_SHORT = "The name you introduced is too short."
CONFIG_MSG_NAME_INVALID = "The name you introduced is invalid."
CONFIG_MSG_UPDATE_FILE = "You need the peach updater to be able to upgrade your peach." & vbCrLf & vbCrLf & "Download it here: http://riplegion.ri.funpic.de/Peach/peachUpdater.exe"

CHAT_COMMAND_SEND = "&Zend"
CHAT_COMMAND_CLEAR = "&Leegmaken"

SF_LABEL_FILENAME = " Bestandsnaam:"
SF_LABEL_SENDING_FILE = "verZenden:"
SF_LABEL_SENT = "0.0% verzonden"
SF_LABEL_SEND_TO = "Send to:"

SF_MSG_USER = "No user selected."
SF_MSG_FILE = "No file selected."
SF_MSG_INCOMMING_FILE_1 = "You are receiving '"
SF_MSG_INCOMMING_FILE_2 = "' from "
SF_MSG_INCOMMING_FILE_3 = ". Do you want to accept?"
SF_MSG_DECILINED = "File transfer was decilined."

SF_COMMAND_BROWSE = "&Zoeken .."
SF_COMMAND_SENDFILE = "&Stuur"
SF_COMMAND_CANCEL = "&Annuleren .."

DESP_TEXT_NEW_MSG = "Nieuw bericht!"
DESP_TEXT_DC_SERVER = "Verbinding verbroken met de server"

LANG_COMMAND_ENTER = "&Openen"
LANG_LABEL_SELLANG = "Selecteer jou taal:"

SOC_FRIEND_LIST = "Vriendenlijst"
SOC_ONLINE_LIST = "Online List"

SOC_COMMAND_ADD = "&Toevoegen"
SOC_COMMAND_REMOVE = "&Verwijderen"

SOC_ASK_DEL_1 = "Do you want to delete '"
SOC_ASK_DEL_2 = "' from your friendlist?"

REG_CAPTION = "Peach - Registration"

REG_FRAME_DETAIL = "Voeg je gegevens in"

REG_LABEL_ACCOUNT_NAME = " Gebruikersnaam:"
REG_LABEL_PASSWORD = " Wachtwoord:"
REG_LABEL_PASSWORD_CONFIRM = " Bevestig Wachtwoord:"
REG_LABEL_PASSWORD_WEAK = "Dit wachtwoord is zwak."
REG_LABEL_PASSWORD_NORMAL = "Dit wachtwoord is redelijk."
REG_LABEL_PASSWORD_STRONG = "Dit wachtwoord is goed."
REG_LABEL_SECRET_QUESTION = "Secret question:"
REG_LABEL_SECRET_ANSWER = "Secret answer:"

REG_CHECK_PASSWORD_SHOW = "Laat wachtwoord zien"

REG_COMMAND_SUBMIT = "&Versturen"
REG_COMMAND_CLOSE = "&Sluiten"

REG_MSG_ACCOUNT_EXIST = "Deze gebruikersnaam is al in gebruik."
REG_MSG_ACCOUNT_INVALID = "Niet bruikbare gebruikersnaam."
REG_MSG_ACCOUNT_NUMERIC = "Een account kan niet uit cijfers bestaan."
REG_MSG_ACCOUNT_EMPTY = "Geen gebruikersnaam ingevoerd."
REG_MSG_ACCOUNT_SHORT = "Gebruikersnaam te kort, minstens 4 teken."

REG_MSG_PASSWORD_MATCH = "Wachtwoorden komen niet overeen."
REG_MSG_PASSWORD_SHORT = "Wachtwoord te kort, minstens 4 tekens."
REG_MSG_PASSWORD_EMPTY = "Geen wachtwoord ingevoerd."

REG_MSG_SECRET_ANSWER_EMPTY = "No secret answered entered."

REG_MSG_SUCCESSFULLY = "Account succesvol aangemaakt."
REG_MSG_ERROR = "Er is een fout opgetreden, probeer het later opnieuw."
REG_MSG_ERROR_OCCURED = "Fout opgetreden ..."
REG_MSG_LOADING = " Laden.. "
REG_MSG_CONNECTION_BROKEN = "Verbinding verbroken."

REG_CMB_SECRET_QUESTION_0 = "What is the name of your pet?"
REG_CMB_SECRET_QUESTION_1 = "Your favorite book?"
REG_CMB_SECRET_QUESTION_2 = "Your favorite movie?"
REG_CMB_SECRET_QUESTION_3 = "Your favorite game?"
REG_CMB_SECRET_QUESTION_4 = "Your favorite singer?"
REG_CMB_SECRET_QUESTION_5 = "The place where your mother was born?"

SET_LABEL_COLOR = "Momenteel gebruikte kleur:"

SET_FRAME_OPTIONS = "Instellingen"
SET_FRAME_CONNECTION = "Verbinding Instellingen"

SET_CHECK_SAVE_ACCOUNT = "Gebruiker opslaan"
SET_CHECK_SAVE_PASSWORD = "Wachtwoord opslaan"
SET_CHECK_ASK_CLOSING = "Ask before closing"
SET_CHECK_MINIMIZE = "Minimize Peach window to system tray"

SET_COMMAND_LANGUAGE = "&Taal"
SET_COMMAND_SAVE = "&Opslaan"

SF_LABEL_SEND_TO = "Verzenden naar:"

SF_MSG_USER = "Geen gebruiker geselecteerd."
SF_MSG_FILE = "Geen bestand geselecteerd."
SF_MSG_INCOMMING_FILE_1 = "You are receiving '"
SF_MSG_INCOMMING_FILE_2 = "' from "
SF_MSG_INCOMMING_FILE_3 = ". Do you want to accept?"
SF_MSG_DECILINED = "Gegevensoverdracht geweigerd."

SF2_COMMAND_OPEN_FILE = "&Open bestandsmap"

FP_FRAME_FORGOT_PASSWORD = "Forgot Passwort"
FP_LABEL_ACCOUNT = " Enter your account name:"
FP_LABEL_SECRET_QUESTION = " Secret Question:"
FP_LABEL_SECRET_ANSWER = " Secret Answer:"
FP_COMMAND_REQUEST = "&Request"
FP_CAPTION = "Peach - Forgot Password"

FP_MSG_SUCCESSFULL = "Your password is "
FP_MSG_WRONG_ANSWER = "The answer is wrong."
End Sub

Public Sub SET_LANG_FRENCH()
CURRENT_LANG = 7

MDI_COMMAND_CHAT = "Ch&at"
MDI_COMMAND_SENDFILE = "&Envoi File"
MDI_COMMAND_SOCIETY = "&Liste contact Online"

MDI_STAT_DISCONNECTED = "Etat: Deconnecté"
MDI_STAT_DISCONNECT = "Etat: Deconnecté du Server"
MDI_STAT_CONNECTED = "Etat: Connecté"
MDI_STAT_CONNECTION_ERROR = "Etat: Deconnecté à cause de problèmes do connection"
MDI_STAT_CONNECTING = "Etat: Connection .."

MDI_MSG_NAME_TAKEN = "Le nom inséré est déjà utilizé."
MDI_MSG_WRONG_ACCOUNT = "The account does not exist or is wrong."
MDI_MSG_WRONG_PASSWORD = "The password is wrong."
MDI_MSG_BANNED = "This account is banned."
MDI_MSG_UNLOAD = "Are you sure you want to close Peach?"

CONFIG_COMMAND_CONNECT = "&Connecté"
CONFIG_COMMAND_DISCONNECT = "&Deconnecté"
CONFIG_COMMAND_SETTINGS = "&Settings"
CONFIG_COMMAND_UPDATE = "&Update"
CONFIG_COMMAND_REGISTER = "&Register Account"
CONFIG_COMMAND_FORGOT_PASSWORD = "Forgot Password"

CONFIG_CHECK_SAVE_PASSWORD = "&Save Password"

LANG_GERMAN = "Alleman"
LANG_ENGLISH = "Anglais"
LANG_SPANISH = "Espagnol"
LANG_SWEDISH = "Suédois"
LANG_ITALIAN = "Italien"
LANG_SERBIAN = "Serbois"
LANG_DUTCH = "Hollandais"
LANG_FRENCH = "Français"

CONFIG_MSG_ACCOUNT = "You did'nt introduce an account."
CONFIG_MSG_PASSWORD = "You did'nt introduce an password."
CONFIG_MSG_NUMERIC = "Tu ne peut pas insérer noms composé de numeros."
CONFIG_MSG_PORT = "Tu n'as pas selectionner une porte valide."
CONFIG_MSG_IP = "Tu n'as pas innecté un IP."
CONFIG_MSG_NAME = "Tu n'as pas innecté un Nom utilizateur."
CONFIG_MSG_NAME_SHORT = "The name you introduced is too short."
CONFIG_MSG_NAME_INVALID = "The name you introduced is invalid."
CONFIG_MSG_UPDATE_FILE = "You need the peach updater to be able to upgrade your peach." & vbCrLf & vbCrLf & "Download it here: http://riplegion.ri.funpic.de/Peach/peachUpdater.exe"

CHAT_COMMAND_SEND = "&Envoi"
CHAT_COMMAND_CLEAR = "&Clear"

SF_LABEL_FILENAME = " Nom file:"
SF_LABEL_SENDING_FILE = "Envoyant:"
SF_LABEL_SENT = "0.0% Envoyé"
SF_LABEL_SEND_TO = "Send to:"

SF_MSG_USER = "No user selected."
SF_MSG_FILE = "No file selected."
SF_MSG_INCOMMING_FILE_1 = "You are receiving '"
SF_MSG_INCOMMING_FILE_2 = "' from "
SF_MSG_INCOMMING_FILE_3 = ". Do you want to accept?"
SF_MSG_DECILINED = "File transfer was decilined."

SF_COMMAND_BROWSE = "&Cherche .."
SF_COMMAND_SENDFILE = "Envoi"
SF_COMMAND_CANCEL = "Annuler .."

DESP_TEXT_NEW_MSG = "New Message!"
DESP_TEXT_DC_SERVER = "Deconnecté du Server"

LANG_COMMAND_ENTER = "&Ouvrir"
LANG_LABEL_SELLANG = "Choisissez votre langue:"

SOC_FRIEND_LIST = "Friend List"
SOC_ONLINE_LIST = "Online List"

SOC_COMMAND_ADD = "&Add"
SOC_COMMAND_REMOVE = "&Remove"

SOC_ASK_DEL_1 = "Do you want to delete '"
SOC_ASK_DEL_2 = "' from your friendlist?"

REG_CAPTION = "Peach - Registration"

REG_FRAME_DETAIL = "Enter your details"

REG_LABEL_ACCOUNT_NAME = " Account Name:"
REG_LABEL_PASSWORD = " Password:"
REG_LABEL_PASSWORD_CONFIRM = " Confirm the Password:"
REG_LABEL_PASSWORD_WEAK = "The password is weak."
REG_LABEL_PASSWORD_NORMAL = "The password is normal."
REG_LABEL_PASSWORD_STRONG = "The password is strong."
REG_LABEL_SECRET_QUESTION = "Secret question:"
REG_LABEL_SECRET_ANSWER = "Secret answer:"

REG_COMMAND_SUBMIT = "&Submit"
REG_COMMAND_CLOSE = "&Close"

REG_CHECK_PASSWORD_SHOW = "&Show Password."

REG_MSG_ACCOUNT_EXIST = "The account name already exists."
REG_MSG_ACCOUNT_INVALID = "Invalid account name."
REG_MSG_ACCOUNT_NUMERIC = "Account can't be made of numeric characters."
REG_MSG_ACCOUNT_EMPTY = "No account entered."
REG_MSG_ACCOUNT_SHORT = "Account name to short, it requieres at least 4 characters."

REG_MSG_PASSWORD_MATCH = "The passwords dont match."
REG_MSG_PASSWORD_SHORT = "Password to short, it requieres at least 6 characters."
REG_MSG_PASSWORD_EMPTY = "No Password entered."

REG_MSG_SECRET_ANSWER_EMPTY = "No secret answered entered."

REG_MSG_SUCCESSFULLY = "The account was successfully registered."
REG_MSG_ERROR = "An error has occured please try later again."
REG_MSG_ERROR_OCCURED = "Error has occured ..."
REG_MSG_LOADING = " Loading .."
REG_MSG_CONNECTION_BROKEN = "Connection is broken please try again later."

REG_CMB_SECRET_QUESTION_0 = "What is the name of your pet?"
REG_CMB_SECRET_QUESTION_1 = "Your favorite book?"
REG_CMB_SECRET_QUESTION_2 = "Your favorite movie?"
REG_CMB_SECRET_QUESTION_3 = "Your favorite game?"
REG_CMB_SECRET_QUESTION_4 = "Your favorite singer?"
REG_CMB_SECRET_QUESTION_5 = "The place where your mother was born?"

SET_LABEL_COLOR = "Current Color:"

SET_FRAME_OPTIONS = "Options"
SET_FRAME_CONNECTION = "Connection Settings"

SET_CHECK_SAVE_ACCOUNT = "Save Account"
SET_CHECK_SAVE_PASSWORD = "Save Password"
SET_CHECK_ASK_CLOSING = "Ask before closing"
SET_CHECK_MINIMIZE = "Minimize Peach window to system tray"

SET_COMMAND_LANGUAGE = "&Language"
SET_COMMAND_SAVE = "&Save"

SF2_COMMAND_OPEN_FILE = "&Open File Folder"

FP_FRAME_FORGOT_PASSWORD = "Forgot Passwort"
FP_LABEL_ACCOUNT = " Enter your account name:"
FP_LABEL_SECRET_QUESTION = " Secret Question:"
FP_LABEL_SECRET_ANSWER = " Secret Answer:"
FP_COMMAND_REQUEST = "&Request"
FP_CAPTION = "Peach - Forgot Password"

FP_MSG_SUCCESSFULL = "Your password is "
FP_MSG_WRONG_ANSWER = "The answer is wrong."
End Sub
