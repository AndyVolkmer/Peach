Attribute VB_Name = "modLang"
Option Explicit

Public CURRENT_LANG                     As Long

'MDI form ..
Public MDI_COMMAND_CHAT                 As String
Public MDI_COMMAND_SENDFILE             As String
Public MDI_COMMAND_SOCIETY              As String

Public MDI_MSG_ERROR_FORM_LOAD          As String
Public MDI_MSG_NAME_TAKEN               As String
Public MDI_MSG_WRONG_ACCOUNT            As String
Public MDI_MSG_WRONG_PASSWORD           As String
Public MDI_MSG_BANNED                   As String
Public MDI_MSG_UNLOAD                   As String

Public MDI_MSG_CANT_ADD_YOU             As String
Public MDI_MSG_ALREADY_IN_IGNORE_LIST   As String
Public MDI_MSG_ALREADY_IN_FRIEND_LIST   As String
Public MDI_MSG_ACCOUNT_NOT_EXIST        As String

Public MDI_MENU                         As String

'Configuration form ..
Public CONFIG_LABEL_ACCOUNT             As String
Public CONFIG_LABEL_PASSWORD            As String

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
Public CONFIG_MSG_UPDATE_FILE           As String

Public CHAT_COMMAND_SEND                As String
Public CHAT_COMMAND_CLEAR               As String

Public SF_LABEL_FILENAME                As String
Public SF_LABEL_SEND_TO                 As String
Public SF_LABEL_TIME                    As String
Public SF_LABEL_KBS                     As String
Public SF_LABEL_KBSS                    As String

Public SF_COMMAND_BROWSE                As String
Public SF_COMMAND_SENDFILE              As String
Public SF_COMMAND_CANCEL                As String

Public SF_MSG_USER                      As String
Public SF_MSG_FILE                      As String
Public SF_MSG_INCOMMING_FILE            As String
Public SF_MSG_DECILINED                 As String

'Language form ..
Public LANG_COMMAND_ENTER               As String
Public LANG_LABEL_SELLANG               As String

Public LANG_QUIT                        As String

Public LANG_GERMAN                      As String
Public LANG_ENGLISH                     As String
Public LANG_SPANISH                     As String
Public LANG_SWEDISH                     As String
Public LANG_ITALIAN                     As String
Public LANG_SERBIAN                     As String
Public LANG_DUTCH                       As String
Public LANG_FRENCH                      As String
Public LANG_NORWEGIAN                   As String

Public SOC_FRIEND_LIST                  As String
Public SOC_ONLINE_LIST                  As String
Public SOC_IGNORE_LIST                  As String

Public SOC_COMMAND_REMOVE               As String
Public SOC_COMMAND_FRIEND               As String
Public SOC_COMMAND_IGNORE               As String
Public SOC_COMMAND_WHISPER              As String

Public SOC_ASK_DEL                      As String

Public SOC_ASK_FRIEND_TEXT              As String
Public SOC_ASK_FRIEND_TITLE             As String
Public SOC_ASK_FRIEND_DEFAULT           As String

Public SOC_ASK_IGNORE_TEXT              As String
Public SOC_ASK_IGNORE_TITLE             As String
Public SOC_ASK_IGNORE_DEFAULT           As String

Public SOC_FRIEND_LIST_STATUS           As String

Public SOC_MSG_CANT_WHISPER             As String

'Create an account form
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

Public REG_MSG_EMAIL_EMPTY              As String
Public REG_MSG_EMAIL_TAKEN              As String

Public REG_CMB_SECRET_QUESTION_0        As String
Public REG_CMB_SECRET_QUESTION_1        As String
Public REG_CMB_SECRET_QUESTION_2        As String
Public REG_CMB_SECRET_QUESTION_3        As String
Public REG_CMB_SECRET_QUESTION_4        As String
Public REG_CMB_SECRET_QUESTION_5        As String

Public REG_LABEL_GENDER                 As String

Public REG_CMB_GENDER_MALE              As String
Public REG_CMB_GENDER_FEMALE            As String

'Settings form
Public SET_LABEL_COLOR                  As String
Public SET_LABEL_FONT                   As String

Public SET_FRAME_STYLE                  As String
Public SET_FRAME_OPTIONS                As String
Public SET_FRAME_CONNECTION             As String

Public SET_CHECK_SAVE_ACCOUNT           As String
Public SET_CHECK_SAVE_PASSWORD          As String
Public SET_CHECK_AUTO_LOGIN             As String
Public SET_CHECK_ASK_CLOSING            As String
Public SET_CHECK_MINIMIZE               As String

Public SET_COMMAND_LANGUAGE             As String
Public SET_COMMAND_SAVE                 As String

Public SF2_COMMAND_OPEN_FILE            As String

Public FP_FRAME_FORGOT_PASSWORD         As String
Public FP_LABEL_EMAIL                   As String
Public FP_LABEL_SECRET_QUESTION         As String
Public FP_LABEL_SECRET_ANSWER           As String
Public FP_COMMAND_REQUEST               As String
Public FP_CAPTION                       As String

Public FP_MSG_SUCCESSFULL               As String
Public FP_MSG_WRONG_ANSWER              As String
Public FP_MSG_WRONG_EMAIL               As String

Public Sub SET_LANG_GERMAN()
CURRENT_LANG = 0

' MDI form ..
MDI_COMMAND_CHAT = "Ch&at"
MDI_COMMAND_SENDFILE = "&Datei senden"
MDI_COMMAND_SOCIETY = "&Gesellschaft"

MDI_MSG_NAME_TAKEN = "Der Name ist bereits vergeben."
MDI_MSG_WRONG_ACCOUNT = "Dieser Konto-Name ist nicht vorhanden oder falsch."
MDI_MSG_WRONG_PASSWORD = "Das Kennwort ist falsch."
MDI_MSG_BANNED = "Dieses Konto wurde gebannt."
MDI_MSG_UNLOAD = "Sind Sie sicher, dass Sie Peach schließen wollen?"

MDI_MSG_CANT_ADD_YOU = "Sie können sich nicht selbst hinzufügen."
MDI_MSG_ALREADY_IN_IGNORE_LIST = "'%u' ist bereits in deiner Ignorier-Liste."
MDI_MSG_ALREADY_IN_FRIEND_LIST = "'%u' ist bereits in deiner Freundesliste."
MDI_MSG_ACCOUNT_NOT_EXIST = "'%u' existiert nicht."

MDI_MENU = "Menü"

CONFIG_LABEL_ACCOUNT = "Konto"
CONFIG_LABEL_PASSWORD = "Kennwort"

CONFIG_COMMAND_CONNECT = "&Verbinden"
CONFIG_COMMAND_DISCONNECT = "&Verbindung trenn."
CONFIG_COMMAND_SETTINGS = "&Einstellungen"
CONFIG_COMMAND_UPDATE = "&Aktualisieren"
CONFIG_COMMAND_REGISTER = "&Konto erstellen"
CONFIG_COMMAND_FORGOT_PASSWORD = "&Kennwort vergessen"

CONFIG_CHECK_SAVE_PASSWORD = "&Kennwort speichern"

CONFIG_FRAME_CONNECTION = "Verbindungsinformationen"

LANG_GERMAN = "Deutsch"
LANG_ENGLISH = "Englisch"
LANG_SPANISH = "Spanisch"
LANG_SWEDISH = "Schwedisch"
LANG_ITALIAN = "Italienisch"
LANG_SERBIAN = "Serbisch"
LANG_DUTCH = "Niederländisch"
LANG_FRENCH = "Französisch"
LANG_NORWEGIAN = "Norwegisch"

CONFIG_MSG_ACCOUNT = "Sie haben keinen Konto-Namen eingegeben."
CONFIG_MSG_PASSWORD = "Sie haben kein Kennwort eingegeben."
CONFIG_MSG_NUMERIC = "Sie können keine Ziffern in ihrem Namen haben."
CONFIG_MSG_PORT = "Sie haben keinen Port eingegeben."
CONFIG_MSG_IP = "Sie haben keine IP eingegeben."
CONFIG_MSG_UPDATE_FILE = "Sie brauchen den Peach Updater um ihr Peach zu updaten." & vbCrLf & vbCrLf & "Sie können es hier downloaden:  http://riplegion.ri.funpic.de/Peach/peachUpdater.exe"

CHAT_COMMAND_SEND = "&Senden"
CHAT_COMMAND_CLEAR = "&Löschen"

SF_LABEL_FILENAME = "Datei Name"
SF_LABEL_SEND_TO = "Sende an"
SF_LABEL_TIME = "Verbleibende Zeit: "
SF_LABEL_KBS = " Kb/Sek, "
SF_LABEL_KBSS = " KBytes gesendet, "

SF_MSG_USER = "Kein Benutzer ausgewählt."
SF_MSG_FILE = "Keine Datei ausgewählt."
SF_MSG_INCOMMING_FILE = "Sie empfangen gerade '%f' von '%u'. Wollen Sie die Datei annehmen?"
SF_MSG_DECILINED = "Der Benutzer hat die Datei abgelehnt."

SF_COMMAND_BROWSE = "&Suchen .."
SF_COMMAND_SENDFILE = "Senden"
SF_COMMAND_CANCEL = "Abbrechen .."

LANG_COMMAND_ENTER = "&Auswählen"
LANG_LABEL_SELLANG = "Wählen Sie Ihre Sprache aus"

LANG_QUIT = "Um die Sprache zu ändern müssen Sie Peach neu starten, möchten Sie dies jetzt tun?"

SOC_FRIEND_LIST = "Freundesliste"
SOC_ONLINE_LIST = "Online-Liste"
SOC_IGNORE_LIST = "Ignorier-Liste"

SOC_COMMAND_REMOVE = "&Entfernen"
SOC_COMMAND_FRIEND = "&Als Freund hinzufügen"
SOC_COMMAND_IGNORE = "&Benutzer ignorieren"
SOC_COMMAND_WHISPER = "&Anflüstern"

SOC_ASK_DEL = "Möchten Sie '%u' von der Liste löschen?"

SOC_ASK_FRIEND_TEXT = "Geben Sie bitte den Konto-Namen Ihres Freundes ein."
SOC_ASK_FRIEND_TITLE = "Freund hinzufügen"
SOC_ASK_FRIEND_DEFAULT = "Konto hier eingeben"

SOC_ASK_IGNORE_TEXT = "Geben Sie bitte den Konto-Namen des Benutzer ein den Sie ignorieren möchten."
SOC_ASK_IGNORE_TITLE = "Benutzer ignorieren"
SOC_ASK_IGNORE_DEFAULT = "Konto hier eingeben"

SOC_FRIEND_LIST_STATUS = "Status"

SOC_MSG_CANT_WHISPER = "Sie können diesen Benutzer nicht anflüstern."

REG_CAPTION = "Peach - Registration"

REG_FRAME_DETAIL = "Geben Sie Ihre Daten an"

REG_LABEL_ACCOUNT_NAME = "Konto-Name"
REG_LABEL_PASSWORD = "Kennwort"
REG_LABEL_PASSWORD_CONFIRM = "Kennwort bestätigen"
REG_LABEL_PASSWORD_WEAK = "Das Kennwort ist schwach."
REG_LABEL_PASSWORD_NORMAL = "Das Kennwort ist gut."
REG_LABEL_PASSWORD_STRONG = "Das Kennwort ist stark."
REG_LABEL_SECRET_QUESTION = "Geheime Frage"
REG_LABEL_SECRET_ANSWER = "Geheime Antwort"

REG_COMMAND_SUBMIT = "&Registrieren"
REG_COMMAND_CLOSE = "&Schließen"

REG_CHECK_PASSWORD_SHOW = "&Kennwort anzeigen"

REG_MSG_ACCOUNT_EXIST = "Dieser Konto-Name ist bereits vergeben."
REG_MSG_ACCOUNT_INVALID = "Ungültiger Konto-Name."
REG_MSG_ACCOUNT_NUMERIC = "Dieser Konto-Name darf nicht aus Ziffern bestehen."
REG_MSG_ACCOUNT_EMPTY = "Kein Konto angegeben."
REG_MSG_ACCOUNT_SHORT = "Dieser Konto-Name ist zu kurz, muss aus wenigstens 4 Zeichen bestehen."

REG_MSG_PASSWORD_MATCH = "Die Kennwörter stimmen nicht überein."
REG_MSG_PASSWORD_SHORT = "Das Kennwort ist zu kurz, muss aus wenigstens 6 Zeichen bestehen."
REG_MSG_PASSWORD_EMPTY = "Kein Kennwort angegeben."

REG_MSG_SECRET_ANSWER_EMPTY = "Keine geheime Antwort angegeben."

REG_MSG_EMAIL_EMPTY = "Keine Email angegeben."
REG_MSG_EMAIL_TAKEN = "Die Email-Adresse wird bereits genutzt, haben Sie Ihr Kennwort vergessen?"

REG_MSG_SUCCESSFULLY = "Ihr Konto wurde erfolgreich erstellt."
REG_MSG_ERROR = "Ein Fehler ist aufgetreten bitte versuchen Sie es später nochmal."
REG_MSG_ERROR_OCCURED = "Fehler aufgetreten ..."
REG_MSG_LOADING = "Lädt .."
REG_MSG_CONNECTION_BROKEN = "Die Verbindung wurde unterbrochen bitte versuchen Sie es später nochmal."

REG_CMB_SECRET_QUESTION_0 = "Wie heißt Ihr Haustier?"
REG_CMB_SECRET_QUESTION_1 = "Ihr Lieblings-Buch?"
REG_CMB_SECRET_QUESTION_2 = "Ihr Lieblings-Film?"
REG_CMB_SECRET_QUESTION_3 = "Ihr Lieblings-Spiel?"
REG_CMB_SECRET_QUESTION_4 = "Ihr Lieblings-Sänger?"
REG_CMB_SECRET_QUESTION_5 = "Geburtsort Ihrer Mutter?"

REG_LABEL_GENDER = "Geschlecht"

REG_CMB_GENDER_MALE = "Männlich"
REG_CMB_GENDER_FEMALE = "Weiblich"

SET_LABEL_COLOR = "Jetzige Farbe"
SET_LABEL_FONT = "Schriftart"

SET_FRAME_STYLE = "Stil"
SET_FRAME_OPTIONS = "Optionen"
SET_FRAME_CONNECTION = "Verbindungseinstellungen"

SET_CHECK_SAVE_ACCOUNT = "Konto-Namen speichern"
SET_CHECK_SAVE_PASSWORD = "Kennwort speichern"
SET_CHECK_AUTO_LOGIN = "Automatisch einloggen"
SET_CHECK_ASK_CLOSING = "Abfragen bevor schließen"
SET_CHECK_MINIMIZE = "Peach-Fenster in die Taskleiste minimieren"

SET_COMMAND_LANGUAGE = "&Sprache"
SET_COMMAND_SAVE = "&Speichern"

SF2_COMMAND_OPEN_FILE = "&Datei Ordner öffnen"

FP_FRAME_FORGOT_PASSWORD = "Kennwort vergessen"
FP_LABEL_EMAIL = "Gebe deine Email ein"
FP_LABEL_SECRET_QUESTION = "Geheime Frage"
FP_LABEL_SECRET_ANSWER = "Geheime Antwort"
FP_COMMAND_REQUEST = "&Abfragen"
FP_CAPTION = "Peach - Kennwort vergessen"

FP_MSG_SUCCESSFULL = "Ihr Account-Name lautet '%u'." & vbCrLf & "Ihr Kennwort lautet '%p'."
FP_MSG_WRONG_ANSWER = "Die Antwort ist falsch."
FP_MSG_WRONG_EMAIL = "Die angegebene Email-Adresse konnte nicht gefunden werden."
End Sub

Public Sub SET_LANG_ENGLISH()
CURRENT_LANG = 1

'MDI form
MDI_COMMAND_CHAT = "Ch&at"
MDI_COMMAND_SENDFILE = "&Send File"
MDI_COMMAND_SOCIETY = "&Society"

'MDImsgbox_errorHandlerFormLoad
MDI_MSG_NAME_TAKEN = "This username is already taken."
MDI_MSG_WRONG_ACCOUNT = "This username does not exist."
MDI_MSG_WRONG_PASSWORD = "The password is wrong."
MDI_MSG_BANNED = "This account is banned."
MDI_MSG_UNLOAD = "Are you sure you want to close Peach?"

MDI_MSG_CANT_ADD_YOU = "You can't add yourself."
MDI_MSG_ALREADY_IN_IGNORE_LIST = "'%u' is already in your ignore list."
MDI_MSG_ALREADY_IN_FRIEND_LIST = "'%u' is already in your friend list."
MDI_MSG_ACCOUNT_NOT_EXIST = "'%u' does not exist."

MDI_MENU = "Menu"

'Configuration form ..
CONFIG_LABEL_ACCOUNT = "Username"
CONFIG_LABEL_PASSWORD = "Password"

CONFIG_COMMAND_CONNECT = "&Login"
CONFIG_COMMAND_DISCONNECT = "&Logout"
CONFIG_COMMAND_SETTINGS = "&Preferences"
CONFIG_COMMAND_UPDATE = "&Update"
CONFIG_COMMAND_REGISTER = "&Create Account"
CONFIG_COMMAND_FORGOT_PASSWORD = "&Forgot My Password"

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
LANG_NORWEGIAN = "Norwegian"

CONFIG_MSG_ACCOUNT = "You did not enter an account."
CONFIG_MSG_PASSWORD = "You did not enter a password."
CONFIG_MSG_NUMERIC = "You can not take numeric usernames."
CONFIG_MSG_PORT = "You did not introduce a port."
CONFIG_MSG_IP = "You did not introduce a IP."
CONFIG_MSG_UPDATE_FILE = "You need the peach updater to be able to upgrade your peach." & vbCrLf & vbCrLf & "Download it here: http://riplegion.ri.funpic.de/Peach/peachUpdater.exe"

' Chat form ..
CHAT_COMMAND_SEND = "&Send"
CHAT_COMMAND_CLEAR = "&Clear"

' Send File form ..
SF_LABEL_FILENAME = "File Name"
SF_LABEL_SEND_TO = "Send to"
SF_LABEL_TIME = " Time left: "
SF_LABEL_KBS = " Kb/Sec, "
SF_LABEL_KBSS = " KBytes sent, "

SF_MSG_USER = "No user selected."
SF_MSG_FILE = "No file selected."
SF_MSG_INCOMMING_FILE = "You are receiving '%f' from '%u'. Do you want to accept?"
SF_MSG_DECILINED = "The file transfer has been refused."

SF_COMMAND_BROWSE = "&Search .."
SF_COMMAND_SENDFILE = "Send"
SF_COMMAND_CANCEL = "Cancel .."

LANG_COMMAND_ENTER = "&Select"
LANG_LABEL_SELLANG = "Select your language"

LANG_QUIT = "In order to change the language you need to restart Peach, do you want to do this now?"

SOC_FRIEND_LIST = "Friend List"
SOC_ONLINE_LIST = "Online List"
SOC_IGNORE_LIST = "Ignore List"

SOC_COMMAND_REMOVE = "&Remove"
SOC_COMMAND_FRIEND = "&Add Friend"
SOC_COMMAND_IGNORE = "&Ignore User"
SOC_COMMAND_WHISPER = "&Whisper"

SOC_ASK_DEL = "Do you want to delete '%u' from the list?"

SOC_ASK_FRIEND_TEXT = "Enter the account name of your friend in the text box below."
SOC_ASK_FRIEND_TITLE = "Adding a friend"
SOC_ASK_FRIEND_DEFAULT = "Enter account here"

SOC_ASK_IGNORE_TEXT = "Enter the account name of the user you want to ignore in the text box below."
SOC_ASK_IGNORE_TITLE = "Ignore"
SOC_ASK_IGNORE_DEFAULT = "Enter account here"

SOC_FRIEND_LIST_STATUS = "Status"

SOC_MSG_CANT_WHISPER = "You can't whisper this user."

REG_CAPTION = "Peach - Registration"

REG_FRAME_DETAIL = "Enter your details"

REG_LABEL_ACCOUNT_NAME = "Username"
REG_LABEL_PASSWORD = "Password"
REG_LABEL_PASSWORD_CONFIRM = "Confirm the Password"
REG_LABEL_PASSWORD_WEAK = "The password is weak."
REG_LABEL_PASSWORD_NORMAL = "The password is normal."
REG_LABEL_PASSWORD_STRONG = "The password is strong."
REG_LABEL_SECRET_QUESTION = "Secret question"
REG_LABEL_SECRET_ANSWER = "Secret answer"

REG_COMMAND_SUBMIT = "&Submit"
REG_COMMAND_CLOSE = "&Close"

REG_CHECK_PASSWORD_SHOW = "&Show Password."

REG_MSG_ACCOUNT_EXIST = "The username is already in use."
REG_MSG_ACCOUNT_INVALID = "Invalid username."
REG_MSG_ACCOUNT_NUMERIC = "Username can not be composed of numeric characters."
REG_MSG_ACCOUNT_EMPTY = "No username entered."
REG_MSG_ACCOUNT_SHORT = "The username is too short, it requieres at least 4 characters."

REG_MSG_PASSWORD_MATCH = "The passwords don't match."
REG_MSG_PASSWORD_SHORT = "The password is too short, it requieres at least 6 characters."
REG_MSG_PASSWORD_EMPTY = "No password entered."

REG_MSG_SECRET_ANSWER_EMPTY = "No secret answered entered."

REG_MSG_EMAIL_EMPTY = "No email adress entered."
REG_MSG_EMAIL_TAKEN = "The email adress you have entered is already beeing used, have you forgotten your password?"

REG_MSG_SUCCESSFULLY = "The username was successfully registered."
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

REG_LABEL_GENDER = "Gender"

REG_CMB_GENDER_MALE = "Male"
REG_CMB_GENDER_FEMALE = "Female"

SET_LABEL_COLOR = "Current Color"
SET_LABEL_FONT = "Font"

SET_FRAME_STYLE = "Style"
SET_FRAME_OPTIONS = "Options"
SET_FRAME_CONNECTION = "Connection Settings"

SET_CHECK_SAVE_ACCOUNT = "Remember Username"
SET_CHECK_SAVE_PASSWORD = "Remember Password"
SET_CHECK_AUTO_LOGIN = "Login automatically"
SET_CHECK_ASK_CLOSING = "Ask before closing"
SET_CHECK_MINIMIZE = "Minimize Peach window to system tray"

SET_COMMAND_LANGUAGE = "&Language"
SET_COMMAND_SAVE = "&Save"

SF2_COMMAND_OPEN_FILE = "&Open File Folder"

FP_FRAME_FORGOT_PASSWORD = "Forgot Password"
FP_LABEL_EMAIL = "Enter your email adress"
FP_LABEL_SECRET_QUESTION = "Secret Question"
FP_LABEL_SECRET_ANSWER = "Secret Answer"
FP_COMMAND_REQUEST = "&Request"
FP_CAPTION = "Peach - Forgot Password"

FP_MSG_SUCCESSFULL = "Your username is '%u'." & vbCrLf & "Your password is '%p'."
FP_MSG_WRONG_ANSWER = "The answer is wrong."
FP_MSG_WRONG_EMAIL = "The entered email adress could not be found."
End Sub

Public Sub SET_LANG_SPANISH()
CURRENT_LANG = 2

MDI_COMMAND_CHAT = "Ch&at"
MDI_COMMAND_SENDFILE = "&Enviar Archivo"
MDI_COMMAND_SOCIETY = "&Sociedad"

MDI_MSG_NAME_TAKEN = "Este nombre ya esta cogido."
MDI_MSG_WRONG_ACCOUNT = "La cuenta no existe o es incorrecta."
MDI_MSG_WRONG_PASSWORD = "La contraseña es incorrecta."
MDI_MSG_BANNED = "Esta cuenta esta baneada."
MDI_MSG_UNLOAD = "¿Esta seguro que quiere cerrar a Peach?"

MDI_MSG_CANT_ADD_YOU = "No te puedes agregar a ti mismo."
MDI_MSG_ALREADY_IN_IGNORE_LIST = "'%u' ya esta en tu lista de ignorados."
MDI_MSG_ALREADY_IN_FRIEND_LIST = "'%u' ya esta en tu lista de amigos."
MDI_MSG_ACCOUNT_NOT_EXIST = "'%u' no existe."

MDI_MENU = "Menu"

CONFIG_LABEL_ACCOUNT = "Cuenta"
CONFIG_LABEL_PASSWORD = "Contraseña"

CONFIG_COMMAND_CONNECT = "&Conectar"
CONFIG_COMMAND_DISCONNECT = "&Desconectar"
CONFIG_COMMAND_SETTINGS = "&Ajustes"
CONFIG_COMMAND_UPDATE = "&Actualizar"
CONFIG_COMMAND_REGISTER = "&Crear una cuenta"
CONFIG_COMMAND_FORGOT_PASSWORD = "¿Ha olvidado contraseña?"

CONFIG_CHECK_SAVE_PASSWORD = "&Guardar contraseña"

CONFIG_FRAME_CONNECTION = "Informaciónes de conexión"

LANG_GERMAN = "Aleman"
LANG_ENGLISH = "Inglés"
LANG_SPANISH = "Español"
LANG_SWEDISH = "Sueco"
LANG_ITALIAN = "Italiano"
LANG_DUTCH = "Holandés"
LANG_SERBIAN = "Serbio"
LANG_FRENCH = "Frances"
LANG_NORWEGIAN = "Noruego"

CONFIG_MSG_ACCOUNT = "No has introducido una cuenta."
CONFIG_MSG_PASSWORD = "No has introducido una contraseña."
CONFIG_MSG_NUMERIC = "No puedes coger nombres con numeros."
CONFIG_MSG_PORT = "No has introducido un puerto."
CONFIG_MSG_IP = "No has introducido una direccion."
CONFIG_MSG_UPDATE_FILE = "Necesitas el Peach Updater para actualizar tu Peach." & vbCrLf & vbCrLf & "Descargalo aqui http://riplegion.ri.funpic.de/Peach/peachUpdater.exe"

CHAT_COMMAND_SEND = "&Enviar"
CHAT_COMMAND_CLEAR = "&Limpiar"

SF_LABEL_FILENAME = "Nombre del archivo"
SF_LABEL_SEND_TO = "Enviar a"
SF_LABEL_TIME = " Time left: "
SF_LABEL_KBS = " Kb/Sec, "
SF_LABEL_KBSS = " KBytes sent, "

SF_MSG_USER = "No ha seleccionado a una persona."
SF_MSG_FILE = "No ha seleccionado a un archivo."
SF_MSG_INCOMMING_FILE = "Esta recibiendo '%f' de '%u'. ¿Quiere aceptar?"
SF_MSG_DECILINED = "El envio ha sido rechazado."

SF_COMMAND_BROWSE = "&Buscar .."
SF_COMMAND_SENDFILE = "Enviar"
SF_COMMAND_CANCEL = "Cancelar .."

LANG_COMMAND_ENTER = "&Seleccionar"
LANG_LABEL_SELLANG = "Elige tu idioma"

LANG_QUIT = "¿Para cambiar el idioma tienes que reiniciar Peach, deseas hacerlo ahora?"

SOC_FRIEND_LIST = "Lista de contactos"
SOC_ONLINE_LIST = "Lista de online"
SOC_IGNORE_LIST = "Lista de ignorados"

SOC_COMMAND_REMOVE = "&Quitar"
SOC_COMMAND_FRIEND = "&Añardir a amigos"
SOC_COMMAND_IGNORE = "&Ignorar al usuario"
SOC_COMMAND_WHISPER = "&Susurrar"

SOC_ASK_DEL = "¿Estas seguro que quieres borrar a '%u' de la lista?"

SOC_ASK_FRIEND_TEXT = "Inserta el nombre de la cuenta de tu amigo aqui."
SOC_ASK_FRIEND_TITLE = "Añadir amigo"
SOC_ASK_FRIEND_DEFAULT = "Cuenta de tu amigo"

SOC_ASK_IGNORE_TEXT = "Inserta el nombre del usuario que quieres ignorar aqui."
SOC_ASK_IGNORE_TITLE = "Ignorar a usuario"
SOC_ASK_IGNORE_DEFAULT = "Cuenta de la persona que quieres ignorar"

SOC_FRIEND_LIST_STATUS = "Estatus"

SOC_MSG_CANT_WHISPER = "No puedes susurrar a este usuario."

REG_CAPTION = "Peach - Registración"

REG_FRAME_DETAIL = "Inserta su detalles"

REG_LABEL_ACCOUNT_NAME = "Nombre de cuenta"
REG_LABEL_PASSWORD = "Contraseña"
REG_LABEL_PASSWORD_CONFIRM = "Confirmar contraseña"
REG_LABEL_PASSWORD_WEAK = "La contraseña es floja."
REG_LABEL_PASSWORD_NORMAL = "La contraseña es normal."
REG_LABEL_PASSWORD_STRONG = "La contraseña es fuerte."
REG_LABEL_SECRET_QUESTION = "Pregunta secreta"
REG_LABEL_SECRET_ANSWER = "Respuesta secreta"

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

REG_MSG_EMAIL_EMPTY = "No ha introducido un correo electronico."
REG_MSG_EMAIL_TAKEN = "El correo electronico ya esta usado, ha olvidado su contraseña?"

REG_MSG_SUCCESSFULLY = "La cuenta ha sido registrada con exito."
REG_MSG_ERROR = "Un error ha occurido intenten de nuevo despues."
REG_MSG_ERROR_OCCURED = "Error occurido ..."
REG_MSG_LOADING = "Cargando .."
REG_MSG_CONNECTION_BROKEN = "La conexión se ha roto, intenten de nuevo despues."

REG_CMB_SECRET_QUESTION_0 = "¿Cual es el nombre de tu mascota?"
REG_CMB_SECRET_QUESTION_1 = "¿Tu libro favorito?"
REG_CMB_SECRET_QUESTION_2 = "¿Tu pelicula favorita?"
REG_CMB_SECRET_QUESTION_3 = "¿Tu juego favorito?"
REG_CMB_SECRET_QUESTION_4 = "¿Tu cantante favorito?"
REG_CMB_SECRET_QUESTION_5 = "¿El lugar de nacimiento de su madre?"

REG_LABEL_GENDER = "Sexo"

REG_CMB_GENDER_MALE = "Masculino"
REG_CMB_GENDER_FEMALE = "Femenino"

SET_LABEL_COLOR = "Color activo"
SET_LABEL_FONT = "Fuente"

SET_FRAME_STYLE = "Estilo"
SET_FRAME_OPTIONS = "Opciones"
SET_FRAME_CONNECTION = "Confgiuración de conexión"

SET_CHECK_SAVE_ACCOUNT = "Guardar cuenta"
SET_CHECK_SAVE_PASSWORD = "Guardar contraseña"
SET_CHECK_AUTO_LOGIN = "Logear automáticamente"
SET_CHECK_ASK_CLOSING = "Preguntar antes de cerrar"
SET_CHECK_MINIMIZE = "Minimizar ventana de Peach en la bandeja del sistema"

SET_COMMAND_LANGUAGE = "&Idioma"
SET_COMMAND_SAVE = "&Guardar"

SF2_COMMAND_OPEN_FILE = "&Abrir carpeta"

FP_FRAME_FORGOT_PASSWORD = "¿Ha olvidado contraseña?"
FP_LABEL_EMAIL = "Introduce su correo electronico"
FP_LABEL_SECRET_QUESTION = "Pregunta secreta"
FP_LABEL_SECRET_ANSWER = "Respuesta secreta"
FP_COMMAND_REQUEST = "&Solicitar"
FP_CAPTION = "Peach - Recuperar contraseña"

FP_MSG_SUCCESSFULL = "El nombre de su cuenta es '%u'." & vbCrLf & "Su contraseña es '%p'."
FP_MSG_WRONG_ANSWER = "La respuesta es incorrecta."
FP_MSG_WRONG_EMAIL = "El correo electronico no se ha encontrado."
End Sub

Public Sub SET_LANG_SWEDISH()
CURRENT_LANG = 3

' MDI form ..
MDI_COMMAND_CHAT = "Ch&att"
MDI_COMMAND_SENDFILE = "&Sänd fil"
MDI_COMMAND_SOCIETY = "&Samhälle"

'MDImsgbox_errorHandlerFormLoad
MDI_MSG_NAME_TAKEN = "Namnet är upptaget."
MDI_MSG_WRONG_ACCOUNT = "Kontot finns inte eller är felaktig."
MDI_MSG_WRONG_PASSWORD = "Lösenordet är fel."
MDI_MSG_BANNED = "Detta konto är förbjuden."
MDI_MSG_UNLOAD = "Är du säker på att du vill stänga Peach?"

MDI_MSG_CANT_ADD_YOU = "Du kan inte lägga till dig själv."
MDI_MSG_ALREADY_IN_IGNORE_LIST = "'%u' finns redan på din ignorerings lista."
MDI_MSG_ALREADY_IN_FRIEND_LIST = "'%u' finns redan på din kompis lista."
MDI_MSG_ACCOUNT_NOT_EXIST = "'%u' finns inte."

MDI_MENU = "Menu"

' Config form
CONFIG_LABEL_ACCOUNT = "Konto"
CONFIG_LABEL_PASSWORD = "Lösenord"

CONFIG_COMMAND_CONNECT = "&Anslut"
CONFIG_COMMAND_DISCONNECT = "&Frånkoppla"
CONFIG_COMMAND_SETTINGS = "&Inställningar"
CONFIG_COMMAND_REGISTER = "&Skapa konto"
CONFIG_COMMAND_UPDATE = "&Updatering"
CONFIG_COMMAND_FORGOT_PASSWORD = "Glömt lösenord"

LANG_GERMAN = "Tyska"
LANG_ENGLISH = "Engelska"
LANG_SPANISH = "Spanska"
LANG_SWEDISH = "Svenska"
LANG_ITALIAN = "Italienska"
LANG_SERBIAN = "Serbiska"
LANG_DUTCH = "Holländska"
LANG_FRENCH = "Franska"
LANG_NORWEGIAN = "Norska"

CONFIG_MSG_ACCOUNT = "Du skrev inte in en användare."
CONFIG_MSG_PASSWORD = "Du skrev inte in ett lösenord."
CONFIG_MSG_NUMERIC = "Du kan inte använda siffror i namnet."
CONFIG_MSG_PORT = "Du angav inget portnummer."
CONFIG_MSG_IP = "Du angav inte ett IP."
CONFIG_MSG_UPDATE_FILE = "Du behöver Peach updater för att uppgradera ditt peach." & vbCrLf & vbCrLf & "Ladda ner det här: http://riplegion.ri.funpic.de/Peach/peachUpdater.exe"

' Chat form ..
CHAT_COMMAND_SEND = "&Sänd"
CHAT_COMMAND_CLEAR = "&Rensa"

' Send file form ..
SF_LABEL_FILENAME = "Fil Namn"
SF_LABEL_SEND_TO = "Skicka till"
SF_LABEL_TIME = " Time left: "
SF_LABEL_KBS = " Kb/Sec, "
SF_LABEL_KBSS = " KBytes sent, "

SF_MSG_USER = "Ingen användare vald."
SF_MSG_FILE = "Ingen fil vald."
SF_MSG_INCOMMING_FILE = "Du har tagit emot '%f' från '%u'. Vill du acceptera?"
SF_MSG_DECILINED = "Filöverföringen var nekad."

SF_COMMAND_BROWSE = "&Sök .."
SF_COMMAND_SENDFILE = "Sänd"

SF2_COMMAND_OPEN_FILE = "&Öppna fil map"

LANG_COMMAND_ENTER = "&Öppna"
LANG_LABEL_SELLANG = "Välj språk"

LANG_QUIT = "För att ändra språk måste du starta om Peach, vill du göra det nu?"

SOC_FRIEND_LIST = "Kompis Lista"
SOC_ONLINE_LIST = "Online Lista"
SOC_IGNORE_LIST = "Ignorerings Lista"

SOC_COMMAND_REMOVE = "&Ta bort"
SOC_COMMAND_FRIEND = "&Lägg till vänner"
SOC_COMMAND_IGNORE = "&Ignorera user"
SOC_COMMAND_WHISPER = "&Viska"

SOC_ASK_DEL = "Vill du ta bort '%u' från listan?"

SOC_ASK_FRIEND_TEXT = "Ange kontonamn på din vän i textrutan nedan."
SOC_ASK_FRIEND_TITLE = "Lägg till en vän"
SOC_ASK_FRIEND_DEFAULT = "Ange konto här"

SOC_ASK_IGNORE_TEXT = "Skriv in kontonamnet för användaren som du vill ignorera i textrutan nedan."
SOC_ASK_IGNORE_TITLE = "Ignorera en användare"
SOC_ASK_IGNORE_DEFAULT = "Ange konto här"

SOC_FRIEND_LIST_STATUS = "Status"

SOC_MSG_CANT_WHISPER = "Du kan inte viska här användaren."

REG_CAPTION = "Peach - Registrering"

REG_FRAME_DETAIL = "Ange dina detaljer"

REG_LABEL_ACCOUNT_NAME = "Användar Namn"
REG_LABEL_PASSWORD = "Lösenord"
REG_LABEL_PASSWORD_CONFIRM = "Bekräfta lösenord"
REG_LABEL_PASSWORD_WEAK = "Lösenordet är lätt."
REG_LABEL_PASSWORD_NORMAL = "Lösenordet är normalt."
REG_LABEL_PASSWORD_STRONG = "Lösenordet är svårt."
REG_LABEL_SECRET_QUESTION = "Säkerhetsfråga"
REG_LABEL_SECRET_ANSWER = "Säkerhet besvara"

REG_COMMAND_SUBMIT = "&Acceptera"
REG_COMMAND_CLOSE = "&Ständ"

REG_CHECK_PASSWORD_SHOW = "&Visa lösenord."

REG_MSG_ACCOUNT_EXIST = "Namnet är upptagvet."
REG_MSG_ACCOUNT_INVALID = "Ogiltigt namn."
REG_MSG_ACCOUNT_NUMERIC = "Namnet kan inte bestå av number."
REG_MSG_ACCOUNT_EMPTY = "Inget namn angivet."
REG_MSG_ACCOUNT_SHORT = "För kort namn, det kräver åtminstone 4 bokstäver."

REG_MSG_PASSWORD_MATCH = "Ogiltigt lösenord."
REG_MSG_PASSWORD_SHORT = "För kort lösenord, det kräver åtminstone 6 bokstäver."
REG_MSG_PASSWORD_EMPTY = "Inget lösenord angivet."

REG_MSG_SECRET_ANSWER_EMPTY = "Inga hemliga svaret infördes."

REG_MSG_EMAIL_EMPTY = "Inga hemliga email."
REG_MSG_EMAIL_TAKEN = "The email adress you have entered is already beeing used, have you forgot your password?"

REG_MSG_SUCCESSFULLY = "Kontot har skapats."
REG_MSG_ERROR = "Ett fel har uppstått var snäll och försök igen."
REG_MSG_ERROR_OCCURED = "Ett fel har uppstått ..."
REG_MSG_LOADING = "Laddar .."
REG_MSG_CONNECTION_BROKEN = "Anslutnings fel, var snäll och försök igen."

REG_CMB_SECRET_QUESTION_0 = "Vad heter ditt husdjur?"
REG_CMB_SECRET_QUESTION_1 = "Vilken är din favoritbok?"
REG_CMB_SECRET_QUESTION_2 = "Vilken är din favoritfilm?"
REG_CMB_SECRET_QUESTION_3 = "Vilket är ditt favoritspel?"
REG_CMB_SECRET_QUESTION_4 = "Vilken är din favorit sångare?"
REG_CMB_SECRET_QUESTION_5 = "Var är den plats där din mor föddes?"

REG_LABEL_GENDER = "Kön"

REG_CMB_GENDER_MALE = "Manlig"
REG_CMB_GENDER_FEMALE = "Kvinna"

SET_LABEL_COLOR = "Nuvarande färg"
SET_LABEL_FONT = "Textsnitt"

SET_FRAME_STYLE = "Stilart"
SET_FRAME_OPTIONS = "Alternativ"
SET_FRAME_CONNECTION = "Anslutnings inställningar"

SET_CHECK_SAVE_ACCOUNT = "Spara konto"
SET_CHECK_SAVE_PASSWORD = "Spara lösenord"
SET_CHECK_AUTO_LOGIN = "Logga in automatiskt"
SET_CHECK_ASK_CLOSING = "Fråga innan stäng"
SET_CHECK_MINIMIZE = "Minimera Peach-fönstret till Aktivitetsfältet"

SET_COMMAND_LANGUAGE = "&Språk"
SET_COMMAND_SAVE = "&Spara"

FP_FRAME_FORGOT_PASSWORD = "Glömt lösenord"
FP_LABEL_EMAIL = "Ange ditt email"
FP_LABEL_SECRET_QUESTION = "Säkerhetsfråga"
FP_LABEL_SECRET_ANSWER = "Säkerhet besvara"
FP_COMMAND_REQUEST = "&Begära"
FP_CAPTION = "Peach - Glömt lösenord"

FP_MSG_SUCCESSFULL = "Your username is '%u'." & vbCrLf & "Your password is '%p'."
FP_MSG_WRONG_ANSWER = "Svaret är fel."
FP_MSG_WRONG_EMAIL = "Den angivna e-postadressen kunde inte hittas."
End Sub

Public Sub SET_LANG_ITALIAN()
CURRENT_LANG = 4

' Mdi form
MDI_COMMAND_CHAT = "Ch&at"
MDI_COMMAND_SENDFILE = "&Invia File"
MDI_COMMAND_SOCIETY = "&Società"

'MDImsgbox_errorHandlerFormLoad
MDI_MSG_NAME_TAKEN = "Il nome immesso e' gia' in uso."
MDI_MSG_WRONG_ACCOUNT = "L'account non esiste o è sbagliato."
MDI_MSG_WRONG_PASSWORD = "La password è errata."
MDI_MSG_BANNED = "Questo account è vietata.."
MDI_MSG_UNLOAD = "Sei sicuro di voler chiudere Peach?"

MDI_MSG_CANT_ADD_YOU = "Non puoi aggiungere te."
MDI_MSG_ALREADY_IN_IGNORE_LIST = "'%u' è già nella vostra lista ignorare."
MDI_MSG_ALREADY_IN_FRIEND_LIST = "'%u' è già nella vostra lista amici."
MDI_MSG_ACCOUNT_NOT_EXIST = "'%u' non esiste."

MDI_MENU = "Menu"

'Config form ..
CONFIG_LABEL_ACCOUNT = "Conto"
CONFIG_LABEL_PASSWORD = "Password"

CONFIG_COMMAND_CONNECT = "&Connesso"
CONFIG_COMMAND_DISCONNECT = "&Disconnesso"
CONFIG_COMMAND_SETTINGS = "&Impostazioni    "
CONFIG_COMMAND_UPDATE = "&Aggiornamento"
CONFIG_COMMAND_REGISTER = "&Crea un account"
CONFIG_COMMAND_FORGOT_PASSWORD = "Hai dimenticato la password"

CONFIG_CHECK_SAVE_PASSWORD = "&Salva password"

LANG_GERMAN = "Tedesco"
LANG_ENGLISH = "Inglese"
LANG_SPANISH = "Spagnolo"
LANG_SWEDISH = "Svedese"
LANG_ITALIAN = "Italiano"
LANG_SERBIAN = "Serbo"
LANG_DUTCH = "Olandese"
LANG_FRENCH = "Francese"
LANG_NORWEGIAN = "Norvegese"

CONFIG_MSG_ACCOUNT = "Non hai inserito un account."
CONFIG_MSG_PASSWORD = "Non hai inserito una password."
CONFIG_MSG_NUMERIC = "Non puoi immettere nomi composti da numeri."
CONFIG_MSG_PORT = "Non hai selezionato una porta valida."
CONFIG_MSG_IP = "Non hai immesso un IP."
CONFIG_MSG_UPDATE_FILE = "You need the peach updater to be able to upgrade your peach." & vbCrLf & vbCrLf & "Download it here: http://riplegion.ri.funpic.de/Peach/peachUpdater.exe"

' Chat form ..
CHAT_COMMAND_SEND = "&Invia"
CHAT_COMMAND_CLEAR = "&Chiaro"

' Send file form ..
SF_LABEL_FILENAME = "Nome file"
SF_LABEL_SEND_TO = "Send to"
SF_LABEL_TIME = " Time left: "
SF_LABEL_KBS = " Kb/Sec, "
SF_LABEL_KBSS = " KBytes sent, "

SF_MSG_USER = "Nessun utente selezionato."
SF_MSG_FILE = "Nessun file selezionato."
SF_MSG_INCOMMING_FILE = "Stai ricevendo '%f' da '%u'. Volete accettare?"
SF_MSG_DECILINED = "Il trasferimento file è stato rifiutato."

SF_COMMAND_BROWSE = "&Cerca .."
SF_COMMAND_SENDFILE = "Invia"
SF_COMMAND_CANCEL = "Annulla .."

LANG_COMMAND_ENTER = "&Apri"
LANG_LABEL_SELLANG = "Seleziona la tua lingua"

LANG_QUIT = "Al fine di modificare la lingua è necessario riavviare Peach, vuoi farlo ora?"

SOC_FRIEND_LIST = "Lista di amici"
SOC_ONLINE_LIST = "Elenco di persone online"
SOC_IGNORE_LIST = "Elenco degli utenti ignorati"

SOC_COMMAND_REMOVE = "&Rimuovere"
SOC_COMMAND_FRIEND = "&Aggiungi ai tuoi amici"
SOC_COMMAND_IGNORE = "&Ignora utente"
SOC_COMMAND_WHISPER = "&Sussurro"

SOC_ASK_DEL = "Vuoi eliminare '%u' dalla lista?"

SOC_ASK_FRIEND_TEXT = "Inserire il nome account del tuo amico nella casella di testo sottostante."
SOC_ASK_FRIEND_TITLE = "Aggiunta di un amico"
SOC_ASK_FRIEND_DEFAULT = "Conto inserisci qui"

SOC_ASK_IGNORE_TEXT = "Inserire il nome account dell'utente che si desidera ignorare nella casella di testo sottostante."
SOC_ASK_IGNORE_TITLE = "Ignorare un utente"
SOC_ASK_IGNORE_DEFAULT = "Conto inserisci qui"

SOC_FRIEND_LIST_STATUS = "Stato"

SOC_MSG_CANT_WHISPER = "Non si può sussurrare questo utente."

REG_CAPTION = "Peach - Registrazione"

REG_FRAME_DETAIL = "Inserisci i tuoi dati"

REG_LABEL_ACCOUNT_NAME = "Nome account"
REG_LABEL_PASSWORD = "Parola d'ordine"
REG_LABEL_PASSWORD_CONFIRM = "Confermare la password"
REG_LABEL_PASSWORD_WEAK = "La password è debole."
REG_LABEL_PASSWORD_NORMAL = "La password è normale."
REG_LABEL_PASSWORD_STRONG = "La password è forte."
REG_LABEL_SECRET_QUESTION = "Domanda segreta"
REG_LABEL_SECRET_ANSWER = "Risposta segreta"

REG_COMMAND_SUBMIT = "&Inoltrare"
REG_COMMAND_CLOSE = "&Chiudere"

REG_CHECK_PASSWORD_SHOW = "&Visualizzare Password"

REG_MSG_ACCOUNT_EXIST = "Il nome di account già esistente."
REG_MSG_ACCOUNT_INVALID = "Nome non valido account."
REG_MSG_ACCOUNT_NUMERIC = "Conto non può essere composta da caratteri numerici."
REG_MSG_ACCOUNT_EMPTY = "Nessun account è stato inserito."
REG_MSG_ACCOUNT_SHORT = "Il nome dell'account è troppo breve."

REG_MSG_PASSWORD_MATCH = "Le password non corrispondono."
REG_MSG_PASSWORD_SHORT = "La password è troppo corta."
REG_MSG_PASSWORD_EMPTY = "Nessuna password è stata inserita."

REG_MSG_SECRET_ANSWER_EMPTY = "Risposta segreta non è stato iscritto."

REG_MSG_EMAIL_EMPTY = "Email non è stato iscritto."
REG_MSG_EMAIL_TAKEN = "L'indirizzo e-mail che hai inserito è già utilizzato beeing, hai dimenticato la parola d'ordine?"

REG_MSG_SUCCESSFULLY = "L'account è stato registrato con successo."
REG_MSG_ERROR = "Un errore si è verificato per favore riprova più tardi di nuovo."
REG_MSG_ERROR_OCCURED = "È verificato un errore ..."
REG_MSG_LOADING = "Carico .."
REG_MSG_CONNECTION_BROKEN = "Connessione viene interrotta per favore riprova più tardi."

REG_CMB_SECRET_QUESTION_0 = "Qual è il nome del vostro animale domestico?"
REG_CMB_SECRET_QUESTION_1 = "Qual è il tuo libro preferito?"
REG_CMB_SECRET_QUESTION_2 = "Qual è il vostro film preferito?"
REG_CMB_SECRET_QUESTION_3 = "Qual è il tuo gioco preferito?"
REG_CMB_SECRET_QUESTION_4 = "Qual è il vostro cantante preferito?"
REG_CMB_SECRET_QUESTION_5 = "Dove si trova il luogo in cui tua madre è nata?"

REG_LABEL_GENDER = "Genere"

REG_CMB_GENDER_MALE = "Maschio"
REG_CMB_GENDER_FEMALE = "Femminile"

SET_LABEL_COLOR = "Colore corrente"
SET_LABEL_FONT = "Fonte"

SET_FRAME_STYLE = "Stile"
SET_FRAME_OPTIONS = "Opzioni"
SET_FRAME_CONNECTION = "Impostazioni di connessione"

SET_CHECK_SAVE_ACCOUNT = "Salva conto"
SET_CHECK_SAVE_PASSWORD = "Salva parola d'ordine"
SET_CHECK_AUTO_LOGIN = "Login automatically"
SET_CHECK_ASK_CLOSING = "Chiedi prima di chiudere"
SET_CHECK_MINIMIZE = "Contrai la finestra di Peach nella barra delle applicazioni"

SET_COMMAND_LANGUAGE = "&Lingua"
SET_COMMAND_SAVE = "&Salva"

SF2_COMMAND_OPEN_FILE = "&Aprire la cartella File"

FP_FRAME_FORGOT_PASSWORD = "Dimenticato la password"
FP_LABEL_EMAIL = "Inserisci il tuo email"
FP_LABEL_SECRET_QUESTION = "Domanda segreta"
FP_LABEL_SECRET_ANSWER = "Risposta segreta"
FP_COMMAND_REQUEST = "&Richiesta"
FP_CAPTION = "Peach - Dimenticato la password"

FP_MSG_SUCCESSFULL = "Il tuo nome account è '%u'." & vbCrLf & " La password è '%p'."
FP_MSG_WRONG_ANSWER = "La risposta è sbagliata."
FP_MSG_WRONG_EMAIL = "L'indirizzo email inserito non è stato trovato."
End Sub

Public Sub SET_LANG_DUTCH()
CURRENT_LANG = 5

MDI_COMMAND_CHAT = "Ch&at"
MDI_COMMAND_SENDFILE = "&Bestand Verzenden"
MDI_COMMAND_SOCIETY = "&Gezelschap"

'MDImsgbox_errorHandlerFormLoad
MDI_MSG_NAME_TAKEN = "Deze naam is niet beschikbaar."
MDI_MSG_WRONG_ACCOUNT = "De account bestaat niet of is verkeerd."
MDI_MSG_WRONG_PASSWORD = "Het wachtwoord is onjuist."
MDI_MSG_BANNED = "Deze account is verboden."
MDI_MSG_UNLOAD = "Weet u zeker dat u wilt Peach sluiten?"

MDI_MSG_CANT_ADD_YOU = "Je kunt jezelf niet toevoegen."
MDI_MSG_ALREADY_IN_IGNORE_LIST = "'%u' is al in uw negeerlijst."
MDI_MSG_ALREADY_IN_FRIEND_LIST = "'%u' is al in uw vriendenlijst."
MDI_MSG_ACCOUNT_NOT_EXIST = "'%u' bestaat niet."

MDI_MENU = "Menu"

CONFIG_LABEL_ACCOUNT = "Rekening"
CONFIG_LABEL_PASSWORD = "Wachtwoord"

CONFIG_COMMAND_CONNECT = "&Verbind"
CONFIG_COMMAND_DISCONNECT = "&Verbinding verbreken"
CONFIG_COMMAND_SETTINGS = "&Instellingen"
CONFIG_COMMAND_REGISTER = "&Account aanmaken"
CONFIG_COMMAND_UPDATE = "&Update"
CONFIG_COMMAND_FORGOT_PASSWORD = "Wachtwoord vergeten"

CONFIG_CHECK_SAVE_PASSWORD = "&Wachtwoord opslaan"

LANG_GERMAN = "Duits"
LANG_ENGLISH = "Engels"
LANG_SPANISH = "Spaans"
LANG_SWEDISH = "Zweeds"
LANG_ITALIAN = "Italiaans"
LANG_SERBIAN = "Serbisch"
LANG_DUTCH = "Nederlands"
LANG_FRENCH = "Frans"
LANG_NORWEGIAN = "Noors"

CONFIG_MSG_ACCOUNT = "Je hebt geen gebruikersnaam ingevuld."
CONFIG_MSG_PASSWORD = "Je hebt geen wachtwoord ingevuld."
CONFIG_MSG_NUMERIC = "U kan geen naam nemen dat nummers bevat."
CONFIG_MSG_PORT = "U hebt geen poort ingesteld."
CONFIG_MSG_IP = "U hebt geen IP gegoven."
CONFIG_MSG_UPDATE_FILE = "Om je Peach versie up-to-date te houden moet je de Peach updater downloaden." & vbCrLf & vbCrLf & "Klik hier om te downloaden: http://riplegion.ri.funpic.de/Peach/peachUpdater.exe"

CHAT_COMMAND_SEND = "&Zend"
CHAT_COMMAND_CLEAR = "&Leegmaken"

SF_LABEL_FILENAME = "Bestandsnaam"
SF_LABEL_SEND_TO = "Stuur naar"
SF_LABEL_TIME = " Resterende tijd: "
SF_LABEL_KBS = " Kb/Sec, "
SF_LABEL_KBSS = " KBytes sent, "

SF_MSG_USER = "Geen gebruiker geselecteerd."
SF_MSG_FILE = "Geen bestand geselecteerd."
SF_MSG_INCOMMING_FILE = "U ontvangt '%f' van '%u'. Wil je accepteren?"
SF_MSG_DECILINED = "De bestandsoverdracht is geweigerd."

SF_COMMAND_BROWSE = "&Zoeken .."
SF_COMMAND_SENDFILE = "&Stuur"
SF_COMMAND_CANCEL = "&Annuleren .."

LANG_COMMAND_ENTER = "&Openen"
LANG_LABEL_SELLANG = "Selecteer jou taal"

LANG_QUIT = "Om je taal te verranderen is het nodig om Peach opnieuw op te starten. Wil je dit nu doen?"

SOC_FRIEND_LIST = "Vriendenlijst"
SOC_ONLINE_LIST = "Onlinelijst"
SOC_IGNORE_LIST = "Negeerlijst"

SOC_COMMAND_REMOVE = "&Verwijderen"
SOC_COMMAND_FRIEND = "&Voeg toe aan vrienden"
SOC_COMMAND_IGNORE = "&Gebruiker negeren"
SOC_COMMAND_WHISPER = "&Fluister"

SOC_ASK_DEL = "Wilt u '%u' verwijderen uit de lijst?"

SOC_ASK_FRIEND_TEXT = "Vul de naam in van je vriend in de kader hieronder."
SOC_ASK_FRIEND_TITLE = "Vriend aan het toevoegen"
SOC_ASK_FRIEND_DEFAULT = "Voeg hier uw accountnaam in"

SOC_ASK_IGNORE_TEXT = "Voeg hier de naam van de gebruiker in die je wilt negeren."
SOC_ASK_IGNORE_TITLE = "Negeer een gebruiker"
SOC_ASK_IGNORE_DEFAULT = "Voeg hier uw accountnaam in"

SOC_FRIEND_LIST_STATUS = "Staat"

SOC_MSG_CANT_WHISPER = "Je kan deze persoon niet fluisteren."

REG_CAPTION = "Peach - Registratie"

REG_FRAME_DETAIL = "Voeg je gegevens in"

REG_LABEL_ACCOUNT_NAME = "Gebruikersnaam"
REG_LABEL_PASSWORD = "Wachtwoord"
REG_LABEL_PASSWORD_CONFIRM = "Bevestig Wachtwoord"
REG_LABEL_PASSWORD_WEAK = "Dit wachtwoord is zwak."
REG_LABEL_PASSWORD_NORMAL = "Dit wachtwoord is redelijk."
REG_LABEL_PASSWORD_STRONG = "Dit wachtwoord is goed."
REG_LABEL_SECRET_QUESTION = "Geheime vraag"
REG_LABEL_SECRET_ANSWER = "Geheim antwoord"

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

REG_MSG_SECRET_ANSWER_EMPTY = "Geen geheim beantwoord opgenomen."

REG_MSG_EMAIL_EMPTY = "Geen email opgenomen."
REG_MSG_EMAIL_TAKEN = "De email adres die je ingevoerd hebt wordt al gebruikt, ben je je passwoord vergeten?"

REG_MSG_SUCCESSFULLY = "Account succesvol aangemaakt."
REG_MSG_ERROR = "Er is een fout opgetreden, probeer het later opnieuw."
REG_MSG_ERROR_OCCURED = "Fout opgetreden ..."
REG_MSG_LOADING = "Laden.. "
REG_MSG_CONNECTION_BROKEN = "Verbinding verbroken."

REG_CMB_SECRET_QUESTION_0 = "Wat is de naam van uw huisdier?"
REG_CMB_SECRET_QUESTION_1 = "Wat is uw favoriete boek?"
REG_CMB_SECRET_QUESTION_2 = "Wat is je favoriete film?"
REG_CMB_SECRET_QUESTION_3 = "Wat is je favoriete spel?"
REG_CMB_SECRET_QUESTION_4 = "Wat is uw favoriete zanger?"
REG_CMB_SECRET_QUESTION_5 = "Waar is de plaats waar je moeder is geboren?"

REG_LABEL_GENDER = "Geslacht"

REG_CMB_GENDER_MALE = "Man"
REG_CMB_GENDER_FEMALE = "Vrouw"

SET_LABEL_COLOR = "Momenteel gebruikte kleur"
SET_LABEL_FONT = "Doopvont"

SET_FRAME_STYLE = "Stijl"
SET_FRAME_OPTIONS = "Instellingen"
SET_FRAME_CONNECTION = "Verbinding Instellingen"

SET_CHECK_SAVE_ACCOUNT = "Gebruiker opslaan"
SET_CHECK_SAVE_PASSWORD = "Wachtwoord opslaan"
SET_CHECK_AUTO_LOGIN = "Login automatically"
SET_CHECK_ASK_CLOSING = "Vraag voordat sluiten"
SET_CHECK_MINIMIZE = "Peach venster minimaliseren naar het systeemvak"

SET_COMMAND_LANGUAGE = "&Taal"
SET_COMMAND_SAVE = "&Opslaan"

SF_LABEL_SEND_TO = "Verzenden naar"

SF_MSG_USER = "Geen gebruiker geselecteerd."
SF_MSG_FILE = "Geen bestand geselecteerd."
SF_MSG_INCOMMING_FILE = "Je bent nu '%f' aan het ontvangen van '%u' Accepteer je?"
SF_MSG_DECILINED = "Gegevensoverdracht geweigerd."

SF2_COMMAND_OPEN_FILE = "&Open bestandsmap"

FP_FRAME_FORGOT_PASSWORD = "Wachtwoord vergeten"
FP_LABEL_EMAIL = "Voer uw email"
FP_LABEL_SECRET_QUESTION = "Geheime vraag"
FP_LABEL_SECRET_ANSWER = "Geheim antwoord"
FP_COMMAND_REQUEST = "&Verzoeken"
FP_CAPTION = "Peach - Wachtwoord vergeten"

FP_MSG_SUCCESSFULL = "Your account name is '%u'." & vbCrLf & "Your password is '%p'."
FP_MSG_WRONG_ANSWER = "Het antwoord is fout."
FP_MSG_WRONG_EMAIL = "The entered email adress could not be found."
End Sub

Public Sub SET_LANG_SERBIAN()
CURRENT_LANG = 6

' Mdi form ..
MDI_COMMAND_CHAT = "Ch&at"
MDI_COMMAND_SENDFILE = "&Slanje fajla"
MDI_COMMAND_SOCIETY = "&Onlajn lista"

'MDImsgbox_errorHandlerFormLoad
MDI_MSG_NAME_TAKEN = "Ime je vec zauzeto."
MDI_MSG_WRONG_ACCOUNT = "Akaunt ne postoji ili je pogresan."
MDI_MSG_WRONG_PASSWORD = "Sifra je pogresna."
MDI_MSG_BANNED = "Profil je banovan."
MDI_MSG_UNLOAD = "Da li si siguran da zelis da zatvoris Peach?"

MDI_MSG_CANT_ADD_YOU = "Ne mozes dodati sebe."
MDI_MSG_ALREADY_IN_IGNORE_LIST = "'%u' je vec u tvojoj ignor listi."
MDI_MSG_ALREADY_IN_FRIEND_LIST = "'%u' je vec u tvojim prijateljima."
MDI_MSG_ACCOUNT_NOT_EXIST = "'%u' ne postoji."

MDI_MENU = "Menu"

'Config form ..
CONFIG_LABEL_ACCOUNT = "Profil"
CONFIG_LABEL_PASSWORD = "Sifra"

CONFIG_COMMAND_CONNECT = "&Povezi se"
CONFIG_COMMAND_DISCONNECT = "&Veza je prekinuta"
CONFIG_COMMAND_SETTINGS = "&Podesavanje"
CONFIG_COMMAND_UPDATE = "&Update"
CONFIG_COMMAND_REGISTER = "&Napravi profil"
CONFIG_COMMAND_FORGOT_PASSWORD = "Zaboravio sifru"

CONFIG_CHECK_SAVE_PASSWORD = "&Sacuvaj sifru"

LANG_GERMAN = "Nemacki"
LANG_ENGLISH = "Engleski"
LANG_SPANISH = "Spanski"
LANG_SWEDISH = "Svedski"
LANG_ITALIAN = "Italijanski"
LANG_SERBIAN = "Srpski"
LANG_DUTCH = "Holandski"
LANG_FRENCH = "Francuski"
LANG_NORWEGIAN = "Norwegian"

CONFIG_MSG_ACCOUNT = "Nisi ukucao lozinku."
CONFIG_MSG_PASSWORD = "Nisi ukucao sifru."
CONFIG_MSG_NUMERIC = "Ne mozete uzeti numericka imena."
CONFIG_MSG_PORT = "Niste uneli port."
CONFIG_MSG_IP = "Niste uneli IP"
CONFIG_MSG_UPDATE_FILE = "You need the peach updater to be able to upgrade your peach." & vbCrLf & vbCrLf & "Download it here: http://riplegion.ri.funpic.de/Peach/peachUpdater.exe"

'Chat form ..
CHAT_COMMAND_SEND = "&Posalji"
CHAT_COMMAND_CLEAR = "&Obrisi"

'Send file form ..
SF_LABEL_FILENAME = "Ime  arhive"
SF_LABEL_SEND_TO = "Send to"
SF_LABEL_TIME = " Time left: "
SF_LABEL_KBS = " Kb/Sec, "
SF_LABEL_KBSS = " KBytes sent, "

SF_MSG_USER = "Nije izabran korisnik."
SF_MSG_FILE = "Podatak nije izabran."
SF_MSG_INCOMMING_FILE = "You are receiving '%f' from '%u'. Do you want to accept?"
SF_MSG_DECILINED = "The file transfer has been refused."

SF_COMMAND_BROWSE = "Trazi .."
SF_COMMAND_SENDFILE = "Posalji"
SF_COMMAND_CANCEL = "Otkazhi .."

LANG_COMMAND_ENTER = "&Otvori"
LANG_LABEL_SELLANG = "Dodaj svoj jezik"

LANG_QUIT = "In order to change the language you need to restart Peach, do you want to do this now?"

SOC_FRIEND_LIST = "Lista prijatelja"
SOC_ONLINE_LIST = "Lista onlajn"

SOC_COMMAND_REMOVE = "&Ukloni"
SOC_COMMAND_FRIEND = "&Dodaj u Prijatelje"
SOC_COMMAND_IGNORE = "&Ignorishi korisnika"
SOC_COMMAND_WHISPER = "&Privatna Poruka"

SOC_ASK_DEL = "Do you want to delete '%u' from the list?"

SOC_ASK_FRIEND_TEXT = "Enter the account name of your friend in the text box below."
SOC_ASK_FRIEND_TITLE = "Adding a friend"
SOC_ASK_FRIEND_DEFAULT = "Enter account here"

SOC_ASK_IGNORE_TEXT = "Enter the account name of the user you want to ignore in the text box below."
SOC_ASK_IGNORE_TITLE = "Ignore a user"
SOC_ASK_IGNORE_DEFAULT = "Enter account here"

SOC_FRIEND_LIST_STATUS = "Status"

SOC_MSG_CANT_WHISPER = "You can't whisper this user."

REG_CAPTION = "Peach - Registration"

REG_FRAME_DETAIL = "Enter your details"

REG_LABEL_ACCOUNT_NAME = "Account Name"
REG_LABEL_PASSWORD = "Password"
REG_LABEL_PASSWORD_CONFIRM = "Confirm the Password"
REG_LABEL_PASSWORD_WEAK = "The password is weak."
REG_LABEL_PASSWORD_NORMAL = "The password is normal."
REG_LABEL_PASSWORD_STRONG = "The password is strong."
REG_LABEL_SECRET_QUESTION = "Secret question"
REG_LABEL_SECRET_ANSWER = "Secret answer"

REG_COMMAND_SUBMIT = "&Submit"
REG_COMMAND_CLOSE = "&Close"

REG_CHECK_PASSWORD_SHOW = "&Show Password."

REG_MSG_ACCOUNT_EXIST = "The account name already exists."
REG_MSG_ACCOUNT_INVALID = "Invalid account name."
REG_MSG_ACCOUNT_NUMERIC = "Account can not be composed of numeric characters."
REG_MSG_ACCOUNT_EMPTY = "No account entered."
REG_MSG_ACCOUNT_SHORT = "Account name to short, it requieres at least 4 characters."

REG_MSG_PASSWORD_MATCH = "The passwords dont match."
REG_MSG_PASSWORD_SHORT = "Password to short, it requieres at least 6 characters."
REG_MSG_PASSWORD_EMPTY = "No Password entered."

REG_MSG_SECRET_ANSWER_EMPTY = "No secret answered entered."

REG_MSG_EMAIL_EMPTY = "No email entered."
REG_MSG_EMAIL_TAKEN = "The email adress you have entered is already beeing used, have you forgot your password?"

REG_MSG_SUCCESSFULLY = "The account was successfully registered."
REG_MSG_ERROR = "An error has occured please try later again."
REG_MSG_ERROR_OCCURED = "Error has occured ..."
REG_MSG_LOADING = "Loading .."
REG_MSG_CONNECTION_BROKEN = "Connection is broken please try again later."

REG_CMB_SECRET_QUESTION_0 = "What is the name of your pet?"
REG_CMB_SECRET_QUESTION_1 = "Your favorite book?"
REG_CMB_SECRET_QUESTION_2 = "Your favorite movie?"
REG_CMB_SECRET_QUESTION_3 = "Your favorite game?"
REG_CMB_SECRET_QUESTION_4 = "Your favorite singer?"
REG_CMB_SECRET_QUESTION_5 = "The place where your mother was born?"

REG_LABEL_GENDER = "Gender"

REG_CMB_GENDER_MALE = "Male"
REG_CMB_GENDER_FEMALE = "Female"

SET_LABEL_COLOR = "Current Color"
SET_LABEL_FONT = "Font"

SET_FRAME_STYLE = "Style"
SET_FRAME_OPTIONS = "Options"
SET_FRAME_CONNECTION = "Connection Settings"

SET_CHECK_SAVE_ACCOUNT = "Save Account"
SET_CHECK_SAVE_PASSWORD = "Save Password"
SET_CHECK_AUTO_LOGIN = "Login automatically"
SET_CHECK_ASK_CLOSING = "Ask before closing"
SET_CHECK_MINIMIZE = "Minimize Peach window to system tray"

SET_COMMAND_LANGUAGE = "&Language"
SET_COMMAND_SAVE = "&Save"

SF2_COMMAND_OPEN_FILE = "&Open File Folder"

FP_FRAME_FORGOT_PASSWORD = "Forgot Password"
FP_LABEL_EMAIL = "Enter your email adress"
FP_LABEL_SECRET_QUESTION = "Secret Question"
FP_LABEL_SECRET_ANSWER = "Secret Answer"
FP_COMMAND_REQUEST = "&Request"
FP_CAPTION = "Peach - Forgot Password"

FP_MSG_SUCCESSFULL = "Your account name is '%u'." & vbCrLf & "Your password is '%p'."
FP_MSG_WRONG_ANSWER = "The answer is wrong."
FP_MSG_WRONG_EMAIL = "The entered email adress could not be found."
End Sub

Public Sub SET_LANG_FRENCH()
CURRENT_LANG = 7

MDI_COMMAND_CHAT = "Ch&at"
MDI_COMMAND_SENDFILE = "&Envoyer un fichier"
MDI_COMMAND_SOCIETY = "&Société"

MDI_MSG_NAME_TAKEN = "Le nom inséré est déjà utilizé."
MDI_MSG_WRONG_ACCOUNT = "The account does not exist or is wrong."
MDI_MSG_WRONG_PASSWORD = "The password is wrong."
MDI_MSG_BANNED = "This account is banned."
MDI_MSG_UNLOAD = "Are you sure you want to close Peach?"

MDI_MSG_CANT_ADD_YOU = "You can't add yourself."
MDI_MSG_ALREADY_IN_IGNORE_LIST = "'%u' is already in your ignore list."
MDI_MSG_ALREADY_IN_FRIEND_LIST = "'%u' is already in your friend list."
MDI_MSG_ACCOUNT_NOT_EXIST = "'%u' does not exist."

MDI_MENU = "Menu"

CONFIG_LABEL_ACCOUNT = "Compte"
CONFIG_LABEL_PASSWORD = "Mot de passe"

CONFIG_COMMAND_CONNECT = "&Connecté"
CONFIG_COMMAND_DISCONNECT = "&Deconnecté"
CONFIG_COMMAND_SETTINGS = "&Paramètres"
CONFIG_COMMAND_UPDATE = "&Mettre à jour"
CONFIG_COMMAND_REGISTER = "&Créer un compte"
CONFIG_COMMAND_FORGOT_PASSWORD = "&Mot de passe perdu"

CONFIG_CHECK_SAVE_PASSWORD = "&Sauvegarder mot de passe"

LANG_GERMAN = "Alleman"
LANG_ENGLISH = "Anglais"
LANG_SPANISH = "Espagnol"
LANG_SWEDISH = "Suédois"
LANG_ITALIAN = "Italien"
LANG_SERBIAN = "Serbois"
LANG_DUTCH = "Hollandais"
LANG_FRENCH = "Français"
LANG_NORWEGIAN = "Norvégien"

CONFIG_MSG_ACCOUNT = "Vous n'avez pas introduit un compte."
CONFIG_MSG_PASSWORD = "Vous n'avez pas introduit un mot de passe."
CONFIG_MSG_NUMERIC = "Tu ne peut pas insérer noms composé de numeros."
CONFIG_MSG_PORT = "Tu n'as pas selectionner une porte valide."
CONFIG_MSG_IP = "Tu n'as pas innecté un IP."
CONFIG_MSG_UPDATE_FILE = "You need the peach updater to be able to upgrade your peach." & vbCrLf & vbCrLf & "Download it here: http://riplegion.ri.funpic.de/Peach/peachUpdater.exe"

CHAT_COMMAND_SEND = "&Envoi"
CHAT_COMMAND_CLEAR = "&Clair"

SF_LABEL_FILENAME = "Nom file"
SF_LABEL_SEND_TO = "Send to"
SF_LABEL_TIME = " Time left: "
SF_LABEL_KBS = " Kb/Sec, "
SF_LABEL_KBSS = " KBytes sent, "

SF_MSG_USER = "Pas d'utilisateur sélectionné."
SF_MSG_FILE = "Aucun fichier sélectionné."
SF_MSG_INCOMMING_FILE = "You are receiving '%f' from '%u'. Do you want to accept?"
SF_MSG_DECILINED = "Le transfert de fichier a été refusée."

SF_COMMAND_BROWSE = "&Cherche .."
SF_COMMAND_SENDFILE = "Envoi"
SF_COMMAND_CANCEL = "Annuler .."

LANG_COMMAND_ENTER = "&Ouvrir"
LANG_LABEL_SELLANG = "Choisissez votre langue"

SOC_FRIEND_LIST = "Liste D'amis"
SOC_ONLINE_LIST = "Liste des onlines"
SOC_IGNORE_LIST = "Stop-Liste"

SOC_COMMAND_REMOVE = "&Supprimer"
SOC_COMMAND_FRIEND = "&Ajouter aux amis"
SOC_COMMAND_IGNORE = "&Ignore user"
SOC_COMMAND_WHISPER = "&Whisper"

SOC_ASK_DEL = "Voulez-vous supprimer '%u' de la liste?"

SOC_ASK_FRIEND_TEXT = "Enter the account name of your friend in the text box below."
SOC_ASK_FRIEND_TITLE = "Adding a friend"
SOC_ASK_FRIEND_DEFAULT = "Enter account here"

SOC_ASK_IGNORE_TEXT = "Enter the account name of the user you want to ignore in the text box below."
SOC_ASK_IGNORE_TITLE = "Ignore a user"
SOC_ASK_IGNORE_DEFAULT = "Enter account here"

SOC_FRIEND_LIST_STATUS = "État"

SOC_MSG_CANT_WHISPER = "You can't whisper this user."

REG_CAPTION = "Peach - D'enregistrement"

REG_FRAME_DETAIL = "Entrez vos coordonnées"

REG_LABEL_ACCOUNT_NAME = "Nom du compte"
REG_LABEL_PASSWORD = "Mot de passe"
REG_LABEL_PASSWORD_CONFIRM = "Confirmer le mot de passe"
REG_LABEL_PASSWORD_WEAK = "Le mot de passe est faible."
REG_LABEL_PASSWORD_NORMAL = "Le mot de passe est normal."
REG_LABEL_PASSWORD_STRONG = "Le mot de passe est forte."
REG_LABEL_SECRET_QUESTION = "Question secrète"
REG_LABEL_SECRET_ANSWER = "Réponse secrète"

REG_COMMAND_SUBMIT = "&Envoyer"
REG_COMMAND_CLOSE = "&Fermer"

REG_CHECK_PASSWORD_SHOW = "&Afficher mot de passe."

REG_MSG_ACCOUNT_EXIST = "Le nom du compte qui existe déjà."
REG_MSG_ACCOUNT_INVALID = "Nom de compte non valide."
REG_MSG_ACCOUNT_NUMERIC = "Compte ne peut pas être composé de caractères numériques."
REG_MSG_ACCOUNT_EMPTY = "Pas de compte soumis."
REG_MSG_ACCOUNT_SHORT = "Nom du compte à court."

REG_MSG_PASSWORD_MATCH = "Les mots de passe ne correspondent pas."
REG_MSG_PASSWORD_SHORT = "Mot de passe à court."
REG_MSG_PASSWORD_EMPTY = "Aucun mot de passe soumis."

REG_MSG_SECRET_ANSWER_EMPTY = "Pas de secret répondu ajouté."

REG_MSG_EMAIL_EMPTY = "No email entered."
REG_MSG_EMAIL_TAKEN = "The email adress you have entered is already beeing used, have you forgot your password?"

REG_MSG_SUCCESSFULLY = "Le compte a été enregistré avec succès."
REG_MSG_ERROR = "Une erreur s'est produite, s'il vous plaît essayer à nouveau plus tard.."
REG_MSG_ERROR_OCCURED = "Une erreur s'est produite ..."
REG_MSG_LOADING = "Charge .."
REG_MSG_CONNECTION_BROKEN = "La connexion est perdue, s'il vous plaît essayer à nouveau plus tard."

REG_CMB_SECRET_QUESTION_0 = "Quel est le nom de votre animal de compagnie?"
REG_CMB_SECRET_QUESTION_1 = "Quel est votre livre préféré?"
REG_CMB_SECRET_QUESTION_2 = "Quel est votre film préféré?"
REG_CMB_SECRET_QUESTION_3 = "Quel est votre jeu préféré?"
REG_CMB_SECRET_QUESTION_4 = "Quel est votre chanteur préféré?"
REG_CMB_SECRET_QUESTION_5 = "Quel est l'endroit où votre mère est née?"

REG_LABEL_GENDER = "Gender"

REG_CMB_GENDER_MALE = "Male"
REG_CMB_GENDER_FEMALE = "Female"

SET_LABEL_COLOR = "Couleur Courante"
SET_LABEL_FONT = "Fonte"

SET_FRAME_STYLE = "Style"
SET_FRAME_OPTIONS = "Options"
SET_FRAME_CONNECTION = "Paramètres de connexion"

SET_CHECK_SAVE_ACCOUNT = "Sauvegarder le compte"
SET_CHECK_SAVE_PASSWORD = "Sauvegarder mot de passe"
SET_CHECK_AUTO_LOGIN = "Login automatically"
SET_CHECK_ASK_CLOSING = "Demander, avant fermeture"
SET_CHECK_MINIMIZE = "Réduire la fenêtre de barre d'état système"

SET_COMMAND_LANGUAGE = "&Langue"
SET_COMMAND_SAVE = "&Sauvegarder"

SF2_COMMAND_OPEN_FILE = "&Ouvrez le dossier de fichiers"

FP_FRAME_FORGOT_PASSWORD = "Mot de passe oublié"
FP_LABEL_EMAIL = "Entrez votre e-mail"
FP_LABEL_SECRET_QUESTION = "Question secrète"
FP_LABEL_SECRET_ANSWER = "Réponse secrète"
FP_COMMAND_REQUEST = "&Demande"
FP_CAPTION = "Peach - Mot de passe oublié"

FP_MSG_SUCCESSFULL = "Votre compte est '%u'." & vbCrLf & "Votre mot de passe est '%p'."
FP_MSG_WRONG_ANSWER = "La réponse est fausse."
FP_MSG_WRONG_EMAIL = "The entered email adress could not be found."
End Sub
