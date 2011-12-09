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
Public REG_MSG_EMAIL_INVALID            As String

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

Public CH_MSG_PASSWORD                  As String

Public MSG_USER_ONLINE                  As String
Public MSG_USER_OFFLINE                 As String
Public MSG_ANNOUNCE                     As String
Public MSG_TABLE_RELOAD                 As String
Public MSG_TABLE_CANT_RELOAD            As String
Public MSG_CONFIG_RELOAD                As String
Public MSG_INCORRECT_SYNTAX             As String
Public MSG_TABLE_NOT_EXIST              As String
Public MSG_USER_NOT_FOUND               As String
Public MSG_DELETED_ACCOUNT              As String
Public MSG_GM_FLAG_ENABLE               As String
Public MSG_GM_FLAG_DISABLE              As String
Public MSG_UNKNOWN_COMMAND              As String
Public MSG_MUTED                        As String
Public MSG_FLOOD_PROTECTION             As String
Public MSG_ROLL                         As String
Public MSG_NOT_AFK                      As String
Public MSG_AFK                          As String
Public MSG_ONLINE_TIME                  As String
Public MSG_VALID_CHANNEL                As String
Public MSG_ALREADY_IN_CHANNEL           As String
Public MSG_NOT_IN_CHANNEL               As String
Public MSG_CHANNEL_ANNOUNCEMENTS        As String
Public MSG_CHANNEL_PASSWORD             As String
Public MSG_CHANNEL_WRONG_PASSWORD       As String
Public MSG_NOT_CHANNEL_LEADER           As String
Public MSG_MESSAGE_BLOCKED              As String
Public MSG_CANT_WHISPER_SELF            As String
Public MSG_IS_IGNORING_YOU              As String
Public MSG_YOU_WHISPER_TO               As String
Public MSG_TARGET_IS_AFK                As String
Public MSG_WHISPER                      As String
Public MSG_USER_ALREADY_MUTED           As String
Public MSG_IS_NOT_MUTED                 As String
Public MSG_MUTED_BY                     As String
Public MSG_UNMUTED_BY                   As String
Public MSG_MUTED_BY_REASON              As String
Public MSG_UNMUTED_BY_REASON            As String
Public MSG_ALREADY_BANNED               As String
Public MSG_ALREADY_UNBANNED             As String
Public MSG_BANNED_BY                    As String
Public MSG_UNBANNED_BY                  As String
Public MSG_BANNED_BY_REASON             As String
Public MSG_UNBANNED_BY_REASON           As String
Public MSG_SUCCESSFULL_RENAME           As String
Public MSG_RENAMED_YOU_TO               As String
Public MSG_USER_ALREADY_USED            As String
Public MSG_LEVEL_INCORRECT_VALUE        As String
Public MSG_SUCCESSFULL_LEVEL            As String
Public MSG_CHANGED_YOUR_LEVEL           As String
Public MSG_GENDER_INCORRECT_VALUE       As String
Public MSG_SUCCESSFULL_GENDER           As String
Public MSG_CHANGED_YOUR_GENDER          As String
Public MSG_SUCCESSFULL_PASSWORD         As String
Public MSG_CHANGED_YOUR_PASSWORD        As String
Public MSG_SUCCESSFULL_EMAIL            As String
Public MSG_CHANGED_YOUR_EMAIL           As String
Public MSG_JOINED_CHANNEL               As String
Public MSG_CHANNEL_USER_JOIN            As String
Public MSG_CHANNEL_USER_LEAVE           As String
Public MSG_MESSAGE_SENT_OFFLINE         As String
Public MSG_OFFLINE_MESSAGE              As String
Public MSG_SUCCESSFULL_SUDO_LOGIN       As String

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
REG_MSG_EMAIL_INVALID = "Die Email-Adresse verfügt nicht über das korrekte Format."

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

CH_MSG_PASSWORD = "Bitte geben Sie das Passwort für den Channel '%c' an."

MSG_USER_ONLINE = "%u hat sich angemeldet."
MSG_USER_OFFLINE = "%u hat sich abgemeldet."
MSG_ANNOUNCE = "%f[%u kündigt an]: %m"
MSG_TABLE_RELOAD = "%u hat die '%t'-Tabelle neu geladen. ( %i ) "
MSG_TABLE_CANT_RELOAD = "Diese Tabelle kann nicht neu geladen werden."
MSG_CONFIG_RELOAD = "%u hat die Konfigurations-Dateien neu geladen. ( %t )"
MSG_INCORRECT_SYNTAX = "Falscher Syntax. Benutze das folgende Format %s."
MSG_TABLE_NOT_EXIST = "Diese Tabelle existiert nicht."
MSG_USER_NOT_FOUND = "Benutzer '%u' wurde nicht gefunden."
MSG_DELETED_ACCOUNT = "Benutzer '%u' (%id) wurde erfolgreich gelöscht."
MSG_GM_FLAG_ENABLE = "GM-Modus wurde aktiviert. Benutzen Sie .gm off um Ihn wieder zu deaktivieren."
MSG_GM_FLAG_DISABLE = "GM-Modus wurde deaktiviert. Benutzen Sie .gm on uhm Ihn wieder zu aktivieren."
MSG_UNKNOWN_COMMAND = "Unbekannter Befehl benutzt."
MSG_MUTED = "Sie sind gemuted."
MSG_FLOOD_PROTECTION = "Ihre Nachricht hat den serverseitigen Spamfilter aktiviert. Bitte wiederholen Sie sich nicht."
MSG_ROLL = "%u würfelt eine %r. (%minR - %maxR)"
MSG_NOT_AFK = "Sie sind nicht mehr (Away from Keyboard – nicht an der Tastatur)."
MSG_AFK = "Sie sind (Away from Keyboard – nicht an der Tastatur)."
MSG_ONLINE_TIME = "Sie sind Online für %t."
MSG_VALID_CHANNEL = "Bitte geben Sie einen gültigen Channel-Namen ein."
MSG_ALREADY_IN_CHANNEL = "Sie sind bereits im Channel '%c'."
MSG_NOT_IN_CHANNEL = "Sie sind nicht im Channel '%c'."
MSG_CHANNEL_ANNOUNCEMENTS = "[%c] Channelankündigungen wurden deaktiviert von %u."
MSG_CHANNEL_PASSWORD = "Das Passwort des Channels '%c' wurde erfolgreich zu '%p' geändert."
MSG_CHANNEL_WRONG_PASSWORD = "Falsches Passwort für Channel '%c'."
MSG_NOT_CHANNEL_LEADER = "Sie sind nicht der Leiter dieses Channels."
MSG_MESSAGE_BLOCKED = "Die Nachtricht wurde geblockt. Bitte schreiben Sie nicht mehr als 75% in Caps."
MSG_CANT_WHISPER_SELF = "Sie können sich nicht selber anflüstern."
MSG_IS_IGNORING_YOU = "%t ignoriert Sie."
MSG_YOU_WHISPER_TO = "[Sie flüstern zu %t]: %m"
MSG_TARGET_IS_AFK = "%t ist (Away from Keyboard – nicht an der Tastatur)."
MSG_WHISPER = "%f[%u flüstert]: %m"
MSG_USER_ALREADY_MUTED = "%u ist bereits gemuted."
MSG_IS_NOT_MUTED = "%u ist nicht gemuted."
MSG_MUTED_BY = "%t wurde von %u gemuted."
MSG_UNMUTED_BY = "%t wurde von %u unmuted."
MSG_MUTED_BY_REASON = "%t wurde von %u gemuted. (%r)"
MSG_UNMUTED_BY_REASON = "%t wurde von %u unmuted. (%r)"
MSG_ALREADY_BANNED = "Account '%u' ist bereits gebannt."
MSG_ALREADY_UNBANNED = "Account '%u' ist nicht gebannt."
MSG_BANNED_BY = "%t wurde gebannt von %u."
MSG_UNBANNED_BY = "%t wurde unbannt von %u."
MSG_BANNED_BY_REASON = "%t wurde gebannt von %u. (%r)"
MSG_UNBANNED_BY_REASON = "%t wurde unbannt von %u. (%r)"
MSG_SUCCESSFULL_RENAME = "Sie haben '%u' erfolgreich zu '%t' umbenannt."
MSG_RENAMED_YOU_TO = "%u hat Sie zu '%t' umbenannt."
MSG_USER_ALREADY_USED = "Der Name '%u' wird bereits verwendet."
MSG_LEVEL_INCORRECT_VALUE = "Das Level enthält inkorrekte Werte, es muss im Bereich von 0-2 liegen."
MSG_SUCCESSFULL_LEVEL = "Sie haben das Level von '%u' erfolgreich zu '%l' geändert."
MSG_CHANGED_YOUR_LEVEL = "%u hat Ihr Level auf '%l' geändert."
MSG_GENDER_INCORRECT_VALUE = "Das Geschlecht enthält inkorrekte Werte, bitte benutzten Sie 'male' oder 'female'."
MSG_SUCCESSFULL_GENDER = "Sie haben das Geschlecht von '%u' erfolgreich zu '%g' geändert."
MSG_CHANGED_YOUR_GENDER = "%u hat Ihr Geschlecht zu '%g' geändert."
MSG_SUCCESSFULL_PASSWORD = "Sie haben das Passwort von '%u' erfolgreich zu '%p' geändert."
MSG_CHANGED_YOUR_PASSWORD = "%u hat Ihr Passwort zu '%p' geändert."
MSG_SUCCESSFULL_EMAIL = "Sie haben die Email-Adresse von '%u' erfolgreich zu '%e' geändert."
MSG_CHANGED_YOUR_EMAIL = "%u hat Ihre Emil-Adresse zu '%e' geändert."
MSG_JOINED_CHANNEL = "Sie sind dem Channel '%c' beigetreten."
MSG_CHANNEL_USER_JOIN = "[%c] %u ist dem Channel beigetreten."
MSG_CHANNEL_USER_LEAVE = "[%c] %u hat den Channel verlassen."
MSG_MESSAGE_SENT_OFFLINE = "Nachricht gesendet. %t ist offline, die Nachricht wird ausgerichtet sobald %t sich einloggt hat."
MSG_OFFLINE_MESSAGE = "(%time)[%from flüstert]: %message"
MSG_SUCCESSFULL_SUDO_LOGIN = "Sie haben sich erfolgreich als Administrator eingeloggt."
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
MDI_MSG_WRONG_PASSWORD = "The information you have entered is not valid. Please check spelling of the account name and password."
MDI_MSG_BANNED = "This account is banned."
MDI_MSG_UNLOAD = "Are you sure you want to close Peach?"

MDI_MSG_CANT_ADD_YOU = "You can't put yourself on your friend list."
MDI_MSG_ALREADY_IN_IGNORE_LIST = "%u is already being ignored."
MDI_MSG_ALREADY_IN_FRIEND_LIST = "%u is already your friend."
MDI_MSG_ACCOUNT_NOT_EXIST = "User %u not found."

MDI_MENU = "Menu"

'Configuration form ..
CONFIG_LABEL_ACCOUNT = "Username"
CONFIG_LABEL_PASSWORD = "Password"

CONFIG_COMMAND_CONNECT = "&Login"
CONFIG_COMMAND_DISCONNECT = "&Logout"
CONFIG_COMMAND_SETTINGS = "&Preferences"
CONFIG_COMMAND_UPDATE = "&Update"
CONFIG_COMMAND_REGISTER = "&Create Account"
CONFIG_COMMAND_FORGOT_PASSWORD = "&Forgot my password"

CONFIG_CHECK_SAVE_PASSWORD = "&Save Password"

CONFIG_FRAME_CONNECTION = "Connection Informeation"

LANG_GERMAN = "German"
LANG_ENGLISH = "English"
LANG_SPANISH = "Spanish"

CONFIG_MSG_ACCOUNT = "You did not enter an account."
CONFIG_MSG_PASSWORD = "You did not enter a password."
CONFIG_MSG_NUMERIC = "You can not take numeric usernames."
CONFIG_MSG_PORT = "You did not introduce a port."
CONFIG_MSG_IP = "You did not introduce an IP."
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
REG_LABEL_PASSWORD_CONFIRM = "Confirm the password"
REG_LABEL_PASSWORD_WEAK = "The password is weak."
REG_LABEL_PASSWORD_NORMAL = "The password is normal."
REG_LABEL_PASSWORD_STRONG = "The password is strong."
REG_LABEL_SECRET_QUESTION = "Secret question"
REG_LABEL_SECRET_ANSWER = "Secret answer"

REG_COMMAND_SUBMIT = "&Submit"
REG_COMMAND_CLOSE = "&Close"

REG_CHECK_PASSWORD_SHOW = "&Show password."

REG_MSG_ACCOUNT_EXIST = "This username is already in use."
REG_MSG_ACCOUNT_INVALID = "Invalid username."
REG_MSG_ACCOUNT_NUMERIC = "Username can not be composed of numeric characters."
REG_MSG_ACCOUNT_EMPTY = "No username entered."
REG_MSG_ACCOUNT_SHORT = "The username is too short, it requieres at least 4 characters."

REG_MSG_PASSWORD_MATCH = "The passwords do not match."
REG_MSG_PASSWORD_SHORT = "The password is too short, it requieres at least 6 characters."
REG_MSG_PASSWORD_EMPTY = "No password entered."

REG_MSG_SECRET_ANSWER_EMPTY = "No secret answered entered."

REG_MSG_EMAIL_EMPTY = "No email adress entered."
REG_MSG_EMAIL_TAKEN = "The email adress you have entered is already beeing used, have you forgotten your password?"
REG_MSG_EMAIL_INVALID = "The email adress you have entered is invalid."

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
FP_MSG_WRONG_EMAIL = "The e-mail adress could not be found."

CH_MSG_PASSWORD = "Please enter the password for '%c'."

MSG_USER_ONLINE = "%u has come online."
MSG_USER_OFFLINE = "%u has gone offline."
MSG_ANNOUNCE = "%f[%u announces]: %m"
MSG_TABLE_RELOAD = "%u initiated the reload of '%t' table. ( %i ) "
MSG_TABLE_CANT_RELOAD = "This table can't be reloaded."
MSG_CONFIG_RELOAD = "%u initiated the reload of configuration files. ( %t )"
MSG_INCORRECT_SYNTAX = "Incorrect Syntax. Use the following format %s."
MSG_TABLE_NOT_EXIST = "This table does not exist."
MSG_USER_NOT_FOUND = "User '%u' was not found."
MSG_DELETED_ACCOUNT = "Successfully deleted account '%u' (%id)."
MSG_GM_FLAG_ENABLE = "Enabled [GM] flag. Use .gm off to disable."
MSG_GM_FLAG_DISABLE = "Disabled [GM] flag. Use .gm on to enable."
MSG_UNKNOWN_COMMAND = "Unknown command used."
MSG_MUTED = "You are muted."
MSG_FLOOD_PROTECTION = "Your message has triggered serverside flood protection. Please don't repeat yourself."
MSG_ROLL = "%u rolls %r. (%minR - %maxR)"
MSG_NOT_AFK = "You are not away from keyboard anymore."
MSG_AFK = "You are away from keyboard now."
MSG_ONLINE_TIME = "You are online for %t."
MSG_VALID_CHANNEL = "Please enter a valid channel name."
MSG_ALREADY_IN_CHANNEL = "You are already in '%c'."
MSG_NOT_IN_CHANNEL = "You are not in channel '%c'."
MSG_CHANNEL_ANNOUNCEMENTS = "[%c] Channel announcements got disabled by %u."
MSG_CHANNEL_PASSWORD = "Successfully changed password of '%c' to '%p'."
MSG_CHANNEL_WRONG_PASSWORD = "Wrong password for channel '%c'."
MSG_NOT_CHANNEL_LEADER = "You are not the leader of this channel."
MSG_MESSAGE_BLOCKED = "Message blocked. Please do not write more than 75% in caps."
MSG_CANT_WHISPER_SELF = "You can't whisper yourself."
MSG_IS_IGNORING_YOU = "%t is ignoring you."
MSG_YOU_WHISPER_TO = "[You whisper to %t]: %m"
MSG_TARGET_IS_AFK = "%t is away from keyboard."
MSG_WHISPER = "%f[%u whispers]: %m"
MSG_USER_ALREADY_MUTED = "%u is already muted."
MSG_IS_NOT_MUTED = "%u is not muted."
MSG_MUTED_BY = "%t got muted by %u."
MSG_UNMUTED_BY = "%t got unmuted by %u."
MSG_MUTED_BY_REASON = "%t got muted by %u. (%r)"
MSG_UNMUTED_BY_REASON = "%t got unmuted by %u. (%r)"
MSG_ALREADY_BANNED = "Account '%u' is already banned."
MSG_ALREADY_UNBANNED = "Account '%u' is not banned."
MSG_BANNED_BY = "%t was banned by %u."
MSG_UNBANNED_BY = "%t was unbanned by %u."
MSG_BANNED_BY_REASON = "%t was banned by %u. (%r)"
MSG_UNBANNED_BY_REASON = "%t was unbanned by %u. (%r)"
MSG_SUCCESSFULL_RENAME = "Successfully renamed '%u' to '%t'."
MSG_RENAMED_YOU_TO = "%u renamed you to '%t'."
MSG_USER_ALREADY_USED = "Username '%u' is already beeing used."
MSG_LEVEL_INCORRECT_VALUE = "Level contains incorrect values, must be in range of 0-2."
MSG_SUCCESSFULL_LEVEL = "Successfully changed level of '%u' to '%l'."
MSG_CHANGED_YOUR_LEVEL = "%u changed your level to '%l'."
MSG_GENDER_INCORRECT_VALUE = "Incorrect gender format use 'male' or 'female'."
MSG_SUCCESSFULL_GENDER = "Successfully changed gender of '%u' to '%g'."
MSG_CHANGED_YOUR_GENDER = "%u changed your gender to '%g'."
MSG_SUCCESSFULL_PASSWORD = "Sucessfully changed password of '%u' to '%p'."
MSG_CHANGED_YOUR_PASSWORD = "%u changed your password to '%p'."
MSG_SUCCESSFULL_EMAIL = "Successfully changed email of '%u' to '%e'."
MSG_CHANGED_YOUR_EMAIL = "%u changed your email to '%e'."
MSG_JOINED_CHANNEL = "You joined channel '%c'."
MSG_CHANNEL_USER_JOIN = "[%c] %u joined channel."
MSG_CHANNEL_USER_LEAVE = "[%c] %u left channel."
MSG_MESSAGE_SENT_OFFLINE = "Message sent. %t will recieve the message upon login."
MSG_OFFLINE_MESSAGE = "(%time)[%from whispers]: %message"
MSG_SUCCESSFULL_SUDO_LOGIN = "You have successfully logged in as administrator."
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
REG_MSG_EMAIL_INVALID = "El correo electronico que ha introducido es incorrecto."

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

CH_MSG_PASSWORD = "Por favor introduzca la contraseña del canal '%c'."

MSG_USER_ONLINE = "%u se ha conectado."
MSG_USER_OFFLINE = "%u se ha desconectado."
MSG_ANNOUNCE = "%f[%u anuncia]: %m"
MSG_TABLE_RELOAD = "%u ha iniciado la recarga de la tabla '%t'. ( %i ) "
MSG_TABLE_CANT_RELOAD = "Esta tabla no puede ser recargada."
MSG_CONFIG_RELOAD = "%u ha iniciado la recarga de los archivos de configuración. ( %t )"
MSG_INCORRECT_SYNTAX = "Sintaxis incorrecta. Por favor usa el formato adecuado %s."
MSG_TABLE_NOT_EXIST = "Esta tabla no existe."
MSG_USER_NOT_FOUND = "El usuario '%u' no ha sido encontrado."
MSG_DELETED_ACCOUNT = "Has borrado con éxito la cuenta '%u' (%id)."
MSG_GM_FLAG_ENABLE = "Has activado el modo de [GM]. Use .gm off para desactivarlo."
MSG_GM_FLAG_DISABLE = "Has desactivado el modo de [GM]. Use .gm on para activarlo."
MSG_UNKNOWN_COMMAND = "Has usado un comando desconocido."
MSG_MUTED = "Estas muted."
MSG_FLOOD_PROTECTION = "Tu mensaje ha activado la protección contra Spam. Por favor, no te repitas."
MSG_ROLL = "%u rolls %r. (%minR - %maxR)"
MSG_NOT_AFK = "Has vuelto de la ausencia"
MSG_AFK = "Estas ausente."
MSG_ONLINE_TIME = "Estas conectado desde %t."
MSG_VALID_CHANNEL = "Por favor introduzca un nombre de canal valido."
MSG_ALREADY_IN_CHANNEL = "Ya estas en el canal '%c'."
MSG_NOT_IN_CHANNEL = "No estas en el canal '%c'."
MSG_CHANNEL_ANNOUNCEMENTS = "[%c] Los anuncios para este canal han sido desactivados por %u."
MSG_CHANNEL_PASSWORD = "Has cambiado la contraseña del canal '%c' a '%p' con éxito."
MSG_CHANNEL_WRONG_PASSWORD = "La contraseña del canal '%c' es incorrecta."
MSG_NOT_CHANNEL_LEADER = "No eres el lider de este canal."
MSG_MESSAGE_BLOCKED = "Mensaje bloqueado. Por favor no escribe mas de 75% en mayúscula."
MSG_CANT_WHISPER_SELF = "No te puedes sussurar a ti mismo."
MSG_IS_IGNORING_YOU = "%t te esta ignorando."
MSG_YOU_WHISPER_TO = "[Estas sussurando a %t]: %m"
MSG_TARGET_IS_AFK = "%t esta ausente."
MSG_WHISPER = "%f[%u sussura]: %m"
MSG_USER_ALREADY_MUTED = "%u ya esta silenciado/a."
MSG_IS_NOT_MUTED = "%u no esta silenciado/a."
MSG_MUTED_BY = "%t ha sido silenciado/a por %u."
MSG_UNMUTED_BY = "%t ha sido dessilenciado/a por %u."
MSG_MUTED_BY_REASON = "%t ha sido dessilenciado/a por %u. (%r)"
MSG_UNMUTED_BY_REASON = "%t ha sido dessilenciado/a por %u. (%r)"
MSG_ALREADY_BANNED = "La cuenta '%u' ya esta baneada."
MSG_ALREADY_UNBANNED = "La cuenta '%u' no esta baneada."
MSG_BANNED_BY = "%t ha sido baneado/a por %u."
MSG_UNBANNED_BY = "%t ha sido desbaneado/a por %u."
MSG_BANNED_BY_REASON = "%t ha sido baneado/a by %u. (%r)"
MSG_UNBANNED_BY_REASON = "%t ha sido desbaneado/a %u. (%r)"
MSG_SUCCESSFULL_RENAME = "Has renombrado '%u' a '%t' con éxito."
MSG_RENAMED_YOU_TO = "%u te ha cambiado el nombre a '%t'."
MSG_USER_ALREADY_USED = "El nombre '%u' ya esta cogido."
MSG_LEVEL_INCORRECT_VALUE = "El nivel contiene valores incorrectos, tienen que estar entre 0-2."
MSG_SUCCESSFULL_LEVEL = "Has cambiado el nivel de '%u' a '%l' con éxito."
MSG_CHANGED_YOUR_LEVEL = "%u ha cambiado tu nivel a '%l'."
MSG_GENDER_INCORRECT_VALUE = "Formato de sexo incorrecto, por favor use 'male' o 'female'."
MSG_SUCCESSFULL_GENDER = "Has cambiado el sexo de '%u' a '%g' con éxito."
MSG_CHANGED_YOUR_GENDER = "%u ha cambiado tu sexo a '%g'."
MSG_SUCCESSFULL_PASSWORD = "Has cambiado la contraseña de '%u' a '%p' con éxito."
MSG_CHANGED_YOUR_PASSWORD = "%u ha cambiado tu contraseña a '%p'."
MSG_SUCCESSFULL_EMAIL = "Has cambiado el correo electronico de '%u' a '%e' con éxito."
MSG_CHANGED_YOUR_EMAIL = "%u ha cambiado tu correo electronico a '%e'."
MSG_JOINED_CHANNEL = "Has entrado en el canal '%c'."
MSG_CHANNEL_USER_JOIN = "[%c] %u ha entrado en el canal."
MSG_CHANNEL_USER_LEAVE = "[%c] %u ha salido del canal."
MSG_MESSAGE_SENT_OFFLINE = "Mensaje enviado. %t esta desconectado, su mensaje sera enviado encuanto se conecte."
MSG_OFFLINE_MESSAGE = "(%time)[%from sussura]: %message"
MSG_SUCCESSFULL_SUDO_LOGIN = "Te has conectaco con exito como administrador."
End Sub
