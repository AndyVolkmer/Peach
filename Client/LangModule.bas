Attribute VB_Name = "LangModule"
Option Explicit

'Start variable support for languages
' MDI form ..
Public MDIcommand_config                As String
Public MDIcommand_chat                  As String
Public MDIcommand_sendfile              As String
Public MDIcommand_onlinelist            As String

Public MDIstatusbar_disconnected        As String
Public MDIstatusbar_dcfromserver        As String
Public MDIstatusbar_connected           As String
Public MDIstatusbar_connectionproblem   As String
Public MDIstatusbar_connecting          As String

Public MDImsgbox_errorHandlerFormLoad   As String
Public MDImsgbox_config_notify          As String
Public MDImsgbox_nametaken              As String
Public MDImsgbox_wrong_account          As String
Public MDImsgbox_wrong_password         As String
Public MDImsgbox_banned                 As String

' Configuration form ..
Public CONFIGcommand_connect            As String
Public CONFIGcommand_disconnect         As String
Public CONFIGcommand_language           As String
Public CONFIGcommand_update             As String
Public CONFIGcommand_register           As String

Public CONFIGcheck_savepassword         As String

Public CONFIGlabel_CI_name              As String
Public CONFIGlabel_selectlanguage       As String

Public CONFIGframe_personal             As String
Public CONFIGframe_connection           As String

Public CONFIGcombo_german               As String
Public CONFIGcombo_english              As String
Public CONFIGcombo_spanish              As String
Public CONFIGcombo_swedish              As String
Public CONFIGcombo_italian              As String
Public CONFIGcombo_greek                As String
Public CONFIGcombo_serbian              As String
Public CONFIGcombo_russian              As String
Public CONFIGcombo_dutch                As String
Public CONFIGcombo_french               As String

Public CONFIGmsgbox_account             As String
Public CONFIGmsgbox_password            As String
Public CONFIGmsgbox_nonumeric           As String
Public CONFIGmsgbox_portnoempty         As String
Public CONFIGmsgbox_namenoempty         As String
Public CONFIGmsgbox_ipnoempty           As String

' Chat form ..
Public CHATcommand_send                 As String
Public CHATcommand_clear                As String

Public CHATtimetext                     As String

' List form ..
Public LISTcaption                      As String
Public LISTcommand_close                As String

' Send File form ..
Public SFlabel_filename                 As String
Public SFlabel_sendingfile              As String
Public SFlabel_sent                     As String
Public SFlabel_sendto                   As String

Public SFcommand_browse                 As String
Public SFcommand_sendfile               As String
Public SFcommand_cancelsending          As String

Public SFmsgbox_nousersel               As String
Public SFmsgbox_nofilesel               As String
Public SFmsgbox_incfile                 As String
Public SFmsgbox_filedecilined           As String

' Desp form ..
Public DESPtext_newmsg                  As String
Public DESPtext_dcserver                As String

Public Sub SetLangGerman()

' MDI form ..
MDIcommand_config = "&Einstellungen"
MDIcommand_chat = "Ch&at"
MDIcommand_sendfile = "&Sende Datei"
MDIcommand_onlinelist = "&Online Liste"

MDIstatusbar_disconnected = "Status: Getrennt"
MDIstatusbar_dcfromserver = "Status: Getrennt vom Server"
MDIstatusbar_connected = "Status: Verbunden mit "
MDIstatusbar_connectionproblem = "Status: Getrennt aufgrund eines Verbindungsfehlers"
MDIstatusbar_connecting = "Status: Verbinden mit "

MDImsgbox_config_notify = "Einige Konfigurationsdateien sind veraltet oder wurden beschädigt, Peach fand den Fehler und wird es mit dem nächsten Neustart korrigieren."
MDImsgbox_nametaken = "Der Name ist bereits vergeben."
MDImsgbox_wrong_account = "Der Account ist nicht vorhanden oder falsch."
MDImsgbox_wrong_password = "Das Passwort ist falsch."
MDImsgbox_banned = "Dieser Account wurde gebannt."

CONFIGcommand_connect = "&Verbinden"
CONFIGcommand_disconnect = "&Verbindung trenn."
CONFIGcommand_language = "&Sprache"
CONFIGcommand_update = "&Aktualisieren"
CONFIGcommand_register = "&Account registrieren"

CONFIGcheck_savepassword = "&Password Speichern"

CONFIGlabel_CI_name = "Name: "
CONFIGlabel_selectlanguage = "Wähle deine Sprache aus:"

CONFIGframe_personal = "Personelle Informationen"
CONFIGframe_connection = "Verbindungs Informationen"

CONFIGcombo_german = "Deutsch"
CONFIGcombo_english = "Englisch"
CONFIGcombo_spanish = "Spanisch"
CONFIGcombo_swedish = "Schwedisch"
CONFIGcombo_italian = "Italienisch"
CONFIGcombo_greek = "Griechisch"
CONFIGcombo_serbian = "Serbisch"
CONFIGcombo_russian = "Russisch"
CONFIGcombo_dutch = "Niederländisch"
CONFIGcombo_french = "Französisch"

CONFIGmsgbox_account = "Du hast keinen Account eingegeben."
CONFIGmsgbox_password = "Du hast kein Passwort eingegeben."
CONFIGmsgbox_nonumeric = "Du kannst keine Ziffern in deinem Namen haben."
CONFIGmsgbox_portnoempty = "Du hast keinen Port eingegeben."
CONFIGmsgbox_namenoempty = "Du hast keinen Namen eingegeben."
CONFIGmsgbox_ipnoempty = "Du hast keine IP eingegeben."

CHATcommand_send = "&Senden"
CHATcommand_clear = "&Löschen"

CHATtimetext = " Die Zeit beträgt "

LISTcaption = "Online Liste"
LISTcommand_close = "&Schliessen"

SFlabel_filename = " Datei Name:"
SFlabel_sendingfile = "Sende:"
SFlabel_sent = "0.0% Gesendet"
SFlabel_sendto = "Sende an:"

SFmsgbox_nousersel = "Kein Benutzer ausgewählt."
SFmsgbox_nofilesel = "Keine Datei ausgewählt."
SFmsgbox_incfile = "Sie empfangen eine Datei, möchsten sie annehmen?"
SFmsgbox_filedecilined = "Der Benutzer hat die Datei abgelehnt."

SFcommand_browse = "&Suchen .."
SFcommand_sendfile = "Senden"
SFcommand_cancelsending = "Abbrechen .."

DESPtext_newmsg = "Neue Nachricht!"
DESPtext_dcserver = "Verbindung unterbrochen!"
End Sub

Public Sub SetLangEnglish()

'MDI form
MDIcommand_config = "&Configuration"
MDIcommand_chat = "Ch&at"
MDIcommand_sendfile = "&Send File"
MDIcommand_onlinelist = "&Online List"

MDIstatusbar_disconnected = "Status: Disconnected"
MDIstatusbar_dcfromserver = "Status: Disconnected from Server"
MDIstatusbar_connected = "Status: Connected to "
MDIstatusbar_connectionproblem = "Status: Disconnected due connection problem"
MDIstatusbar_connecting = "Status: Connecting to "

'MDImsgbox_errorHandlerFormLoad
MDImsgbox_config_notify = "Some configuration files are outdated or got damaged, Peach found the problem and will fix it on next program launch."
MDImsgbox_nametaken = "This name is already taken."
MDImsgbox_wrong_account = "The account does not exist or is wrong."
MDImsgbox_wrong_password = "The password is wrong."
MDImsgbox_banned = "This account is banned."

' Configuration form ..
CONFIGcommand_connect = "&Connect"
CONFIGcommand_disconnect = "&Disconnect"
CONFIGcommand_language = "&Language"
CONFIGcommand_update = "&Update"
CONFIGcommand_register = "&Register Account"

CONFIGcheck_savepassword = "&Save Password"

CONFIGlabel_CI_name = "Name: "
CONFIGlabel_selectlanguage = "Select your language:"

CONFIGframe_personal = "Personal Information"
CONFIGframe_connection = "Connection Information"

CONFIGcombo_german = "German"
CONFIGcombo_english = "English"
CONFIGcombo_spanish = "Spanish"
CONFIGcombo_swedish = "Swedish"
CONFIGcombo_italian = "Italian"
CONFIGcombo_greek = "Greek"
CONFIGcombo_serbian = "Serbian"
CONFIGcombo_russian = "Russian"
CONFIGcombo_dutch = "Dutch"
CONFIGcombo_french = "French"

CONFIGmsgbox_account = "You didnt introduce an account."
CONFIGmsgbox_password = "You didnt introduce an password."
CONFIGmsgbox_nonumeric = "You cant take numeric names."
CONFIGmsgbox_portnoempty = "You didnt introduce an port."
CONFIGmsgbox_namenoempty = "You didnt introduce an name."
CONFIGmsgbox_ipnoempty = "You didnt introduce an IP."

' Chat form ..
CHATcommand_send = "&Send"
CHATcommand_clear = "&Clear"

CHATtimetext = " The time is "

' List form ..
LISTcaption = "Online List"
LISTcommand_close = "&Close"

' Send File form ..
SFlabel_filename = " File Name:"
SFlabel_sendingfile = "Sending:"
SFlabel_sent = "0.0% Sent"
SFlabel_sendto = "Send to:"

SFmsgbox_nousersel = "No user selected."
SFmsgbox_nofilesel = "No file selected."
SFmsgbox_incfile = "You are getting an incomming file, do you want to accept?"
SFmsgbox_filedecilined = "File transfer was decilined."

SFcommand_browse = "&Search .."
SFcommand_sendfile = "Send"
SFcommand_cancelsending = "Cancel .."

DESPtext_newmsg = "New Message!"
DESPtext_dcserver = "Disconnected from Server!"
End Sub

Public Sub SetLangSpanish()

MDIcommand_config = "&Configuración"
MDIcommand_chat = "Ch&at"
MDIcommand_sendfile = "&Enviar Archivo"
MDIcommand_onlinelist = "&Lista Online"

MDIstatusbar_disconnected = "Estado: Desconectado"
MDIstatusbar_dcfromserver = "Estado: Desconectado del servidor"
MDIstatusbar_connected = "Estado: Disponible "
MDIstatusbar_connectionproblem = "Estado: Desconectado por problemas de conexión"
MDIstatusbar_connecting = "Estado: Conectando "

MDImsgbox_config_notify = "Alguna configuración de archivos estan caducados o dañados, Peach busca el problema y lo arreglará en el siguiente lanzamiento del programa."
MDImsgbox_nametaken = "Este nombre ya esta cogido."
MDImsgbox_wrong_account = "La cuenta no existe o es incorrecta."
MDImsgbox_wrong_password = "La contraseña es incorrecta."
MDImsgbox_banned = "Esta cuenta esta baneada."

CONFIGcommand_connect = "&Conectar"
CONFIGcommand_disconnect = "&Desconectar"
CONFIGcommand_language = "&Idioma"
CONFIGcommand_update = "&Actualizar"
CONFIGcommand_register = "&Registrar cuenta"

CONFIGcheck_savepassword = "&Save Password"

CONFIGlabel_CI_name = "Nombre: "
CONFIGlabel_selectlanguage = "Elige tu idioma:"

CONFIGframe_personal = "Informaciónes personales"
CONFIGframe_connection = "Informaciónes de conexión"

CONFIGcombo_german = "Aleman"
CONFIGcombo_english = "Inglés"
CONFIGcombo_spanish = "Español"
CONFIGcombo_swedish = "Sueco"
CONFIGcombo_italian = "Italiano"
CONFIGcombo_dutch = "Holandés"
CONFIGcombo_serbian = "Serbio"
CONFIGcombo_french = "Frances"

CONFIGmsgbox_account = "You didnt introduce an account."
CONFIGmsgbox_password = "You didnt introduce an password."
CONFIGmsgbox_nonumeric = "No puedes coger nombres con numeros."
CONFIGmsgbox_portnoempty = "No has introducido un puerto."
CONFIGmsgbox_namenoempty = "No has introducido un nombre."
CONFIGmsgbox_ipnoempty = "No has introducido una direccion."

CHATcommand_send = "&Enviar"
CHATcommand_clear = "&Limpiar"

CHATtimetext = " El tiempo es "

LISTcaption = "Lista Online"
LISTcommand_close = "&Cerrar"

SFlabel_filename = " Nombre del archivo:"
SFlabel_sendingfile = "Enviando:"
SFlabel_sent = "0.0% Enviado"
SFlabel_sendto = "Enviar a:"

SFmsgbox_nousersel = "No has seleccionado a una persona."
SFmsgbox_nofilesel = "No has seleccionado a un archivo."
SFmsgbox_incfile = "Estas recibiendo un archivo, quieres aceptar?"
SFmsgbox_filedecilined = "El envio ha sido rechazado."

SFcommand_browse = "&Buscar .."
SFcommand_sendfile = "Enviar"
SFcommand_cancelsending = "Cancelar .."

DESPtext_newmsg = "Nuevo mensaje!"
DESPtext_dcserver = "Desconectado del servidor!"
End Sub

Public Sub SetLangSwedish()
' MDI form ..
MDIcommand_config = "&Inställningar"
MDIcommand_chat = "Ch&att"
MDIcommand_sendfile = "&Sänd fil"
MDIcommand_onlinelist = "&Online Lista"

MDIstatusbar_disconnected = "Status: Frånkopplad"
MDIstatusbar_dcfromserver = "Status: Koppla ifrån servern"
MDIstatusbar_connected = "Status: Anslut till "
MDIstatusbar_connectionproblem = "Status: Avkopplad på grund av anslutningsproblem"
MDIstatusbar_connecting = "Status: Ansluter till "

'MDImsgbox_errorHandlerFormLoad
MDImsgbox_config_notify = "Några Konfiguration filer är gamla eller skadade, Peach hittade problemet och det kommer bli reparerat nästa gång du kör programmet."
MDImsgbox_nametaken = "Namnet är upptaget."
MDImsgbox_wrong_account = "The account does not exist or is wrong."
MDImsgbox_wrong_password = "The password is wrong."
MDImsgbox_banned = "This account is banned."

' Config form
CONFIGcommand_connect = "&Anslut"
CONFIGcommand_disconnect = "&Frånkoppla"
CONFIGcommand_language = "&Språk"
CONFIGcommand_update = "&Update"
CONFIGcommand_register = "&Register Account"

CONFIGcheck_savepassword = "&Save Password"

CONFIGlabel_CI_name = "Namn: "
CONFIGlabel_selectlanguage = "Välj språk:"

CONFIGframe_personal = "Personal Information"
CONFIGframe_connection = "Connection Information"

CONFIGcombo_german = "Tyska"
CONFIGcombo_english = "Engelska"
CONFIGcombo_spanish = "Spanska"
CONFIGcombo_swedish = "Svenska"
CONFIGcombo_italian = "Italienska"
CONFIGcombo_greek = "Grekiska"
CONFIGcombo_serbian = "Serbiska"
CONFIGcombo_russian = "Ryska"
CONFIGcombo_dutch = "Holländska"
CONFIGcombo_french = "Franska"

CONFIGmsgbox_account = "You didnt introduce an account."
CONFIGmsgbox_password = "You didnt introduce an password."
CONFIGmsgbox_nonumeric = "Du kan inte använda siffror i namnet."
CONFIGmsgbox_portnoempty = "Du angav inget portnummer."
CONFIGmsgbox_namenoempty = "Du angav inte ett namn."
CONFIGmsgbox_ipnoempty = "Du angav inte ett IP."

' Chat form ..
CHATcommand_send = "&Sänd"
CHATcommand_clear = "&Rensa"

CHATtimetext = " Tiden är "

' List form ..
LISTcaption = "Online Lista"
LISTcommand_close = "&Stäng"

' Send file form ..
SFlabel_filename = " Fil Namn:"
SFlabel_sendingfile = "Sänder:"
SFlabel_sent = "0.0% Sänt"
SFlabel_sendto = "Send to:"

SFmsgbox_nousersel = "No user selected."
SFmsgbox_nofilesel = "No file selected."
SFmsgbox_incfile = "You are getting an incomming file, do you want to accept?"
SFmsgbox_filedecilined = "File transfer was decilined."

SFcommand_browse = "&Sök .."
SFcommand_sendfile = "Sänd"

DESPtext_newmsg = "New Message!"
DESPtext_dcserver = "Koppla ifrån servern!"
End Sub

Public Sub SetLangItalian()
' Mdi form
MDIcommand_config = "&Configurazione"
MDIcommand_chat = "Ch&at"
MDIcommand_sendfile = "&Invia File"
MDIcommand_onlinelist = "&Lista contatti Online"

MDIstatusbar_disconnected = "Stato: Disconnesso"
MDIstatusbar_dcfromserver = "Stato: Disconnesso dal Server"
MDIstatusbar_connected = "Stato: Connesso a "
MDIstatusbar_connectionproblem = "Stato: Disconnesso a causa di problemi di connessione"
MDIstatusbar_connecting = "Stato: Connessione a "

'MDImsgbox_errorHandlerFormLoad
MDImsgbox_config_notify = "Alcuni file della configurazione potrebbero essere obsoleti o danneggiati, Peach ha riscontrato il problema e lo corregera' al prossimo avvio."
MDImsgbox_nametaken = "Il nome immesso e' gia' in uso."
MDImsgbox_wrong_account = "The account does not exist or is wrong."
MDImsgbox_wrong_password = "The password is wrong."
MDImsgbox_banned = "This account is banned."

' Config form ..
CONFIGcommand_connect = "&Connesso"
CONFIGcommand_disconnect = "&Disconnesso"
CONFIGcommand_language = "&Lingua"
CONFIGcommand_update = "&Update"
CONFIGcommand_register = "&Register Account"

CONFIGcheck_savepassword = "&Save Password"

CONFIGlabel_CI_name = "Nome: "
CONFIGlabel_selectlanguage = "Seleziona la tua lingua:"

CONFIGframe_personal = "Personal Information"
CONFIGframe_connection = "Connection Information"

CONFIGcombo_german = "Tedesco"
CONFIGcombo_english = "Inglese"
CONFIGcombo_spanish = "Spagnolo"
CONFIGcombo_swedish = "Svedese"
CONFIGcombo_italian = "Italiano"
CONFIGcombo_greek = "Greco"
CONFIGcombo_serbian = "Serbo"
CONFIGcombo_russian = "Russo"
CONFIGcombo_dutch = "Olandese"
CONFIGcombo_french = "Francese"

CONFIGmsgbox_account = "You didnt introduce an account."
CONFIGmsgbox_password = "You didnt introduce an password."
CONFIGmsgbox_nonumeric = "Non puoi immettere nomi composti da numeri."
CONFIGmsgbox_portnoempty = "Non hai selezionato una porta valida."
CONFIGmsgbox_namenoempty = "Non hai immesso un Nome utente."
CONFIGmsgbox_ipnoempty = "Non hai immesso un IP."

' Chat form ..
CHATcommand_send = "&Invia"
CHATcommand_clear = "&Clear"

CHATtimetext = " L'ora e' "

' List form ..
LISTcaption = "Lista Online"
LISTcommand_close = "&Chiudi"

' Send file form ..
SFlabel_filename = " Nome file:"
SFlabel_sendingfile = "Inviando:"
SFlabel_sent = "0.0% Inviato"
SFlabel_sendto = "Send to:"

SFmsgbox_nousersel = "No user selected."
SFmsgbox_nofilesel = "No file selected."
SFmsgbox_incfile = "You are getting an incomming file, do you want to accept?"
SFmsgbox_filedecilined = "File transfer was decilined."

SFcommand_browse = "&Cerca .."
SFcommand_sendfile = "Invia"
SFcommand_cancelsending = "Annulla .."

DESPtext_newmsg = "New Message!"
DESPtext_dcserver = "Disconnesso dal Server!"
End Sub

Public Sub SetLangSerbian()
' Mdi form ..
MDIcommand_config = "&Konfiguracija"
MDIcommand_chat = "Ch&at"
MDIcommand_sendfile = "&Slanje fajla"
MDIcommand_onlinelist = "&Onlajn lista"

MDIstatusbar_disconnected = "Status: Veza je prekinuta"
MDIstatusbar_dcfromserver = "Status: Veza sa serverom je prekinuta"
MDIstatusbar_connected = "Status: Povezi se "
MDIstatusbar_connectionproblem = "Status: Problem sa konekcijom veza je prekinuta "
MDIstatusbar_connecting = "Status: Povezi se "

'MDImsgbox_errorHandlerFormLoad
MDImsgbox_config_notify = "Datoteka konfigurac. Je zastarela ili ostecena, problem ce biti pronadjen i popravljen sledecim pokretanjem programa."
MDImsgbox_nametaken = "Ime je vec zauzeto."
MDImsgbox_wrong_account = "The account does not exist or is wrong."
MDImsgbox_wrong_password = "The password is wrong."
MDImsgbox_banned = "This account is banned."

' Config form ..
CONFIGcommand_connect = "&Povezi se"
CONFIGcommand_disconnect = "&Veza je prekinuta"
CONFIGcommand_language = "&Jezik"
CONFIGcommand_update = "&Update"
CONFIGcommand_register = "&Register Account"

CONFIGcheck_savepassword = "&Save Password"

CONFIGlabel_CI_name = "Ime :"
CONFIGlabel_selectlanguage = "Dodaj svoj jezik:"

CONFIGframe_personal = "Personal Information"
CONFIGframe_connection = "Connection Information"

CONFIGcombo_german = "Nemacki"
CONFIGcombo_english = "Engleski"
CONFIGcombo_spanish = "Spanski"
CONFIGcombo_swedish = "Svedski"
CONFIGcombo_italian = "Italijanski"
CONFIGcombo_greek = "Crcki"
CONFIGcombo_serbian = "Srpski"
CONFIGcombo_russian = "Ruski"
CONFIGcombo_dutch = "Holandski"
CONFIGcombo_french = "Francuski"

CONFIGmsgbox_account = "You didnt introduce an account."
CONFIGmsgbox_password = "You didnt introduce an password."
CONFIGmsgbox_nonumeric = "Ne mozete uzeti numericka imena."
CONFIGmsgbox_portnoempty = "Niste uneli port."
CONFIGmsgbox_namenoempty = "Niste uneli ime"
CONFIGmsgbox_ipnoempty = "Niste uneli IP"

' Chat form ..
CHATcommand_send = "&Posalji"
CHATcommand_clear = "&Obrisi"

CHATtimetext = " Vreme je "

' List form ..
LISTcaption = "Onlajn lista"
LISTcommand_close = "&Zatvori"

' Send file form ..
SFlabel_filename = " Ime  arhive:"
SFlabel_sendingfile = "Slanje:"
SFlabel_sent = "0.0% Poslato"
SFlabel_sendto = "Send to:"

SFmsgbox_nousersel = "No user selected."
SFmsgbox_nofilesel = "No file selected."
SFmsgbox_incfile = "You are getting an incomming file, do you want to accept?"
SFmsgbox_filedecilined = "File transfer was decilined."

SFcommand_browse = "Trazi .."
SFcommand_sendfile = "Posalji"
SFcommand_cancelsending = "Otkazhi .."

DESPtext_newmsg = "New Message!"
DESPtext_dcserver = "Veza sa serverom je prekinuta!"
End Sub

Public Sub SetLangDutch()
MDIcommand_config = "&Configuratie"
MDIcommand_chat = "Ch&at"
MDIcommand_sendfile = "&Bestand Verzenden"
MDIcommand_onlinelist = "&Online List"

MDIstatusbar_disconnected = "Status: Verbinding verbroken"
MDIstatusbar_dcfromserver = "Status: verbinding verbroken met de server"
MDIstatusbar_connected = "Status: verbonden met "
MDIstatusbar_connectionproblem = "Status: verbinding verbroken wegens connectie problemen"
MDIstatusbar_connecting = "Status: verbinden met "

'MDImsgbox_errorHandlerFormLoad
MDImsgbox_config_notify = "Enkele bestanden zijn oud of beschadigd, Peach heeft het probleem gevonden en zal het herstellen bij de volgende herstart."
MDImsgbox_nametaken = "Deze naam is niet beschikbaar."
MDImsgbox_wrong_account = "The account does not exist or is wrong."
MDImsgbox_wrong_password = "The password is wrong."
MDImsgbox_banned = "This account is banned."

CONFIGcommand_connect = "&Verbind"
CONFIGcommand_disconnect = "&Verbreek de verbinding"
CONFIGcommand_language = "&Taal"
CONFIGcommand_update = "&Update"
CONFIGcommand_register = "&Register Account"

CONFIGcheck_savepassword = "&Save Password"

CONFIGlabel_CI_name = "Naam: "
CONFIGlabel_selectlanguage = "Selecteer jou taal:"

CONFIGframe_personal = "Personal Information"
CONFIGframe_connection = "Connection Information"

CONFIGcombo_german = "Duits"
CONFIGcombo_english = "Engels"
CONFIGcombo_spanish = "Spaans"
CONFIGcombo_swedish = "Zweeds"
CONFIGcombo_italian = "Italiaans"
CONFIGcombo_greek = "Grieks"
CONFIGcombo_serbian = "Serbisch"
CONFIGcombo_russian = "Russisch"
CONFIGcombo_dutch = "Nederlands"
CONFIGcombo_french = "Frans"

CONFIGmsgbox_account = "You didnt introduce an account."
CONFIGmsgbox_password = "You didnt introduce an password."
CONFIGmsgbox_nonumeric = "U kan geen naam nemen dat nummers bevat."
CONFIGmsgbox_portnoempty = "U hebt geen poort ingesteld."
CONFIGmsgbox_namenoempty = "U hebt geen naam gegoven."
CONFIGmsgbox_ipnoempty = "U hebt geen IP gegoven."

CHATcommand_send = "&Zend"
CHATcommand_clear = "&Leegmaken"

CHATtimetext = " De Tijd is: "

LISTcaption = "Online List"
LISTcommand_close = "&Sluiten"

SFlabel_filename = " Bestandsnaam:"
SFlabel_sendingfile = "verZenden:"
SFlabel_sent = "0.0% verzonden"
SFlabel_sendto = "Send to:"

SFmsgbox_nousersel = "No user selected."
SFmsgbox_nofilesel = "No file selected."
SFmsgbox_incfile = "You are getting an incomming file, do you want to accept?"
SFmsgbox_filedecilined = "File transfer was decilined."

SFcommand_browse = "&Zoeken .."
SFcommand_sendfile = "&Stuur"
SFcommand_cancelsending = "&Annuleren .."

DESPtext_newmsg = "New Message!"
DESPtext_dcserver = "Verbinding verbroken met de server"
End Sub

Public Sub SetLangFrench()
MDIcommand_config = "&Configuration"
MDIcommand_chat = "Ch&at"
MDIcommand_sendfile = "&Envoi File"
MDIcommand_onlinelist = "&Liste contact Online"

MDIstatusbar_disconnected = "Etat: Deconnecté"
MDIstatusbar_dcfromserver = "Etat: Deconnecté du Server"
MDIstatusbar_connected = "Etat: Connecté à "
MDIstatusbar_connectionproblem = "Etat: Deconnecté à cause de problèmes do connection"
MDIstatusbar_connecting = "Etat: Connection à "

MDImsgbox_config_notify = "Quelques files de la configuration pourrait etre daumagés ou obsolète , Peach a trouvé le problème et le corrigerà au prochain envoi."
MDImsgbox_nametaken = "Le nom inséré est déjà utilizé."
MDImsgbox_wrong_account = "The account does not exist or is wrong."
MDImsgbox_wrong_password = "The password is wrong."
MDImsgbox_banned = "This account is banned."

CONFIGcommand_connect = "&Connecté"
CONFIGcommand_disconnect = "&Deconnecté"
CONFIGcommand_language = "&Langue"
CONFIGcommand_update = "&Update"
CONFIGcommand_register = "&Register Account"

CONFIGcheck_savepassword = "&Save Password"

CONFIGlabel_CI_name = "Nome: "
CONFIGlabel_selectlanguage = "Choisissez votre langue:"

CONFIGframe_personal = "Personal Information"
CONFIGframe_connection = "Connection Information"

CONFIGcombo_german = "Alleman"
CONFIGcombo_english = "Anglais"
CONFIGcombo_spanish = "Espagnol"
CONFIGcombo_swedish = "Suédois"
CONFIGcombo_italian = "Italien"
CONFIGcombo_greek = "Grèque"
CONFIGcombo_serbian = "Serbois"
CONFIGcombo_russian = "Russe"
CONFIGcombo_dutch = "Hollandais"
CONFIGcombo_french = "Français"

CONFIGmsgbox_account = "You didnt introduce an account."
CONFIGmsgbox_password = "You didnt introduce an password."
CONFIGmsgbox_nonumeric = "Tu ne peut pas insérer noms composé de numeros."
CONFIGmsgbox_portnoempty = "Tu n'as pas selectionner une porte valide."
CONFIGmsgbox_namenoempty = "Tu n'as pas innecté un Nom utilizateur."
CONFIGmsgbox_ipnoempty = "Tu n'as pas innecté un IP."

CHATcommand_send = "&Envoi"
CHATcommand_clear = "&Clear"

CHATtimetext = " L'heure est "

LISTcaption = "Liste Online"
LISTcommand_close = "&Ferme"

SFlabel_filename = " Nom file:"
SFlabel_sendingfile = "Envoyant:"
SFlabel_sent = "0.0% Envoyé"
SFlabel_sendto = "Send to:"

SFmsgbox_nousersel = "No user selected."
SFmsgbox_nofilesel = "No file selected."
SFmsgbox_incfile = "You are getting an incomming file, do you want to accept?"
SFmsgbox_filedecilined = "File transfer was decilined."

SFcommand_browse = "&Cherche .."
SFcommand_sendfile = "Envoi"
SFcommand_cancelsending = "Annuler .."

DESPtext_newmsg = "New Message!"
DESPtext_dcserver = "Deconnecté du Server"
End Sub
