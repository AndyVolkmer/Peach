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

' Configuration form ..
Public CONFIGcommand_connect            As String
Public CONFIGcommand_disconnect         As String
Public CONFIGcommand_language           As String

Public CONFIGlabel_CI_name              As String
Public CONFIGlabel_selectlanguage       As String

Public CONFIGframe_config               As String
Public CONFIGframe_client               As String
Public CONFIGframe_server               As String

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

Public SFcommand_browse                 As String
Public SFcommand_sendfile               As String
Public SFcommand_cancelsending          As String

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

'MDImsgbox_errorHandlerFormLoad
MDImsgbox_config_notify = "Einige Konfigurationsdateien sind veraltet oder wurden besch�digt, Peach fand den Fehler und wird es mit dem n�chsten Neustart korrigieren."
MDImsgbox_nametaken = "Der Name ist bereits vergeben."

CONFIGcommand_connect = "&Verbinden"
CONFIGcommand_disconnect = "&Verbindung trenn."
CONFIGcommand_language = "&Sprache"

CONFIGlabel_CI_name = "Name: "
CONFIGlabel_selectlanguage = "W�hle deine Sprache aus:"

CONFIGframe_config = "Einstellungen"
CONFIGframe_client = "Client Informationen: "
CONFIGframe_server = "Server Informationen: "

CONFIGcombo_german = "Deutsch"
CONFIGcombo_english = "Englisch"
CONFIGcombo_spanish = "Spanisch"
CONFIGcombo_swedish = "Schwedisch"
CONFIGcombo_italian = "Italienisch"
CONFIGcombo_greek = "Griechisch"
CONFIGcombo_serbian = "Serbisch"
CONFIGcombo_russian = "Russisch"
CONFIGcombo_dutch = "Niederl�ndisch"
CONFIGcombo_french = "Franz�sisch"

CONFIGmsgbox_nonumeric = "Du kannst keine Ziffern in deinem Namen haben."
CONFIGmsgbox_portnoempty = "Du hast keinen Port eingeben."
CONFIGmsgbox_namenoempty = "Du hast keinen Namen eingeben."
CONFIGmsgbox_ipnoempty = "Du hast keine IP eingeben."

CHATcommand_send = "&Senden"
CHATcommand_clear = "&L�schen"

CHATtimetext = " Die Zeit betr�gt "

LISTcaption = "Online Liste"
LISTcommand_close = "&Schliessen"

SFlabel_filename = " Datei Name:"
SFlabel_sendingfile = "Sende:"
SFlabel_sent = "Gesendet"

SFcommand_browse = "&Suchen .."
SFcommand_sendfile = "Senden"
SFcommand_cancelsending = "Abbrechen .."
End Sub

Public Sub SetLangEnglish()

' MDI form ..
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

' Configuration form ..
CONFIGcommand_connect = "&Connect"
CONFIGcommand_disconnect = "&Disconnect"
CONFIGcommand_language = "&Language"

CONFIGlabel_CI_name = "Name: "
CONFIGlabel_selectlanguage = "Select your language:"

CONFIGframe_config = "Configuration"
CONFIGframe_client = "Client Information: "
CONFIGframe_server = "Server Information: "

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

CONFIGmsgbox_nonumeric = "You cant take numeric names."
CONFIGmsgbox_portnoempty = "You didnt introduce a port."
CONFIGmsgbox_namenoempty = "You didnt introduce a name."
CONFIGmsgbox_ipnoempty = "You didnt introduce a IP."

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
SFlabel_sent = "Sent"

SFcommand_browse = "&Search .."
SFcommand_sendfile = "Send"
SFcommand_cancelsending = "Cancel .."
End Sub

Public Sub SetLangSpanish()

' MDI form ..
MDIcommand_config = "&Configuraci�n"
MDIcommand_chat = "Ch&at"
MDIcommand_sendfile = "&Enviar archivo"
MDIcommand_onlinelist = "&Lista de conectados"

MDIstatusbar_disconnected = "Status: Desconectado"
MDIstatusbar_dcfromserver = "Status: Desconectado del Servidor"
MDIstatusbar_connected = "Status: Conectado con "
MDIstatusbar_connectionproblem = "Status: Desconectado por problemas de conexi�n"
MDIstatusbar_connecting = "Status: Conectando con "

'MDImsgbox_errorHandlerFormLoad
MDImsgbox_config_notify = "Algunos archivos estaban da�ados o borrados, Peach iniciara con datos por defecto."
MDImsgbox_nametaken = "Este nombre ya esta cogido."

' Configuration form ..
CONFIGcommand_connect = "&Conectar"
CONFIGcommand_disconnect = "&Desconectar"
CONFIGcommand_language = "&Idioma"

CONFIGlabel_CI_name = "Nombre: "
CONFIGlabel_selectlanguage = "Elige tu idioma:"

CONFIGframe_config = "Configuraci�n"
CONFIGframe_client = "Informci�n del cliente: "
CONFIGframe_server = "Informci�n del servidor: "

CONFIGcombo_german = "Aleman"
CONFIGcombo_english = "Ingl�s"
CONFIGcombo_spanish = "Espa�ol"
CONFIGcombo_swedish = "Sueco"
CONFIGcombo_italian = "Italiano"
CONFIGcombo_greek = "Griego"
CONFIGcombo_serbian = "Serbio"
CONFIGcombo_russian = "Ruso"
CONFIGcombo_dutch = "Holand�s"
CONFIGcombo_french = "Franc�s"

CONFIGmsgbox_nonumeric = "No puedes cojer numeros como nombre."
CONFIGmsgbox_portnoempty = "No ha introducido ningun puerto."
CONFIGmsgbox_namenoempty = "No ha introducido ningun nombre."
CONFIGmsgbox_ipnoempty = "No ha introducido ninguna IP."

' Chat form ..
CHATcommand_send = "&Enviar"
CHATcommand_clear = "&Borrar"

CHATtimetext = " Son las "

' List form ..
LISTcaption = "Lista de conectados"
LISTcommand_close = "&Cerrar"

' Send File form ..
SFlabel_filename = " Nombre del Archivo:"
SFlabel_sendingfile = "Enviando:"
SFlabel_sent = "Enviado"

SFcommand_browse = "&Buscar .."
SFcommand_sendfile = "Enviar"
SFcommand_cancelsending = "Abortar .."
End Sub

Public Sub SetLangSwedish()
' MDI form ..
MDIcommand_config = "&Inst�llningar"
MDIcommand_chat = "Ch&att"
MDIcommand_sendfile = "&S�nd fil"
MDIcommand_onlinelist = "&Online Lista"

MDIstatusbar_disconnected = "Status: Fr�nkopplad"
MDIstatusbar_dcfromserver = "Status: Koppla ifr�n servern"
MDIstatusbar_connected = "Status: Anslut till "
MDIstatusbar_connectionproblem = "Status: Avkopplad p� grund av anslutningsproblem"
MDIstatusbar_connecting = "Status: Ansluter till "

'MDImsgbox_errorHandlerFormLoad
MDImsgbox_config_notify = "N�gra Konfiguration filer �r gamla eller skadade, Peach hittade problemet och det kommer bli reparerat n�sta g�ng du k�r programmet."
MDImsgbox_nametaken = "Namnet �r upptaget."

' Config form
CONFIGcommand_connect = "&Anslut"
CONFIGcommand_disconnect = "&Fr�nkoppla"
CONFIGcommand_language = "&Spr�k"

CONFIGlabel_CI_name = "Namn: "
CONFIGlabel_selectlanguage = "V�lj spr�k:"

CONFIGframe_config = "Inst�llningar"
CONFIGframe_client = "Anv�ndar Information: "
CONFIGframe_server = "Server Information: "

CONFIGcombo_german = "Tyska"
CONFIGcombo_english = "Engelska"
CONFIGcombo_spanish = "Spanska"
CONFIGcombo_swedish = "Svenska"
CONFIGcombo_italian = "Italienska"
CONFIGcombo_greek = "Grekiska"
CONFIGcombo_serbian = "Serbiska"
CONFIGcombo_russian = "Ryska"
CONFIGcombo_dutch = "Holl�ndska"
CONFIGcombo_french = "Franska"

CONFIGmsgbox_nonumeric = "Du kan inte anv�nda siffror i namnet."
CONFIGmsgbox_portnoempty = "Du angav inget portnummer."
CONFIGmsgbox_namenoempty = "Du angav inte ett namn."
CONFIGmsgbox_ipnoempty = "Du angav inte ett IP."

' Chat form ..
CHATcommand_send = "&S�nd"
CHATcommand_clear = "&Rensa"

CHATtimetext = " Tiden �r "

' List form ..
LISTcaption = "Online Lista"
LISTcommand_close = "&St�ng"

' Send file form ..
SFlabel_filename = " Fil Namn:"
SFlabel_sendingfile = "S�nder:"
SFlabel_sent = "S�nt"

SFcommand_browse = "&S�k .."
SFcommand_sendfile = "S�nd"
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

' Config form ..
CONFIGcommand_connect = "&Connesso"
CONFIGcommand_disconnect = "&Disconnesso"
CONFIGcommand_language = "&Lingua"

CONFIGlabel_CI_name = "Nome: "
CONFIGlabel_selectlanguage = "Seleziona la tua lingua:"

CONFIGframe_config = "Configurazione"
CONFIGframe_client = "Informazioni sul Client: "
CONFIGframe_server = "Informazioni sul Server: "

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
SFlabel_sent = "Inviato"

SFcommand_browse = "&Cerca .."
SFcommand_sendfile = "Invia"
SFcommand_cancelsending = "Annulla .."
End Sub

Public Sub SetLangSerbian()
' Mdi form ..
MDIcommand_config = "&Konfiguracija"
MDIcommand_chat = "Ch&at"
MDIcommand_sendfile = "&Slanje fajla"
MDIcommand_onlinelist = "&Onlajn lista"

MDIstatusbar_disconnected = "Status: Veza je prekinuta"
MDIstatusbar_dcfromserver = "Status: veza sa serverom je prekinuta"
MDIstatusbar_connected = "Status: Povezi se "
MDIstatusbar_connectionproblem = "Status: Problem sa konekcijom veza je prekinuta "
MDIstatusbar_connecting = "Status: Povezi se "

'MDImsgbox_errorHandlerFormLoad
MDImsgbox_config_notify = "Datoteka konfigurac. Je zastarela ili ostecena, problem ce biti pronadjen i popravljen sledecim pokretanjem programa."
MDImsgbox_nametaken = "Ime je vec zauzeto."

' Config form ..
CONFIGcommand_connect = "&Povezi se"
CONFIGcommand_disconnect = "&Veza je prekinuta"
CONFIGcommand_language = "&Jezik"

CONFIGlabel_CI_name = "Ime :"
CONFIGlabel_selectlanguage = "Dodaj svoj jezik:"

CONFIGframe_config = "Konfiguracija"
CONFIGframe_client = "Client informacije: "
CONFIGframe_server = "Server informacije: "

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
SFlabel_sent = "Poslato "

SFcommand_browse = "Trazi .."
SFcommand_sendfile = "Posalji"
SFcommand_cancelsending = "Otkazhi .."
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

CONFIGcommand_connect = "&Verbind"
CONFIGcommand_disconnect = "&Verbreek de verbinding"
CONFIGcommand_language = "&Taal"

CONFIGlabel_CI_name = "Naam: "
CONFIGlabel_selectlanguage = "Selecteer jou taal:"

CONFIGframe_config = "Configuratie"
CONFIGframe_client = "Client Informatie: "
CONFIGframe_server = "Server Informatie: "

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
SFlabel_sent = "verzonden"

SFcommand_browse = "&Zoeken .."
SFcommand_sendfile = "Stuur"
SFcommand_cancelsending = "Annuleren .."
End Sub

Public Sub SetLangFrench()
MDIcommand_config = "&Configuration"
MDIcommand_chat = "Ch&at"
MDIcommand_sendfile = "&Envoi File"
MDIcommand_onlinelist = "&Liste contact Online"

MDIstatusbar_disconnected = "Etat: Deconnect�"
MDIstatusbar_dcfromserver = "Etat: Deconnect� du Server"
MDIstatusbar_connected = "Etat: Connect� � "
MDIstatusbar_connectionproblem = "Etat: Deconnect� � cause de probl�mes do connection"
MDIstatusbar_connecting = "Etat: Connection � "

'MDImsgbox_errorHandlerFormLoad
MDImsgbox_config_notify = "Quelques files de la configuration pourrait etre daumag�s ou obsol�te , Peach a trouv� le probl�me et le corriger� au prochain envoi."
MDImsgbox_nametaken = "Le nom ins�r� est d�j� utiliz�."

CONFIGcommand_connect = "&Connect�"
CONFIGcommand_disconnect = "&Deconnect�"
CONFIGcommand_language = "&Langue"

CONFIGlabel_CI_name = "Nome: "
CONFIGlabel_selectlanguage = "Choisissez votre langue:"

CONFIGframe_config = "Configuration"
CONFIGframe_client = "Informations sur Client: "
CONFIGframe_server = "Informations sur Server: "

CONFIGcombo_german = "Alleman"
CONFIGcombo_english = "Anglais"
CONFIGcombo_spanish = "Espagnol"
CONFIGcombo_swedish = "Su�dois"
CONFIGcombo_italian = "Italien"
CONFIGcombo_greek = "Gr�que"
CONFIGcombo_serbian = "Serbois"
CONFIGcombo_russian = "Russe"
CONFIGcombo_dutch = "Hollandais"
CONFIGcombo_french = "Fran�ais"

CONFIGmsgbox_nonumeric = "Tu ne peut pas ins�rer noms compos� de numeros."
CONFIGmsgbox_portnoempty = "Tu n'as pas selectionner une porte valide."
CONFIGmsgbox_namenoempty = "Tu n'as pas innect� un Nom utilizateur."
CONFIGmsgbox_ipnoempty = "Tu n'as pas innect� un IP."

CHATcommand_send = "&Envoi"
CHATcommand_clear = "&Clear"

CHATtimetext = " L'heure est "

LISTcaption = "Liste Online"
LISTcommand_close = "&Ferme"

SFlabel_filename = " Nom file:"
SFlabel_sendingfile = "Envoyant:"
SFlabel_sent = "Envoy�"

SFcommand_browse = "&Cherche .."
SFcommand_sendfile = "Envoi"
SFcommand_cancelsending = "Annuler .."
End Sub
