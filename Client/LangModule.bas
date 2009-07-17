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

Public CONFIGframe_config               As String
Public CONFIGframe_client               As String
Public CONFIGframe_server               As String

Public CONFIGcombo_german               As String
Public CONFIGcombo_english              As String
Public CONFIGcombo_spanish              As String

Public CONFIGmsgbox_nonumeric           As String
Public CONFIGmsgbox_portnoempty         As String
Public CONFIGmsgbox_namenoempty         As String
Public CONFIGmsgbox_ipnoempty           As String

' Chat form ..
Public CHATcommand_send                 As String
Public CHATcommand_clear                As String

' List form ..
Public LISTcaption                      As String
Public LISTcommand_close                As String

' Send File form ..
Public SFlabel_filename                 As String
Public SFlabel_sendingfile              As String
Public SFlabel_sent                     As String

Public SFcommand_browse                 As String
Public SFcommand_sendfile               As String

Public Sub SetLangGerman()

' MDI form ..
MDIcommand_config = "&Einstellungen"
MDIcommand_chat = "Ch&at"
MDIcommand_sendfile = "&Datei senden"
MDIcommand_onlinelist = "&Online Liste"

MDIstatusbar_disconnected = "Status: Nicht verbunden"
MDIstatusbar_dcfromserver = "Status: Verbindung mit Server unterbrochen"
MDIstatusbar_connected = "Status: Verbunden mit "
MDIstatusbar_connectionproblem = "Status: Verbindungs Problem"
MDIstatusbar_connecting = "Status: Verbinden mit "

'MDImsgbox_errorHandlerFormLoad = ""
MDImsgbox_config_notify = "Einige konfigurations Dateien sind beschädigt oder wurden gelöscht, einige Werte wurden auf einen Standard wert gesetzt."
MDImsgbox_nametaken = "Der Name wird bereits genutzt."

' Configuration form ..
CONFIGcommand_connect = "&Verbinden"
CONFIGcommand_disconnect = "&Verbindung trenn."
CONFIGcommand_language = "&Sprache"

CONFIGlabel_CI_name = "Name: "

CONFIGframe_config = "Einstellungen"
CONFIGframe_client = "Client Informationen: "
CONFIGframe_server = "Server Informationen: "

CONFIGcombo_german = "Deutsch"
CONFIGcombo_english = "Englisch"
CONFIGcombo_spanish = "Spanisch"

CONFIGmsgbox_nonumeric = "Dein Name kann nicht aus Nummern bestehen."
CONFIGmsgbox_portnoempty = "Sie haben keinen Port angegeben."
CONFIGmsgbox_namenoempty = "Sie haben keinen Namen angegeben."
CONFIGmsgbox_ipnoempty = "Sie haben keine IP angegeben."

' Chat form ..
CHATcommand_send = "&Senden"
CHATcommand_clear = "&Löschen"

' List form ..
LISTcaption = "Online Liste"
LISTcommand_close = "&Schliessen"

' Send File form ..
SFlabel_filename = " Datei Name:"
SFlabel_sendingfile = "Sende:"
SFlabel_sent = "Gesendet"

SFcommand_browse = "&Suchen .."
SFcommand_sendfile = "Senden"
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

CONFIGframe_config = "Configuration"
CONFIGframe_client = "Client Information: "
CONFIGframe_server = "Server Information: "

CONFIGcombo_german = "German"
CONFIGcombo_english = "English"
CONFIGcombo_spanish = "Spanish"

CONFIGmsgbox_nonumeric = "You cant take numeric names."
CONFIGmsgbox_portnoempty = "You didnt introduced a port."
CONFIGmsgbox_namenoempty = "You didnt introduced a name."
CONFIGmsgbox_ipnoempty = "You didnt introduced a IP."

' Chat form ..
CHATcommand_send = "&Send"
CHATcommand_clear = "&Clear"

' List form ..
LISTcaption = "Online List"
LISTcommand_close = "&Close"

' Send File form ..
SFlabel_filename = " File Name:"
SFlabel_sendingfile = "Sending:"
SFlabel_sent = "Sent"

SFcommand_browse = "&Search .."
SFcommand_sendfile = "Send"
End Sub

Public Sub SetLangSpanish()

' MDI form ..
MDIcommand_config = "&Configuración"
MDIcommand_chat = "Ch&at"
MDIcommand_sendfile = "&Enviar archivo"
MDIcommand_onlinelist = "&Lista de conectados"

MDIstatusbar_disconnected = "Status: Desconectado"
MDIstatusbar_dcfromserver = "Status: Desconectado del Servidor"
MDIstatusbar_connected = "Status: Conectado con "
MDIstatusbar_connectionproblem = "Status: Desconectado por problemas de conexión"
MDIstatusbar_connecting = "Status: Conectando con "

'MDImsgbox_errorHandlerFormLoad
MDImsgbox_config_notify = "Algunos archivos estaban dañados o borrados, Peach iniciara con datos por defecto."
MDImsgbox_nametaken = "Este nombre ya esta cogido."

' Configuration form ..
CONFIGcommand_connect = "&Conectar"
CONFIGcommand_disconnect = "&Desconectar"
CONFIGcommand_language = "&Idioma"

CONFIGlabel_CI_name = "Nombre: "

CONFIGframe_config = "Configuración"
CONFIGframe_client = "Informción del cliente: "
CONFIGframe_server = "Informción del servidor: "

CONFIGcombo_german = "Aleman"
CONFIGcombo_english = "Inglés"
CONFIGcombo_spanish = "Español"

CONFIGmsgbox_nonumeric = "No puedes cojer numeros como nombre."
CONFIGmsgbox_portnoempty = "No ha introducido ningun puerto."
CONFIGmsgbox_namenoempty = "No ha introducido ningun nombre."
CONFIGmsgbox_ipnoempty = "No ha introducido ninguna IP."

' Chat form ..
CHATcommand_send = "&Enviar"
CHATcommand_clear = "&Borrar"

' List form ..
LISTcaption = "Lista de conectados"
LISTcommand_close = "&Cerrar"

' Send File form ..
SFlabel_filename = " Nombre del Archivo:"
SFlabel_sendingfile = "Enviando:"
SFlabel_sent = "Enviado"

SFcommand_browse = "&Buscar .."
SFcommand_sendfile = "Enviar"
End Sub
