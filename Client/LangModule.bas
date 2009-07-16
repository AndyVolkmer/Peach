Attribute VB_Name = "LangModule"
Option Explicit

'Start variable support for languages
' MDI form ..
Dim MDIcommand_config               As String
Dim MDIcommand_chat                 As String
Dim MDIcommand_sendfile             As String
Dim MDIcommand_onlinelist           As String

Dim MDIstatusbar_disconnected       As String
Dim MDIstatusbar_dcfromserver       As String
Dim MDIstatusbar_connected          As String
Dim MDIstatusbar_connectionproblem  As String
Dim MDIstatusbar_connecting         As String

Dim MDImsgbox_errorHandlerFormLoad  As String
Dim MDImsgbox_config_notify         As String
Dim MDImsgbox_nametaken             As String

' Configuration form ..
Dim CONFIGcommand_connect           As String
Dim CONFIGcommand_disconnect        As String

Dim CONFIGlabel_CI_name             As String

