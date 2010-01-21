Attribute VB_Name = "modLang"
Option Explicit

Public MSG_COME_ONLINE      As String
Public MSG_GONE_OFFLINE     As String
Public MSG_ANNOUNCE         As String

Public Sub SetLanguageByID(pID As Long)
Select Case pID
    Case 0 'German
        MSG_COME_ONLINE = "%u ist online gekommen."
        MSG_GONE_OFFLINE = "%u ist offline gegangen."
        MSG_ANNOUNCE = "[%u kündigt an]: "
        
    Case 1 'English
        MSG_COME_ONLINE = "%u has come online."
        MSG_GONE_OFFLINE = "%u has gone offline."
        MSG_ANNOUNCE = "[%u announces]: "
        
    Case 2 'Spanish
        MSG_COME_ONLINE = "%u se ha conectado."
        MSG_GONE_OFFLINE = "%u se ha desconectado."
        MSG_ANNOUNCE = "[%u anuncia]: "
        
    Case 3 'Swedish
        MSG_COME_ONLINE = "%u har kommit online."
        MSG_GONE_OFFLINE = "%u har gått offline."
        MSG_ANNOUNCE = "[%u meddelar]: "
        
    Case 4 'Italian
        MSG_COME_ONLINE = "%u e' online."
        MSG_GONE_OFFLINE = "%u ora e' offline."
        MSG_ANNOUNCE = "[%u annuncia]: "
        
    Case 5 'Dutch
        MSG_COME_ONLINE = "%u has come online."
        MSG_GONE_OFFLINE = "%u is offline gegaan."
        MSG_ANNOUNCE = "[%u kondigt]: "
        
    Case 6 'Serbian
        MSG_COME_ONLINE = "*ERROR*"
        MSG_GONE_OFFLINE = "*ERROR*"
        MSG_ANNOUNCE = "*ERROR*"
        
    Case 7 'French
        MSG_COME_ONLINE = "%u has come online."
        MSG_GONE_OFFLINE = "%u s'est déconnecté."
        MSG_ANNOUNCE = "[%u announcer]: "
        
    Case Else
        MSG_COME_ONLINE = "%u has come online."
        MSG_GONE_OFFLINE = "%u has gone offline."
        MSG_ANNOUNCE = "[%u announces]: "
        
End Select
End Sub
