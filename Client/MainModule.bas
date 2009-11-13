Attribute VB_Name = "modMain"
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32" ()

Sub Main()
'If the application is already open then close it
If App.PrevInstance Then
    End
End If

'Load Windows own style
Call InitCommonControls

'Start loading the ini values
LoadIniValue

If Setting.VALIDATE = 0 Then
    Select Case Setting.LANGUAGE
    Case 0 'German
        SET_LANG_GERMAN
    Case 1 'English
        SET_LANG_ENGLISH
    Case 2 'Spanish
        SET_LANG_SPANISH
    Case 3 'Swedish
        SET_LANG_SWEDISH
    Case 4 'Italian
        SET_LANG_ITALIAN
    Case 5 'Dutch
        SET_LANG_DUTCH
    Case 6 'Serbian
        SET_LANG_SERBIAN
    Case 7 'French
        SET_LANG_FRENCH
    End Select
    frmMain.Show
    SetScheme True
Else
    'Set language default -> english
    SET_LANG_ENGLISH
    frmLanguage.Show
End If

WriteIniValue App.Path & "\Config.ini", "Revision", "Number", pRev
End Sub

Private Sub LoadIniValue()
'Open Settings variable
With Setting

'Read 'IP' from .ini file
If Len(Trim$(ReadIniValue(App.Path & "\Config.ini", "Connection", "IP"))) <> 0 Then
    .SERVER_IP = ReadIniValue(App.Path & "\Config.ini", "Connection", "IP")
Else
    .SERVER_IP = "127.0.0.1"
End If
 
'Read 'Port' from .ini file
If Len(Trim$(ReadIniValue(App.Path & "\Config.ini", "Connection", "Port"))) <> 0 Then
    .SERVER_PORT = ReadIniValue(App.Path & "\Config.ini", "Connection", "Port")
Else
    .SERVER_PORT = 4728
End If

'Read 'Nickname' from .ini file
If Len(Trim$(ReadIniValue(App.Path & "\Config.ini", "Private", "Nickname"))) <> 0 Then
    .NICKNAME = ReadIniValue(App.Path & "\Config.ini", "Private", "Nickname")
End If

'Account Tick
If Len(Trim$(ReadIniValue(App.Path & "\Config.ini", "Private", "AccountTick"))) = 0 Then
    .ACCOUNT_TICK = False
Else
    .ACCOUNT_TICK = CBool(ReadIniValue(App.Path & "\Config.ini", "Private", "AccountTick"))
End If

'If Account Tick = True then read account
If .ACCOUNT_TICK = True Then
    If Len(Trim$(ReadIniValue(App.Path & "\Config.ini", "Private", "Account"))) <> 0 Then
        .ACCOUNT = ReadIniValue(App.Path & "\Config.ini", "Private", "Account")
    End If
End If

'Password Tick
If Len(Trim$(ReadIniValue(App.Path & "\Config.ini", "Private", "PasswordTick"))) = 0 Then
    .PASSWORD_TICK = False
Else
    .PASSWORD_TICK = CBool(ReadIniValue(App.Path & "\Config.ini", "Private", "PasswordTick"))
End If

'If Password Tick = True then read password
If .PASSWORD_TICK = True Then
    .PASSWORD = DeCode(DeCode(ReadIniValue(App.Path & "\Config.ini", "Private", "Password")))
Else
    .PASSWORD = vbNullString
End If

'Load configuration from .ini file into switch variables
If Len(Trim$(ReadIniValue(App.Path & "\Config.ini", "Language", "Validate"))) = 0 Then
    .VALIDATE = 1 'Default choose language
Else
    .VALIDATE = ReadIniValue(App.Path & "\Config.ini", "Language", "Validate")
End If

If Len(Trim$(ReadIniValue(App.Path & "\Config.ini", "Language", "Language"))) = 0 Then
    .LANGUAGE = 1  'Default language english
Else
    .LANGUAGE = ReadIniValue(App.Path & "\Config.ini", "Language", "Language")
End If

'Load 'Top' position from ini, if there is non take default value ( 1200 )
If Len(Trim$(ReadIniValue(App.Path & "\Config.ini", "Position", "Top"))) = 0 Then
    Setting.MAIN_TOP = 1200
Else
    Setting.MAIN_TOP = ReadIniValue(App.Path & "\Config.ini", "Position", "Top")
End If

'Load 'Left' position from ini, if there is non take default value ( 1200 )
If Len(Trim$(ReadIniValue(App.Path & "\Config.ini", "Position", "Left"))) = 0 Then
    Setting.MAIN_LEFT = 1200
Else
    Setting.MAIN_LEFT = ReadIniValue(App.Path & "\Config.ini", "Position", "Left")
End If

'Load scheme color
If Len(Trim$(ReadIniValue(App.Path & "\Config.ini", "Private", "SchemeColor"))) = 0 Then
    Setting.SCHEME_COLOR = &HF4F4F4
Else
    Setting.SCHEME_COLOR = ReadIniValue(App.Path & "\Config.ini", "Private", "SchemeColor")
End If

'Close Settings variable
End With
End Sub

Public Sub SetScheme(IsFirst As Boolean)
Dim SC As String
SC = Setting.SCHEME_COLOR

With frmMain
    .BackColor = SC
    .Picture1.BackColor = SC
End With

frmChat.BackColor = SC

With frmConfig
    .BackColor = SC
    .Label1.BackColor = SC
    .Label2.BackColor = SC
    .lblAccount.BackColor = SC
    .lblAuthor.BackColor = SC
    .lblNickname.BackColor = SC
    .lblPassword.BackColor = SC
    .lblVersion.BackColor = SC
End With

With frmSendFile
    .BackColor = SC
    .picProgress.BackColor = SC
    .Label1.BackColor = SC
    .Label4.BackColor = SC
    .lblFileToSend.BackColor = SC
    .lblProgress.BackColor = SC
    .lblSendSpeed.BackColor = SC
End With

If IsFirst = False Then
    With frmSettings
        .BackColor = SC
        .Frame1.BackColor = SC
        .Frame2.BackColor = SC
        .Frame3.BackColor = SC
        .Label1.BackColor = SC
        .Label2.BackColor = SC
        .Label3.BackColor = SC
        .SaveAccount.BackColor = SC
        .SavePassword.BackColor = SC
    End With
End If
With frmSociety
    .BackColor = SC
    .SSTab1.BackColor = SC
End With
End Sub
