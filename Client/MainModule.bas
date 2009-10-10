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
    SetScheme
Else
    'Set language default -> english
    SET_LANG_ENGLISH
    frmLanguage.Show
End If

WriteIniValue App.Path & "\Config.ini", "Revision", "Number", Rev
End Sub

Private Sub LoadIniValue()
'Open Settings variable
With Setting

'Read 'IP' from .ini file
If Len(Trim$(ReadIniValue(App.Path & "\Config.ini", "Connection", "IP"))) <> 0 Then
    .SERVER_IP = ReadIniValue(App.Path & "\Config.ini", "Connection", "IP")
End If

'Read 'Port' from .ini file
If Len(Trim$(ReadIniValue(App.Path & "\Config.ini", "Connection", "Port"))) <> 0 Then
    .SERVER_PORT = ReadIniValue(App.Path & "\Config.ini", "Connection", "Port")
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

Public Sub SetScheme()
With frmMain
    .BackColor = Setting.SCHEME_COLOR
    .Picture1.BackColor = Setting.SCHEME_COLOR
End With

With frmAbout
    .BackColor = Setting.SCHEME_COLOR
    .Label1.BackColor = Setting.SCHEME_COLOR
End With

frmChat.BackColor = Setting.SCHEME_COLOR

With frmLanguage
    .BackColor = Setting.SCHEME_COLOR
    .Label1.BackColor = Setting.SCHEME_COLOR
End With

With frmConfig
    .BackColor = Setting.SCHEME_COLOR
    .Label1.BackColor = Setting.SCHEME_COLOR
    .lblAccount.BackColor = Setting.SCHEME_COLOR
    .lblAuthor.BackColor = Setting.SCHEME_COLOR
    .lblNickname.BackColor = Setting.SCHEME_COLOR
    .lblPassword.BackColor = Setting.SCHEME_COLOR
    .lblVersion.BackColor = Setting.SCHEME_COLOR
End With

With frmRegistration
    .BackColor = Setting.SCHEME_COLOR
    .Check1.BackColor = Setting.SCHEME_COLOR
    .Frame1.BackColor = Setting.SCHEME_COLOR
    .Label1.BackColor = Setting.SCHEME_COLOR
    .Label2.BackColor = Setting.SCHEME_COLOR
    .Label3.BackColor = Setting.SCHEME_COLOR
    .Label4.BackColor = Setting.SCHEME_COLOR
    .Label5.BackColor = Setting.SCHEME_COLOR
    .Picture1.BackColor = Setting.SCHEME_COLOR
End With

With frmSendFile
    .BackColor = Setting.SCHEME_COLOR
    .picProgress.BackColor = Setting.SCHEME_COLOR
    .Label1.BackColor = Setting.SCHEME_COLOR
    .Label4.BackColor = Setting.SCHEME_COLOR
    .lblFileToSend.BackColor = Setting.SCHEME_COLOR
    .lblProgress.BackColor = Setting.SCHEME_COLOR
    .lblSendSpeed.BackColor = Setting.SCHEME_COLOR
End With

frmSendFile2.BackColor = Setting.SCHEME_COLOR

With frmSettings
    .BackColor = Setting.SCHEME_COLOR
    .Frame1.BackColor = Setting.SCHEME_COLOR
    .Frame2.BackColor = Setting.SCHEME_COLOR
    .Frame3.BackColor = Setting.SCHEME_COLOR
    .Label1.BackColor = Setting.SCHEME_COLOR
    .Label2.BackColor = Setting.SCHEME_COLOR
    .Label3.BackColor = Setting.SCHEME_COLOR
    .SaveAccount.BackColor = Setting.SCHEME_COLOR
    .SavePassword.BackColor = Setting.SCHEME_COLOR
End With

With frmSociety
    .BackColor = Setting.SCHEME_COLOR
    .SSTab1.BackColor = Setting.SCHEME_COLOR
End With
End Sub
