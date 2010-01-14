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

'Start loading registry values
LoadRegistryValue

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
    'Set language default ( english )
    SET_LANG_ENGLISH
    frmLanguage.Show
End If

InsertIntoRegistry "Number", pRev
End Sub

Private Sub LoadRegistryValue()
With Setting
    If Len(ReadFromRegistry("IP")) = 0 Then
        .SERVER_IP = "127.0.0.1"
    Else
        .SERVER_IP = ReadFromRegistry("IP")
    End If
    
    If Len(ReadFromRegistry("Port")) = 0 Then
        .SERVER_PORT = 4728
    Else
        .SERVER_PORT = ReadFromRegistry("Port")
    End If
    
    .NICKNAME = ReadFromRegistry("Nickname")
    
    If Len(ReadFromRegistry("AccountTick")) = 0 Then
        .ACCOUNT_TICK = False
    Else
        .ACCOUNT_TICK = CBool(ReadFromRegistry("AccountTick"))
    End If
    
    .ACCOUNT = ReadFromRegistry("Account")
    
    If Len(ReadFromRegistry("PasswordTick")) = 0 Then
        .PASSWORD_TICK = False
    Else
        .PASSWORD_TICK = CBool(ReadFromRegistry("PasswordTick"))
    End If
    
    If .PASSWORD_TICK Then
        .PASSWORD = DeCode(DeCode(ReadFromRegistry("Password")))
    End If
    
    If Len(ReadFromRegistry("AskTick")) = 0 Then
        .ASK_TICK = False
    Else
        .ASK_TICK = CBool(ReadFromRegistry("AskTick"))
    End If
    
    If Len(ReadFromRegistry("MinimizeTray")) = 0 Then
        .MIN_TICK = False
    Else
        .MIN_TICK = CBool(ReadFromRegistry("MinimizeTray"))
    End If
    
    If Len(ReadFromRegistry("Validate")) = 0 Then
        .VALIDATE = 1
    Else
        .VALIDATE = ReadFromRegistry("Validate")
    End If
    
    If Len(ReadFromRegistry("Language")) = 0 Then
        .LANGUAGE = 1 'English as default
    Else
        .LANGUAGE = ReadFromRegistry("Language")
    End If
    
    If Len(ReadFromRegistry("Top")) = 0 Then
        .MAIN_TOP = 1200
    Else
        .MAIN_TOP = ReadFromRegistry("Top")
    End If
    
    If Len(ReadFromRegistry("Left")) = 0 Then
        .MAIN_LEFT = 1200
    Else
        .MAIN_LEFT = ReadFromRegistry("Left")
    End If
    
    If Len(ReadFromRegistry("SchemeColor")) = 0 Then
        .SCHEME_COLOR = 15724527
    Else
        .SCHEME_COLOR = ReadFromRegistry("SchemeColor")
    End If
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
    .lblSendStatus.BackColor = SC
    .lblSendSpeed.BackColor = SC
End With

If Not IsFirst Then
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
        .CheckAsk.BackColor = SC
        .CheckMin.BackColor = SC
        .lblMinimizeTray.BackColor = SC
    End With
End If

With frmSociety
    .BackColor = SC
    .SSTab1.BackColor = SC
End With
End Sub
