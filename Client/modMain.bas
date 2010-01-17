Attribute VB_Name = "modMain"
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32" ()

Sub Main()
'If the application is already open then close it
If App.PrevInstance And Not App.EXEName = "peachClient_DEBUG" Then
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

InsertIntoRegistry "Client\Revision", "Number", pRev
End Sub

Private Sub LoadRegistryValue()
With Setting
    If Len(ReadFromRegistry("Client\Configuration", "IP")) = 0 Then
        .SERVER_IP = "127.0.0.1"
    Else
        .SERVER_IP = ReadFromRegistry("Client\Configuration", "IP")
    End If
    
    If Len(ReadFromRegistry("Client\Configuration", "Port")) = 0 Then
        .SERVER_PORT = 4728
    Else
        .SERVER_PORT = ReadFromRegistry("Client\Configuration", "Port")
    End If
    
    .NICKNAME = ReadFromRegistry("Client\Configuration", "Nickname")
    
    If Len(ReadFromRegistry("Client\Configuration", "AccountTick")) = 0 Then
        .ACCOUNT_TICK = False
    Else
        .ACCOUNT_TICK = CBool(ReadFromRegistry("Client\Configuration", "AccountTick"))
    End If
    
    .ACCOUNT = ReadFromRegistry("Client\Configuration", "Account")
    
    If Len(ReadFromRegistry("Client\Configuration", "PasswordTick")) = 0 Then
        .PASSWORD_TICK = False
    Else
        .PASSWORD_TICK = CBool(ReadFromRegistry("Client\Configuration", "PasswordTick"))
    End If
    
    If .PASSWORD_TICK Then
        .PASSWORD = DeCode(DeCode(ReadFromRegistry("Client\Configuration", "Password")))
    End If
    
    If Len(ReadFromRegistry("Client\Configuration", "AskTick")) = 0 Then
        .ASK_TICK = False
    Else
        .ASK_TICK = CBool(ReadFromRegistry("Client\Configuration", "AskTick"))
    End If
    
    If Len(ReadFromRegistry("Client\Configuration", "MinimizeTray")) = 0 Then
        .MIN_TICK = False
    Else
        .MIN_TICK = CBool(ReadFromRegistry("Client\Configuration", "MinimizeTray"))
    End If
    
    If Len(ReadFromRegistry("Client\Configuration", "Validate")) = 0 Then
        .VALIDATE = 1
    Else
        .VALIDATE = ReadFromRegistry("Client\Configuration", "Validate")
    End If
    
    If Len(ReadFromRegistry("Client\Configuration", "Language")) = 0 Then
        .LANGUAGE = 1 'English as default
    Else
        .LANGUAGE = ReadFromRegistry("Client\Configuration", "Language")
    End If
    
    If Len(ReadFromRegistry("Client\Configuration", "Top")) = 0 Then
        .MAIN_TOP = 1200
    Else
        .MAIN_TOP = ReadFromRegistry("Client\Configuration", "Top")
    End If
    
    If Len(ReadFromRegistry("Client\Configuration", "Left")) = 0 Then
        .MAIN_LEFT = 1200
    Else
        .MAIN_LEFT = ReadFromRegistry("Client\Configuration", "Left")
    End If
    
    If Len(ReadFromRegistry("Client\Configuration", "SchemeColor")) = 0 Then
        .SCHEME_COLOR = 15724527
    Else
        .SCHEME_COLOR = ReadFromRegistry("Client\Configuration", "SchemeColor")
    End If
End With

With Fonts
    If Len(ReadFromRegistry("Client\Font", "FontBold")) = 0 Then
        .FONT_BOLD = False
    Else
        .FONT_BOLD = CBool(ReadFromRegistry("Client\Font", "FontBold"))
    End If
    
    If Len(ReadFromRegistry("Client\Font", "FontItalic")) = 0 Then
        .FONT_ITALIC = False
    Else
        .FONT_ITALIC = CBool(ReadFromRegistry("Client\Font", "FontItalic"))
    End If
    
    If Len(ReadFromRegistry("Client\Font", "FontName")) = 0 Then
        .FONT_NAME = "Segoe UI"
    Else
        .FONT_NAME = ReadFromRegistry("Client\Font", "FontName")
    End If
    
    If Len(ReadFromRegistry("Client\Font", "FontSize")) = 0 Then
        .FONT_SIZE = 9
    Else
        .FONT_SIZE = CLng(ReadFromRegistry("Client\Font", "FontSize"))
    End If
    
    If Len(ReadFromRegistry("Client\Font", "FontStrike")) = 0 Then
        .FONT_STRIKE = False
    Else
        .FONT_STRIKE = CBool(ReadFromRegistry("Client\Font", "FontStrike"))
    End If
    
    If Len(ReadFromRegistry("Client\Font", "FontUnder")) = 0 Then
        .FONT_UNDER = False
    Else
        .FONT_UNDER = CBool(ReadFromRegistry("Client\Font", "FontStrike"))
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
        .lblColor.BackColor = SC
        .lblFont.BackColor = SC
        .Label2.BackColor = SC
        .Label3.BackColor = SC
        .SaveAccount.BackColor = SC
        .SavePassword.BackColor = SC
        .CheckAsk.BackColor = SC
        .CheckMin.BackColor = SC
        .lblMinimizeTray.BackColor = SC
    End With
End If

With frmChat.txtConver.Font
    .Name = Fonts.FONT_NAME
    .Bold = Fonts.FONT_BOLD
    .Italic = Fonts.FONT_ITALIC
    .Size = Fonts.FONT_SIZE
    .Strikethrough = Fonts.FONT_STRIKE
    .Underline = Fonts.FONT_UNDER
End With

With frmSociety
    .BackColor = SC
    .SSTab1.BackColor = SC
End With
End Sub
