Attribute VB_Name = "modFunctions"
Option Explicit

Public Const aPort              As Long = 6123
Public Const bPort              As Long = 6124
Public Const rPort              As Long = 6222

Public ACC_SWITCH               As String
Public pCaption                 As String
Public Setting                  As CONFIG
Public Fonts                    As FNT
Public NID                      As NOTIFYICONDATA

Type CONFIG
    MAIN_TOP                    As Long
    MAIN_LEFT                   As Long
    VALIDATE                    As Long
    LANGUAGE                    As Long
    ACCOUNT_TICK                As Boolean
    PASSWORD_TICK               As Boolean
    AUTO_LOGIN                  As Boolean
    ASK_TICK                    As Boolean
    MIN_TICK                    As Boolean
    SCHEME_COLOR                As String
    SERVER_IP                   As String
    SERVER_PORT                 As String
    ACCOUNT                     As String
    PASSWORD                    As String
    NICKNAME                    As String
End Type

Type FNT
    Name                        As String
    Bold                        As Boolean
    Italic                      As Boolean
    Size                        As Long
    Strike                      As Boolean
    Under                       As Boolean
End Type

Public Type NOTIFYICONDATA
    cbSize                      As Long
    hwnd                        As Long
    uId                         As Long
    uFlags                      As Long
    uCallBackMessage            As Long
    hIcon                       As Long
    szTip                       As String * 64
End Type

Public Type POINTAPI
    X                           As Long
    Y                           As Long
End Type

Public Const NIM_ADD            As Long = &H0
Public Const NIM_MODIFY         As Long = &H1
Public Const NIM_DELETE         As Long = &H2
Public Const WM_MOUSEMOVE       As Long = &H200
Public Const NIF_MESSAGE        As Long = &H1
Public Const NIF_ICON           As Long = &H2
Public Const NIF_TIP            As Long = &H4
Public Const WM_LBUTTONDBLCLK   As Long = &H203 'Double-click
Public Const WM_LBUTTONDOWN     As Long = &H201 'Button down
Public Const WM_LBUTTONUP       As Long = &H202 'Button up
Public Const WM_RBUTTONDBLCLK   As Long = &H206 'Double-click
Public Const WM_RBUTTONDOWN     As Long = &H204 'Button down
Public Const WM_RBUTTONUP       As Long = &H205 'Button up

Public Declare Function SendMessage2 Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function GetActiveWindow Lib "user32" () As Long

Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal binvert As Long) As Long

Public Sub SendMessage(pMessage As String)
With frmMain.Winsock1
    If .State = 7 Then
        .SendData pMessage & Chr(24) & Chr(25)
        DoEvents
    End If
End With
End Sub

Public Sub FlashTitle(Handle As Long, ReturnOrig As Boolean)
Call FlashWindow(Handle, ReturnOrig)
End Sub

Public Sub MinimizeToTray(pForm As Form)
pForm.Hide
With NID
    .cbSize = Len(NID)
    .hwnd = pForm.hwnd
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallBackMessage = WM_MOUSEMOVE
    .hIcon = pForm.Icon ' the icon will be your Form1 project icon
    .szTip = "Peach" & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, NID
End Sub

Public Sub Disconnect()
Dim WiSk As Winsock

'Clear the online user list
With frmSociety
    .lvFriendList.ListItems.Clear
    .lvIgnoreList.ListItems.Clear
    .lvOnlineList.ListItems.Clear
End With

frmSendFile.Combo1.Clear

With frmSendFile2
    For Each WiSk In .SckReceiveFile
        If WiSk.State = 7 Then
            WiSk.Close
            Unload WiSk
        End If
    Next
    .SckReceiveFile(0).Close
End With

With frmMain
    For Each WiSk In .FSocket2
        If WiSk.State = 7 Then
            WiSk.Close
            Unload WiSk
        End If
    Next
    
    .FSocket2(0).Close
    .Winsock1.Close
    
    'Reset RunOnce variable
    .RunOnce = False
End With

SetupForm frmMain

SwitchButtons True
End Sub

'If an error occurs, this function returns False
Public Function FileExists(FileName As String) As Boolean
On Error GoTo ErrorHandler
FileExists = (GetAttr(FileName) And vbDirectory) = 0
ErrorHandler:
End Function

Public Function IsInvalid(pString As String) As Boolean
Dim SIGN_STRING     As String
Dim SIGN_ARRAY()    As String
Dim i               As Long

SIGN_STRING = " 1F.1F*1F#1F{1F}1F,1F(1F)1F&1F!1F@1F?1F/1F¬1F=1F<1F>1F[1F]1F'1F¿1Fº1Fª1F\1F|1F~1F´1F`1F+1F-1F^1F_1F·1F€"
SIGN_ARRAY = Split(SIGN_STRING, "1F")

For i = LBound(SIGN_ARRAY) To UBound(SIGN_ARRAY)
    If InStr(1, pString, SIGN_ARRAY(i)) <> 0 Then
        IsInvalid = True
        Exit Function
    End If
Next i
End Function

Public Sub CloseThis()
'Write data entries into registry
InsertIntoRegistry "Client\Configuration", "Password", Encode(Encode(frmMain.txtPassword.Text))
InsertIntoRegistry "Client\Configuration", "Account", frmMain.txtAccount.Text
InsertIntoRegistry "Client\Configuration", "Nickname", frmMain.txtName.Text

If frmMain.WindowState = vbMinimized Then 'vbMinimized = 1
    InsertIntoRegistry "Client\Configuration", "Top", Screen.Height / Screen.TwipsPerPixelY
    InsertIntoRegistry "Client\Configuration", "Left", Screen.Width / Screen.TwipsPerPixelX
Else
    InsertIntoRegistry "Client\Configuration", "Top", frmMain.Top
    InsertIntoRegistry "Client\Configuration", "Left", frmMain.Left
End If

Shell_NotifyIcon NIM_DELETE, NID  'Del tray icon
End
End Sub

Public Sub SwitchButtons(pSwitch As Boolean)
With frmMain
    .lblAccount.Enabled = pSwitch
    .lblCreateAccount.Enabled = pSwitch
    .lblForgotPassword.Enabled = pSwitch
    .lblName.Enabled = pSwitch
    .lblPassword.Enabled = pSwitch
    
    .txtAccount.Enabled = pSwitch
    .txtName.Enabled = pSwitch
    .txtPassword.Enabled = pSwitch
    
    .menuConfig.Enabled = pSwitch
    .menuUpdate.Enabled = pSwitch
    
    If pSwitch Then
        .cmdConnect.Caption = CONFIG_COMMAND_CONNECT
    Else
        .cmdConnect.Caption = CONFIG_COMMAND_DISCONNECT
    End If
End With
End Sub

Public Sub SetScheme()
Dim SC As String
    SC = Setting.SCHEME_COLOR

With frmMain
    .BackColor = SC
    .lblAccount.BackColor = SC
    .lblCreateAccount.BackColor = SC
    .lblForgotPassword.BackColor = SC
    .lblName.BackColor = SC
    .lblPassword.BackColor = SC
    .lblAuthor.BackColor = SC
    .lblVersion.BackColor = SC
End With

With frmSettings
    .BackColor = SC
    .Frame1.BackColor = SC
    .Frame2.BackColor = SC
    .Frame3.BackColor = SC
    .lblColor.BackColor = SC
    .lblFont.BackColor = SC
    .Label2.BackColor = SC
    .Label3.BackColor = SC
    .chkSaveAccount.BackColor = SC
    .chkSavePassword.BackColor = SC
    .chkAutoLogin.BackColor = SC
    .chkAskClosing.BackColor = SC
    .chkMinimize.BackColor = SC
    .lblMinimizeTray.BackColor = SC
End With

frmContainer.BackColor = SC
frmContainer.Picture1.BackColor = SC

With frmSendFile
    .BackColor = SC
    .picProgress.BackColor = SC
    .Label1.BackColor = SC
    .Label4.BackColor = SC
    .lblSendStatus.BackColor = SC
    .lblSendSpeed.BackColor = SC
End With

frmChat.BackColor = SC

With frmChat.txtConver
    .Font.Name = Fonts.Name
    .Font.Bold = Fonts.Bold
    .Font.Italic = Fonts.Italic
    .Font.Size = Fonts.Size
    .Font.Strikethrough = Fonts.Strike
    .Font.Underline = Fonts.Under
End With

With frmChat.txtToSend
    .Font.Name = Fonts.Name
    .Font.Bold = Fonts.Bold
    .Font.Italic = Fonts.Italic
    .Font.Size = Fonts.Size
    .Font.Strikethrough = Fonts.Strike
    .Font.Underline = Fonts.Under
End With

With frmSociety
    .BackColor = SC
    .SSTab1.BackColor = SC
End With
End Sub

Public Sub SetupChildForm(pForm As Form)
Dim pNewForm As Form

For Each pNewForm In Forms
    If pNewForm.Name <> frmContainer.Name Then pNewForm.Hide
Next

pForm.Show
End Sub

Public Sub SetupForm(pForm As Form)
Dim pNewForm As Form

For Each pNewForm In Forms
    pNewForm.Hide
Next

pForm.Show
End Sub
