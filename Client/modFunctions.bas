Attribute VB_Name = "modFunctions"
Option Explicit

Public Const aPort              As Long = 6123
Public Const bPort              As Long = 6124
Public Const rPort              As Long = 6222

Public ACC_SWITCH               As String
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

Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal binvert As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long

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

Public Sub MinimizeToTray()
frmMain.Hide
With NID
    .cbSize = Len(NID)
    .hwnd = frmMain.hwnd
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallBackMessage = WM_MOUSEMOVE
    .hIcon = frmMain.Icon ' the icon will be your Form1 project icon
    .szTip = "Peach -  " & frmConfig.txtNick & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, NID
End Sub

Public Sub SwitchButtons(pSwitch As Boolean, IsConnecting As Boolean)
Dim pBool As Boolean
'pSwitch - True     = Disconnected
'pSwitch - False    = Connected
With frmConfig
    .lblAccount.Enabled = pSwitch
    .txtAccount.Enabled = pSwitch
        
    .lblPassword.Enabled = pSwitch
    .txtPassword.Enabled = pSwitch
    
    .lblNickname.Enabled = pSwitch
    .txtNick.Enabled = pSwitch
    
    .Command3.Enabled = pSwitch
    .Command4.Enabled = pSwitch
    .Label1.Enabled = pSwitch
    .Label2.Enabled = pSwitch

    If pSwitch Then
        .cmdConnect.Caption = CONFIG_COMMAND_CONNECT
        frmMain.Connect.Caption = CONFIG_COMMAND_CONNECT
    Else
        .cmdConnect.Caption = CONFIG_COMMAND_DISCONNECT
        frmMain.Connect.Caption = CONFIG_COMMAND_DISCONNECT
    End If
End With

If Not pSwitch And Not IsConnecting Then
    pBool = True
Else
    pBool = False
End If

With frmChat
    .cmdSend.Enabled = pBool
    .cmdClear.Enabled = pBool
    .txtToSend.Enabled = pBool
    .txtConver.Enabled = pBool
End With

With frmSendFile
    .Label1.Enabled = pBool
    .Label4.Enabled = pBool
    .txtFileName.Enabled = pBool
    .Combo1.Enabled = pBool
    .cmdBrowse.Enabled = pBool
    .lblSendSpeed.Enabled = pBool
    .lblSendStatus.Enabled = pBool
    .picProgress.Enabled = pBool
End With

With frmSociety
    .cmdAddFriend.Enabled = pBool
    .cmdAddIgnore.Enabled = pBool
    .cmdRemoveFriend.Enabled = pBool
    .cmdRemoveIgnore.Enabled = pBool
    .cmdAddToFriend.Enabled = pBool
    .cmdAddToIgnore.Enabled = pBool
End With
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
    
    .StatusBar1.Panels(1).Text = MDI_STAT_DISCONNECTED
End With

frmConfig.cmdConnect.Caption = CONFIG_COMMAND_CONNECT
frmMain.Connect.Caption = CONFIG_COMMAND_CONNECT

SwitchButtons True, False
End Sub

'If an error occurs, this function returns False
Public Function FileExists(FileName As String) As Boolean
On Error GoTo ErrorHandler
FileExists = (GetAttr(FileName) And vbDirectory) = 0
ErrorHandler:
End Function

Public Function IsInvalid(pString As String) As Boolean
Dim CHAR As String
Dim SIGN_STRING As String
Dim SIGN_ARRAY() As String

SIGN_STRING = " 1F.1F*1F#1F{1F}1F,1F(1F)1F&1F!1F@1F?1F/1F¬1F=1F<1F>1F[1F]1F'1F¿1Fº1Fª1F\1F|1F~1F´1F`1F+1F-1F^1F_1F·1F€"
SIGN_ARRAY = Split(SIGN_STRING, "1F")

For i = LBound(SIGN_ARRAY) To UBound(SIGN_ARRAY)
    If InStr(1, pString, SIGN_ARRAY(i)) <> 0 Then
        IsInvalid = True
        Exit Function
    End If
Next i
End Function
