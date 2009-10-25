Attribute VB_Name = "modFunctions"
Option Explicit

Public Const Rev      As String = "1.1.6.0"

Public Const aPort    As Long = 6123
Public Const bPort    As Long = 6124
Public Const rPort    As Long = 6222

Public Prefix   As String  'Time Prefix vairbale
Public i        As Long    'Global 'FOR' variable

Private Type CONFIG
    'frmMain values
    MAIN_TOP        As Long
    MAIN_LEFT       As Long
    
    'Language values
    VALIDATE        As Long
    LANGUAGE        As Long
        
    'Peach color scheme
    SCHEME_COLOR    As String
    
    'Ticks
    ACCOUNT_TICK    As Boolean
    PASSWORD_TICK   As Boolean
    
    'Server information
    SERVER_IP       As String
    SERVER_PORT     As String
    
    'frmCOnfig information
    ACCOUNT         As String
    PASSWORD        As String
    NICKNAME        As String
End Type

Public Setting As CONFIG

Public Type NOTIFYICONDATA
    cbSize              As Long
    hwnd                As Long
    uId                 As Long
    uFlags              As Long
    uCallBackMessage    As Long
    hIcon               As Long
    szTip               As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_LBUTTONDBLCLK = &H203 'Double-click
Public Const WM_LBUTTONDOWN = &H201 'Button down
Public Const WM_LBUTTONUP = &H202 'Button up
Public Const WM_RBUTTONDBLCLK = &H206 'Double-click
Public Const WM_RBUTTONDOWN = &H204 'Button down
Public Const WM_RBUTTONUP = &H205 'Button up

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Const GWL_STYLE = (-16)
Public Const WS_SYSMENU = &H80000
Public Const WS_MINIMIZEBOX = &H20000
    
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
    
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2

Private Const GW_HWNDNEXT = 2

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal binvert As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public nid As NOTIFYICONDATA ' trayicon variable

Public Sub SendMsg(iMessage As String)
If frmMain.Winsock1.State <> 7 Then Exit Sub
frmMain.Winsock1.SendData iMessage
DoEvents
End Sub

Public Sub FlashTitle(Handle As Long, ReturnOrig As Boolean)
Call FlashWindow(Handle, ReturnOrig)
End Sub

Public Sub minimize_to_tray()
frmMain.Hide
nid.cbSize = Len(nid)
nid.hwnd = frmMain.hwnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = frmMain.Icon ' the icon will be your Form1 project icon
nid.szTip = "Peach -  " & frmConfig.txtNick & vbNullChar
Shell_NotifyIcon NIM_ADD, nid
End Sub

Public Sub SetTrans(oForm As Form, Optional bytAlpha As Byte = 255, Optional lColor As Long = 0)
Dim lStyle As Long
lStyle = GetWindowLong(oForm.hwnd, GWL_EXSTYLE)
If Not (lStyle And WS_EX_LAYERED) = WS_EX_LAYERED Then _
    SetWindowLong oForm.hwnd, GWL_EXSTYLE, lStyle Or WS_EX_LAYERED
SetLayeredWindowAttributes oForm.hwnd, lColor, bytAlpha, LWA_COLORKEY Or LWA_ALPHA
End Sub

Public Function IsOverCtl(oForm As Form, ByVal X As Long, ByVal Y As Long) As Boolean
Dim ctl As Control, lhWnd As Long, r As RECT, pt As POINTAPI

pt.X = X: pt.Y = Y
ClientToScreen oForm.hwnd, pt

For Each ctl In oForm.Controls
    On Error GoTo ErrHandler
    lhWnd = ctl.hwnd
    On Error GoTo 0
    If lhWnd Then
        GetWindowRect ctl.hwnd, r
        IsOverCtl = (pt.X >= r.Left And pt.X <= r.Right And pt.Y >= r.Top And pt.Y <= r.Bottom)
        If IsOverCtl Then Exit Function
    End If
Next ctl
Exit Function
ErrHandler:
    lhWnd = 0
    Resume Next
End Function

Public Function GetNextWindow(ByVal lhWnd As Long) As Long
    GetNextWindow = GetWindow(lhWnd, GW_HWNDNEXT)
End Function

Public Sub SwitchButtons(pSwitch As Boolean)
'True = Disconnected
'False = Connected
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

    If pSwitch = True Then
        .cmdConnect.Caption = CONFIG_COMMAND_CONNECT
    Else
        .cmdConnect.Caption = CONFIG_COMMAND_DISCONNECT
    End If
    
End With

With frmChat
    .cmdSend.Enabled = Not pSwitch
    .cmdClear.Enabled = Not pSwitch
    .txtToSend.Enabled = Not pSwitch
    .txtConver.Enabled = Not pSwitch
End With

With frmSociety
    .Command1.Enabled = Not pSwitch
    .Command2.Enabled = Not pSwitch
End With
End Sub

Public Sub Disconnect()
Dim WiSk As Winsock

'Reset RunOnce variable
frmMain.RunOnce = False

SwitchButtons True
'Clear the online user list
With frmSociety
    .ListView1.ListItems.Clear
    .ListView2.ListItems.Clear
End With

frmSendFile.Combo1.Clear

'Close and unload this sockets also
With frmMain
    For Each WiSk In .FSocket2
        If WiSk.State = 7 Then
            WiSk.Close
            Unload WiSk
        End If
    Next
    .FSocket2(0).Close
End With

'Close and unload all connected winsocks
With frmSendFile2
    For Each WiSk In .SckReceiveFile
        If WiSk.State = 7 Then
            WiSk.Close
            Unload WiSk
        End If
    Next
    .SckReceiveFile(0).Close
End With
End Sub

Public Function FileExists(FileName As String) As Boolean
    On Error GoTo ErrorHandler
    FileExists = (GetAttr(FileName) And vbDirectory) = 0
ErrorHandler:
    ' if an error occurs, this function returns False
End Function

Public Function CheckString(pString As String) As Boolean
Dim CHAR As String
Dim SIGN_STRING As String
Dim SIGN_ARRAY() As String

SIGN_STRING = " 1F.1F*1F#1F{1F}1F,1F(1F)1F&1F!1F@1F?1F/1F¬1F=1F<1F>1F[1F]1F'1F¿1Fº1Fª1F\1F|1F~1F´1F`1F+1F-1F^1F_1F·"
SIGN_ARRAY = Split(SIGN_STRING, "1F")

For i = LBound(SIGN_ARRAY) To UBound(SIGN_ARRAY)
    If InStr(1, pString, SIGN_ARRAY(i)) <> 0 Then
        CheckString = True
        Exit Function
    End If
Next i
End Function
