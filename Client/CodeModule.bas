Attribute VB_Name = "CodeModule"
Option Explicit

Public Const Rev = "1.0.2.7"

' Create an Icon in System Tray Needs
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

'**** Glass Form Stuff ****'
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

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'**** Glass Form Stuff End ****'

Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal binvert As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long

Public nid As NOTIFYICONDATA ' trayicon variable

Public Sub SendMessage(iMessage As String)
If frmMain.Winsock1.State <> 7 Then Exit Sub
frmMain.Winsock1.SendData iMessage
End Sub

Public Sub VisualizeMessage(Whisper As Boolean, Name As String, Conver As String)
If Whisper = True Then
    frmChat.txtConver.Text = frmChat.txtConver.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "] [You whisper to " & Name & "]: " & Conver
Else
    frmChat.txtConver.Text = frmChat.txtConver.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "] [" & Name & "]: " & Conver
End If
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
