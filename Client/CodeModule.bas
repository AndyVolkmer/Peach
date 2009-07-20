Attribute VB_Name = "CodeModule"
Option Explicit

' Create an Icon in System Tray Needs
Public Type NOTIFYICONDATA
cbSize As Long
hwnd As Long
uId As Long
uFlags As Long
uCallBackMessage As Long
hIcon As Long
szTip As String * 64
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

Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal binvert As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long

Public nid As NOTIFYICONDATA ' trayicon variable

Public Sub SendMessage(iMessage As String)

If frmMain.Winsock1(0).State = 7 Then
    frmMain.Winsock1(0).SendData iMessage
Else
    MsgBox "Not connected!", vbInformation
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
