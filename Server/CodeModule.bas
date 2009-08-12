Attribute VB_Name = "CodeModule"
Option Explicit

Public Const Rev = "1.0.6.9"
Public Const RegPort = 6222

Public GetUser      As String
Public GetConver    As String
Public Prefix       As String
Public Command      As String
Public Message      As String
Public ForWho       As String

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

Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal binvert As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long

Public nid As NOTIFYICONDATA ' trayicon variable

Public Sub SendMessage(Message As String)
Dim WinSk As Winsock

For Each WinSk In frmMain.Winsock1
    If WinSk.State = 7 Then
        WinSk.SendData Message
    End If
Next
End Sub

Public Sub SendSingle(Message As String, Wsk As Winsock)
If Wsk.State = 7 Then
    Wsk.SendData Message
End If
End Sub

Public Function UpdateUsersList() As Integer
Dim i As Integer
Dim tMsg As String
    With frmPanel.ListView1.ListItems
        tMsg = "!listupdate#"
        For i = 1 To .Count
            tMsg = tMsg & .Item(i) & "#"
        Next i
    End With
    If tMsg <> "!listupdate#" Then
        SendMessage tMsg
    End If
End Function

Public Sub VisualizeMessage(Command As String, Name As String, Message As String, Optional ForWho As String)
With frmChat
    Select Case Command
    Case "!msg"
        .txtConver.Text = .txtConver.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "]" & " [" & Command & "] [" & Name & "]: " & Message
    Case "!w"
        .txtConver.Text = .txtConver.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "]" & " [" & Command & "] [" & Name & " - " & ForWho & "]: " & Message
    Case "!namerequest"
        .txtConver.Text = .txtConver.Text & vbCrLf & " '" & Name & "' is requesting Name."
    Case "!connected"
        .txtConver.Text = .txtConver.Text & vbCrLf & " '" & Name & "' connected succesfully."
    Case "!login"
        .txtConver.Text = .txtConver.Text & vbCrLf & " Account '" & Name & "' - '" & Message & "' is logging in."
    Case Else
        .txtConver.Text = .txtConver.Text & vbCrLf & "[" & Command & "] [" & Name & "] [" & ForWho & "] [" & Message & "]"
    End Select
End With
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
