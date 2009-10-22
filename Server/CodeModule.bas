Attribute VB_Name = "CodeModule"
Option Explicit

Public Const Rev    As String = "1.1.5.5"
Public Const rPort  As Long = 6222

Public i        As Long    'Global "FOR" variable
Public HasError As Boolean 'Used in frmMain and frmConfig

Type NOTIFYICONDATA
    cbSize              As Long
    hwnd                As Long
    uId                 As Long
    uFlags              As Long
    uCallBackMessage    As Long
    hIcon               As Long
    szTip               As String * 64
End Type

Type DB
    Database       As String
    User           As String
    Host           As String
    Password       As String
    Friend_Table   As String
    Account_Table  As String
End Type

Public Database As DB

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

Public nid As NOTIFYICONDATA 'trayicon variable

Public Sub WriteLog(Data As String)
With frmConfig
    .txt_log.Text = .txt_log.Text & vbCrLf & "[" & Format$(Time, "hh:mm:ss") & "] " & Data
End With
End Sub

Public Sub SendMessage(Message As String)
Dim WinSk As Winsock
For Each WinSk In frmMain.Winsock1
    If WinSk.State = 7 Then
        WinSk.SendData Message
        DoEvents
    End If
Next
End Sub

Public Sub SendSingle(Message As String, Wsk As Winsock)
If Wsk.State = 7 Then
    Wsk.SendData Message
    DoEvents
End If
End Sub

Public Sub UpdateUsersList()
Dim GetList As String
With frmPanel.ListView1.ListItems
    GetList = "!update_online#"
    For i = 1 To .Count
        If .Item(i).SubItems(7) = "True" Then
            GetList = GetList & "<AFK>" & .Item(i) & "#"
        Else
            GetList = GetList & .Item(i) & "#"
        End If
    Next i
End With
If GetList <> "!listupdate#" Then SendMessage GetList
End Sub

Public Sub UpdateFriendList(pName As String)
Dim buffer As String
Dim a_array() As String

buffer = "!update_friends#"
With frmFriendList.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(1) = pName Then
            a_array = Split(GetAccountStatus(.Item(i).SubItems(2)), "#")
            
            Select Case a_array(0)
            Case "!online"
                buffer = buffer & .Item(i).SubItems(2) & " - " & a_array(1) & "$Online#"
            Case "!offline"
                buffer = buffer & .Item(i).SubItems(2) & "$Offline#"
            End Select
        End If
    Next i
End With

SendSingle buffer, frmMain.Winsock1(GetWinsockID(pName))
End Sub

Private Function GetAccountStatus(pAccount As String) As String
Dim IsAvaible As Boolean
Dim j As Integer
With frmPanel.ListView1.ListItems
    For j = 1 To .Count
        If .Item(j).SubItems(5) = pAccount Then
            GetAccountStatus = "!online#" & .Item(j) & "#"
            IsAvaible = True
            Exit For
        End If
    Next j
    If IsAvaible = False Then
        GetAccountStatus = "!offline#"
    End If
End With
End Function

Private Function GetWinsockID(pAccount As String) As Integer
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(5) = pAccount Then
            GetWinsockID = .Item(i).SubItems(2)
            Exit For
        End If
    Next i
End With
End Function

Public Sub CMSG(Command As String, Optional Name As String, Optional Message As String, Optional ForWho As String)
Dim TimePrefix As String
TimePrefix = "[" & Format(Time, "hh:nn:ss") & "] "
With frmChat.txtConver
    .SelStart = Len(.Text)
    Select Case Command
    Case "!msg"
        .SelRTF = vbCrLf & TimePrefix & "[" & Name & "]: " & Message
    Case "!w"
        .SelRTF = vbCrLf & TimePrefix & "[" & Name & " - " & ForWho & "]: " & Message
    Case "!namerequest"
        .SelRTF = vbCrLf & TimePrefix & "'" & Name & "' is requesting Name."
    Case "!connected"
        .SelRTF = vbCrLf & TimePrefix & "'" & Name & "' connected succesfully."
    Case "!login"
        .SelRTF = vbCrLf & TimePrefix & "Account: '" & Name & "' Password: '" & Message & "' is logging in."
    Case "!nameisfree"
        .SelRTF = vbCrLf & TimePrefix & "Send answer that '" & Name & "' is free to take."
    Case "!nametaken"
        .SelRTF = vbCrLf & TimePrefix & "User '" & Name & "'. tryed to login but failed. (Name already taken)."
    Case "!account"
        .SelRTF = vbCrLf & TimePrefix & "Account '" & Name & "' tryed to login but failed. (Account doesnt exist)."
    Case "!password"
        .SelRTF = vbCrLf & TimePrefix & "Account '" & Name & "' tryed to login but failed. (Wrong Password)."
    Case "!badname"
        .SelRTF = vbCrLf & TimePrefix & "User '" & Name & "' tryed to login but failed . (Badname)."
    Case "!muted"
        .SelRTF = vbCrLf & TimePrefix & "[Muted][" & Name & "]: " & Message
    Case "!repeat"
        .SelRTF = vbCrLf & TimePrefix & "[" & Name & "] activated flood control."
    Case "!long"
        .SelRTF = vbCrLf & TimePrefix & "[" & Name & "] wrote a too long message. (Kicked)"
    Case "!banned"
        .SelRTF = vbCrLf & TimePrefix & "'" & Name & "' tryed to login but failed. (Account banned)."
    Case "!disconnected"
        .SelRTF = vbCrLf & TimePrefix & "[System]: Disconnected due connection problem."
    Case Else
        .SelRTF = vbCrLf & TimePrefix & "[" & Command & "] [" & Name & "] [" & ForWho & "] [" & Message & "]"
    End Select
End With
End Sub

Public Function IsAlphaCharacter(sChar As String) As Boolean
    IsAlphaCharacter = (Not (UCase(sChar) = LCase(sChar))) Or (sChar = " ")
End Function

Public Sub minimize_to_tray()
frmMain.Hide
nid.cbSize = Len(nid)
nid.hwnd = frmMain.hwnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = frmMain.Icon ' the icon will be your Form1 project icon
nid.szTip = "Peach" & vbNullChar
Shell_NotifyIcon NIM_ADD, nid
End Sub
