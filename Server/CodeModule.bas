Attribute VB_Name = "CodeModule"
Option Explicit

Public Const Rev    As String = "1.1.6.7"
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

Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Type DB
    Database       As String
    User           As String
    Host           As String
    Password       As String
    Friend_Table   As String
    Account_Table  As String
    Emote_Table    As String
End Type

Type EMT
    Command         As String
    IsUserText1     As String
    IsUserText2     As String
    IsNotUser       As String
End Type

Public Database As DB
Public Emotes() As EMT

'Tray Constants
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

'Mouseclick constants
Public Const WM_LBUTTONDBLCLK = &H203 'Double-click
Public Const WM_LBUTTONDOWN = &H201 'Button down
Public Const WM_LBUTTONUP = &H202 'Button up
Public Const WM_RBUTTONDBLCLK = &H206 'Double-click
Public Const WM_RBUTTONDOWN = &H204 'Button down
Public Const WM_RBUTTONUP = &H205 'Button up

'Windows version constant
Public Const VER_PLATFORM_WIN32s = 0        ' Win32s on Windows 3.1
Public Const VER_PLATFORM_WIN32_WINDOWS = 1 ' Windows 95, Windows 98, or Windows Me
Public Const VER_PLATFORM_WIN32_NT = 2      ' Windows NT, Windows 2000, Windows XP, or Windows Server 2003 family.

Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
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

Public Sub UPDATE_ONLINE()
Dim GetList As String
With frmPanel.ListView1.ListItems
    GetList = "!update_online#"
    For i = 1 To .Count
        If .Item(i).SubItems(7) = "True" Then
            GetList = GetList & "<AFK>" & .Item(i) & " - [ Account: " & .Item(i).SubItems(5) & " ]#"
        Else
            GetList = GetList & .Item(i) & " - [ Account: " & .Item(i).SubItems(5) & " ]#"
        End If
    Next i
End With
If GetList <> "!update_online#" Then SendMessage GetList
End Sub

Public Sub UPDATE_FRIEND(pName As String)
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
        Else
            If j = .Count Then GetAccountStatus = "!offline#"
        End If
    Next j
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

Public Function GetVersion() As String
'                           Major     Minor
' OS              Platform  Version   Version  Build
' Windows 95      1         4          0
' Windows 98      1         4         10       1998
' Windows 98SE    1         4         10       2222
' Windows Me      1         4         90       3000
' NT 3.51         2         3         51
' NT              2         4          0       1381
' 2000            2         5          0
' XP              2         5          1       2600
' Server 2003     2         5          2
Dim OSInfo As OSVERSIONINFO
Dim retvalue As Integer
  
OSInfo.dwOSVersionInfoSize = 148
OSInfo.szCSDVersion = Space$(128)
retvalue = GetVersionExA(OSInfo)
  
With OSInfo
    Select Case .dwPlatformId
    Case VER_PLATFORM_WIN32s           ' Win32s on Windows 3.1
        GetVersion = "Windows 3.1"

    Case VER_PLATFORM_WIN32_WINDOWS    ' Windows 95, Windows 98,
        Select Case .dwMinorVersion    ' or Windows Me
        Case 0
            GetVersion = "Windows 95"
        Case 10
            If (OSInfo.dwBuildNumber And &HFFFF&) = 2222 Then
                GetVersion = "Windows 98SE"
            Else
                GetVersion = "Windows 98"
            End If
        Case 90
            GetVersion = "Windows Me"
        End Select

    Case VER_PLATFORM_WIN32_NT         ' Windows NT, Windows 2000, Windows XP,
        Select Case .dwMajorVersion    ' or Windows Server 2003 family.
        Case 3
            GetVersion = "Windows NT 3.51"
        Case 4
            GetVersion = "Windows NT 4.0"
        Case 5
            Select Case .dwMinorVersion
            Case 0
                GetVersion = "Windows 2000"
            Case 1
                GetVersion = "Windows XP"
            Case 2
                GetVersion = "Windows Server 2003"
            End Select
        Case 6
            Select Case .dwMinorVersion
            Case 0
                GetVersion = "Windows Vista"
            Case 1
                GetVersion = "Windows 7"
            End Select
    End Select

Case Else
    GetVersion = "Failed"
End Select
End With
End Function
