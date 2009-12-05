Attribute VB_Name = "CodeModule"
Option Explicit

Public Const Rev            As String = "1.2.0.6"
Public Const rPort          As Long = 6222

Public VarTime              As Long    'Time counter variable
Public i                    As Long    'Global "FOR" variable
Public HasError             As Boolean 'Used in frmMain and frmConfig

Type NOTIFYICONDATA
    cbSize                  As Long
    hwnd                    As Long
    uId                     As Long
    uFlags                  As Long
    uCallBackMessage        As Long
    hIcon                   As Long
    szTip                   As String * 64
End Type

Type OSVERSIONINFO
    dwOSVersionInfoSize     As Long
    dwMajorVersion          As Long
    dwMinorVersion          As Long
    dwBuildNumber           As Long
    dwPlatformId            As Long
    szCSDVersion            As String * 128
End Type

Type DB
    Database                As String
    User                    As String
    Host                    As String
    Password                As String
    FriendTable             As String
    IgnoreTable             As String
    AccountTable            As String
    EmoteTable              As String
    DeclinedNameTable       As String
End Type

Type OPT
    ChatLevel               As String
End Type

Type EMT
    Command                 As String
    IsUserText1             As String
    IsUserText2             As String
    IsNotUser               As String
    Description             As String
End Type

Public Options              As OPT
Public Database             As DB
Public Emotes()             As EMT
Public DeclinedNames()      As String

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

Public Sub SendProtectedMessage(pMessage As String)
'/////////////////////////////
'

'/////////////////////////////
'

'/////////////////////////////
'

End Sub

Public Sub SendMessage(pMessage As String)
Dim WinSk As Winsock
For Each WinSk In frmMain.Winsock1
    If WinSk.State = 7 Then
        WinSk.SendData pMessage
        DoEvents
    End If
Next
End Sub

Public Sub SendSingle(pMessage As String, pIndex As Integer)
If frmMain.Winsock1(pIndex).State = 7 Then
    frmMain.Winsock1(pIndex).SendData pMessage
    DoEvents
End If
End Sub

Public Sub UPDATE_ONLINE()
Dim GetList As String
With frmPanel.ListView1.ListItems
    GetList = "!update_online#"
    For i = 1 To .Count
        If .Item(i).SubItems(7) = "True" Then
            GetList = GetList & "<AFK>" & .Item(i) & " ( " & .Item(i).SubItems(5) & " )#"
        Else
            GetList = GetList & .Item(i) & " ( " & .Item(i).SubItems(5) & " )#"
        End If
    Next i
End With
If GetList <> "!update_online#" Then SendMessage GetList
End Sub

Public Sub UPDATE_FRIEND(pName As String, pIndex As Integer)
Dim buffer      As String
Dim p_Array()   As String

buffer = "!update_friends#"
With frmFriendIgnoreList.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(1) = pName Then
            p_Array = Split(GetAccountStatus(.Item(i).SubItems(2)), "#")
            
            Select Case p_Array(0)
            Case "!online"
                buffer = buffer & .Item(i).SubItems(2) & " - " & p_Array(1) & "$Online#"
            Case "!offline"
                buffer = buffer & .Item(i).SubItems(2) & "$Offline#"
            End Select
        End If
    Next i
End With

SendSingle buffer, pIndex
End Sub

Public Sub UPDATE_IGNORE(pName As String, pIndex As Integer)
Dim buffer As String

buffer = "!update_ignore#"
With frmFriendIgnoreList.ListView2.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(1) = pName Then
            buffer = buffer & .Item(i).SubItems(2) & "#"
        End If
    Next i
End With

SendSingle buffer, pIndex
End Sub

Public Sub UPDATE_STATUS_BAR()
frmMain.StatusBar1.Panels(1).Text = "Status: Connected with " & frmMain.Winsock1.Count - 1 & " Client(s)."
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

Public Sub CMSG(Command As String, Optional Name As String, Optional Message As String, Optional ForWho As String)
Dim TimePrefix As String
TimePrefix = "[" & Format(Time, "hh:nn:ss") & "] "
With frmChat.txtConver
    .SelStart = Len(.Text)
    .SelRTF = vbCrLf & TimePrefix
    Select Case Command
    Case "!message"
        .SelRTF = "[" & Name & "]: " & Message
    Case "!w"
        .SelRTF = "[" & Name & " - " & ForWho & "]: " & Message
    Case "!namerequest"
        .SelRTF = "'" & Name & "' is requesting Name."
    Case "!connected"
        .SelRTF = "'" & Name & "' connected succesfully."
    Case "!login"
        .SelRTF = "Account: '" & Name & "' Password: '" & Message & "' is logging in."
    Case "!nameisfree"
        .SelRTF = "Send answer that '" & Name & "' is free to take."
    Case "!nametaken"
        .SelRTF = "User '" & Name & "'. tryed to login but failed. (Name already taken)."
    Case "!account"
        .SelRTF = "Account '" & Name & "' tryed to login but failed. (Account doesnt exist)."
    Case "!password"
        .SelRTF = "Account '" & Name & "' tryed to login but failed. (Wrong Password)."
    Case "!badname"
        .SelRTF = "User '" & Name & "' tryed to login but failed . (Badname)."
    Case "!muted"
        .SelRTF = "[Muted][" & Name & "]: " & Message
    Case "!repeat"
        .SelRTF = "[" & Name & "] activated flood control."
    Case "!long"
        .SelRTF = "[" & Name & "] wrote a too long message. (Kicked)"
    Case "!banned"
        .SelRTF = "'" & Name & "' tryed to login but failed. (Account banned)."
    Case "!disconnected"
        .SelRTF = "[System]: Disconnected due connection problem."
    Case Else
        .SelRTF = "[" & Command & "] [" & Name & "] [" & ForWho & "] [" & Message & "]"
    End Select
End With
End Sub

Public Function IsAlphaCharacter(pChar As String) As Boolean
    IsAlphaCharacter = (Not (UCase(pChar) = LCase(pChar))) Or (pChar = " ")
End Function

Public Function GetRandomNumber(Optional ByVal MIN As Long = 1, Optional ByVal MAX As Long = 100) As Long
Randomize
Do
    GetRandomNumber = Int(Rnd * MAX) + MIN
Loop Until GetRandomNumber <= MAX 'And GetRandomNumber > MIN
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

Public Function GetOS() As String
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
        GetOS = "WinDOS_3.1"

    Case VER_PLATFORM_WIN32_WINDOWS    ' Windows 95, Windows 98,
        Select Case .dwMinorVersion    ' or Windows Me
        Case 0
            GetOS = "Win95"
        Case 10
            If (OSInfo.dwBuildNumber And &HFFFF&) = 2222 Then
                GetOS = "Win98SE"
            Else
                GetOS = "Win98"
            End If
        Case 90
            GetOS = "Windows Me"
        End Select

    Case VER_PLATFORM_WIN32_NT         ' Windows NT, Windows 2000, Windows XP,
        Select Case .dwMajorVersion    ' or Windows Server 2003 family.
        Case 3
            GetOS = "WinNT3.51"
        Case 4
            GetOS = "WinNT4"
        Case 5
            Select Case .dwMinorVersion
            Case 0
                GetOS = "Win2000"
            Case 1
                GetOS = "WinXP"
            Case 2
                GetOS = "Windows Server 2003"
            End Select
        Case 6
            Select Case .dwMinorVersion
            Case 0
                GetOS = "WinVista"
            Case 1
                GetOS = "Win7"
            End Select
    End Select

Case Else
    GetOS = "Failed"
End Select
End With
End Function
