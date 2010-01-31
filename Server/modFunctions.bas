Attribute VB_Name = "modFunctions"
Option Explicit

Global Const rPort                          As Long = 6222
Global Const MAX_INT_VALUE                  As Long = 2147483647
Global Const MIN_INT_VALUE                  As Long = -2147483647

Global VarTime                              As Long     'Time counter variable

Global Const DATABASE_TABLE_ACCOUNTS        As String = "accounts"
Global Const DATABASE_TABLE_EMOTES          As String = "emotes"
Global Const DATABASE_TABLE_COMMANDS        As String = "commands"
Global Const DATABASE_TABLE_DECLINED_NAMES  As String = "declinednames"
Global Const DATABASE_TABLE_FRIENDS         As String = "friends"
Global Const DATABASE_TABLE_IGNORES         As String = "ignores"

Type NOTIFYICONDATA
    cbSize                                  As Long
    hwnd                                    As Long
    uId                                     As Long
    uFlags                                  As Long
    uCallBackMessage                        As Long
    hIcon                                   As Long
    szTip                                   As String * 64
End Type

Type MENUITEMINFO
    cbSize                      As Long
    fMask                       As Long
    fType                       As Long
    fState                      As Long
    wID                         As Long
    hSubMenu                    As Long
    hbmpChecked                 As Long
    hbmpUnchecked               As Long
    dwItemData                  As Long
    dwTypeData                  As String
    cch                         As Long
End Type

Type OSVERSIONINFO
    dwOSVersionInfoSize                     As Long
    dwMajorVersion                          As Long
    dwMinorVersion                          As Long
    dwBuildNumber                           As Long
    dwPlatformId                            As Long
    szCSDVersion                            As String * 128
End Type

Type DB
    Database                                As String
    User                                    As String
    Host                                    As String
    Password                                As String
End Type

Type EMT
    Command                                 As String
    SingleEmote                             As String
    TargetEmote                             As String
End Type

Type GC
    Syntax                                  As String
    Description                             As String
End Type

Public Database                             As DB
Public Commands()                           As GC
Public Emotes()                             As EMT
Public DeclinedNames()                      As String

'Tray constants
Public Const NIM_ADD                        As Long = &H0
Public Const NIM_MODIFY                     As Long = &H1
Public Const NIM_DELETE                     As Long = &H2
Public Const WM_MOUSEMOVE                   As Long = &H200
Public Const NIF_MESSAGE                    As Long = &H1
Public Const NIF_ICON                       As Long = &H2
Public Const NIF_TIP                        As Long = &H4
Public NID                                  As NOTIFYICONDATA   'Trayicon variable

'Mouseclick constants
Public Const WM_LBUTTONDBLCLK               As Long = &H203     'Double-click
Public Const WM_LBUTTONDOWN                 As Long = &H201     'Button down
Public Const WM_LBUTTONUP                   As Long = &H202     'Button up
Public Const WM_RBUTTONDBLCLK               As Long = &H206     'Double-click
Public Const WM_RBUTTONDOWN                 As Long = &H204     'Button down
Public Const WM_RBUTTONUP                   As Long = &H205     'Button up

'Windows version constants
Public Const VER_PLATFORM_WIN32s            As Long = 0         'Win32s on Windows 3.1
Public Const VER_PLATFORM_WIN32_WINDOWS     As Long = 1         'Windows 95, Windows 98, or Windows Me
Public Const VER_PLATFORM_WIN32_NT          As Long = 2         'Windows NT, Windows 2000, Windows XP, Windows Vista, Windows 7 or Windows Server 2003 family.

'Windows frame constants
Public Const GWL_STYLE                      As Long = (-16)
Public Const WS_THICKFRAME                  As Long = &H40000
Public Const WS_MAXIMIZEBOX                 As Long = &H10000
Public Const WS_MINIMIZEBOX                 As Long = &H20000

Private Const WS_SYSMENU                    As Long = &H80000
Private Const WS_CAPTION                    As Long = &HC00000
Private Const SC_MAXIMIZE                   As Long = &HF030&
Private Const SC_MINIMIZE                   As Long = &HF020&
Private Const SC_CLOSE                      As Long = &HF060&
Private Const MIIM_STATE                    As Long = &H1&
Private Const MIIM_ID                       As Long = &H2&
Private Const MFS_GRAYED                    As Long = &H3&
Private Const WM_NCACTIVATE                 As Long = &H86

'Time functions to determine exact ms time
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal binvert As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long

Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SendMessage2 Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Public Sub WriteLog(pData As String)
With frmConfig
    .txt_log.Text = .txt_log.Text & vbCrLf & "[" & Format$(Time, "hh:mm:ss") & "] " & pData
End With
End Sub

Public Function IsIgnoring(pUser As String, pTarget As String) As Boolean
Dim j As Long
With frmFriendIgnoreList.ListView2.ListItems
    For j = 1 To .Count
        If .Item(j).SubItems(1) = pUser Then
            If .Item(j).SubItems(2) = pTarget Then
                IsIgnoring = True
                Exit For
            End If
        End If
    Next j
End With
End Function

Public Sub SendProtectedMessage(pUser As String, pMessage As String)
Dim pAccount As String

'Get account name to interact better with frmFriendIgnoreList
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = pUser Then
            pAccount = .Item(i).SubItems(5)
            Exit For
        End If
    Next i
    
    For i = 1 To .Count
        If IsIgnoring(.Item(i).SubItems(5), pAccount) = False Then
            SendSingle pMessage, .Item(i).SubItems(2)
            DoEvents
        End If
    Next i
End With
End Sub

Public Sub SendMessage(pMessage As String)
Dim WinSk As Winsock
For Each WinSk In frmMain.Winsock1
    With WinSk
        If .State = 7 Then
            .SendData pMessage & Chr(24) & Chr(25)
            DoEvents
        End If
    End With
Next
End Sub

Public Sub SendSingle(pMessage As String, pIndex As Integer)
With frmMain.Winsock1(pIndex)
    If .State = 7 Then
        .SendData pMessage & Chr(24) & Chr(25)
        DoEvents
    End If
End With
End Sub

Public Sub UPDATE_ONLINE()
Dim GetList As String
With frmPanel.ListView1.ListItems
    GetList = "!update_online#"
    For i = 1 To .Count
        GetList = GetList & .Item(i) & " ( " & .Item(i).SubItems(5) & " )#"
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
Dim IsAvaible   As Boolean
Dim j           As Integer

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

Public Sub CMSG(pData As String)
With frmChat.txtConver
    .Text = .Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "] " & pData
    .SelStart = Len(.Text)
End With
End Sub

Public Function IsAlphaCharacter(pChar As String) As Boolean
IsAlphaCharacter = (Not (UCase(pChar) = LCase(pChar))) Or (pChar = " ")
End Function

Public Function GetRandomNumber(Optional ByVal MIN As Long = 1, Optional ByVal MAX As Long = 100) As Long
Randomize
Do
    GetRandomNumber = Int(Rnd * MAX) + MIN
Loop Until GetRandomNumber <= MAX
End Function

Public Sub MinimizeToTray()
frmMain.Hide
NID.cbSize = Len(NID)
NID.hwnd = frmMain.hwnd
NID.uId = vbNull
NID.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
NID.uCallBackMessage = WM_MOUSEMOVE
NID.hIcon = frmMain.Icon        'the icon will be your frmMain project icon
NID.szTip = "Peach" & vbNullChar
Shell_NotifyIcon NIM_ADD, NID
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
            GetOS = "UnknownOS"
    End Select
End With
End Function

Public Sub DisableFormResize(frm As Form)
Dim style           As Long
Dim hMenu           As Long
Dim MII             As MENUITEMINFO
Dim lngMenuID       As Long
Const xSC_MAXIMIZE  As Long = -11

style = GetWindowLong(frm.hwnd, GWL_STYLE)

style = style And Not WS_THICKFRAME
style = style And Not WS_MAXIMIZEBOX

style = SetWindowLong(frm.hwnd, GWL_STYLE, style)

On Error Resume Next

hMenu = GetSystemMenu(frm.hwnd, 0)

With MII
    .cbSize = Len(MII)
    .dwTypeData = String(80, 0)
    .cch = Len(.dwTypeData)
    .fMask = MIIM_STATE
    .wID = SC_MAXIMIZE
End With
If GetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then Exit Sub

With MII
    lngMenuID = .wID
    .wID = xSC_MAXIMIZE
    .fMask = MIIM_ID
End With
If SetMenuItemInfo(hMenu, lngMenuID, False, MII) = 0 Then Exit Sub

With MII
    .fState = (.fState Or MFS_GRAYED)
    .fMask = MIIM_STATE
End With
If SetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then Exit Sub

SendMessage2 frmMain.hwnd, WM_NCACTIVATE, True, 0

frm.Width = frm.Width - 1
frm.Width = frm.Width + 1
End Sub
