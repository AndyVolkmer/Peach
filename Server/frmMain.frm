VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00F4F4F4&
   Caption         =   " Peach Server"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7395
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   ScrollBars      =   0   'False
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   5
      Top             =   4740
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   7938
            MinWidth        =   7938
         EndProperty
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   7395
      TabIndex        =   0
      Top             =   0
      Width           =   7395
      Begin VB.CommandButton Command5 
         Caption         =   "&Friend List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   6
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&User Panel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Acco&unt Panel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ch&at"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Configuration"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GWL_STYLE         As Long = (-16)
Private Const WS_SYSMENU        As Long = &H80000
Private Const WS_CAPTION        As Long = &HC00000
Private Const WS_THICKFRAME     As Long = &H40000
Private Const WS_MAXIMIZEBOX    As Long = &H10000
Private Const WS_MINIMIZEBOX    As Long = &H20000
Private Const SC_MAXIMIZE       As Long = &HF030&
Private Const SC_MINIMIZE       As Long = &HF020&
Private Const SC_CLOSE          As Long = &HF060&
Private Const MIIM_STATE        As Long = &H1&
Private Const MIIM_ID           As Long = &H2&
Private Const MFS_GRAYED        As Long = &H3&
Private Const WM_NCACTIVATE     As Long = &H86

Private Type MENUITEMINFO
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

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage2 Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32" ()

Public xConnection              As New ADODB.Connection
Public xCommand                 As New ADODB.Command
Public xRecordSet               As New ADODB.Recordset

Private Vali                    As Boolean

Private Sub Command1_Click()
SetupForms frmConfig
End Sub

Private Sub Command2_Click()
SetupForms frmChat
End Sub

Private Sub Command3_Click()
SetupForms frmAccountPanel
End Sub

Private Sub Command4_Click()
SetupForms frmPanel
End Sub

Private Sub Command5_Click()
SetupForms frmFriendIgnoreList
End Sub

Private Sub MDIForm_Initialize()
Call InitCommonControls
End Sub

Private Sub MDIForm_Load()
DisableFormResize Me
Dim L As Long
    L = GetWindowLong(Me.hwnd, GWL_STYLE)
'   L = L And Not (WS_MINIMIZEBOX)
    L = L And Not (WS_MAXIMIZEBOX)
    L = SetWindowLong(Me.hwnd, GWL_STYLE, L)

'Open Database variable
With Database

'== Database ==
If LenB(ReadFromRegistry("Server\Database", "Name")) <> 0 Then
    .Database = ReadFromRegistry("Server\Database", "Name")
Else
    WriteLog "No Database value found."
    .Database = InputBox("The configuration file does not cotain a valid database, please insert one in the textbox below.", "Database error ..", "Database Name")
End If

If LenB(ReadFromRegistry("Server\Database", "User")) <> 0 Then
    .User = ReadFromRegistry("Server\Database", "User")
Else
    WriteLog "No User value found."
    .User = InputBox("The configuration file does not cotain a valid user name, please insert one in the textbox below.", "Database error ..", "User Name")
End If

If LenB(DeCode(ReadFromRegistry("Server\Database", "Password"))) <> 0 Then
    .Password = DeCode(ReadFromRegistry("Server\Database", "Password"))
Else
    WriteLog "No Password value found."
    .Password = InputBox("The configuration file does not cotain a valid password, please insert one in the textbox below.", "Database error ..", "Password")
End If

If LenB(ReadFromRegistry("Server\Database", "Host")) <> 0 Then
    .Host = ReadFromRegistry("Server\Database", "Host")
Else
    WriteLog "No Host value found."
    .Host = InputBox("The configuration file does not cotain a host adress, please insert one in the textbox below.", "Database error ..", "Host Adress")
End If

If LenB(ReadFromRegistry("Server\Database", "Account_Table")) <> 0 Then
    .AccountTable = ReadFromRegistry("Server\Database", "Account_Table")
Else
    WriteLog "No Account-Table found."
    .AccountTable = InputBox("The configuration file does not contain a account table, please insert one in the textbox below.", "Database error ..", "Account Table")
End If

If LenB(ReadFromRegistry("Server\Database", "Friend_Table")) <> 0 Then
    .FriendTable = ReadFromRegistry("Server\Database", "Friend_Table")
Else
    WriteLog "No Friends-Table found."
    .FriendTable = InputBox("The configuration file does not contain a friend table, please insert one in the textbox below.", "Database error ..", "Friend Table")
End If

If LenB(ReadFromRegistry("Server\Database", "Ignore_Table")) <> 0 Then
    .IgnoreTable = ReadFromRegistry("Server\Database", "Ignore_Table")
Else
    WriteLog "No Ignore-Table found."
    .IgnoreTable = InputBox("The configuration file does not contain a ignore table, please insert one in the textbox below.", "Database error ..", "Ignore Table")
End If

If LenB(ReadFromRegistry("Server\Database", "Emote_Table")) <> 0 Then
    .EmoteTable = ReadFromRegistry("Server\Database", "Emote_Table")
Else
    WriteLog "No Emotes-Table found."
    .EmoteTable = InputBox("The configuration file does not contain a emote table, please insert one in the textbox below.", "Database error ..", "Emote Table")
End If

If LenB(ReadFromRegistry("Server\Database", "Declined_Name_Table")) <> 0 Then
    .DeclinedNameTable = ReadFromRegistry("Server\Database", "Declined_Name_Table")
Else
    WriteLog "No Declined-Names-Table found."
    .DeclinedNameTable = InputBox("The configuration file does not contain a declined name table, please insert one in the textbox below.", "Database error ..", "Declined Name Table")
End If

If LenB(ReadFromRegistry("Server\Database", "Command_Table")) <> 0 Then
    .CommandsTable = ReadFromRegistry("Server\Database", "Command_Table")
Else
    WriteLog "No Command-Table found."
    .CommandsTable = InputBox("The configuration file does not contain a command table, please insert one in the textbox below.", "Database error ..", "Commands Table")
End If

'== Position ==
If LenB(ReadFromRegistry("Server\Configuration", "Top")) <> 0 Then
    Me.Top = ReadFromRegistry("Server\Configuration", "Top")
Else
    Me.Top = 1200
End If

If LenB(ReadFromRegistry("Server\Configuration", "Left")) <> 0 Then
    Me.Left = ReadFromRegistry("Server\Configuration", "Left")
Else
    Me.Left = 1200
End If

'== Options ==
If LenB(ReadFromRegistry("Server\Configuration", "ChatLevel")) <> 0 Then
    Options.ChatLevel = ReadFromRegistry("Server\Configuration", "ChatLevel")
Else
    Options.ChatLevel = 0
End If

'Connect to MySQL Database
ConnectMySQL .Database, .User, .Password, .Host

'Load Accounts
LoadAccounts .AccountTable

'Load Emotes
LoadEmotes .EmoteTable

'Load Friends
LoadFriends .FriendTable

'Load Ignores
LoadIgnores .IgnoreTable

'Load Declined Names
LoadDeclinedNames .DeclinedNameTable

'Load commands
LoadCommands .CommandsTable

'Close Database variable
End With

frmMain.StatusBar1.Panels(1) = "Status: Disconnected"
SetupForms frmConfig

If HasError = False Then WriteLog "Correctly loaded."
End Sub

Public Sub ConnectMySQL(pDatabase As String, pUser As String, pPassword As String, pIP As String)
On Error GoTo HandleErrorConnection

'Connect with Database (MySQL)
Set xConnection = New ADODB.Connection
xConnection.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
    & "SERVER=" & pIP & ";" _
    & "DATABASE=" & pDatabase & ";" _
    & "UID=" & pUser & ";" _
    & "PWD=" & pPassword & ";" _
    & "OPTION=18475" '-> 1 + 2 + 8 + 32 + 2048 + 16384

Screen.MousePointer = vbHourglass
xConnection.Open

Screen.MousePointer = vbDefault
Set xCommand = New ADODB.Command
Set xCommand.ActiveConnection = xConnection
xCommand.CommandType = adCmdText
WriteLog "Connected with Database"

Exit Sub
HandleErrorConnection:
'Print error
WriteLog Err.Description

'Change mouse pointer
Screen.MousePointer = vbDefault

'Set error flag
HasError = True
End Sub

Private Sub LoadCommands(pTable As String)
Dim SQL     As String
Dim Counter As Long

If HasError Then Exit Sub

SQL = "SELECT * FROM " & pTable
Counter = 0

On Error GoTo HandleErrorEmotes
xCommand.CommandText = SQL
Set xRecordSet = xCommand.Execute

ReDim Commands(0)

With xRecordSet
    Do Until .EOF
        ReDim Preserve Commands(Counter + 1)
        Commands(Counter).Syntax = !Syntax
        Commands(Counter).Description = !Description
        Counter = Counter + 1
        .MoveNext
    Loop
End With

Set xRecordSet = Nothing

WriteLog "Loaded " & Counter & " command(s)."

Exit Sub
HandleErrorEmotes:
'Print error
WriteLog Err.Description & "."

'Set error flag
HasError = True
End Sub

Private Sub LoadDeclinedNames(pTable As String)
Dim SQL     As String
Dim Counter As Long

If HasError Then Exit Sub

SQL = "SELECT * FROM " & pTable
Counter = 0

On Error GoTo HandleErrorEmotes
xCommand.CommandText = SQL
Set xRecordSet = xCommand.Execute

ReDim DeclinedNames(0)

With xRecordSet
    Do Until .EOF
        ReDim Preserve DeclinedNames(Counter + 1)
        DeclinedNames(Counter) = !Name
        Counter = Counter + 1
        .MoveNext
    Loop
End With

Set xRecordSet = Nothing

WriteLog "Loaded " & Counter & " declined name(s)."

Exit Sub
HandleErrorEmotes:
'Print error
WriteLog Err.Description & "."

'Set error flag
HasError = True
End Sub

Private Sub LoadEmotes(pTable As String)
Dim SQL     As String
Dim Counter As Long

If HasError Then Exit Sub

SQL = "SELECT * FROM " & pTable
Counter = 0

On Error GoTo HandleErrorEmotes
xCommand.CommandText = SQL
Set xRecordSet = xCommand.Execute

ReDim Emotes(0)

With xRecordSet
    Do Until .EOF
        ReDim Preserve Emotes(Counter + 1)
        Emotes(Counter).Command = !Command
        Emotes(Counter).SingleEmote = !single_emote
        Emotes(Counter).TargetEmote = !target_emote
        .MoveNext
        Counter = Counter + 1
    Loop
End With

Set xRecordSet = Nothing

WriteLog "Loaded " & Counter & " emote(s)."

Exit Sub
HandleErrorEmotes:
'Print error
WriteLog Err.Description

'Set error flag
HasError = True
End Sub

Private Sub LoadFriends(pTable As String)
Dim SQL     As String
Dim LItem   As ListItem
Dim Counter As Long

If HasError Then Exit Sub

SQL = "SELECT * FROM " & pTable
Counter = 0

On Error GoTo HandleErrorFriends
xCommand.CommandText = SQL
Set xRecordSet = xCommand.Execute

With xRecordSet
    Do Until .EOF
        Set LItem = frmFriendIgnoreList.ListView1.ListItems.Add(, , !ID)
        LItem.SubItems(1) = !Name
        LItem.SubItems(2) = !Friend
        .MoveNext
        Counter = Counter + 1
    Loop
End With

Set LItem = Nothing
Set xRecordSet = Nothing

WriteLog "Loaded " & Counter & " relation(s)."

Exit Sub
HandleErrorFriends:
'Print error
WriteLog Err.Description

'Set error flag
HasError = True
End Sub

Private Sub LoadIgnores(pTable As String)
Dim SQL     As String
Dim LItem   As ListItem
Dim Counter As Long

If HasError Then Exit Sub

SQL = "SELECT * FROM " & pTable
Counter = 0

On Error GoTo HandleErrorFriends
xCommand.CommandText = SQL
Set xRecordSet = xCommand.Execute

With xRecordSet
    Do Until .EOF
        Set LItem = frmFriendIgnoreList.ListView2.ListItems.Add(, , !ID)
        LItem.SubItems(1) = !Name
        LItem.SubItems(2) = !IgnoredName
        .MoveNext
        Counter = Counter + 1
    Loop
End With

Set LItem = Nothing
Set xRecordSet = Nothing

WriteLog "Loaded " & Counter & " ignore(s)."

Exit Sub
HandleErrorFriends:
'Print error
WriteLog Err.Description

'Set error flag
HasError = True
End Sub

Private Sub LoadAccounts(pTable As String)
Dim SQL     As String
Dim LItem   As ListItem
Dim Counter As Long

If HasError Then Exit Sub

SQL = "SELECT * FROM " & pTable
Counter = 0

On Error GoTo HandleErrorTable
xCommand.CommandText = SQL
Set xRecordSet = xCommand.Execute

With frmAccountPanel
    .ListView1.ListItems.Clear
    With .cmbBanned
        .Clear
        .AddItem "False"
        .AddItem "True"
    End With
    With .cmbLevel
        .Clear
        .AddItem "0"
        .AddItem "1"
        .AddItem "2"
    End With
    With .cmbGender
        .Clear
        .AddItem "Male"
        .AddItem "Female"
    End With
End With

With xRecordSet
    Do Until .EOF
        Set LItem = frmAccountPanel.ListView1.ListItems.Add(, , !ID)
        LItem.SubItems(1) = !Name1
        LItem.SubItems(2) = !Password1
        LItem.SubItems(3) = Format$(!Time1, "hh:nn:ss")
        LItem.SubItems(4) = !Date1
        LItem.SubItems(5) = !Banned1
        LItem.SubItems(6) = !Level1
        LItem.SubItems(7) = !SecretQuestion1
        LItem.SubItems(8) = !SecretAnswer1
        LItem.SubItems(9) = !gender1
        .MoveNext
        Counter = Counter + 1
    Loop
End With

Set LItem = Nothing
Set xRecordSet = Nothing

WriteLog "Loaded " & Counter & " account(s)."

Exit Sub
HandleErrorTable:
'Print error
WriteLog Err.Description

'Set error flag
HasError = True
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim msg         As Long
Dim sFilter     As String

msg = X / Screen.TwipsPerPixelX

Select Case msg
    Case WM_LBUTTONDOWN
    Case WM_LBUTTONUP
        Vali = True
        frmMain.Show
        frmMain.WindowState = 0
    Case WM_LBUTTONDBLCLK
    Case WM_RBUTTONDOWN
    Case WM_RBUTTONUP
    Case WM_RBUTTONDBLCLK
End Select
End Sub

Private Sub MDIForm_Resize()
If Me.WindowState = 1 Then
    If Vali = False Then
        MinimizeToTray
    End If
    Vali = False
End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'Remove tray icon
Shell_NotifyIcon NIM_DELETE, nid

'Open Database variable
With Database

'== Database ==
InsertIntoRegistry "Server\Database", "Name", .Database
InsertIntoRegistry "Server\Database", "User", .User
InsertIntoRegistry "Server\Database", "Password", Encode(.Password)
InsertIntoRegistry "Server\Database", "Host", .Host
InsertIntoRegistry "Server\Database", "Account_Table", .AccountTable
InsertIntoRegistry "Server\Database", "Friend_Table", .FriendTable
InsertIntoRegistry "Server\Database", "Ignore_Table", .IgnoreTable
InsertIntoRegistry "Server\Database", "Emote_Table", .EmoteTable
InsertIntoRegistry "Server\Database", "Declined_Name_Table", .DeclinedNameTable
InsertIntoRegistry "Server\Database", "Command_Table", .CommandsTable

'Close Database variable
End With

'== Position ==
InsertIntoRegistry "Server\Configuration", "Top", Me.Top
InsertIntoRegistry "Server\Configuration", "Left", Me.Left

'== Options ==
InsertIntoRegistry "Server\Configuration", "ChatLevel", Options.ChatLevel
End Sub

Private Sub Winsock1_Close(Index As Integer)
Dim pTemp As String
Dim j     As Long

Unload Winsock1(Index)
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(2) = Index Then
            If Not Len(.Item(i)) = 0 Then
                For j = 1 To .Count
                    If Not .Item(j).SubItems(2) = Index Then
                        SetLanguageByID .Item(j).SubItems(7)
                        pTemp = Replace(MSG_GONE_OFFLINE, "%u", .Item(i))
                        SendSingle pTemp, .Item(j).SubItems(2)
                    End If
                Next j
            End If
            .Remove (i)
            Exit For
        End If
    Next i
End With

UPDATE_ONLINE
UPDATE_STATUS_BAR
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim j As Long
j = loadSocket

With Winsock1(j)
    .LocalPort = frmConfig.txtPort.Text
    .Accept requestID
End With

'Add new user to panel without account and name
With frmPanel.ListView1.ListItems
    .Add , , vbNullString
    i = .Count
    If Winsock1(j).RemoteHostIP = "127.0.0.1" Then
        .Item(i).SubItems(1) = Winsock1(0).LocalIP
    Else
        .Item(i).SubItems(1) = Winsock1(j).RemoteHostIP
    End If
    .Item(i).SubItems(2) = j
    .Item(i).SubItems(3) = vbNullString
    .Item(i).SubItems(4) = "False"
    .Item(i).SubItems(5) = vbNullString
    .Item(i).SubItems(6) = Format$(Time, "hh:mm:ss")
End With

UPDATE_STATUS_BAR
End Sub

Private Function socketFree() As Long
On Error GoTo HandleErrorFreeSocket

With Winsock1
    For i = .LBound + 1 To .UBound
        If Winsock1(i).LocalIP Then
        End If
    Next i
    socketFree = .UBound + 1
End With

Exit Function
HandleErrorFreeSocket:
socketFree = i
End Function

Private Function loadSocket() As Integer
Dim theFreeSocket As Integer
theFreeSocket = 0
theFreeSocket = socketFree

Load Winsock1(theFreeSocket)

loadSocket = theFreeSocket
End Function

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim p_PreArray()    As String
Dim k               As Long

Dim p_Message       As String
Dim p_MainArray()   As String   'Whole message string is saved here and split up by # sign
Dim p_Command       As String   'First part of main array ( always the command )
Dim bMatch          As Boolean  'bMatch controls the login
Dim IsMuted         As Boolean  'Mute explains itself

'Get Message
frmMain.Winsock1(Index).GetData p_Message
DoEvents

'Do first array to avoid spam
p_PreArray = Split(p_Message, Chr(24) & Chr(25))

'Start looping through
For k = 0 To UBound(p_PreArray) - 1
    'Print the message
    CMSG p_PreArray(k) & " | Index: " & Index
    
    'We decode (split) the message into an array
    p_MainArray = Split(p_PreArray(k), "#")
    
    'Assign the variable to the array
    If UBound(p_MainArray) > -1 Then
        p_Command = p_MainArray(0)
    End If
    
    Select Case p_Command
        'Select action and execute command
        Case "!friend"
            Select Case p_MainArray(1)
                'Update Friend list
                Case "-get"
                    UPDATE_FRIEND p_MainArray(2), Index
                    
                'Add friend to list by user name
                Case "-add-user"
                    frmFriendIgnoreList.AddFriend p_MainArray(2), GetProperAccountNameByUser(p_MainArray(3)), Index
                    
                'Add friend to list by account name
                Case "-add-account"
                    frmFriendIgnoreList.AddFriend p_MainArray(2), GetProperAccountName(p_MainArray(3)), Index
                    
                'Remove friend from list by user name
                Case "-remove-user"
                    frmFriendIgnoreList.RemoveFriend p_MainArray(2), GetProperAccountNameByUser(p_MainArray(3)), Index
                    
                'Remove friend from list by account name
                Case "-remove-account"
                    frmFriendIgnoreList.RemoveFriend p_MainArray(2), GetProperAccountName(p_MainArray(3)), Index
            End Select
            
        Case "!ignore"
            Select Case p_MainArray(1)
                'Update Ignore list
                Case "-get"
                    UPDATE_IGNORE p_MainArray(2), Index
                    
                'Add ignore to list by user name
                Case "-add-user"
                    frmFriendIgnoreList.AddIgnore p_MainArray(2), GetProperAccountNameByUser(p_MainArray(3)), Index
                    
                'Add ignore to list by account name
                Case "-add-account"
                    frmFriendIgnoreList.AddIgnore p_MainArray(2), GetProperAccountName(p_MainArray(3)), Index
                    
                'Remove ignore from list by user name
                Case "-remove-user"
                    frmFriendIgnoreList.RemoveIgnore p_MainArray(2), GetProperAccountNameByUser(p_MainArray(3)), Index
                    
                'Remove ignore from list by account name
                Case "-remove-account"
                    frmFriendIgnoreList.RemoveIgnore p_MainArray(2), GetProperAccountName(p_MainArray(3)), Index
            End Select
            
        Case "!connected"
            UPDATE_ONLINE
            
            With frmPanel.ListView1.ListItems
                For i = 1 To .Count
                    If Not .Item(i) = p_MainArray(1) Then
                        SetLanguageByID .Item(i).SubItems(7)
                        SendSingle Replace$(MSG_COME_ONLINE, "%u", p_MainArray(1)), .Item(i).SubItems(2)
                    End If
                Next i
            End With
            
        'Send Server information
        Case "!server_info"
            SendSingle "!server_info#" & GetServerInformation, Index
            
        Case "!login"
            With frmAccountPanel.ListView1.ListItems
                For i = 1 To .Count
                    If LCase$(.Item(i).SubItems(1)) = LCase$(p_MainArray(1)) Then
                        'Ban Check
                        If .Item(i).SubItems(5) = "True" Then
                            SendSingle "!login#Banned#", Index
                            Exit Sub
                        End If
                        
                        'Password Check
                        If Not .Item(i).SubItems(2) = p_MainArray(2) Then
                            SendSingle "!login#Password#", Index
                            Exit Sub
                        End If
                        Exit For
                    Else
                        If i = .Count Then
                            SendSingle "!login#Account#", Index
                            Exit Sub
                        End If
                    End If
                Next i
            End With
            
            'Check badname list
            For i = LBound(DeclinedNames) To UBound(DeclinedNames)
                If DeclinedNames(i) = p_MainArray(3) Then
                    SendSingle "!decilined", Index
                    Exit Sub
                End If
            Next i
            
            With frmPanel.ListView1.ListItems
                'Check current online list
                For i = 1 To .Count
                    If .Item(i) = p_MainArray(3) Then
                        SendSingle "!decilined", Index
                        Exit Sub
                    End If
                Next i
                        
                'If the account is already beeing used kick first instance
                For i = 1 To .Count
                    If .Item(i).SubItems(5) = p_MainArray(1) Then
                        Unload Winsock1(.Item(i).SubItems(2))
                        .Remove (i)
                        Exit For
                    End If
                Next i
                
                .Item(.Count).Text = p_MainArray(3)
                .Item(.Count).SubItems(5) = GetProperAccountName(p_MainArray(1))
                .Item(.Count).SubItems(7) = p_MainArray(4)
            End With
            
            UPDATE_STATUS_BAR
            SendSingle "!accepted#", Index
                
        'We get ip request and send ip back
        Case "!iprequest"
            With frmPanel.ListView1.ListItems
                For i = 1 To .Count
                    If .Item(i) = p_MainArray(1) Then
                        SendSingle "!iprequest#" & .Item(i).SubItems(1) & "#", Index
                    End If
                Next i
            End With
            
        Case "!message"
            Dim p_CHAT_ARRAY()      As String
            Dim p_TEXT_FIRST        As String
            Dim p_TEXT_FIRST_PROP   As String
            Dim p_TEXT_SECOND       As String
            Dim p_TEXT_SECOND_PROP  As String
            Dim IsCommand           As Boolean
            Dim IsSlash             As Boolean
                
            'Split the conversation text by spaces
            p_CHAT_ARRAY = Split(p_MainArray(2), " ")
            
            'Check first position of the text for a point indicating command or emote
            Select Case Left$(p_CHAT_ARRAY(0), 1)
                Case Chr(46)
                    If GetLevel(p_MainArray(1)) > 0 Then IsCommand = True
                        
                Case Chr(47)
                    IsSlash = True
                    
            End Select
            
            'Capture first part of the text
            If UBound(p_CHAT_ARRAY) > 0 Then
                p_TEXT_FIRST = p_CHAT_ARRAY(1)
                p_TEXT_FIRST_PROP = StrConv(p_CHAT_ARRAY(1), vbProperCase)
            End If
            
            'Capture second part
            If UBound(p_CHAT_ARRAY) > 1 Then
                p_TEXT_SECOND = p_CHAT_ARRAY(2)
                p_TEXT_SECOND_PROP = StrConv(p_CHAT_ARRAY(2), vbProperCase)
            End If
            
            'Check if user is muted
            With frmPanel.ListView1.ListItems
                For i = 1 To .Count
                    If .Item(i) = p_MainArray(1) Then
                        If .Item(i).SubItems(4) = "True" Then
                            IsMuted = True
                            Exit For
                        End If
                    End If
                Next i
            End With
            
            'If a command is used check out which
            If IsCommand Then
                Dim Reason          As String
                Dim Reason2         As String
                
                'Save the reason
                For i = 2 To UBound(p_CHAT_ARRAY)
                    Reason = Reason & p_CHAT_ARRAY(i) & " "
                Next i
                
                'Save the second reason
                For i = 3 To UBound(p_CHAT_ARRAY)
                    Reason2 = Reason2 & p_CHAT_ARRAY(i) & " "
                Next i
                
                Select Case LCase$(p_CHAT_ARRAY(0))
                    Case ".show"
                        If IsPartOf(p_TEXT_FIRST, "accounts") Then
                            SendSingle "!split_text#" & GetAccountList, Index
                            
                        ElseIf IsPartOf(p_TEXT_FIRST, "users") Then
                            SendSingle "!split_text#" & GetUserList, Index
                        
                        Else
                            SendSingle "Incorrect Syntax, use the following format .show [accounts, users].", Index
                            
                        End If
                    
                    Case ".userinfo", ".uinfo"
                        GetUserInfo p_TEXT_FIRST_PROP, p_CHAT_ARRAY(0), Index
                    
                    Case ".accountinfo", ".accinfo", ".ainfo"
                        GetAccountInfo p_TEXT_FIRST, p_CHAT_ARRAY(0), Index
                    
                    Case ".kick"
                        With frmPanel.ListView1.ListItems
                            For i = 1 To .Count
                                If .Item(i) = p_TEXT_FIRST_PROP Then
                                    KickUser p_TEXT_FIRST_PROP
                                Else
                                    If i = .Count Then
                                        If LenB(p_TEXT_FIRST_PROP) = 0 Then
                                            SendSingle "Incorrect syntax, use following format .kick [User].", Index
                                        Else
                                            SendSingle "User '" & p_TEXT_FIRST_PROP & "' not found.", Index
                                        End If
                                    End If
                                End If
                            Next i
                        End With
                        
                    Case ".ban"
                        If IsPartOf(p_TEXT_FIRST, "user") Then
                            BanUser p_TEXT_SECOND_PROP, p_MainArray(1), True, Index, Trim$(Reason2)
                            
                        ElseIf IsPartOf(p_TEXT_FIRST, "account") Then
                            BanAccount p_TEXT_SECOND, p_MainArray(1), True, Index, Trim$(Reason2)
                        
                        Else
                            SendSingle "Incorrect syntax, use the following format .ban User / Account 'Name' 'Reason'", Index
                            
                        End If
                        
                    Case ".unban"
                        If IsPartOf(p_TEXT_FIRST, "user") Then
                            BanUser p_TEXT_SECOND_PROP, p_MainArray(1), False, Index, Trim$(Reason2)
                            
                        ElseIf IsPartOf(p_TEXT_FIRST, "account") Then
                            BanAccount p_TEXT_SECOND, p_MainArray(1), False, Index, Trim$(Reason2)
                            
                        Else
                            SendSingle "Incorrect syntax, use the following format .unban [User, Account] [Name] [Reason]", Index
                            
                        End If
                    
                    Case ".mute"
                        MuteUser p_TEXT_FIRST_PROP, p_MainArray(1), True, Index, Trim$(Reason)
                    
                    Case ".unmute"
                        MuteUser p_TEXT_FIRST_PROP, p_MainArray(1), False, Index, Trim$(Reason)
                    
                    Case ".announce", ".ann", ".broadcast"
                        Dim p_ANN_MSG As String
                        
                        'Capture announce message
                        If UBound(p_CHAT_ARRAY) > 0 Then
                            p_ANN_MSG = Mid$(p_MainArray(2), Len(p_CHAT_ARRAY(0)) + 2, Len(p_MainArray(2)))
                        End If
                        
                        If LenB(p_ANN_MSG) = 0 Then
                            SendSingle "Incorrect syntax, use the following format " & p_CHAT_ARRAY(0) & " [Text].", Index
                        Else
                            With frmPanel.ListView1.ListItems
                                For i = 1 To .Count
                                    SetLanguageByID .Item(i).SubItems(7)
                                    SendSingle Replace$(MSG_ANNOUNCE, "%u", p_MainArray(1)) & p_ANN_MSG, .Item(i).SubItems(2)
                                Next i
                            End With
                        End If
                    
                    Case ".help", ".command", ".commands"
                        SendSingle GetCommands, Index
                        
                    Case ".reload"
                        With Database
                            Select Case LCase$(p_TEXT_FIRST)
                                Case LCase$(.AccountTable), LCase$(.FriendTable), LCase$(.IgnoreTable)
                                    SendSingle "This table can't be reloaded.", Index
                                
                                Case LCase$(.CommandsTable)
                                    Erase Commands
                                    LoadCommands .CommandsTable
                                    SendMessage p_MainArray(1) & " initiated the reload of '" & .CommandsTable & "' table."
                                    
                                Case LCase$(.DeclinedNameTable)
                                    Erase DeclinedNames
                                    LoadDeclinedNames .DeclinedNameTable
                                    SendMessage p_MainArray(1) & " initiated the reload of '" & .DeclinedNameTable & "' table."
                                    
                                Case LCase$(.EmoteTable)
                                    Erase Emotes
                                    LoadEmotes .EmoteTable
                                    SendMessage p_MainArray(1) & " initiated the reload of '" & .EmoteTable & "' table."
                                
                                Case Else
                                    If LenB(p_TEXT_FIRST) = 0 Then
                                        SendSingle "Incorrect Syntax. Use the following format .reload Table.", Index
                                    Else
                                        SendSingle "This table does not exist.", Index
                                    End If
                            End Select
                        End With
                    
                    Case ".clear"
                        If LenB(p_TEXT_FIRST) = 0 Then
                            SendSingle "Incorrect Syntax. Use the following format .clear [User]", Index
                        Else
                            With frmPanel.ListView1.ListItems
                                For i = 1 To .Count
                                    If .Item(i) = p_TEXT_FIRST_PROP Then
                                        SendSingle "!clear#", .Item(i).SubItems(2)
                                        Exit For
                                    Else
                                        If i = .Count Then SendSingle "User '" & p_TEXT_FIRST & "' was not found.", Index
                                    End If
                                Next i
                            End With
                        End If
                        
                    Case Else
                        SendSingle "Unknown command used. Check .help for more information about commands.", Index
                End Select
                Exit Sub
            End If
            
            If IsSlash Then
                Dim IsUser As Boolean
                
                If IsMuted Then
                    SendSingle "You are muted.", Index
                    Exit Sub
                End If
                
                If GetLevel(p_MainArray(1)) = Options.ChatLevel Then
                    If IsRepeating(p_MainArray(1), p_MainArray(2)) Then
                        SendSingle "Your message has triggered serverside flood protection. Please don't repeat yourself.", Index
                        Exit Sub
                    End If
                End If
                
                With frmPanel.ListView1.ListItems
                    For i = 1 To .Count
                        If .Item(i) = p_TEXT_FIRST_PROP Then
                            IsUser = True
                            Exit For
                        End If
                    Next i
                End With
                
                Select Case LCase(p_CHAT_ARRAY(0))
                    'Roll function
                    Case "/roll"
                        Dim pRoll       As Long
                        Dim pMinRoll    As Long
                        Dim pMaxRoll    As Long
                         
                        If UBound(p_CHAT_ARRAY) > 1 Then
                            If IsNumeric(p_CHAT_ARRAY(1)) Then
                                If p_CHAT_ARRAY(1) > MAX_INT_VALUE Or p_CHAT_ARRAY(1) < MIN_INT_VALUE Then
                                    pMinRoll = 1
                                Else
                                    pMinRoll = p_CHAT_ARRAY(1)
                                End If
                            Else
                                pMinRoll = 1
                            End If
                            
                            If IsNumeric(p_CHAT_ARRAY(2)) Then
                                If p_CHAT_ARRAY(2) > MAX_INT_VALUE Or p_CHAT_ARRAY(2) < MIN_INT_VALUE Then
                                    pMaxRoll = 100
                                Else
                                    pMaxRoll = p_CHAT_ARRAY(2)
                                End If
                            Else
                                pMaxRoll = 100
                            End If
                                                        
                        ElseIf UBound(p_CHAT_ARRAY) > 0 Then
                            pMinRoll = 1
                            
                            If IsNumeric(p_CHAT_ARRAY(1)) Then
                                If p_CHAT_ARRAY(1) > MAX_INT_VALUE Or p_CHAT_ARRAY(1) < MIN_INT_VALUE Then
                                    pMaxRoll = 100
                                Else
                                    pMaxRoll = p_CHAT_ARRAY(1)
                                End If
                            Else
                                pMaxRoll = 100
                            End If
                            
                        Else
                            pMinRoll = 1
                            pMaxRoll = 100
                            
                        End If
                        
                        If pMinRoll > pMaxRoll Then pMaxRoll = 100
                        If pMinRoll = 0 Then pMinRoll = 1
                        If pMaxRoll = 0 Then pMaxRoll = 100
                        
                        pRoll = GetRandomNumber(pMinRoll, pMaxRoll)
                        
                        SendProtectedMessage p_MainArray(1), p_MainArray(1) & " rolls " & pRoll & ". (" & pMinRoll & " - " & pMaxRoll & ")"
                        
                    'Whisper X to Z from Y
                    Case "/w", "/whisper"
                        If IsUser Then
                            Dim pMes As String
                            
                            If UBound(p_CHAT_ARRAY) > 1 Then
                                For i = 2 To UBound(p_CHAT_ARRAY)
                                    pMes = pMes & p_CHAT_ARRAY(i) & " "
                                Next i
                                
                                Whisper p_MainArray(1), p_TEXT_FIRST_PROP, Trim$(pMes), Index
                                SetLastMessage p_MainArray(1), p_MainArray(2)
                            Else
                                Exit Sub
                            End If
                        Else
                            If LenB(p_TEXT_FIRST_PROP) = 0 Then
                                SendSingle "Please use the following format /whisper [Name] [Text]", Index
                                Exit Sub
                            Else
                                SendSingle "No user named '" & p_TEXT_FIRST & "' is currently online.", Index
                            End If
                        End If
                        
                    Case "/online"
                        SendSingle "You are online for " & Trim$(GetOnlineTime(p_MainArray(1))) & ".", Index
                        
                    Case "/logout"
                        KickUser p_MainArray(1)
                        
                    Case Else
                        For i = LBound(Emotes) To UBound(Emotes)
                            If LCase$(Emotes(i).Command) = LCase$(p_CHAT_ARRAY(0)) Then
                                If p_MainArray(1) = p_TEXT_FIRST_PROP Then
                                    IsUser = False
                                End If
                                                                
                                Dim pTemp   As String
                                Dim pGen    As String
                                Dim j       As Long
                                                                
                                With frmPanel.ListView1.ListItems
                                    For j = 1 To .Count
                                        If .Item(j) = p_MainArray(1) Then
                                            pTemp = .Item(j).SubItems(5)
                                            Exit For
                                        End If
                                    Next j
                                End With
                                
                                With frmAccountPanel.ListView1.ListItems
                                    For j = 1 To .Count
                                        If .Item(j).SubItems(1) = pTemp Then
                                            Select Case .Item(j).SubItems(9)
                                                Case "Male"
                                                    pGen = "his"
                                                Case "Female"
                                                    pGen = "her"
                                            End Select
                                            Exit For
                                        End If
                                    Next j
                                End With
                                                                
                                If IsUser Then
                                    pTemp = Replace(Emotes(i).TargetEmote, "%u", p_MainArray(1))
                                    pTemp = Replace(pTemp, "%g", pGen)
                                    pTemp = Replace(pTemp, "%t", p_TEXT_FIRST_PROP)
                                Else
                                    pTemp = Replace(Emotes(i).SingleEmote, "%u", p_MainArray(1))
                                    pTemp = Replace(pTemp, "%g", pGen)
                                End If
                                SendProtectedMessage p_MainArray(1), pTemp
                                Exit For
                            Else
                                If i = UBound(Emotes) Then SendSingle "Unknown command used.", Index
                            End If
                        Next i
                        
                End Select
                SetLastMessage p_MainArray(1), p_MainArray(2)
                Exit Sub
            End If
            
            'Check if user is muted
            If IsMuted Then
                SendSingle "You are muted.", Index
                Exit Sub
            End If
            
            Dim S1  As Long
            Dim E   As String
            
            'Only certain level accounts can by pass special rules
            If GetLevel(p_MainArray(1)) = Options.ChatLevel Then
                'We just bother checking if the text is longer then 5 characters
                If Len(p_MainArray(2)) > 5 Then
                    'If the text is just made of numbers we don't check
                    If IsNumeric(p_MainArray(2)) = False Then
                        'If there are letters and not just signs then check
                        If IsAlphaCharacter(p_MainArray(2)) Then
                            S1 = 0
                            For i = 1 To Len(p_MainArray(2))
                                E = Mid$(p_MainArray(2), i, 1)
                                If UCase$(E) = E Then S1 = S1 + 1
                            Next i
                            E = vbNullString
                            'Exit if there are more then 75% of caps
                            If Format$(100 * S1 / Len(p_MainArray(2)), "0") > 75 Then
                                SendSingle "Message blocked. Please do not write more then 75% in caps.", Index
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                
                If IsRepeating(p_MainArray(1), p_MainArray(2)) Then
                    SendSingle "Your message has triggered serverside flood protection. Please don't repeat yourself.", Index
                    Exit Sub
                End If
            End If
            
            'Send Message and print in chat
            SendProtectedMessage p_MainArray(1), "[" & p_MainArray(1) & "]: " & p_MainArray(2)
            
            'Set last message
            SetLastMessage p_MainArray(1), p_MainArray(2)
            
        Case Else
            SendMessage "Error."
    End Select
Next k
End Sub

Private Sub SetLastMessage(pUser As String, pMessage As String)
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = pUser Then
            .Item(i).SubItems(3) = pMessage
            Exit For
        End If
    Next i
End With
End Sub

Private Function IsRepeating(pUser As String, pMessage As String) As Boolean
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = pUser Then
            If .Item(i).SubItems(3) = pMessage Then
                 IsRepeating = True
            End If
            Exit For
        End If
    Next i
End With
End Function

Private Function IsPartOf(pPart As String, pCommand As String) As Boolean
Dim pTemp1  As String
Dim j       As Long

For j = 1 To Len(pPart)
    pTemp1 = Mid(pPart, 1, j)
    If LCase$(pTemp1) = LCase$(Left$(pCommand, Len(pTemp1))) Then
        IsPartOf = True
    Else
        IsPartOf = False
        Exit For
    End If
Next j
End Function

Private Function GetOnlineTime(pUser As String) As String
Dim TD      As String
Dim TD1()   As String
Dim j       As Long

With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = pUser Then
            TD = TimeSerial(0, 0, DateDiff("s", .Item(i).SubItems(6), Time))
            
            TD1 = Split(TD, ":")
            TD = vbNullString
            
            For j = LBound(TD1) To UBound(TD1)
                Select Case j
                    Case 0
                        If TD1(j) <> 0 Then
                            If Left$(TD1(j), 1) = 0 Then
                                TD = TD & Left$(TD1(j), 1) & " hours"
                            Else
                                TD = TD & TD1(j) & " hours "
                            End If
                        End If
                        
                    Case 1
                        If TD1(j) <> 0 Then
                            If Left$(TD1(j), 1) = 0 Then
                                TD = TD & Right$(TD1(j), 1) & " minutes "
                            Else
                                TD = TD & TD1(j) & " minutes "
                            End If
                        End If
                        
                    Case 2
                        If TD1(j) <> 0 Then
                            If Left$(TD1(j), 1) = 0 Then
                                TD = TD & Right$(TD1(j), 1) & " seconds"
                            Else
                                TD = TD & TD1(j) & " seconds"
                            End If
                        End If
                End Select
            Next j
            
            GetOnlineTime = TD
        End If
    Next i
End With
End Function

Private Function GetProperAccountName(pAccount As String) As String
With frmAccountPanel.ListView1.ListItems
    For i = 1 To .Count
        If LCase$(.Item(i).SubItems(1)) = LCase$(pAccount) Then
            GetProperAccountName = .Item(i).SubItems(1)
            Exit For
        Else
            If i = .Count Then
                GetProperAccountName = pAccount
            End If
        End If
    Next i
End With
End Function

Private Function GetProperAccountNameByUser(pUser As String) As String
Dim pAccount As String

With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = pUser Then
            GetProperAccountNameByUser = .Item(i).SubItems(5)
            Exit For
        End If
    Next i
End With
End Function

Private Function GetServerInformation() As String
GetServerInformation = _
"Welcome to Peach Servers." & "#" & _
"Server: Peach r" & pRev & "/" & GetOS & "#" & _
"Online User: " & frmMain.Winsock1.Count - 1 & "#" & _
frmConfig.Label2.Caption & "#"
End Function

Private Sub Whisper(pUser As String, pTarget As String, pConv As String, Index As Integer)
Dim pAccount As String

'Check if user is whispering itself
If pUser = pTarget Then
    SendSingle "You can't whisper yourself.", Index
    Exit Sub
End If

With frmPanel.ListView1.ListItems
    'Get user account name
    For i = 1 To .Count
        If .Item(i) = pUser Then
            pAccount = .Item(i).SubItems(5)
            Exit For
        End If
    Next i

    'Search target in list and send message
    For i = 1 To .Count
        If pTarget = .Item(i) Then
            If IsIgnoring(.Item(i).SubItems(5), pAccount) = True Then
                SendSingle pTarget & " is ignoring you.", Index
            Else
                SendSingle "[You whisper to " & pTarget & "]: " & pConv, Index
                SendSingle "[" & pUser & " whispers]: " & pConv, .Item(i).SubItems(2)
            End If
            Exit For
        End If
    Next i
End With
End Sub

Private Sub MuteUser(User As String, AdminName As String, IsMuted As Boolean, pIndex As Integer, Reason As String)
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = User Then
            
            If IsMuted And .Item(i).SubItems(4) = "True" Then
                SendSingle User & " is already muted.", pIndex
                Exit Sub
                
            ElseIf Not IsMuted And .Item(i).SubItems(4) = "False" Then
                SendSingle User & " is not muted.", pIndex
                Exit Sub
                
            End If
            
            'Set flag in userlist
            .Item(i).SubItems(4) = CStr(IsMuted)
            
            'Announce the action
            If LenB(Reason) = 0 Then
                If IsMuted Then
                    SendMessage User & " got muted by " & AdminName & "."
                Else
                    SendMessage User & " got unmuted by " & AdminName & "."
                End If
            Else
                If IsMuted Then
                    SendMessage User & " got muted by " & AdminName & ". (" & Reason & ")"
                Else
                    SendMessage User & " got unmuted by " & AdminName & ". (" & Reason & ")"
                End If
            End If
            Exit For
        Else
            If i = .Count Then
                If LenB(Trim$(User)) = 0 Then
                    SendSingle "Incorrect syntax, use the following format .mute [User] [Reason].", pIndex
                Else
                    SendSingle "User '" & User & "' was not found.", pIndex
                End If
            End If
        End If
    Next i
End With
End Sub

Private Function GetAccountList() As String
With frmAccountPanel.ListView1.ListItems
    GetAccountList = "Account List:#"
    For i = 1 To .Count
        GetAccountList = GetAccountList & .Item(i).SubItems(1) & "#"
    Next i
End With
End Function

Private Function GetUserList() As String
With frmPanel.ListView1.ListItems
    GetUserList = "User List:#"
    For i = 1 To .Count
        GetUserList = GetUserList & .Item(i) & "#"
    Next i
End With
End Function

Private Sub BanUser(User As String, AdminName As String, Ban As Boolean, pIndex As Integer, Reason As String)
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = StrConv(User, vbProperCase) Then
            BanAccount .Item(i).SubItems(5), AdminName, Ban, pIndex, Reason
            Exit For
        Else
            If i = .Count Then
                If LenB(User) = 0 Then
                    If Ban Then
                        SendSingle "Incorrect syntax, use the following format .ban user [User] [Reason].", pIndex
                    Else
                        SendSingle "Incorrect syntax, use the following format .unban user [User] [Reason].", pIndex
                    End If
                Else
                    SendSingle "User '" & User & "' not found.", pIndex
                End If
            End If
        End If
    Next i
End With
End Sub

Private Sub BanAccount(Account As String, AdminName As String, Ban As Boolean, pIndex As Integer, Reason As String)
Dim User As String
Dim j    As Long

With frmAccountPanel.ListView1.ListItems
    For i = 1 To .Count
        If LCase(.Item(i).SubItems(1)) = LCase(Account) Then
            'If the account is already banned send feedback
            Account = .Item(i).SubItems(1)
            
            If Ban Then
                If .Item(i).SubItems(5) = "True" Then
                    SendSingle "Account '" & Account & "' is already banned.", pIndex
                    Exit Sub
                End If
            Else
                If .Item(i).SubItems(5) = "False" Then
                    SendSingle "Account '" & Account & "' is not banned.", pIndex
                    Exit Sub
                End If
            End If
            
            'Ban account in database
            frmAccountPanel.ModifyAccount Account, .Item(i).SubItems(2), Ban, .Item(i).SubItems(6), .Item(i), .Item(i).Index, .Item(i).SubItems(9)
            
            'Determine user from account
            With frmPanel.ListView1.ListItems
                For j = 1 To .Count
                    If LCase(.Item(j).SubItems(5)) = LCase(Account) Then
                        User = .Item(j)
                        Exit For
                    End If
                Next j
            End With
            
            If LenB(Trim$(User)) = 0 Then
                User = Account
            End If
            
            'Announce the action
            If LenB(Reason) = 0 Then
                If Ban Then
                    SendMessage User & " was account banned by " & AdminName & "."
                Else
                    SendMessage User & " was account unbanned by " & AdminName & "."
                End If
            Else
                If Ban Then
                    SendMessage User & " was account banned by " & AdminName & ". (" & Reason & ")"
                Else
                    SendMessage User & " was account unbanned by " & AdminName & ". (" & Reason & ")"
                End If
            End If
            Exit For
        Else
            If i = .Count Then
                If LenB(Trim$(Account)) = 0 Then
                    SendSingle "Incorrect syntax, use the following format .ban account [Account] [Reason].", pIndex
                Else
                    SendSingle "Account '" & Account & "' not found.", pIndex
                End If
            End If
        End If
    Next i
End With
End Sub

Private Sub KickUser(pUser As String)
Dim j       As Long
Dim pTemp   As String

With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = pUser Then
            Unload frmMain.Winsock1(.Item(i).SubItems(2))
            
            For j = 1 To .Count
                If Not .Item(j) = pUser Then
                    SetLanguageByID .Item(j).SubItems(7)
                    pTemp = Replace(MSG_GONE_OFFLINE, "%u", .Item(i))
                    SendSingle pTemp, .Item(j).SubItems(2)
                End If
            Next j
            
            .Remove (i)
            
            UPDATE_ONLINE
            UPDATE_STATUS_BAR
            Exit For
        End If
    Next i
End With
End Sub

Private Sub GetAccountInfo(Account As String, pUsedSyntax As String, pIndex As Integer)
With frmAccountPanel.ListView1.ListItems
    For i = 1 To .Count
        If LCase(.Item(i).SubItems(1)) = LCase(Account) Then
            SendSingle vbCrLf & " Account information about '" & Account & "'" & vbCrLf & " Name: " & .Item(i).SubItems(1) & vbCrLf & " Password: " & .Item(i).SubItems(2) & vbCrLf & " Registration Time: " & .Item(i).SubItems(3) & vbCrLf & " Registration Date: " & .Item(i).SubItems(4) & vbCrLf & " Banned: " & .Item(i).SubItems(5) & vbCrLf & " Level: " & .Item(i).SubItems(6), pIndex
            Exit For
        Else
            If i = .Count Then
                If LenB(Trim$(Account)) = 0 Then
                    SendSingle "Incorrect syntax, use following format " & pUsedSyntax & " [Account].", pIndex
                Else
                    SendSingle "Account '" & Account & "' not found.", pIndex
                End If
            End If
        End If
    Next i
End With
End Sub

Private Sub GetUserInfo(pUser As String, pUsedSyntax As String, pIndex As Integer)
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = StrConv(pUser, vbProperCase) Then
            SendSingle vbCrLf & "User information about '" & pUser & "'" & vbCrLf & " IP : " & .Item(i).SubItems(1) & vbCrLf & " Winsock ID: " & .Item(i).SubItems(2) & vbCrLf & " Last Message: " & .Item(i).SubItems(3) & vbCrLf & " Muted: " & .Item(i).SubItems(4) & vbCrLf & " Account: " & .Item(i).SubItems(5) & vbCrLf & " Login Time: " & .Item(i).SubItems(6), pIndex
            Exit For
        Else
            If i = .Count Then
                If LenB(pUser) = 0 Then
                    SendSingle "Incorrect syntax, use following format " & pUsedSyntax & " [User].", pIndex
                Else
                    SendSingle "User '" & pUser & "' was not found.", pIndex
                End If
            End If
        End If
    Next i
End With
End Sub

Private Function GetCommands() As String
GetCommands = vbCrLf & "*********************************************" & vbCrLf & "* List of all avaible commands:" & vbCrLf
For i = 0 To UBound(Commands) - 1
    GetCommands = GetCommands & "* " & Commands(i).Syntax & " (" & Commands(i).Description & ")" & vbCrLf
Next i
GetCommands = GetCommands & "*********************************************"
End Function

Private Function GetLevel(ByVal Name As String) As Long
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = Name Then
            Name = .Item(i).SubItems(5)
            Exit For
        End If
    Next i
End With

With frmAccountPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(1) = Name Then
            GetLevel = .Item(i).SubItems(6)
            Exit For
        End If
    Next i
End With
End Function

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim j       As Long
Dim pTemp   As String

Unload Winsock1(Index)
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(2) = Index Then
            If Not Len(.Item(i)) = 0 Then
                For j = 1 To .Count
                    If Not .Item(j).SubItems(2) = Index Then
                        SetLanguageByID .Item(j).SubItems(7)
                        pTemp = Replace(MSG_GONE_OFFLINE, "%u", .Item(i))
                        SendSingle pTemp, .Item(j).SubItems(2)
                    End If
                Next j
            End If
            .Remove (i)
            Exit For
        End If
    Next i
End With

UPDATE_ONLINE
UPDATE_STATUS_BAR
End Sub

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

SendMessage2 hwnd, WM_NCACTIVATE, True, 0

frm.Width = frm.Width - 1
frm.Width = frm.Width + 1
End Sub

Public Sub SetupForms(pForm As Form)
frmChat.Hide
frmFriendIgnoreList.Hide
frmConfig.Hide
frmAccountPanel.Hide
frmPanel.Hide
pForm.Show
End Sub
