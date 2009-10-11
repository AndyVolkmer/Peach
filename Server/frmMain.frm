VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00F4F4F4&
   Caption         =   " Peach Server"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7395
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
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

Private Const GWL_STYLE = (-16)
Private Const WS_SYSMENU = &H80000
Private Const WS_CAPTION = &HC00000
Private Const WS_THICKFRAME = &H40000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const SC_CLOSE As Long = &HF060&
Private Const SC_MAXIMIZE = &HF030&
Private Const SC_MINIMIZE = &HF020&
Private Const MIIM_STATE As Long = &H1&
Private Const MIIM_ID As Long = &H2&
Private Const MFS_GRAYED As Long = &H3&
Private Const WM_NCACTIVATE As Long = &H86

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
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

Public xConnection     As New ADODB.Connection
Public xCommand        As New ADODB.Command
Public xRecordSet      As New ADODB.Recordset

Public intCounter   As Long
Dim Muted, Vali     As Boolean

Private Sub Command1_Click()
SetupForms frmConfig
End Sub

Private Sub Command2_Click()
SetupForms frmChat
End Sub

Public Sub SetupForms(Nix As Form)
frmChat.Hide
frmFriendList.Hide
frmConfig.Hide
frmAccountPanel.Hide
frmPanel.Hide
Nix.Show
End Sub

Private Sub Command3_Click()
SetupForms frmAccountPanel
End Sub

Private Sub Command4_Click()
SetupForms frmPanel
End Sub

Private Sub Command5_Click()
SetupForms frmFriendList
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

'Assign Variables to the .ini strings
If Len(Trim$(ReadIniValue(App.Path & "\Config.ini", "Database", "Database"))) <> 0 Then
    .Database = ReadIniValue(App.Path & "\Config.ini", "Database", "Database")
Else
    WriteLog "No Database value found."
    .Database = InputBox("The configuration file does not cotain a valid database, please insert one in the textbox below.", "Database error ..", "Database Name")
End If

If Len(Trim$(ReadIniValue(App.Path & "\Config.ini", "Database", "User"))) <> 0 Then
    .User = ReadIniValue(App.Path & "\Config.ini", "Database", "User")
Else
    WriteLog "No User value found."
    .User = InputBox("The configuration file does not cotain a valid user name, please insert one in the textbox below.", "Database error ..", "User Name")
End If

If Len(Trim$(DeCode(ReadIniValue(App.Path & "\Config.ini", "Database", "Password")))) <> 0 Then
    .Password = DeCode(ReadIniValue(App.Path & "\Config.ini", "Database", "Password"))
Else
    WriteLog "No Password value found."
    .Password = InputBox("The configuration file does not cotain a valid password, please insert one in the textbox below.", "Database error ..", "Password")
End If

If Len(Trim$(ReadIniValue(App.Path & "\Config.ini", "Database", "Host"))) <> 0 Then
    .Host = ReadIniValue(App.Path & "\Config.ini", "Database", "Host")
Else
    WriteLog "No Host value found."
    .Host = InputBox("The configuration file does not cotain a host adress, please insert one in the textbox below.", "Database error ..", "Host Adress")
End If

If Len(Trim$(ReadIniValue(App.Path & "\Config.ini", "Database", "A_Table"))) <> 0 Then
    .Account_Table = ReadIniValue(App.Path & "\Config.ini", "Database", "A_Table")
Else
    WriteLog "No Account-Table found."
    .Account_Table = InputBox("The configuration file does not contain a account table, please insert one in the textbox below.", "Database error ..", "Account Table")
End If

If Len(Trim$(ReadIniValue(App.Path & "\Config.ini", "Database", "F_Table"))) <> 0 Then
    .Friend_Table = ReadIniValue(App.Path & "\Config.ini", "Database", "F_Table")
Else
    WriteLog "No Friends-Table found."
    .Friend_Table = InputBox("The configuration file does not contain a friend table, please insert one in the textbox below.", "Database error ..", "Friend Table")
End If

WriteLog "Values from configuration:" _
        & vbCrLf & vbTab & "Database: " & .Database _
        & vbCrLf & vbTab & "Host: " & .Host _
        & vbCrLf & vbTab & "Password: " & .Password _
        & vbCrLf & vbTab & "Account Table: " & .Account_Table _
        & vbCrLf & vbTab & "Friend Table: " & .Friend_Table

'Connect to MySQL Database
CONNECT_MYSQL .Database, .User, .Password, .Host

'Load Accounts
LoadAccounts .Account_Table

'Load Friends
LoadFriends .Friend_Table

'Close Database variable
End With

StatusBar1.Panels(1).Text = "Status : Disconnected"
SetupForms frmConfig

If HasError = False Then
    WriteLog "Correctly loaded!"
End If
End Sub

Public Sub CONNECT_MYSQL(qDatabase As String, qUser As String, qPassword As String, qIP As String)
StatusBar1.Panels(1).Text = "Status : Connecting to database .."

On Error GoTo HandleErrorConnection
'Connect with Database (MySQL)
Set xConnection = New ADODB.Connection
xConnection.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
    & "SERVER=" & qIP & ";" _
    & "DATABASE=" & qDatabase & ";" _
    & "UID=" & qUser & ";" _
    & "PWD=" & qPassword & ";" _
    & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384

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

Private Sub LoadFriends(qTable As String)
Dim SQL     As String
Dim LItem   As ListItem
Dim Counter As Long

If HasError = True Then Exit Sub
If Len(qTable) = 0 Then Exit Sub

SQL = "SELECT * FROM " & qTable
Counter = 0

StatusBar1.Panels(1).Text = "Status : Loading friends .."

On Error GoTo HandleErrorFriends
xCommand.CommandText = SQL
Set xRecordSet = xCommand.Execute

With xRecordSet
    Do Until .EOF
        Set LItem = frmFriendList.ListView1.ListItems.Add(, , !ID)
        LItem.SubItems(1) = !Name
        LItem.SubItems(2) = !Friend
        .MoveNext
        Counter = Counter + 1
    Loop
End With

Set LItem = Nothing
Set xRecordSet = Nothing

WriteLog "Loaded " & Counter & " relation(s) from '" & qTable & "'."

Exit Sub
HandleErrorFriends:
'Print error
WriteLog Err.Description & "."

'Set error flag
HasError = True
End Sub

Private Sub LoadAccounts(qTable As String)
Dim SQL     As String
Dim LItem   As ListItem
Dim Counter As Long

If HasError = True Then Exit Sub
If Len(qTable) = 0 Then Exit Sub

SQL = "SELECT * FROM " & qTable
Counter = 0

StatusBar1.Panels(1).Text = "Status : Loading accounts .."

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
End With

With xRecordSet
    Do Until .EOF
        Set LItem = frmAccountPanel.ListView1.ListItems.Add(, , !ID)
        LItem.SubItems(1) = !Name1
        LItem.SubItems(2) = !Password1
        LItem.SubItems(3) = Format$(!Time1, "hh:mm:ss")
        LItem.SubItems(4) = !Date1
        LItem.SubItems(5) = !Banned1
        LItem.SubItems(6) = !Level1
        .MoveNext
        Counter = Counter + 1
    Loop
End With

Set LItem = Nothing
Set xRecordSet = Nothing

WriteLog "Loaded " & Counter & " account(s) from '" & qTable & "'."

Exit Sub
HandleErrorTable:
'Print error
WriteLog Err.Description & "."

'Set error flag
HasError = True
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As Long
Dim sFilter As String
msg = X / Screen.TwipsPerPixelX
Select Case msg
Case WM_LBUTTONDOWN
Case WM_LBUTTONUP
Case WM_LBUTTONDBLCLK
    Vali = True
    frmMain.Show 'show form
Case WM_RBUTTONDOWN
Case WM_RBUTTONUP
Case WM_RBUTTONDBLCLK
End Select
End Sub

Private Sub MDIForm_Resize()
If Me.WindowState = 1 Then
    If Vali = False Then
        minimize_to_tray
    End If
    Vali = False
End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'Open Database variable
With Database

'Write values into .ini file
WriteIniValue App.Path & "\Config.ini", "Database", "Database", .Database
WriteIniValue App.Path & "\Config.ini", "Database", "User", .User
WriteIniValue App.Path & "\Config.ini", "Database", "Password", .Password
WriteIniValue App.Path & "\Config.ini", "Database", "Host", .Host
WriteIniValue App.Path & "\Config.ini", "Database", "A_Table", .Account_Table
WriteIniValue App.Path & "\Config.ini", "Database", "F_Table", .Friend_Table

'Close Database variable
End With
End Sub

Private Sub Winsock1_Close(Index As Integer)
Unload Winsock1(Index)
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(2) = Index Then
            .Remove (i)
            UpdateUsersList
            Exit For
        End If
    Next i
End With
StatusBar1.Panels(1).Text = "Status: Connected with " & Winsock1.Count - 1 & " Client(s)."
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim MAX_CONNECTION As Long
MAX_CONNECTION = loadSocket

Winsock1(MAX_CONNECTION).LocalPort = frmConfig.txtPort.Text
Winsock1(MAX_CONNECTION).Accept requestID

'Add new user to panel without account and name
With frmPanel.ListView1.ListItems
    .Add , , vbNullString
    i = .Count
    .Item(i).SubItems(1) = Winsock1(MAX_CONNECTION).RemoteHostIP
    .Item(i).SubItems(2) = MAX_CONNECTION
    .Item(i).SubItems(3) = vbNullString
    .Item(i).SubItems(4) = "False"
    .Item(i).SubItems(5) = vbNullString
    .Item(i).SubItems(6) = Format$(Time, "hh:mm:ss")
    .Item(i).SubItems(7) = "False"
End With

StatusBar1.Panels(1).Text = "Status: Connected with " & Winsock1.Count - 1 & " Client(s)."
End Sub

Private Function socketFree() As Integer
On Error GoTo HandleErrorFreeSocket

For i = Winsock1.LBound + 1 To Winsock1.UBound
    If Winsock1(i).LocalIP Then
    End If
Next i

socketFree = Winsock1.UBound + 1
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
Dim array1()        As String
Dim array2()        As String
Dim GetUser         As String
Dim GetMessage      As String
Dim GetConver       As String
Dim GetCommand      As String
Dim GetTarget       As String
Dim pGetTarget      As String
Dim GetLastMessage  As String
Dim bMatch, Mute    As Boolean

'Get Message
frmMain.Winsock1(Index).GetData GetMessage

'We decode (split) the message into an array
array1 = Split(GetMessage, "#")

'Assign the variables to the array
GetCommand = array1(0)
GetUser = array1(1)
GetConver = array1(2)

With frmPanel.ListView1.ListItems
    'Get the latest message
    For i = 1 To .Count
        If GetUser = .Item(i) Then
            GetLastMessage = .Item(i).SubItems(3)
            Exit For
        End If
    Next i

    'Check if user is muted
    For i = 1 To .Count
        If .Item(i) = GetUser Then
            If .Item(i).SubItems(4) = "True" Then
                Mute = True
                Exit For
            End If
        End If
    Next i
    
    'Get proper account name
    For i = 1 To .Count
        If LCase(.Item(i).SubItems(5)) = LCase(GetUser) Then
            pGetTarget = .Item(i).SubItems(5)
            Exit For
        End If
    Next i
End With

'Validate: If message is to long then kick
If Len(GetConver) > 200 Then
    CMSG "!long", GetUser
    frmPanel.ListView1.ListItems.Remove (Index) 'Remove from list
    Winsock1(Index).Close 'Close connection
    Unload Winsock1(Index) 'Remove socket
    StatusBar1.Panels(1).Text = "Status: Connected with " & Winsock1.Count - 1 & " Client(s)."
    UpdateUsersList
    Exit Sub
End If

Select Case GetCommand
'Send Server information
Case "!server_info"
    SendSingle "!split_text#" & GetServerInformation, frmMain.Winsock1(Index)
    
'Update Friend list
Case "!get_friends"
    UpdateFriendList pGetTarget
    
'Check if friends exist and save
Case "!add_friend"
    frmFriendList.AddFriend pGetTarget, GetConver, Index

'Remove friend from list
Case "!remove_friend"
    frmFriendList.RemoveFriend pGetTarget, GetConver, Index
   
'Add Name & Account to frmPanel ListView
Case "!connected"
    Dim ORIGINAL_ACCOUNT As String
    
    'If the account is already beeing used kick first instance
    With frmPanel.ListView1.ListItems
        For i = 1 To .Count
            If .Item(i).SubItems(5) = GetConver Then
                CMSG "!dinstance", GetConver
                Winsock1(.Item(i).SubItems(2)).Close
                Unload Winsock1(.Item(i).SubItems(2))
                .Remove (i)
                
                StatusBar1.Panels(1).Text = "Status: Connected with " & Winsock1.Count - 1 & " Client(s)."
                Exit For
            End If
        Next i
    End With
    
    'Get the proper written account name
    With frmAccountPanel.ListView1.ListItems
        For i = 1 To .Count
            If LCase(.Item(i).SubItems(1)) = LCase(GetConver) Then
                ORIGINAL_ACCOUNT = .Item(i).SubItems(1)
                Exit For
            End If
        Next i
    End With
    
    CMSG GetCommand, GetUser
    With frmPanel.ListView1.ListItems
        i = .Count
        .Item(i).Text = GetUser
        .Item(i).SubItems(5) = ORIGINAL_ACCOUNT
    End With
    UpdateUsersList

Case "!namerequest"
    CMSG GetCommand, GetUser
    'Check badname list
    For i = 0 To frmPanel.List1.ListCount
        If frmPanel.List1.List(i) = GetUser Then
            bMatch = True
            CMSG "!badname", GetUser
            Exit For
        End If
    Next i
    
    'Check current online list
    With frmPanel.ListView1.ListItems
        For i = 1 To .Count
            If .Item(i) = GetUser Then
                bMatch = True
                CMSG "!nametaken", GetUser
                Exit For
            End If
        Next i
    End With
    
    'Return answer to client
    If bMatch = True Then
        SendSingle "!decilined", frmMain.Winsock1(Index)
    Else
        SendSingle "!accepted", frmMain.Winsock1(Index)
        CMSG "!nameisfree", GetUser
    End If
    
Case "!login"
    With frmAccountPanel.ListView1.ListItems
        For i = 1 To .Count
            If LCase(.Item(i).SubItems(1)) = LCase(GetUser) Then
                'Ban Check
                If .Item(i).SubItems(5) = "True" Then
                    SendSingle "!login#Banned#", frmMain.Winsock1(Index)
                    CMSG "!banned", GetUser
                    Exit Sub
                End If
                
                'Password Check
                If GetConver = .Item(i).SubItems(2) Then
                    'Send back confirmation
                    SendSingle "!login#Yes#", frmMain.Winsock1(Index)
                    CMSG GetCommand, GetUser, GetConver
                Else
                    SendSingle "!login#Password#", frmMain.Winsock1(Index)
                    CMSG "!password", GetUser
                End If
                Exit For
            Else
                If i = .Count Then
                    SendSingle "!login#Account#", frmMain.Winsock1(Index)
                    CMSG "!account", GetUser, GetConver
                End If
            End If
        Next i
    End With
    
'We get ip request and send ip back
Case "!iprequest"
    For i = 1 To frmPanel.ListView1.ListItems.Count
        If GetUser = frmPanel.ListView1.ListItems.Item(i) Then
            SendSingle "!iprequest#" & frmPanel.ListView1.ListItems.Item(i).SubItems(1) & "#", frmMain.Winsock1(Index)
        End If
    Next i
    
Case "!msg"
    Dim IsCommand   As Boolean
    Dim IsSlash     As Boolean
    Dim ANN_MSG     As String
    Dim Reason      As String
    
    'Split the conversation text by spaces
    array2 = Split(GetConver, " ")
        
    'Check first position of the text for a point indicating command
    If Left$(GetConver, 1) = "." Then
        If GetLevel(GetUser) <> 0 Then
            IsCommand = True
        End If
    End If
    
    'Check first position of the text for an slash indicating emote
    If Left$(GetConver, 1) = "/" Then
        IsSlash = True
    End If
    
    'Lets try to work without GoTo's
    If UBound(array2) > 0 Then
        GetTarget = StrConv(array2(1), vbProperCase)
        
        'Just capture the non propercased name if it's a command, emotes dont use accounts
        If IsCommand = True Then
            pGetTarget = array2(1)
            ANN_MSG = Mid$(GetConver, Len(array2(0)) + 2, Len(GetConver))
        End If
    End If
   
    'If a command is used check out which
    If IsCommand = True Then
        'Save the reason
        For i = 2 To UBound(array2)
            Reason = Reason & array2(i) & " "
        Next i
        
        Select Case array2(0)
        Case ".show"
            Select Case LCase$(GetTarget)
            Case "accounts", "account", "a"
                SendSingle "!split_text#" & GetAccountList, frmMain.Winsock1(Index)
            Case "users", "user", "u"
                SendSingle "!split_text#" & GetUserList, frmMain.Winsock1(Index)
            Case Else
                SendSingle "Incorrect Syntax. Use .show - account / user", frmMain.Winsock1(Index)
            End Select
            
        Case ".userinfo", ".uinfo"
            GetUserInfo GetTarget, Index
                        
        Case ".accountinfo", ".accinfo", ".ainfo"
            GetAccountInfo pGetTarget, Index
                        
        Case ".kick"
            KickUser GetTarget, Index
            
        Case ".banaccount"
            BanAccount pGetTarget, GetUser, True, Index, Trim$(Reason)
                        
        Case ".banuser"
            BanUser GetTarget, GetUser, True, Index, Trim$(Reason)
            
        Case ".unbanuser"
            BanUser GetTarget, GetUser, False, Index, Trim$(Reason)
                        
        Case ".unbanaccount"
            BanAccount pGetTarget, GetUser, False, Index, Trim$(Reason)
            
        Case ".mute"
            MuteUser GetTarget, GetUser, True, Index, Trim$(Reason)
                    
        Case ".unmute"
            MuteUser GetTarget, GetUser, False, Index, Trim$(Reason)
                    
        Case ".announce", ".ann", ".broadcast"
            If Len(Trim$(ANN_MSG)) = 0 Then
                SendSingle "Incorrect syntax. Use .help for more explanation.", frmMain.Winsock1(Index)
            Else
                SendMessage GetTag(GetUser) & "[" & GetUser & "] announces: " & ANN_MSG
            End If
            ANN_MSG = vbNullString
                    
        Case ".help", ".command", ".commands"
            SendSingle GetCommands, frmMain.Winsock1(Index)
                    
        Case Else
            SendSingle "Unknown command used. Check .help for more information about commands.", frmMain.Winsock1(Index)
                        
        End Select
        Exit Sub
    End If
    
    If IsSlash = True Then
        If Mute = True Then
            SendSingle "You are muted.", frmMain.Winsock1(Index)
            Exit Sub
        End If
        
        Dim IsUser As Boolean
            
        With frmPanel.ListView1.ListItems
            For i = 1 To .Count
                If GetTarget = .Item(i) Then
                    IsUser = True
                    Exit For
                End If
            Next i
        End With
            
        Select Case LCase(array2(0))
        Case "/lol", "/laugh"
            If IsUser = True Then
                SendMessage GetUser & " laughs at " & GetTarget & "."
            Else
                SendMessage GetUser & " laughs."
            End If
            
        Case "/rofl"
            If IsUser = True Then
                SendMessage GetUser & " rolls on the floor laughing at " & GetTarget & "."
            Else
                SendMessage GetUser & " rolls on the floor laughing."
            End If
            
        Case "/beer"
            SendMessage GetUser & " takes a beer from the fridge."
            
        Case "/fart"
            If IsUser = True Then
                SendMessage GetUser & " brushes up and farts loudly against " & GetTarget & "."
            Else
                SendMessage GetUser & " farts loudly."
            End If
            
        Case "/lmao"
            If IsUser = True Then
                SendMessage GetUser & " is laughing his / her ass off at " & GetTarget & "."
            Else
                SendMessage GetUser & " is laughing his / her ass off."
            End If
            
        Case "/facepalm"
            If IsUser = True Then
                SendMessage GetUser & " takes a look on " & GetTarget & " and covers his face with a palm."
            Else
                SendMessage GetUser & " covers his face with his palm."
            End If
            
        Case "/violin"
            If IsUser = True Then
                SendMessage GetUser & " looks at " & GetTarget & " and starts playing the world smallest violin."
            Else
                SendMessage GetUser & " plays the world smallest violin."
            End If
            
        Case "/insult"
            If IsUser = True Then
                SendMessage GetUser & " starts insulting " & GetTarget & "  heavily."
            Else
                SendMessage GetUser & " insults him / her-self as bitch."
            End If
            
        Case "/smile"
            If IsUser = True Then
                SendMessage GetUser & " smiles at " & GetTarget & "."
            Else
                SendMessage GetUser & " smiles."
            End If
            
        Case "/love", "/<3"
            If IsUser = True Then
                SendMessage GetUser & " loves " & GetTarget & "."
            Else
                SendMessage GetUser & " feels the love."
            End If
            
        Case "/cheer"
            If IsUser = True Then
                SendMessage GetUser & " cheers at " & GetTarget & "."
            Else
                SendMessage GetUser & " cheers."
            End If
            
        Case "/kiss"
            If IsUser = True Then
                SendMessage GetUser & " blows a kiss to " & GetTarget & "."
            Else
                SendMessage GetUser & " blows a kiss into the wind."
            End If
            
        Case "/hug"
            If IsUser = True Then
                SendMessage GetUser & " hugs " & GetTarget & "."
            Else
                SendMessage GetUser & " needs a hug!"
            End If
            
        Case "/facepalm", "/fp"
            If IsUser = True Then
                SendMessage GetUser & " covers his face with the palm."
            Else
                SendMessage GetUser & " looks at " & GetTarget & " and covers his face with the palm."
            End If
            
        Case "/w", "/whisper"
            'Whisper
            If IsUser = True Then
                Whisper GetUser, GetTarget, array2(2), Index
            Else
                If Len(Trim$(GetTarget)) = 0 Then
                    Exit Sub
                Else
                    SendSingle "No user named '" & GetTarget & "' is currently online.", frmMain.Winsock1(Index)
                End If
            End If
        
        Case "/afk"
            'Set AFK Flag
            With frmPanel.ListView1.ListItems
                For i = 1 To .Count
                    If .Item(i) = GetUser Then
                        If .Item(i).SubItems(7) = "True" Then
                            .Item(i).SubItems(7) = "False"
                        Else
                            .Item(i).SubItems(7) = "True"
                        End If
                        Exit For
                    End If
                Next i
                UpdateUsersList
            End With
        
        Case "/logout"
            KickUser GetUser, Index
        
        Case Else
            SendSingle "Unknown command used.", frmMain.Winsock1(Index)
                    
        End Select
        CMSG "!emote", GetUser, GetConver
        Exit Sub
    End If
    
    Dim S1&, E$
  
    If Len(GetConver) > 5 Then
        If IsNumeric(GetConver) = False Then
            If IsAlphaCharacter(GetConver) = True Then
                S1 = 0
                For i = 1 To Len(GetConver)
                    E = Mid$(GetConver, i, 1)
                    If UCase$(E) = E Then S1 = S1 + 1
                Next i
                E = vbNullString
                If Format$(100 * S1 / Len(GetConver), "0.00") > 75 Then
                    SendSingle "Message blocked. Please do not write more then 75% in caps.", frmMain.Winsock1(Index)
                    Exit Sub
                End If
            End If
        End If
    End If
        
    'Check if user is muted
    If Mute = True Then
        SendSingle "You are muted.", frmMain.Winsock1(Index)
        CMSG "!muted", GetUser, GetConver
        Exit Sub
    End If
    
    'Check if user is repeating
    If GetConver = GetLastMessage Then
        SendSingle "Your message has triggered serverside flood protection. Please don't repeat yourself.", frmMain.Winsock1(Index)
        CMSG "!repeat", GetUser
        Exit Sub
    End If
    
    'Send Message
    SendMessage "[" & GetUser & "]: " & GetConver
    CMSG GetCommand, GetUser, GetConver
    
    'Set last message
    With frmPanel.ListView1.ListItems
        For i = 1 To .Count
            If GetUser = .Item(i) Then
                frmPanel.ListView1.ListItems.Item(i).SubItems(3) = GetConver
            End If
        Next i
    End With
    
Case Else
    SendMessage "Unknown operation."
    
End Select
End Sub

Private Function GetServerInformation() As String
GetServerInformation = _
"Welcome to Peach Servers." & "#" & _
"Server: Peach " & Rev & " - Win32" & "#" & _
"Online User: " & frmMain.Winsock1.Count - 1 & "#" & _
frmConfig.Label2.Caption & "#"
End Function

Private Sub Whisper(User As String, Target As String, Conversation As String, Index As Integer)
'Check if user is whispering itself
Select Case User
Case Target, "<AFK>" & User
    SendSingle "You can't whisper yourself.", frmMain.Winsock1(Index)
    Exit Sub
End Select

'Search target in list and send message
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        Select Case Target
        Case .Item(i)
            SendSingle "[You whisper to " & Target & "]: " & Conversation, frmMain.Winsock1(Index)
            SendSingle "[" & User & " whispers]: " & Conversation, frmMain.Winsock1(.Item(i).SubItems(2))
            CMSG "!w", User, Conversation, Target
            Exit For
        Case "<AFK>" & .Item(i)
            SendSingle .Item(i) & " is away from keyboard.", frmMain.Winsock1(Index)
            SendSingle "[" & User & " whispers]: " & Conversation, frmMain.Winsock1(.Item(i).SubItems(2))
            CMSG "!w", User, Conversation, Target
            Exit For
        End Select
    Next i
End With
End Sub

Private Function GetTag(User As String) As String
Select Case GetLevel(User)
Case 1
    GetTag = "<GM> "
Case 2
    GetTag = "<Admin> "
End Select
End Function

Private Sub MuteUser(User As String, AdminName As String, Mute As Boolean, SIndex As Integer, Reason As String)
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = User Then
            
            If Mute = True Then
                'If the user is already muted then return feedback
                If .Item(i).SubItems(4) = "True" Then
                    SendSingle User & " is already muted.", frmMain.Winsock1(SIndex)
                    Exit Sub
                End If
            Else
                If .Item(i).SubItems(4) = "False" Then
                    SendSingle User & " is not muted.", frmMain.Winsock1(SIndex)
                    Exit Sub
                End If
            End If
        
            'Set flag in userlist
            .Item(i).SubItems(4) = CStr(Mute)
            
            'Announce the action
            If Len(Reason) = 0 Then
                If Mute = True Then
                    SendMessage User & " got muted by " & GetTag(AdminName) & AdminName & "."
                Else
                    SendMessage User & " got unmuted by " & GetTag(AdminName) & AdminName & "."
                End If
            Else
                If Mute = True Then
                    SendMessage User & " got muted by " & GetTag(AdminName) & AdminName & ". (" & Reason & ")"
                Else
                    SendMessage User & " got unmuted by " & GetTag(AdminName) & AdminName & ". (" & Reason & ")"
                End If
            End If
            
            Exit For
        Else
            If i = .Count Then
                If Len(Trim$(User)) = 0 Then
                    SendSingle "Incorrect syntax. Use .help for more explanation.", frmMain.Winsock1(SIndex)
                Else
                    SendSingle "User '" & User & "' was not found.", Winsock1(SIndex)
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

Private Sub BanUser(User As String, AdminName As String, Ban As Boolean, SIndex As Integer, Reason As String)
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = StrConv(User, vbProperCase) Then
            BanAccount .Item(i).SubItems(5), AdminName, Ban, SIndex, Reason
            Exit For
        Else
            If Len(Trim$(User)) = 0 Then
                If i = .Count Then SendSingle "Incorrect syntax. Use .help for more explanation.", frmMain.Winsock1(SIndex)
            Else
                If i = .Count Then SendSingle "User '" & User & "' not found.", Winsock1(SIndex)
            End If
        End If
    Next i
End With
End Sub

Private Sub BanAccount(Account As String, AdminName As String, Ban As Boolean, SIndex As Integer, Reason As String)
Dim User    As String
Dim j       As Long

With frmAccountPanel.ListView1.ListItems
    For i = 1 To .Count
        If LCase(.Item(i).SubItems(1)) = LCase(Account) Then
            'If the account is already banned send feedback
            If Ban = True Then
                If .Item(i).SubItems(5) = "True" Then
                    SendSingle "Account '" & Account & "' is already banned.", frmMain.Winsock1(SIndex)
                    Exit Sub
                End If
            Else
                If .Item(i).SubItems(5) = "False" Then
                    SendSingle "Account '" & Account & "' is not banned.", frmMain.Winsock1(SIndex)
                    Exit Sub
                End If
            End If
        
            'Ban account in database
            frmAccountPanel.ModifyAccount Account, .Item(i).SubItems(2), Ban, .Item(i).SubItems(6), .Item(i), .Item(i).Index
            
            'Determine user from account
            For j = 1 To frmPanel.ListView1.ListItems.Count
                If LCase(frmPanel.ListView1.ListItems.Item(j).SubItems(5)) = LCase(Account) Then
                    User = frmPanel.ListView1.ListItems.Item(j)
                    Exit For
                End If
            Next j
            
            'Announce the action
            If Len(Trim$(Reason)) = 0 Then
                If Ban = True Then
                    SendMessage User & " was account banned by " & GetTag(AdminName) & AdminName & "."
                Else
                    SendMessage User & " was unbanned by " & GetTag(AdminName) & AdminName & "."
                End If
            Else
                If Ban = True Then
                    SendMessage User & " was account banned by " & GetTag(AdminName) & AdminName & ". (" & Reason & ")"
                Else
                    SendMessage User & " was account unbanned by " & GetTag(AdminName) & AdminName & ". (" & Reason & ")"
                End If
            End If
            Exit For
        Else
            If i = .Count Then
                If Len(Trim$(Account)) = 0 Then
                    SendSingle "Incorrect syntax. Use .help for more explanation.", frmMain.Winsock1(SIndex)
                Else
                    SendSingle "Account '" & Account & "' not found.", Winsock1(SIndex)
                End If
            End If
        End If
    Next i
End With
End Sub

Private Sub KickUser(User As String, SIndex As Integer)
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = User Then
            'Disconnect and unload the socket
            frmMain.Winsock1(.Item(i).SubItems(2)).Close
            Unload frmMain.Winsock1(.Item(i).SubItems(2))
            
            'Remove from userlist
            .Remove (i)
            
            'Update userlist and statusbar
            UpdateUsersList
            frmMain.StatusBar1.Panels(1).Text = "Status: Connected with " & frmMain.Winsock1.Count - 1 & " Client(s)."
            
            Exit For
        Else
            If Len(Trim$(User)) = 0 Then
                If i = .Count Then SendSingle "Incorrect syntax. Use .help for more explanation.", frmMain.Winsock1(SIndex)
            Else
                If i = .Count Then SendSingle "User '" & User & "' not found.", Winsock1(SIndex)
            End If
        End If
    Next i
End With
End Sub

Private Sub GetAccountInfo(Account As String, SIndex As Integer)
With frmAccountPanel.ListView1.ListItems
    For i = 1 To .Count
        If LCase(.Item(i).SubItems(1)) = LCase(Account) Then
            SendSingle vbCrLf & " Account information about '" & Account & "'" & vbCrLf & " Name: " & .Item(i).SubItems(1) & vbCrLf & " Password: " & .Item(i).SubItems(2) & vbCrLf & " Registration Time: " & .Item(i).SubItems(3) & vbCrLf & " Registration Date: " & .Item(i).SubItems(4) & vbCrLf & " Banned: " & .Item(i).SubItems(5) & vbCrLf & " Level: " & .Item(i).SubItems(6), Winsock1(SIndex)
            Exit For
        Else
            If Len(Trim$(Account)) = 0 Then
                If i = .Count Then SendSingle "Incorrect syntax. Use .help for more explanation.", frmMain.Winsock1(SIndex)
            Else
                If i = .Count Then SendSingle "Account '" & Account & "' not found.", Winsock1(SIndex)
            End If
        End If
    Next i
End With
End Sub

Private Sub GetUserInfo(User As String, SIndex As Integer)
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = StrConv(User, vbProperCase) Then
            SendSingle vbCrLf & "User information about '" & User & "'" & vbCrLf & " IP : " & .Item(i).SubItems(1) & vbCrLf & " Winsock ID: " & .Item(i).SubItems(2) & vbCrLf & " Last Message: " & .Item(i).SubItems(3) & vbCrLf & " Muted: " & .Item(i).SubItems(4) & vbCrLf & " Account: " & .Item(i).SubItems(5) & vbCrLf & " Login Time: " & .Item(i).SubItems(6) & vbCrLf & " AFK: " & .Item(i).SubItems(7), Winsock1(SIndex)
            Exit For
        Else
            If Len(Trim$(User)) = 0 Then
                If i = .Count Then SendSingle "Incorrect syntax. Use .help for more explanation.", frmMain.Winsock1(SIndex)
            Else
                If i = .Count Then SendSingle "User '" & User & " was not found.", Winsock1(SIndex)
            End If
        End If
    Next i
End With
End Sub

Private Function GetCommands() As String
GetCommands = vbCrLf & _
"*********************************************" & vbCrLf & _
"* List of all avaible commands:" & vbCrLf & _
"* .announce 'Text' ( Send an server side tagged announced )" & vbCrLf & _
"* .banuser 'Name' ( Bans users account )" & vbCrLf & _
"* .banaccount 'Account' ( Bans the account )" & vbCrLf & _
"* .unbanuser 'Name' ( Removes ban from 'Name' )" & vbCrLf & _
"* .unbanaccount 'Account' ( Removes ban from 'Account )" & vbCrLf & _
"* .kick 'Name' ( Kicks 'Name' from Server )" & vbCrLf & _
"* .mute 'Name' ( Mutes 'Name' until unmute )" & vbCrLf & _
"* .unmute 'Name' ( Removes mute from 'Name' )" & vbCrLf & _
"* .userinfo 'Name' ( Shows all information about 'Name' )" & vbCrLf & _
"* .accountinfo / .accinfo ( Shows all information about that account )" & vbCrLf & _
"* .show accounts / users ( Shows a list of all accounts / user )" & vbCrLf & _
"* .help ( shows this list of all avaible commands )" & vbCrLf & _
"*********************************************"
End Function

Private Function GetLevel(Name As String) As Long
Dim GetAccount As String
For i = 1 To frmPanel.ListView1.ListItems.Count
    If Name = frmPanel.ListView1.ListItems.Item(i) Then
        GetAccount = frmPanel.ListView1.ListItems.Item(i).SubItems(5)
        Exit For
    End If
Next i
For i = 1 To frmAccountPanel.ListView1.ListItems.Count
    If GetAccount = frmAccountPanel.ListView1.ListItems.Item(i).SubItems(1) Then
        GetLevel = frmAccountPanel.ListView1.ListItems.Item(i).SubItems(6)
        Exit For
    End If
Next i
End Function

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1(Index).Close
Unload Winsock1(Index)

CMSG "!disconnected"

StatusBar1.Panels(1).Text = "[System]: Disconnected due connection problem."

frmPanel.ListView1.ListItems.Clear

StatusBar1.Panels(1).Text = "Status: Connected with " & Winsock1.Count - 1 & " Client(s)."
End Sub

Public Sub DisableFormResize(frm As Form)
Dim style As Long
Dim hMenu As Long
Dim MII As MENUITEMINFO
Dim lngMenuID As Long
Const xSC_MAXIMIZE As Long = -11

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
