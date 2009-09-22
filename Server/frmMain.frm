VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00F4F4F4&
   Caption         =   " Peach (Server)"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7470
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
      Width           =   7470
      _ExtentX        =   13176
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
      BackColor       =   &H00F4F4F4&
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
      ScaleHeight     =   555
      ScaleWidth      =   7410
      TabIndex        =   0
      Top             =   0
      Width           =   7470
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
Dim MAX_CONNECTION  As Long
Dim Vali            As Boolean
Dim Acc             As Boolean
Dim Muted           As Boolean
Dim Avaible         As Boolean

'Database variables
Dim INI_DATABASE            As String
Dim INI_USER                As String
Dim INI_PASSWORD            As String
Dim INI_IP                  As String
Dim INI_ACCOUNT_TABLE       As String
Dim INI_FRIENDS_TABLE       As String

Public HasError             As Boolean

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

'Assign Variables to the .ini strings
If Len(ReadIniValue(App.Path & "\Config.ini", "Database", "Database")) <> 0 Then
    INI_DATABASE = ReadIniValue(App.Path & "\Config.ini", "Database", "Database")
    WriteLog "Database loaded. (" & INI_DATABASE & ")"
Else
    WriteLog "No Database value found."
End If

If Len(ReadIniValue(App.Path & "\Config.ini", "Database", "User")) <> 0 Then
    INI_USER = ReadIniValue(App.Path & "\Config.ini", "Database", "User")
    WriteLog "User loaded. (" & INI_USER & ")"
Else
    WriteLog "No User value found."
End If

If Len(DeCode(ReadIniValue(App.Path & "\Config.ini", "Database", "Password"))) <> 0 Then
    INI_PASSWORD = DeCode(ReadIniValue(App.Path & "\Config.ini", "Database", "Password"))
    WriteLog "Password loaded."
Else
    WriteLog "No Password value found."
    INI_PASSWORD = InputBox("Enter your MySQL Password", "MySQL Password")
End If

If Len(ReadIniValue(App.Path & "\Config.ini", "Database", "IP")) <> 0 Then
    INI_IP = ReadIniValue(App.Path & "\Config.ini", "Database", "IP")
    WriteLog "IP loaded. (" & INI_IP & ")"
Else
    WriteLog "No IP value found."
End If

If Len(ReadIniValue(App.Path & "\Config.ini", "Database", "A_Table")) <> 0 Then
    INI_ACCOUNT_TABLE = ReadIniValue(App.Path & "\Config.ini", "Database", "A_Table")
    WriteLog "Account-Table loaded. (" & INI_ACCOUNT_TABLE & ")"
Else
    WriteLog "No Account-Table found."
End If

If Len(ReadIniValue(App.Path & "\Config.ini", "Database", "F_Table")) <> 0 Then
    INI_FRIENDS_TABLE = ReadIniValue(App.Path & "\Config.ini", "Database", "F_Table")
    WriteLog "Friends-Table loaded. (" & INI_FRIENDS_TABLE & ")"
Else
    WriteLog "No Friends-Table found."
End If

'Connect to MySQL Database
ConnectMySQL INI_DATABASE, INI_USER, INI_PASSWORD, INI_IP

'Load Accounts
LoadAccounts INI_ACCOUNT_TABLE

'Load Friends
LoadFriends INI_FRIENDS_TABLE

StatusBar1.Panels(1).Text = "Status : Disconnected"
SetupForms frmConfig

If HasError = False Then
    WriteLog "Correctly loaded!"
End If
End Sub

Public Sub ConnectMySQL(qDatabase As String, qUser As String, qPassword As String, qIP As String)
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
WriteLog "[MySQL] Connected with Database"

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

If HasError = True Then Exit Sub
If Len(qTable) = 0 Then Exit Sub

SQL = "SELECT * FROM " & qTable

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
    Loop
End With

Set LItem = Nothing
Set xRecordSet = Nothing

WriteLog "[MySQL] Loaded table '" & qTable & "' data -> Friends."

Exit Sub
HandleErrorFriends:
'Print error
WriteLog Err.Description & "."

'Set error flag
HasError = True
End Sub

Private Sub LoadAccounts(qTable As String)
Dim SQL         As String
Dim LItem   As ListItem

If HasError = True Then Exit Sub
If Len(qTable) = 0 Then Exit Sub

SQL = "SELECT * FROM " & qTable

StatusBar1.Panels(1).Text = "Status : Loading accounts .."

On Error GoTo HandleErrorTable
xCommand.CommandText = SQL
Set xRecordSet = xCommand.Execute

With frmAccountPanel
    .ListView1.ListItems.Clear
    With .cmbBanned
        .Clear
        .AddItem "Yes"
        .AddItem "No"
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
    Loop
End With

Set LItem = Nothing
Set xRecordSet = Nothing

WriteLog "[MySQL] Loaded table '" & qTable & "' data -> Accounts."

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
WriteIniValue App.Path & "\Config.ini", "Database", "Database", INI_DATABASE
WriteIniValue App.Path & "\Config.ini", "Database", "User", INI_USER
WriteIniValue App.Path & "\Config.ini", "Database", "Password", INI_PASSWORD
WriteIniValue App.Path & "\Config.ini", "Database", "IP", INI_IP
WriteIniValue App.Path & "\Config.ini", "Database", "A_Table", INI_ACCOUNT_TABLE
WriteIniValue App.Path & "\Config.ini", "Database", "F_Table", INI_FRIENDS_TABLE
End Sub

Private Sub Winsock1_Close(Index As Integer)
Unload Winsock1(Index)
For i = 1 To frmPanel.ListView1.ListItems.Count + 1
    'Update user lists ( server and client )
    If frmPanel.ListView1.ListItems.Item(i).SubItems(2) = Index Then
    
        'Pick the user
        frmPanel.ListView1.ListItems.Remove (i)
        
        'Update Users List
        UpdateUsersList
        Exit For
    End If
Next i
StatusBar1.Panels(1).Text = "Status: Connected with " & Winsock1.Count - 1 & " Client(s)."
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
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
    .Item(i).SubItems(4) = "No"
    .Item(i).SubItems(5) = vbNullString
    .Item(i).SubItems(6) = Format$(Time, "hh:mm:ss")
    .Item(i).SubItems(7) = "No"
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
Dim GetTarget       As String
Dim pGetTarget      As String
Dim GetLastMessage  As String
Dim bMatch, Mute    As Boolean

'Get Message
frmMain.Winsock1(Index).GetData GetMessage

'We decode (split) the message into an array
array1 = Split(GetMessage, "#")

'Assign the variables to the array
Command = array1(0)
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
            If .Item(i).SubItems(4) = "Yes" Then
                Mute = True
            End If
        End If
    Next i
End With

'Validate: If message is to long then kick
If Len(GetConver) > 200 Then
    SMSG "!long", GetUser
    frmPanel.ListView1.ListItems.Remove (Index) ' Remove from list
    Winsock1(Index).Close 'Close connection
    Unload Winsock1(Index) 'Remove socket
    StatusBar1.Panels(1).Text = "Status: Connected with " & Winsock1.Count - 1 & " Client(s)."
    UpdateUsersList
    Exit Sub
End If

Select Case Command
'Update Friend list
Case "!get_friends"
    UpdateFriendList GetUser
    
'Check if friends exist and save
Case "!add_friend"
    frmFriendList.AddFriend GetUser, GetConver, Index

'Remove friend from list
Case "!remove_friend"
    frmFriendList.RemoveFriend GetUser, GetConver, Index
   
'Add Name & Account to frmPanel ListView
Case "!connected"
    'If the account is already beeing used kick first instance
    With frmPanel.ListView1.ListItems
        For i = 1 To .Count
            If .Item(i).SubItems(5) = GetConver Then
                SMSG "!dinstance", GetConver
                Winsock1(.Item(i).SubItems(2)).Close
                Unload Winsock1(.Item(i).SubItems(2))
                .Remove (i)
                
                StatusBar1.Panels(1).Text = "Status: Connected with " & Winsock1.Count - 1 & " Client(s)."
                Exit For
            End If
        Next i
    End With
        
    SMSG Command, GetUser
    With frmPanel.ListView1.ListItems
        .Item(.Count).Text = GetUser
        .Item(.Count).SubItems(5) = GetConver
    End With
    UpdateUsersList

Case "!namerequest"
    SMSG Command, GetUser
    'Check badname list
    For i = 0 To frmPanel.List1.ListCount
        If frmPanel.List1.List(i) = GetUser Then
            bMatch = True
            SMSG "!badname", GetUser
            Exit For
        End If
    Next i
    
    'Check current online list
    With frmPanel.ListView1.ListItems
        For i = 1 To .Count
            If .Item(i) = GetUser Then
                bMatch = True
                SMSG "!nametaken", GetUser
                Exit For
            End If
        Next i
    End With
    
    'Return answer to client
    If bMatch = True Then
        SendSingle "!decilined", frmMain.Winsock1(Index)
    Else
        SendSingle "!accepted", frmMain.Winsock1(Index)
        SMSG "!nameisfree", GetUser
    End If
    
Case "!w"
    If Mute = True Then
        SendSingle " You are muted.", frmMain.Winsock1(Index)
        SMSG "!wmuted", GetUser, GetConver
    Else
        'Split name into user name and normal name
        array2 = Split(GetUser, "|")
        GetUser = array2(0)
        ForWho = array2(1)
        
        'Check if user is whispering itself
        Select Case GetUser
        Case ForWho
            SendSingle " You can't whisper yourself.", frmMain.Winsock1(Index)
            Exit Sub
        Case "<AFK>" & GetUser
            SendSingle " You can't whisper yourself.", frmMain.Winsock1(Index)
            Exit Sub
        End Select
        
        'Check in listitems if forwho name is in the list and get the socket id
        With frmPanel.ListView1.ListItems
            For i = 1 To .Count
                Select Case ForWho
                Case .Item(i)
                    SendSingle " [You whisper to " & ForWho & "]: " & GetConver, frmMain.Winsock1(Index)
                    SendSingle " [" & GetUser & " whispers]: " & GetConver, frmMain.Winsock1(.Item(i).SubItems(2))
                    SMSG Command, GetUser, GetConver, ForWho
                    Exit For
                Case "<AFK>" & .Item(i)
                    SendSingle " " & .Item(i) & " is away from keyboard.", frmMain.Winsock1(Index)
                    SendSingle " [" & GetUser & " whispers]: " & GetConver, frmMain.Winsock1(.Item(i).SubItems(2))
                    SMSG Command, GetUser, GetConver, ForWho
                    Exit For
                End Select
            Next i
        End With
    End If
    
Case "!login"
    With frmAccountPanel.ListView1.ListItems
        For i = 1 To .Count
            Select Case GetUser
            Case .Item(i).SubItems(1)
                
                'Ban Check
                If .Item(i).SubItems(5) = "Yes" Then
                    SendSingle "!login#Banned#", frmMain.Winsock1(Index)
                    SMSG "!banned", GetUser
                    Exit Sub
                End If
                
                'Password Check
                If GetConver = .Item(i).SubItems(2) Then
                    'Send back confirmation and account level
                    SendSingle "!login#Yes#", frmMain.Winsock1(Index)
                    SMSG Command, GetUser, GetConver
                Else
                    SendSingle "!login#Password#", frmMain.Winsock1(Index)
                    SMSG "!password", GetUser
                End If
                
                Acc = True
                Exit For
            Case Else
                Acc = False
            End Select
        Next i
    End With
    
    If Acc = False Then
        SendSingle "!login#Account#", frmMain.Winsock1(Index)
        SMSG "!account", GetUser, GetConver
    End If
        
Case "!iprequest"
    For i = 1 To frmPanel.ListView1.ListItems.Count
        If GetUser = frmPanel.ListView1.ListItems.Item(i) Then
            SendSingle "!iprequest#" & frmPanel.ListView1.ListItems.Item(i).SubItems(1) & "#", frmMain.Winsock1(Index)
        End If
    Next i
    
Case "!msg"
    Dim IsCommand   As Boolean
    Dim IsEmote     As Boolean
    
    'Split the conversation text by spaces
    array2 = Split(GetConver, " ")
        
    'Check first position of the text for an point indicating command
    If Left$(GetConver, 1) = "." Then
        If GetLevel(GetUser) <> 0 Then
            IsCommand = True
        End If
    End If
    
    'Check first position of the text for an slash indicating emote
    If Left$(GetConver, 1) = "/" Then
        IsEmote = True
    End If
    
    Dim ANN_MSG As String
    On Error GoTo TargetErrorHandler
    GetTarget = StrConv(array2(1), vbProperCase)
    pGetTarget = array2(1) 'Account, not propercased
    ANN_MSG = Mid$(GetConver, Len(array2(0)) + 2, Len(GetConver))

Commands:
    
    'If an command is used check out which
    If IsCommand = True Then
        Select Case array2(0)
        Case ".show"
            Select Case LCase$(GetTarget)
            Case "accounts"
                SendSingle "!accountlist#" & GetAccountList, frmMain.Winsock1(Index)
            Case "users"
                SendSingle "!userlist#" & GetUserList, frmMain.Winsock1(Index)
            Case Else
                SendSingle " You can just use .list account or user.", frmMain.Winsock1(Index)
            End Select
            
        Case ".userinfo", ".uinfo"
            GetUserInfo GetTarget, Index
                        
        Case ".accountinfo", ".accinfo", ".ainfo"
            GetAccountInfo pGetTarget, Index
                        
        Case ".kick"
            KickUser GetTarget, Index
                        
        Case ".banaccount"
            BanAccount pGetTarget, GetUser, "Yes", Index
                        
        Case ".banuser"
            BanUser GetTarget, GetUser, "Yes", Index
                        
        Case ".unbanuser"
            BanUser GetTarget, GetUser, "No", Index
                        
        Case ".unbanaccount"
            BanAccount pGetTarget, GetUser, "No", Index
                        
        Case ".mute"
            MuteUser GetTarget, GetUser, "Yes", Index
            
        Case ".unmute"
            MuteUser GetTarget, GetUser, "No", Index
            
        Case ".announce", ".ann", ".broadcast"
            CMSG GetUser, ANN_MSG
                    
        Case ".help", ".command", ".commands"
            SendSingle GetCommands, frmMain.Winsock1(Index)
            
        Case Else
            SendSingle " Unknown command used. Check .help for more information about commands.", frmMain.Winsock1(Index)
                        
        End Select
        Exit Sub
    End If
    
Emotes:
    If IsEmote = True Then
        If Mute = False Then
            Dim IsUser As Boolean
            
            With frmPanel.ListView1.ListItems
                For i = 1 To .Count
                    If GetTarget = .Item(i) Then
                        IsUser = True
                        Exit For
                    End If
                Next i
            End With
            
            If Len(GetTarget) = 0 Then
                Select Case LCase(array2(0))
                Case "/lol", "/laugh"
                    SendMessage " " & GetUser & " laughs."
                Case "/rofl"
                    SendMessage " " & GetUser & " rolls on the floor laughing."
                Case "/beer"
                    SendMessage " " & GetUser & " takes a beer from the fridge."
                Case "/fart"
                    SendMessage " " & GetUser & " farts loudly."
                Case "/lmao"
                    SendMessage " " & GetUser & " is laughing his / her ass off."
                Case "/facepalm"
                    SendMessage " " & GetUser & " covers his face with his palm."
                Case "/violin"
                    SendMessage " " & GetUser & " plays the world smallest violin."
                Case "/insult"
                    SendMessage " " & GetUser & " insults him / her-self as bitch."
                Case "/smile"
                    SendMessage " " & GetUser & " smiles."
                Case "/love", "/<3"
                    SendMessage " " & GetUser & " feels the love."
                Case "/cheer"
                    SendMessage " " & GetUser & " cheers."
                Case "/kiss"
                    SendMessage " " & GetUser & " blows a kiss into the wind."
                Case "/afk"
                    'Set afk flag
                    With frmPanel.ListView1.ListItems
                        For i = 1 To .Count
                            If .Item(i) = GetUser Then
                                If .Item(i).SubItems(7) = "Yes" Then
                                    .Item(i).SubItems(7) = "No"
                                Else
                                    .Item(i).SubItems(7) = "Yes"
                                End If
                                Exit For
                            End If
                        Next i
                        UpdateUsersList
                        Exit Sub
                    End With
                Case Else
                    GoTo Message
                End Select
            Else
                If IsUser = True Then
                    Select Case LCase(array2(0))
                    Case "/lol"
                        SendMessage " " & GetUser & " laughs at " & GetTarget & "."
                    Case "/rofl"
                        SendMessage " " & GetUser & " rolls on the floor laughing at " & GetTarget & "."
                    Case "/beer"
                        SendMessage " " & GetUser & " takes a beer from the fridge."
                    Case "/fart"
                        SendMessage " " & GetUser & " brushes up and farts loudly against " & GetTarget & "."
                    Case "/lmao"
                        SendMessage " " & GetUser & " is laughing his / her ass off at " & GetTarget & "."
                    Case "/facepalm"
                        SendMessage " " & GetUser & " takes a look on " & GetTarget & " and covers his face with a palm."
                    Case "/violin"
                        SendMessage " " & GetUser & " looks at " & GetTarget & " and starts playing the world smallest violin."
                    Case "/insult"
                        SendMessage " " & GetUser & " starts insulting " & GetTarget & "  heavily."
                    Case "/smile"
                        SendMessage " " & GetUser & " smiles at " & GetTarget & "."
                    Case "/love"
                        SendMessage " " & GetUser & " loves " & GetTarget & "."
                    Case "/cheer"
                        SendMessage " " & GetUser & " cheers at " & GetTarget & "."
                    Case "/kiss"
                        SendMessage " " & GetUser & " blows a kiss to " & GetTarget & "."
                    End Select
                Else
                    SendSingle " User '" & GetTarget & "' not found.", frmMain.Winsock1(Index)
                End If
            End If
            SMSG "!emote", GetUser, GetConver
        End If
        Exit Sub
    End If
            
Message:
    'Check if user is muted
    If Mute = True Then
        SendSingle " You are muted.", frmMain.Winsock1(Index)
        SMSG "!muted", GetUser, GetConver
        Exit Sub
    End If
    
    'Check if user is repeating
    If GetConver = GetLastMessage Then
        SendSingle " Your message has triggered serverside flood protection. Please don't repeat yourself.", frmMain.Winsock1(Index)
        SMSG "!repeat", GetUser, ""
        Exit Sub
    End If
        
    'Send Message
    SendMessage " [" & GetUser & "]: " & GetConver
    SMSG Command, GetUser, GetConver
    
Case Else
    SendMessage " Unknown operation."

End Select
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If GetUser = .Item(i) Then
            frmPanel.ListView1.ListItems.Item(i).SubItems(3) = GetConver
        End If
    Next i
End With
Exit Sub
TargetErrorHandler:
    Select Case Err.Number
    Case 9
        If IsCommand = True Then
            Select Case array2(0)
            Case _
                ".userinfo", _
                ".uinfo", _
                ".accountinfo", _
                ".accinfo", _
                ".ainfo", _
                ".kick", _
                ".banaccount", _
                ".unbanaccount", _
                ".list", _
                ".banuser", _
                ".unbanuser", _
                ".mute", _
                ".unmute"
                
                SendSingle " Incorrect Syntax. Use .help for more information about commands.", frmMain.Winsock1(Index)
        
            Case _
                ".help", _
                ".command", _
                ".commands"
                            
                GoTo Commands
                
            Case Else
                SendSingle " Unknown command used. Check .help for more information about commands.", frmMain.Winsock1(Index)
                
            End Select
        Else
            If IsEmote = True Then
                GoTo Emotes
            Else
                GoTo Message
            End If
        End If
    End Select
End Sub

Private Sub CMSG(pName As String, pMessage As String)
If GetLevel(pName) = 1 Then
    SendMessage " " & "<GM> [" & pName & "] announces: " & pMessage
Else
    SendMessage " " & "<Admin> [" & pName & "] announces: " & pMessage
End If
End Sub

Private Sub MuteUser(User As String, AdminName As String, Mute As String, SIndex As Integer)
User = StrConv(User, vbProperCase)
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = User Then
            .Item(i).SubItems(4) = Mute
            Avaible = True
            Exit For
        Else
            Avaible = False
        End If
    Next i
End With

If Avaible = False Then
    SendSingle " User '" & User & "' was not found.", Winsock1(SIndex)
Else
    If Mute = "Yes" Then
        SendMessage " " & User & " got muted by " & AdminName & "."
    Else
        SendMessage " " & User & " got unmuted by " & AdminName & "."
    End If
End If
End Sub

Private Function GetAccountList() As String
With frmAccountPanel.ListView1.ListItems
    For i = 1 To .Count
        GetAccountList = GetAccountList & .Item(i).SubItems(1) & " "
    Next i
End With
End Function

Private Function GetUserList() As String
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        GetUserList = GetUserList & .Item(i) & " "
    Next i
End With
End Function

Private Sub BanUser(User As String, AdminName As String, Ban As String, SIndex As Integer)
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = StrConv(User, vbProperCase) Then
            BanAccount .Item(i).SubItems(5), AdminName, Ban
            Avaible = True
            Exit For
        Else
            Avaible = False
        End If
    Next i
End With

If Avaible = False Then
    SendSingle " User '" & User & "' not found.", Winsock1(SIndex)
End If
End Sub

Private Sub BanAccount(Account As String, AdminName As String, Ban As String, Optional SIndex As Integer)
Dim User As String

'Ban account in database
With frmAccountPanel.ListView1.ListItems
    For i = 1 To .Count
        If LCase(.Item(i).SubItems(1)) = LCase(Account) Then
            frmAccountPanel.ModifyAccount Account, .Item(i).SubItems(2), Ban, .Item(i).SubItems(6), .Item(i), .Item(i).Index
            Avaible = True
            Exit For
        Else
            Avaible = False
        End If
    Next i
End With

'Find the user in the list by account
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If LCase(.Item(i).SubItems(5)) = LCase(Account) Then
            User = .Item(i)
            Exit For
        End If
    Next i
End With

'Announce or give feedback
If Avaible = False Then
    SendSingle " Account '" & Account & "' not found.", Winsock1(SIndex)
Else
    If Ban = "Yes" Then
        SendMessage " " & User & " was account banned by " & AdminName & "."
    Else
        SendMessage " " & User & " was unbanned by " & AdminName & "."
    End If
End If
End Sub

Private Sub KickUser(User As String, SIndex As Integer)
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = StrConv(User, vbProperCase) Then
            frmMain.Winsock1(.Item(i).SubItems(2)).Close
            Unload frmMain.Winsock1(.Item(i).SubItems(2))
            .Remove (i)
            Avaible = True
            Exit For
        Else
            Avaible = False
        End If
    Next i
End With

'If the user was found and kicked then update the users list else advice the Admin that user was not found
If Avaible = True Then
    UpdateUsersList
Else
    SendSingle " User '" & User & "' not found.", Winsock1(SIndex)
End If

frmMain.StatusBar1.Panels(1).Text = "Status: Connected with " & frmMain.Winsock1.Count - 1 & " Client(s)."
        
End Sub

Private Sub GetAccountInfo(Account As String, SIndex As Integer)
With frmAccountPanel.ListView1.ListItems
    For i = 1 To .Count
        If LCase(.Item(i).SubItems(1)) = LCase(Account) Then
            SendSingle vbCrLf & " Account information about '" & Account & "'" & vbCrLf & " Name: " & .Item(i).SubItems(1) & vbCrLf & " Password: " & .Item(i).SubItems(2) & vbCrLf & " Registration Time: " & .Item(i).SubItems(3) & vbCrLf & " Registration Date: " & .Item(i).SubItems(4) & vbCrLf & " Banned: " & .Item(i).SubItems(5) & vbCrLf & " Level: " & .Item(i).SubItems(6), Winsock1(SIndex)
            Avaible = True
            Exit For
        Else
            Avaible = False
        End If
    Next i
End With
If Avaible = False Then
    SendSingle " Account '" & Account & "' not found.", Winsock1(SIndex)
End If
End Sub

Private Sub GetUserInfo(User As String, SIndex As Integer)
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = StrConv(User, vbProperCase) Then
            SendSingle vbCrLf & " User information about '" & User & "'" & vbCrLf & " IP : " & .Item(i).SubItems(1) & vbCrLf & " Winsock ID: " & .Item(i).SubItems(2) & vbCrLf & " Last Message: " & .Item(i).SubItems(3) & vbCrLf & " Muted: " & .Item(i).SubItems(4) & vbCrLf & " Account: " & .Item(i).SubItems(5) & vbCrLf & " Login Time: " & .Item(i).SubItems(6) & vbCrLf & " AFK: " & .Item(i).SubItems(7), Winsock1(SIndex)
            Avaible = True
            Exit For
        Else
            Avaible = False
        End If
    Next i
End With

If Avaible = False Then
    SendSingle " User '" & User & " was not found.", Winsock1(SIndex)
End If
End Sub

Private Function GetCommands() As String
GetCommands = vbCrLf & _
" *********************************************" & vbCrLf & _
" * List of all avaible commands:" & vbCrLf & _
" * .announce 'Text' ( Send an server side tagged announced )" & vbCrLf & _
" * .banuser 'Name' ( Bans users account )" & vbCrLf & _
" * .banaccount 'Account' ( Bans the account )" & vbCrLf & _
" * .unbanuser 'Name' ( Removes ban from 'Name' )" & vbCrLf & _
" * .unbanaccount 'Account' ( Removes ban from 'Account )" & vbCrLf & _
" * .kick 'Name' ( Kicks 'Name' from Server )" & vbCrLf & _
" * .mute 'Name' ( Mutes 'Name' until unmute )" & vbCrLf & _
" * .unmute 'Name' ( Removes mute from 'Name' )" & vbCrLf & _
" * .userinfo 'Name' ( Shows all information about 'Name' )" & vbCrLf & _
" * .accountinfo / .accinfo ( Shows all information about that account )" & vbCrLf & _
" * .show accounts / users ( Shows a list of all accounts / user )" & vbCrLf & _
" * .help ( shows this list of all avaible commands )" & vbCrLf & _
" *********************************************"
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
Prefix = "[" & Format(Time, "hh:nn:ss") & "]"

Winsock1(Index).Close
Unload Winsock1(Index)

frmChat.txtConver.Text = frmChat.txtConver.Text & vbCrLf & Prefix & "[System]: Disconnected due connection problem."
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
