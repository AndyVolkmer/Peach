VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00F4F4F4&
   Caption         =   " Peach (Server)"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7545
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
      Width           =   7545
      _ExtentX        =   13309
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
      ScaleWidth      =   7485
      TabIndex        =   0
      Top             =   0
      Width           =   7545
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
         Left            =   5520
         TabIndex        =   4
         Top             =   120
         Width           =   1815
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
         Left            =   3720
         TabIndex        =   3
         Top             =   120
         Width           =   1815
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
         Left            =   1920
         TabIndex        =   2
         Top             =   120
         Width           =   1815
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
         Width           =   1815
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

Public xConnection     As ADODB.Connection
Public xCommand        As ADODB.Command
Public xRecordSet      As ADODB.Recordset

Public intCounter   As Integer
Dim i               As Integer 'Global "FOR" variable
Dim Vali            As Boolean
Dim Acc             As Boolean

Private Sub Command1_Click()
    SetupForms frmConfig
End Sub

Private Sub Command2_Click()
    SetupForms frmChat
End Sub

Public Sub SetupForms(Nix As Form)
    frmChat.Hide
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
    
ConnectToDB
LoadCustomerListView

StatusBar1.Panels(1).Text = "Status : Disconnected"
SetupForms frmConfig
End Sub

Private Sub ConnectToDB()
    
Set xConnection = New ADODB.Connection
xConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Accounts.mdb"
xConnection.Open

Set xCommand = New ADODB.Command
Set xCommand.ActiveConnection = xConnection
xCommand.CommandType = adCmdText

End Sub

Private Sub LoadCustomerListView()
                                 
Dim strSQL      As String
Dim xListItem   As ListItem
                                 
strSQL = "SELECT * FROM accounts"

xCommand.CommandText = strSQL
Set xRecordSet = xCommand.Execute

With frmAccountPanel
    .ListView1.ListItems.Clear
    .cmbBanned.Clear
    .cmbBanned.AddItem "Yes"
    .cmbBanned.AddItem "No"
    .cmbLevel.Clear
    .cmbLevel.AddItem "0"
    .cmbLevel.AddItem "1"
    .cmbLevel.AddItem "2"
End With

With xRecordSet
    Do Until .EOF
        Set xListItem = frmAccountPanel.ListView1.ListItems.Add(, , !id)
        xListItem.SubItems(1) = !Name1
        xListItem.SubItems(2) = !password1
        xListItem.SubItems(3) = !time1
        xListItem.SubItems(4) = !date1
        xListItem.SubItems(5) = !banned1
        xListItem.SubItems(6) = !level1
        .MoveNext
    Loop
End With

Set xListItem = Nothing
Set xRecordSet = Nothing

End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim msg As Long
Dim sFilter As String
msg = x / Screen.TwipsPerPixelX
Select Case msg
Case WM_LBUTTONDOWN
Case WM_LBUTTONUP
Case WM_LBUTTONDBLCLK
    Vali = True
    frmMain.Show ' show form
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

Private Sub Winsock1_Close(Index As Integer)
Dim x As Integer
Unload Winsock1(Index)
For x = 1 To frmPanel.ListView1.ListItems.Count + 1
    ' Update user lists ( server and client )
    If frmPanel.ListView1.ListItems.Item(x).SubItems(2) = Index Then
        ' Pick the user
        frmPanel.ListView1.ListItems.Remove (x)
        
        ' Update Users List
        UpdateUsersList
        Exit For
    End If
Next x
StatusBar1.Panels(1).Text = "Status: Connected with  " & Winsock1.Count - 1 & " Client(s)."
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim RR As Integer
RR = frmPanel.ListView1.ListItems.Count + 1
intCounter = loadSocket
Winsock1(intCounter).LocalPort = frmConfig.txtPort.Text
Winsock1(intCounter).Accept requestID

'New user should be listed in the panel
With frmPanel.ListView1
    .ListItems.Add RR, , "Unknown"
    .ListItems.Item(RR).SubItems(1) = Winsock1(intCounter).RemoteHostIP
    .ListItems.Item(RR).SubItems(2) = intCounter
    .ListItems.Item(RR).SubItems(3) = Format(Time, "hh:nn:ss")
    .ListItems.Item(RR).SubItems(4) = "No"
End With
StatusBar1.Panels(1).Text = "Status: Connected with  " & Winsock1.Count - 1 & " Client(s)."
End Sub

Private Function socketFree() As Integer
On Error GoTo HandleErrorFreeSocket
    Dim theIP As Variant
    Dim p As Integer
    For p = Winsock1.LBound + 1 To Winsock1.UBound
        theIP = Winsock1(p).LocalIP
    Next
    socketFree = Winsock1.UBound + 1
Exit Function
HandleErrorFreeSocket:
socketFree = p
End Function

Private Function loadSocket() As Integer
Dim theFreeSocket As Integer
theFreeSocket = 0
theFreeSocket = socketFree

Load Winsock1(theFreeSocket)

loadSocket = theFreeSocket
End Function

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim array1()        As String 'Message array
Dim array2()        As String 'Whisper & Command array
Dim strMessage      As String
Dim RR              As Integer
Dim bMatch          As Boolean
RR = frmPanel.ListView1.ListItems.Count

'Get Message
frmMain.Winsock1(Index).GetData strMessage

'We decode (split) the message into an array
array1 = Split(strMessage, "#")

'Assign the variables to the array
Command = array1(0)
GetUser = array1(1)
GetConver = array1(2)

'Check if user is muted
For i = 1 To frmPanel.ListView1.ListItems.Count
    If frmPanel.ListView1.ListItems.Item(i) = GetUser Then
        If frmPanel.ListView1.ListItems.Item(i).SubItems(4) = "Yes" Then
            SendSingle "!muted" & "#", frmMain.Winsock1(Index)
            Exit Sub
        End If
    End If
Next i

'Validate: If message is to long then kick
If Len(GetConver) > 200 Then
    frmPanel.ListView1.ListItems.Remove (Index) ' Remove from list
    Winsock1(Index).Close ' Close connection
    Unload Winsock1(Index) ' Remove socket
    StatusBar1.Panels(1).Text = "Status: Connected with  " & Winsock1.Count - 1 & " Client(s)."
    UpdateUsersList
End If

Select Case Command

'Announce connected player and send to user online list
Case "!connected"
    frmPanel.ListView1.ListItems.Item(RR).Text = GetUser
    frmPanel.ListView1.ListItems.Item(RR).SubItems(5) = GetConver
    UpdateUsersList
    
Case "!namerequest"

    'Check badname list
    For i = 0 To frmPanel.List1.ListCount
        If StrConv(frmPanel.List1.List(i), vbProperCase) = StrConv(GetUser, vbProperCase) Then
            bMatch = True
            Exit For
        End If
    Next i
    
    'Check current online list
    For i = 1 To frmPanel.ListView1.ListItems.Count
        If StrConv(frmPanel.ListView1.ListItems.Item(i), vbProperCase) = StrConv(GetUser, vbProperCase) Then
            bMatch = True
            Exit For
        End If
    Next i
    
    'Return answer to client
    If bMatch = True Then
        bMatch = False
        SendSingle "!decilined", frmMain.Winsock1(Index)
    Else
        SendSingle "!accepted", frmMain.Winsock1(Index)
    End If
    
Case "!w"
    'Split name into user name and normal name
    array2 = Split(GetUser, "|")
    GetUser = array2(0)
    ForWho = array2(1)
    
    'Check in listitems if forwho name is in the list and get the socket id
    For i = 1 To frmPanel.ListView1.ListItems.Count
        If ForWho = frmPanel.ListView1.ListItems.Item(i) Then
            SendSingle " [" & GetUser & "] whispers: " & GetConver, frmMain.Winsock1(frmPanel.ListView1.ListItems.Item(i).SubItems(2))
            VisualizeMessage Command, GetUser, GetConver, ForWho
            Exit For
        End If
    Next i
Case "!login"
    For i = 1 To frmAccountPanel.ListView1.ListItems.Count
        Select Case GetUser
        Case frmAccountPanel.ListView1.ListItems.Item(i).SubItems(1)
            
            'Ban Check
            Select Case frmAccountPanel.ListView1.ListItems.Item(i).SubItems(5)
            Case "Yes"
                SendSingle "!login" & "#" & "Banned" & "#", frmMain.Winsock1(Index)
                Exit For
            End Select
            
            'Password Check
            If GetConver = frmAccountPanel.ListView1.ListItems.Item(i).SubItems(2) Then
                'Send back confirmation and account level
                SendSingle "!login" & "#" & "Yes" & "#" & frmAccountPanel.ListView1.ListItems.Item(i).SubItems(6) & "#", frmMain.Winsock1(Index)
                Acc = True
                Exit For
            Else
                SendSingle "!login" & "#" & "Password" & "#", frmMain.Winsock1(Index)
            End If
        Case Else
            Acc = False
        End Select
    Next i
    If Acc = False Then SendSingle "!login" & "#" & "Account" & "#", frmMain.Winsock1(Index)
        
Case "!iprequest"
    For i = 1 To frmPanel.ListView1.ListItems.Count
        If GetUser = frmPanel.ListView1.ListItems.Item(i) Then
            SendSingle "!iprequest" & "#" & frmPanel.ListView1.ListItems.Item(i).SubItems(1), frmMain.Winsock1(Index)
        End If
    Next i
    
Case "!emote"
    Select Case GetConver
    Case "/lol", "/LOL", "/Lol", "/Laugh", "/laugh"
        SendMessage " " & GetUser & " laughs."
    Case "/Rofl", "/rofl", "/ROFL"
        SendMessage " " & GetUser & " rolls on the floor laughing."
    Case "/Beer", "/beer"
        SendMessage " " & GetUser & " takes a beer from the fridge."
    Case "/fart", "/Fart"
        SendMessage " " & GetUser & " farts loudly."
    Case "/lmao", "/LMAO"
        SendMessage " " & GetUser & " is laughing his / her ass off."
    Case "/facepalm", "/Facepalm"
        SendMessage " " & GetUser & " covers his face with his palm."
    Case "/violin"
        SendMessage " " & GetUser & " plays the world smallest violin."
    End Select

Case "!msg"
    'Split it
    array2 = Split(GetConver, " ")
    
    'Check if there is any special command
    Select Case array2(0)
    Case ".list"
        Select Case array2(1)
        Case "account"
            SendSingle "!accountlist" & "#" & GetAccountList, frmMain.Winsock1(Index)
        Case "user"
            SendSingle "!userlist" & "#" & GetUserList, frmMain.Winsock1(Index)
        End Select
        
    Case ".userinfo"
        If GetLevel(GetUser) <> "0" Then
            SendSingle GetUserInfo(array2(1)), Winsock1(Index)
        Else
            SendMessage " [" & GetUser & "]: " & GetConver
        End If
        
    Case ".accountinfo", ".accinfo"
        If GetLevel(GetUser) <> "0" Then
            SendSingle GetAccountInfo(array2(1)), Winsock1(Index)
        Else
            SendMessage " [" & GetUser & "]: " & GetConver
        End If
        
    Case ".kick"
        If GetLevel(GetUser) <> "0" Then
            KickUser (array2(1))
        Else
            SendMessage " [" & GetUser & "]: " & GetConver
        End If
        
    Case ".banaccount"
        If GetLevel(GetUser) <> "0" Then
            BanAccount array2(1), "Yes"
        Else
            SendMessage " [" & GetUser & "]: " & GetConver
        End If
        
    Case ".banuser"
        If GetLevel(GetUser) <> "0" Then
            BanUser array2(1), "Yes"
        Else
            SendMessage " [" & GetUser & "]: " & GetConver
        End If
        
    Case ".unbanuser"
        If GetLevel(GetUser) <> "0" Then
            BanUser array2(1), "No"
        Else
            SendMessage " [" & GetUser & "]: " & GetConver
        End If
        
    Case ".unbanaccount"
        If GetLevel(GetUser) <> "0" Then
            BanAccount array2(1), "No"
        Else
            SendMessage " [" & GetUser & "]: " & GetConver
        End If
        
    Case Else
        Select Case GetLevel(GetUser)
        Case "0"
            SendMessage " [" & GetUser & "]: " & GetConver
        Case "1"
            SendMessage " [GM][" & GetUser & "]: " & GetConver
        Case "2"
            SendMessage " [Admin][" & GetUser & "]: " & GetConver
        End Select
        
    End Select
       
Case Else
    SendMessage " Unknown operation."

End Select
'We want to read the message also , different then others tho
VisualizeMessage Command, GetUser, GetConver
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

Private Sub BanUser(User As String, Ban As String)
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = StrConv(User, vbProperCase) Then
            BanAccount .Item(i).SubItems(5), Ban
        End If
        Exit For
    Next i
End With
End Sub

Private Sub BanAccount(User As String, Ban As String)
With frmAccountPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(1) = User Then
            frmAccountPanel.ModifyAccount User, .Item(i).SubItems(2), Ban, .Item(i).SubItems(6), .Item(i)
            Exit For
        End If
    Next i
End With
End Sub

Private Sub KickUser(User As String)
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = User Then
            frmMain.Winsock1(.Item(i).SubItems(2)).Close
            Unload frmMain.Winsock1(.Item(i).SubItems(2))
            .Remove (i)
        End If
        Exit For
    Next i
End With

'Update clients user online list
UpdateUsersList

'Update Statusbar
frmMain.StatusBar1.Panels(1).Text = "Status: Connected with  " & frmMain.Winsock1.Count - 1 & " Client(s)."
        
End Sub

Private Function GetAccountInfo(Account As String) As String
For i = 1 To frmAccountPanel.ListView1.ListItems.Count
    With frmAccountPanel.ListView1.ListItems
        If .Item(i).SubItems(1) = Account Then
            GetAccountInfo = vbCrLf & " Account information about '" & Account & "'" & vbCrLf & " Name: " & .Item(i).SubItems(1) & vbCrLf & " Password: " & .Item(i).SubItems(2) & vbCrLf & " Registration Time: " & .Item(i).SubItems(3) & vbCrLf & " Registration Date: " & .Item(i).SubItems(4) & vbCrLf & " Banned: " & .Item(i).SubItems(5) & vbCrLf & " Level: " & .Item(i).SubItems(6)
            Exit For
        End If
    End With
Next i
End Function

Private Function GetUserInfo(User As String) As String
For i = 1 To frmPanel.ListView1.ListItems.Count
    With frmPanel.ListView1.ListItems
        If .Item(i) = StrConv(User, vbProperCase) Then
            GetUserInfo = vbCrLf & " User information about '" & StrConv(User, vbProperCase) & "'" & vbCrLf & " IP : " & .Item(i).SubItems(1) & vbCrLf & " Winsock ID: " & .Item(i).SubItems(2) & vbCrLf & " Login Time: " & .Item(i).SubItems(3) & vbCrLf & " Muted: " & .Item(i).SubItems(4) & vbCrLf & " Account: " & .Item(i).SubItems(5)
            Exit For
        End If
    End With
Next i
End Function

Private Function GetLevel(iName As String) As String
Dim GetAccount As String
For i = 1 To frmPanel.ListView1.ListItems.Count
    If iName = frmPanel.ListView1.ListItems.Item(i) Then
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
If Index < 0 Then
    frmChat.txtConver.Text = frmChat.txtConver.Text & vbCrLf & Prefix & " [System]: Disconnected with a host."
    For i = 1 To frmPanel.ListView1.ListItems.Count
        frmPanel.ListView1.ListItems.Remove (i)
    Next i
Else
    frmChat.txtConver.Text = frmChat.txtConver.Text & vbCrLf & Prefix & "[System]: Disconnected due connection problem."
    StatusBar1.Panels(1).Text = "[System]: Disconnected due connection problem."
    For i = 1 To frmPanel.ListView1.ListItems.Count
        frmPanel.ListView1.ListItems.Remove (i)
    Next i
End If
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
