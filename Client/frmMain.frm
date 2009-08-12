VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{342261DD-4D19-481B-8BF9-F24E643D0C20}#2.0#0"; "Hyperlink.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00F4F4F4&
   Caption         =   " Peach"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7470
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   ScrollBars      =   0   'False
   Begin MSWinsockLib.Winsock FSocket2 
      Index           =   0
      Left            =   960
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer STimer 
      Enabled         =   0   'False
      Left            =   480
      Top             =   1080
   End
   Begin MSWinsockLib.Winsock FSocket 
      Left            =   480
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin proHyp.Hyperlink Hyperlink1 
      Left            =   0
      Top             =   1560
      _ExtentX        =   2328
      _ExtentY        =   2514
   End
   Begin VB.Timer UpdateListPosition 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   1080
   End
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
            Object.Width           =   11466
            MinWidth        =   11466
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
      Left            =   0
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00F4F4F4&
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
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   7470
      TabIndex        =   0
      Top             =   0
      Width           =   7470
      Begin VB.CommandButton Command4 
         BackColor       =   &H00F4F4F4&
         Caption         =   "&Online List"
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
         BackColor       =   &H00F4F4F4&
         Caption         =   "&Send File"
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
         ToolTipText     =   "Work in progess .."
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00F4F4F4&
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
         BackColor       =   &H00F4F4F4&
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
   Begin VB.Menu myPOP 
      Caption         =   "myPOP"
      Visible         =   0   'False
      Begin VB.Menu Show 
         Caption         =   "&Show"
      End
      Begin VB.Menu Close 
         Caption         =   "&Exit"
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
    cbSize          As Long
    fMask           As Long
    fType           As Long
    fState          As Long
    wID             As Long
    hSubMenu        As Long
    hbmpChecked     As Long
    hbmpUnchecked   As Long
    dwItemData      As Long
    dwTypeData      As String
    cch             As Long
End Type
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32" ()

Dim Vali            As Boolean

Private Sub Close_Click()
    Unload frmList
    Unload frmLanguage
    Unload frmBlank
    Unload frmDESP
    Unload frmAbout
End Sub

Private Sub Command1_Click()
    SetupForms frmConfig
End Sub

Private Sub Command2_Click()
    SetupForms frmChat
End Sub

Private Sub SetupForms(Nix As Form)
    frmChat.Hide
    frmConfig.Hide
    frmSendFile.Hide
    Nix.Show
End Sub

Private Sub Command3_Click()
    SetupForms frmSendFile
End Sub

Private Sub Command4_Click()
frmMain.UpdateListPosition.Enabled = True
With frmList
    .Left = frmMain.Left + .Width * 2 + 20
    .Top = frmMain.Top
    .Height = frmMain.Height - 400
    .Show
End With
End Sub

Public Sub LoadMDIForm()
Command1.Caption = MDIcommand_config
Command2.Caption = MDIcommand_chat
Command3.Caption = MDIcommand_sendfile
Command4.Caption = MDIcommand_onlinelist
End Sub

Private Sub FSocket_Close()
Disconnect
End Sub

Private Sub FSocket_DataArrival(ByVal bytesTotal As Long)
Dim strMsg As String
Dim StrArr() As String
Dim CommX As String

FSocket.GetData strMsg

StrArr = Split(strMsg, "#")
CommX = StrArr(0)

Select Case CommX
Case "!acceptfile"
    frmSendFile.SendF FSocket.RemoteHost
Case "!denyfile"
    MsgBox SFmsgbox_filedecilined, vbInformation
End Select
End Sub

Private Sub FSocket2_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim intCounter As Integer
intCounter = loadSocket
With FSocket2(intCounter)
    .LocalPort = aPort
    .Accept requestID
End With
End Sub

Private Function socketFree() As Integer
On Error GoTo HandleErrorFreeSocket
    Dim theIP As Variant
    Dim p As Integer
    For p = FSocket2.LBound + 1 To FSocket2.UBound
        theIP = FSocket2(p).LocalIP
    Next
    socketFree = FSocket2.UBound + 1
Exit Function
HandleErrorFreeSocket:
socketFree = p
End Function

Private Function loadSocket() As Integer
Dim theFreeSocket As Integer
theFreeSocket = 0
theFreeSocket = socketFree

Load FSocket2(theFreeSocket)

loadSocket = theFreeSocket
End Function

Private Sub FSocket2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim ArrD As String
Dim ArrX() As String
Dim Comm As String

FSocket2(Index).GetData ArrD

ArrX = Split(ArrD, "#")
Comm = ArrX(0)

Select Case Comm
Case "!filerequest"
    If MsgBox(SFmsgbox_incfile, vbYesNo + vbQuestion) = vbYes Then
        FSocket2(Index).SendData "!acceptfile" & "#"
        frmSendFile2.Show
    Else
        FSocket2(Index).SendData "!denyfile" & "#"
    End If
End Select
End Sub

Private Sub MDIForm_Initialize()
Call InitCommonControls
End Sub

Private Sub MDIForm_Load()
On Error GoTo HandleErrorFile

LoadMDIForm
DisableFormResize Me

Dim L As Long
    L = GetWindowLong(Me.hwnd, GWL_STYLE)
'   L = L And Not (WS_MINIMIZEBOX)
    L = L And Not (WS_MAXIMIZEBOX)
    L = SetWindowLong(Me.hwnd, GWL_STYLE, L)

StatusBar1.Panels(1).Text = MDIstatusbar_disconnected

'Load 'Top' position from ini, if there is non take default value ( 1200 )
If ReadIniValue(App.Path & "\Config.ini", "Position", "Top") = "" Then
    Me.Top = 1200
Else
    Me.Top = ReadIniValue(App.Path & "\Config.ini", "Position", "Top")
End If
    
'Load 'Left' position from ini, if there is non take default value ( 1200 )
If ReadIniValue(App.Path & "\Config.ini", "Position", "Left") = "" Then
    Me.Left = 1200
Else
    Me.Left = ReadIniValue(App.Path & "\Config.ini", "Position", "Left")
End If

SetupForms frmConfig

Exit Sub
HandleErrorFile:
Select Case Err.Number
Case Else
    MsgBox "Error: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbInformation
    Exit Sub
End Select
End Sub


Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Msg As Long
Dim sFilter As String
Msg = X / Screen.TwipsPerPixelX
Select Case Msg
Case WM_LBUTTONDOWN
Case WM_LBUTTONUP
Case WM_LBUTTONDBLCLK
    Vali = True
    frmMain.Show ' show form
Case WM_RBUTTONDOWN
    frmMain.PopupMenu myPOP
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
    Unload frmList
    Unload frmLanguage
    Unload frmBlank
    Unload frmDESP
    Unload frmAbout
    Unload frmSendFile2
Shell_NotifyIcon NIM_DELETE, nid ' del tray icon
End Sub

Private Sub Show_Click()
Vali = True
frmMain.Show
End Sub

Private Sub STimer_Timer()
With FSocket
    If .State <> 7 Then
        'Do Nothing
    Else
        STimer.Enabled = False
        .SendData "!filerequest" & "#"
    End If
End With
End Sub

Private Sub UpdateListPosition_Timer()
With frmList
    .Left = frmMain.Left + .Width * 2 + 20
    .Top = frmMain.Top
    .Height = frmMain.Height - 400
End With
End Sub

Private Sub Winsock1_Close()
Prefix = "[" & Format(Time, "hh:nn:ss") & "]"

Winsock1.Close

Disconnect

LastMsg = ""

StatusBar1.Panels(1).Text = MDIstatusbar_dcfromserver
frmChat.txtConver.Text = frmChat.txtConver.Text & vbCrLf & Prefix & " [System]: You got disconnected from Server."
frmDESP.DisplayMessage DESPtext_dcserver

SetupForms frmConfig

End Sub

Private Sub Winsock1_Connect()
'Do the enable stuff
With frmConfig
    .Command2.Enabled = True
    .Command1.Enabled = False
    .txtNick.Enabled = False
    .txtIP.Enabled = False
    .txtPort.Enabled = False
    .txtAccount.Enabled = False
    .txtPassword.Enabled = False
    .SPT.Enabled = False
End With
With frmChat
    .cmdSend.Enabled = True
    .cmdClear.Enabled = True
    .txtToSend.Enabled = True
End With

frmConfig.Hide
frmChat.Show
frmChat.txtToSend.SetFocus

SendMsg "!login" & "#" & frmConfig.txtAccount & "#" & frmConfig.txtPassword & "#"
End Sub

Private Sub ConnectIsTrue()
StatusBar1.Panels(1).Text = MDIstatusbar_connected & frmConfig.txtIP.Text & ":" & frmConfig.txtPort.Text

SendMsg "!connected" & "#" & frmConfig.txtNick.Text & "#" & frmConfig.txtAccount.Text & "#"

frmDESP.DisplayMessage "Hello " & frmConfig.txtNick
End Sub

Private Sub ConnectIsFalse()
With frmConfig
    .Command2_Click
    .txtNick = ""
frmChat.Hide
    .Show
    .txtNick.SetFocus
End With
MsgBox MDImsgbox_nametaken, vbInformation
End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Prefix = "[" & Format(Time, "hh:nn:ss") & "]"
Dim GetCommand  As String
Dim StrArr()    As String
Dim StrArr2()   As String
Dim i           As Integer
Dim GetMessage  As String

'We get the message
Winsock1.GetData GetMessage

'We split the message into an array
StrArr = Split(GetMessage, "#")
    
'Assign the variables to the array
GetCommand = StrArr(0)

Select Case GetCommand

'We can't login ( choose other name )
Case "!accountlist"
    StrArr2 = Split(StrArr(1), " ")
    With frmChat
        .txtConver.Text = .txtConver.Text & vbCrLf & " Account List :"
        For i = LBound(StrArr2) To UBound(StrArr2) - 1
            .txtConver.Text = .txtConver.Text & vbCrLf & " - '" & StrArr2(i) & "'"
        Next i
    End With
        
Case "!userlist"
    StrArr2 = Split(StrArr(1), " ")
    With frmChat
        .txtConver.Text = .txtConver.Text & vbCrLf & " User List :"
        For i = LBound(StrArr2) To UBound(StrArr2) - 1
            .txtConver.Text = .txtConver.Text & vbCrLf & " - '" & StrArr2(i) & "'"
        Next i
    End With
    
Case "!decilined"
    ConnectIsFalse
    
'We can login
Case "!accepted"
    ConnectIsTrue
    
'Wipe out current list and insert new values
Case "!listupdate"
    frmList.ListView1.ListItems.Clear
    frmSendFile.Combo1.Clear
    For i = LBound(StrArr) + 1 To UBound(StrArr) - 1
        frmList.ListView1.ListItems.Add , , StrArr(i)
        frmSendFile.Combo1.AddItem StrArr(i)
    Next i

'We get login answer here
Case "!login"
    Select Case StrArr(1)
    Case "Yes"
        GetLevel = StrArr(2)
        SendMsg "!namerequest" & "#" & frmConfig.txtNick.Text & "#"
    Case "Password"
        With frmConfig
            .Command2_Click
            .txtPassword = ""
        frmChat.Hide
            .Show
            .txtPassword.SetFocus
        End With
        MsgBox MDImsgbox_wrong_password, vbInformation
    Case "Account"
        With frmConfig
            .Command2_Click
            .txtAccount = ""
        frmChat.Hide
            .Show
            .txtAccount.SetFocus
        End With
        MsgBox MDImsgbox_wrong_account, vbInformation
    Case "Banned"
        With frmConfig
            .Command2_Click
        frmChat.Hide
            .Show
        End With
        MsgBox MDImsgbox_banned, vbInformation
    End Select

'We get ip here
Case "!iprequest"
    
    'Connect new winsock to client
    FSocket.Close
    With FSocket
        .RemoteHost = StrArr(1)
        .RemotePort = aPort
        .Connect
    End With
    
    'Start timer to send file
    With STimer
        .Interval = 5
        .Enabled = True
    End With

Case Else
    frmChat.txtConver.Text = frmChat.txtConver.Text & vbCrLf & Prefix & GetMessage
    If frmMain.WindowState = 1 Then frmDESP.DisplayMessage DESPtext_newmsg
    
End Select
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim WiSk As Winsock
Prefix = "[" & Format(Time, "hh:nn:ss") & "]"

'Close connecting winsock ( state = 0 )
Winsock1.Close

'Call Disconnect function to DC all sockets and do buttons
Disconnect

'Last Message is empty
LastMsg = ""

'Change status to connection problem
StatusBar1.Panels(1).Text = MDIstatusbar_connectionproblem

frmConfig.Show
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
    
    SendMessage hwnd, WM_NCACTIVATE, True, 0
    
    frm.Width = frm.Width - 1
    frm.Width = frm.Width + 1
End Sub
