VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00F4F4F4&
   Caption         =   " Peach"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7350
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
      Left            =   0
      Top             =   1080
   End
   Begin MSWinsockLib.Winsock FSocket 
      Left            =   480
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   5
      Top             =   4740
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   14994
            MinWidth        =   14994
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
      ScaleWidth      =   7350
      TabIndex        =   0
      Top             =   0
      Width           =   7350
      Begin VB.CommandButton Command4 
         BackColor       =   &H00F4F4F4&
         Caption         =   "&Society"
         BeginProperty Font 
            Name            =   "Segoe UI"
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
            Name            =   "Segoe UI"
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
         BackColor       =   &H00F4F4F4&
         Caption         =   "Ch&at"
         BeginProperty Font 
            Name            =   "Segoe UI"
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
         Caption         =   "&Login"
         BeginProperty Font 
            Name            =   "Segoe UI"
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
   Begin MSWinsockLib.Winsock RegSock 
      Left            =   1440
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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

Dim Vali        As Boolean
Public RunOnce  As Boolean

Private Sub Close_Click()
End
End Sub

Private Sub Command1_Click()
SetupForms frmConfig
End Sub

Private Sub Command2_Click()
SetupForms frmChat
End Sub

Private Sub SetupForms(Nix As Form)
frmSociety.Hide
frmChat.Hide
frmConfig.Hide
frmSendFile.Hide
Nix.Show
End Sub

Private Sub Command3_Click()
SetupForms frmSendFile
End Sub

Private Sub Command4_Click()
SetupForms frmSociety
End Sub

Public Sub LoadMDIForm()
Command2.Caption = MDI_COMMAND_CHAT
Command3.Caption = MDI_COMMAND_SENDFILE
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
    MsgBox SF_MSG_DECILINED, vbInformation
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
    Dim theIP As String
    For i = FSocket2.LBound + 1 To FSocket2.UBound
        theIP = FSocket2(i).LocalIP
    Next i
    socketFree = FSocket2.UBound + 1
Exit Function
HandleErrorFreeSocket:
socketFree = i
End Function

Private Function loadSocket() As Integer
Dim theFreeSocket As Integer
theFreeSocket = 0
theFreeSocket = socketFree

Load FSocket2(theFreeSocket)

loadSocket = theFreeSocket
End Function

Private Sub FSocket2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim GetMessage As String
Dim array1() As String
Dim GetCommand As String

FSocket2(Index).GetData GetMessage

array1 = Split(GetMessage, "#")
GetCommand = array1(0)

Select Case GetCommand
Case "!filerequest"
    If MsgBox(SF_MSG_INCOMMING_FILE, vbYesNo + vbQuestion) = vbYes Then
        FSocket2(Index).SendData "!acceptfile#"
        frmSendFile2.Show
    Else
        FSocket2(Index).SendData "!denyfile#"
    End If
End Select
End Sub

Private Sub MDIForm_Initialize()
Call InitCommonControls
End Sub

Private Sub MDIForm_Load()
LoadMDIForm
DisableFormResize Me

Dim l As Long
    l = GetWindowLong(Me.hwnd, GWL_STYLE)
'   L = L And Not (WS_MINIMIZEBOX)
    l = l And Not (WS_MAXIMIZEBOX)
    l = SetWindowLong(Me.hwnd, GWL_STYLE, l)

StatusBar1.Panels(1).Text = MDI_STAT_DISCONNECTED

Me.Top = Setting.MAIN_TOP
Me.Left = Setting.MAIN_LEFT

SetupForms frmConfig
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
    With frmMain
        .Show
        .WindowState = 0
    End With
    Shell_NotifyIcon NIM_DELETE, nid     'Del tray icon
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
Shell_NotifyIcon NIM_DELETE, nid 'Del tray icon
End
End Sub

Private Sub Show_Click()
Vali = True
With frmMain
    .Show
    .WindowState = 0
End With
Shell_NotifyIcon NIM_DELETE, nid 'Del tray icon
End Sub

Private Sub STimer_Timer()
With FSocket
    If .State = 7 Then
        STimer.Enabled = False
        .SendData "!filerequest#"
    End If
End With
End Sub

Private Sub Winsock1_Close()
Prefix = " [" & Format$(Time, "hh:nn:ss") & "]"

Winsock1.Close
Disconnect

StatusBar1.Panels(1).Text = MDI_STAT_DISCONNECT
With frmChat.txtConver
    .SelStart = Len(.Text)
    .SelRTF = vbCrLf & Prefix & " [System]: You got disconnected from Server."
End With
    
SetupForms frmConfig
End Sub

Private Sub Winsock1_Connect()
SwitchButtons False
SendMsg "!login#" & frmConfig.txtAccount & "#" & frmConfig.txtPassword & "#"
End Sub

Private Sub Connection(Args As Boolean)
If Args = True Then
    frmConfig.Hide
    frmChat.Show
    frmChat.txtToSend.SetFocus
    
    StatusBar1.Panels(1).Text = MDI_STAT_CONNECTED
    
    SendMsg "!connected#" & frmConfig.txtNick & "#" & frmConfig.txtAccount & "#"
Else
    With frmConfig
        .cmdConnect_Click
        .txtNick = vbNullString
        frmChat.Hide
        .Show
        .txtNick.SetFocus
    End With
    MsgBox MDI_MSG_NAME_TAKEN, vbInformation
End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Prefix = "[" & Format$(Time, "hh:nn:ss") & "]"
Dim GetCommand  As String
Dim StrArr()    As String
Dim StrArr2()   As String
Dim GetMessage  As String
Dim Buffer      As String

'We get the message
Winsock1.GetData GetMessage
DoEvents

'We split the message into an array
StrArr = Split(GetMessage, "#")
    
'Assign the variables to the array
GetCommand = StrArr(0)

Select Case GetCommand

Case "!split_text"
    For i = 1 To UBound(StrArr)
        Buffer = Buffer & vbCrLf & " " & StrArr(i)
    Next i
    With frmChat.txtConver
        .SelStart = Len(.Text)
        .SelRTF = Buffer
    End With
   
'We can't login
Case "!decilined"
    Connection False
    
'We can login
Case "!accepted"
    Connection True
    
'Wipe out current friend list and insert new values
Case "!update_friends"
    Dim f_array() As String
    Dim j As Long
    
    With frmSociety.ListView2.ListItems
        .Clear
        For i = LBound(StrArr) + 1 To UBound(StrArr) - 1
            f_array = Split(StrArr(i), "$")
            
            'Add account name of friend
            .Add , , f_array(0)
            j = .Count
            
            'Write down status
            .Item(j).SubItems(1) = f_array(1)
            .Item(j).ListSubItems(1).Bold = True
            
            'Color the listview with apropiate color
            If Len(f_array(1)) = 6 Then
                .Item(j).ListSubItems(1).ForeColor = RGB(0, 132, 43)
            Else
                .Item(j).ListSubItems(1).ForeColor = RGB(132, 0, 0)
            End If
        Next i
        
        'Ask for server information
        If RunOnce = False Then
            RunOnce = True
            SendMsg "!server_info#"
        End If
    End With
    
'Wipe out current list and insert new values
Case "!update_online"
    Dim User As String
    
    'Clear the current list so we don't get multiply entries
    frmSociety.ListView1.ListItems.Clear
    frmSendFile.Combo1.Clear
    
    'Go through array and added users
    For i = LBound(StrArr) + 1 To UBound(StrArr) - 1
        frmSociety.ListView1.ListItems.Add , , StrArr(i)
        User = Left(StrArr(i), InStr(1, StrArr(i), " ") - 1)
        frmSendFile.Combo1.AddItem User
    Next i
    
    'Ask for friendlist
    SendMsg "!friend_get#" & frmConfig.txtAccount & "#"
    
'We get login answer here
Case "!login"
    Select Case StrArr(1)
    Case "Yes"
        frmMain.StatusBar1.Panels(1).Text = "Status : Authenticating.."
        SendMsg "!namerequest#" & frmConfig.txtNick & "#"
        
    Case "Password"
        With frmConfig
            .cmdConnect_Click
            .txtPassword = vbNullString
            frmChat.Hide
            .Show
            .txtPassword.SetFocus
        End With
        MsgBox MDI_MSG_WRONG_PASSWORD, vbInformation
        
    Case "Account"
        With frmConfig
            .cmdConnect_Click
            .txtAccount = vbNullString
            frmChat.Hide
            .Show
            .txtAccount.SetFocus
        End With
        MsgBox MDI_MSG_WRONG_ACCOUNT, vbInformation
        
    Case "Banned"
        With frmConfig
            .cmdConnect_Click
            frmChat.Hide
            .Show
        End With
        MsgBox MDI_MSG_BANNED, vbInformation
        
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

'Function to display information messageboxes
Case "!msgbox"
    MsgBox StrArr(1), vbInformation

Case Else
    With frmChat.txtConver
        .SelStart = Len(.Text)
        .SelRTF = vbCrLf & Space(1) & Prefix & Space(1) & GetMessage
    End With
    
End Select
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim WiSk As Winsock

Winsock1.Close
Disconnect
StatusBar1.Panels(1).Text = MDI_STAT_CONNECTION_ERROR

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

Private Sub RegSock_Close()
If ACC_SWITCH = "REG" Then
    With frmRegistration
        .Caption = REG_MSG_ERROR_OCCURED
        .Label4.Caption = REG_MSG_ERROR
        .Command1.Caption = REG_COMMAND_CLOSE
        .Command1.Visible = True
        .Frame1.Visible = False
        .Check1.Visible = False
        .cmbSecretQuestion.Visible = False
        .txtSecretAnswer.Visible = False
    End With
Else
    With frmForgotPassword
        .Caption = REG_MSG_ERROR_OCCURED
        .lblStatus.Caption = REG_MSG_ERROR
        .cmdRequest.Caption = REG_COMMAND_CLOSE
        .cmdRequest.Visible = True
        .Frame1.Visible = False
        .txtAccount.Visible = False
        .cmbSecretQuestion.Visible = False
        .txtSecretAnswer.Visible = False
    End With
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub RegSock_Connect()
If ACC_SWITCH = "REG" Then
    With frmRegistration
        .Caption = REG_CAPTION
        .Frame1.Visible = True
        .Check1.Visible = True
        .cmbSecretQuestion.Visible = True
        .txtSecretAnswer.Visible = True
        .Command1.Visible = True
        .Command1.Caption = REG_COMMAND_SUBMIT
    End With
Else
    With frmForgotPassword
        .Caption = FP_CAPTION
        .Frame1.Visible = True
        .txtAccount.Visible = True
        .txtSecretAnswer.Visible = True
        .cmbSecretQuestion.Visible = True
        .cmdRequest.Visible = True
        .cmdRequest.Caption = FP_COMMAND_REQUEST
    End With
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub RegSock_DataArrival(ByVal bytesTotal As Long)
Dim GetMessage As String
Dim array1() As String

RegSock.GetData GetMessage

array1 = Split(GetMessage, "#")

Select Case array1(0)
Case "!nameexist"
    MsgBox REG_MSG_ACCOUNT_EXIST, vbInformation
    With frmRegistration
        .txtAccount = vbNullString
        .txtAccount.SetFocus
    End With
    
Case "!done"
    MsgBox REG_MSG_SUCCESSFULLY, vbInformation
    Unload frmRegistration
    
Case "!error_fp"
    MsgBox FP_MSG_WRONG_ANSWER, vbInformation
    With frmForgotPassword
        .txtSecretAnswer = vbNullString
        .txtSecretAnswer.SetFocus
    End With
    
Case "!successfull"
    MsgBox FP_MSG_SUCCESSFULL & array1(1), vbInformation
    Unload frmForgotPassword
    
Case "!account_not_exist"
    MsgBox MDI_MSG_WRONG_ACCOUNT, vbInformation
    With frmForgotPassword
        .txtAccount = vbNullString
        .txtAccount.SetFocus
    End With
    
End Select
End Sub

Private Sub RegSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If ACC_SWITCH = "REG" Then
    With frmRegistration
        .Caption = REG_MSG_ERROR_OCCURED
        .Label4.Caption = REG_MSG_ERROR
        .Command1.Caption = REG_COMMAND_CLOSE
        .Command1.Visible = True
        .Frame1.Visible = False
        .Check1.Visible = False
        .cmbSecretQuestion.Visible = False
        .txtSecretAnswer.Visible = False
    End With
Else
    With frmForgotPassword
        .Caption = REG_MSG_ERROR_OCCURED
        .lblStatus.Caption = REG_MSG_ERROR
        .cmdRequest.Caption = REG_COMMAND_CLOSE
        .cmdRequest.Visible = True
        .Frame1.Visible = False
        .txtAccount.Visible = False
        .cmbSecretQuestion.Visible = False
        .txtSecretAnswer.Visible = False
    End With
End If
Screen.MousePointer = vbDefault
End Sub

