VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{342261DD-4D19-481B-8BF9-F24E643D0C20}#2.0#0"; "Hyperlink.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   " Peach (Client)"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7470
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   ScrollBars      =   0   'False
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
      Left            =   0
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
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
Private Declare Function SendMessage2 Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Public Prefix       As String
Public NameText     As String
Public ConverText   As String
Public Message      As String
Dim Vali            As Boolean

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

Public Sub MDIForm_Load()
On Error GoTo HandleErrorFile
Dim TSSO As TypeSSO
TSSO = ReadConfigFile(App.Path & "\bin.conf")

With TSSO
    Me.Top = Trim(.TopPos)
    Me.Left = Trim(.LeftPos)
End With

Command1.Caption = MDIcommand_config
Command2.Caption = MDIcommand_chat
Command3.Caption = MDIcommand_sendfile
Command4.Caption = MDIcommand_onlinelist

DisableFormResize Me

Dim L As Long
    L = GetWindowLong(Me.hwnd, GWL_STYLE)
'   L = L And Not (WS_MINIMIZEBOX)
    L = L And Not (WS_MAXIMIZEBOX)
    L = SetWindowLong(Me.hwnd, GWL_STYLE, L)

StatusBar1.Panels(1).Text = MDIstatusbar_disconnected

SetupForms frmConfig

Exit Sub
':::::::::::::::::::::::::::::::::::
'Error Handler
':::::::::::::::::::::::::::::::::::
HandleErrorFile:
Select Case Err.Number
Case 13
    MsgBox MDImsgbox_config_notify, vbInformation
    With Me
        .Top = 1200
        .Left = 1200
    End With
    With frmConfig
        .txtIP = "0.0.0.0"
        .txtPort = "4728"
        .txtNick = "DefaultNick"
        .Show
    End With
Case Else
    MsgBox "Error: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbCritical
    Exit Sub
End Select
End Sub


Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim msg As Long
Dim sFilter As String
msg = x / Screen.TwipsPerPixelX
Select Case msg
Case WM_LBUTTONDOWN
Case WM_LBUTTONUP
Case WM_LBUTTONDBLCLK
    Vali = True
    frmMain.Show ' show form
    'Shell_NotifyIcon NIM_DELETE, nid ' del tray icon
Case WM_RBUTTONDOWN
Case WM_RBUTTONUP
Case WM_RBUTTONDBLCLK
End Select
End Sub

Private Sub MDIForm_Resize()
If Me.WindowState = vbMinimized Then
    If Vali = False Then
        minimize_to_tray
    End If
    Vali = False
End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Unload frmList
    Unload frmLanguage
Shell_NotifyIcon NIM_DELETE, nid ' del tray icon
End Sub

Private Sub UpdateListPosition_Timer()
With frmList
    .Left = frmMain.Left + .Width * 2 + 20
    .Top = frmMain.Top
    .Height = frmMain.Height - 400
End With
End Sub

Private Sub Winsock1_Close(Index As Integer)
Prefix = "[" & Format(Time, "hh:nn:ss") & "]"

Winsock1(0).Close
StatusBar1.Panels(1).Text = MDIstatusbar_dcfromserver
frmChat.txtConver.Text = frmChat.txtConver.Text & vbCrLf & Prefix & " [System]: You got disconnected from Server."
frmList.List1.Clear

' Do the buttons
With frmConfig
    .Command1.Enabled = True
    .Command2.Enabled = False
    .txtNick.Enabled = True
    .txtIP.Enabled = True
    .txtPort.Enabled = True
    .Label5.Caption = "IP: "
    .Label6.Caption = "Port: "
End With
With frmChat
    .cmdSend.Enabled = False
    .cmdClear.Enabled = False
    .txtToSend.Enabled = False
End With

End Sub

Private Sub Winsock1_Connect(Index As Integer)
' Send request to check if name is avaible
SendMessage "!namerequest" & "#" & frmConfig.txtNick.Text & "#"
frmDESP.Show
frmDESP.YouAreOnline
End Sub

Private Sub ConnectIsTrue()
StatusBar1.Panels(1).Text = MDIstatusbar_connected & frmConfig.txtIP.Text & ":" & frmConfig.txtPort.Text
With frmMain
    .NameText = frmConfig.txtNick
    .Message = "!connected" & "#" & .NameText & "#"
SendMessage .Message
End With
End Sub

Private Sub ConnectIsFalse()
MsgBox MDImsgbox_nametaken, vbInformation
With frmConfig
    .Command2_Click
    .txtNick = ""
    SetupForms frmConfig
    .txtNick.SetFocus
End With
End Sub


Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Prefix = "[" & Format(Time, "hh:nn:ss") & "]"
Dim Command         As String
Dim arr()           As String
Dim i               As Integer
Dim Message         As String
    'We get the message
    Winsock1(Index).GetData Message
    
    'We decode (split) the message into an array
    arr = Split(Message, "#")
        
    'Assign the variables to the array
    Command = arr(0)
    
    Select Case Command
    Case "!decilineD" ' We cannot login ( choose other name )
        ConnectIsFalse
    Case "!accepteD" ' We can login
        ConnectIsTrue
    Case "!listupdate" ' Update the list
        frmList.List1.Clear
        For i = LBound(arr) + 1 To UBound(arr) - 1
                frmList.List1.AddItem arr(i)
        Next i
    Case Else ' Normal message
        frmChat.txtConver.Text = frmChat.txtConver.Text & vbCrLf & Prefix & Message
    End Select
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Prefix = "[" & Format(Time, "hh:nn:ss") & "]"
Winsock1(Index).Close
frmChat.txtConver.Text = frmChat.txtConver.Text & vbCrLf & Prefix & " [System]: Disconnected due connection problem."
StatusBar1.Panels(1).Text = MDIstatusbar_connectionproblem
frmList.List1.Clear
With frmConfig
    .Command1.Enabled = True
    .Command2.Enabled = False
    .txtNick.Enabled = True
    .txtIP.Enabled = True
    .txtPort.Enabled = True
    .Label5.Caption = "IP: "
    .Label6.Caption = "Port: "
End With
With frmChat
    .Hide
    .cmdClear.Enabled = False
    .cmdSend.Enabled = False
    .txtToSend.Enabled = False
End With
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
    
    SendMessage2 hwnd, WM_NCACTIVATE, True, 0
    
    frm.Width = frm.Width - 1
    frm.Width = frm.Width + 1
End Sub
