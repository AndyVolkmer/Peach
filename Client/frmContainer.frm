VERSION 5.00
Begin VB.MDIForm frmContainer 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H00F4F4F4&
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7395
   Icon            =   "frmContainer.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   ScrollBars      =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00F4F4F4&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
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
      ScaleWidth      =   7395
      TabIndex        =   0
      Top             =   0
      Width           =   7395
      Begin VB.CommandButton cmdSwitch 
         Caption         =   "v"
         Height          =   375
         Left            =   6840
         TabIndex        =   4
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdSociety 
         Caption         =   "cmdSociety"
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
         Left            =   4650
         TabIndex        =   3
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdSendFile 
         Caption         =   "cmdSendFile"
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
         Left            =   2850
         TabIndex        =   2
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdChat 
         Caption         =   "cmdChat"
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
         Left            =   1050
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type MENUITEMINFO
    cbSize                          As Long
    fMask                           As Long
    fType                           As Long
    fState                          As Long
    wID                             As Long
    hSubMenu                        As Long
    hbmpChecked                     As Long
    hbmpUnchecked                   As Long
    dwItemData                      As Long
    dwTypeData                      As String
    cch                             As Long
End Type

Private Vali                        As Boolean

'Windows frame constants
Private Const GWL_STYLE             As Long = (-16)

Private Const WS_MAXIMIZEBOX        As Long = &H10000
Private Const WS_MINIMIZEBOX        As Long = &H20000
Private Const WS_THICKFRAME         As Long = &H40000
Private Const WS_SYSMENU            As Long = &H80000
Private Const WS_CAPTION            As Long = &HC00000

Private Const SC_MAXIMIZE           As Long = &HF030&
Private Const SC_MINIMIZE           As Long = &HF020&
Private Const SC_CLOSE              As Long = &HF060&
Private Const MIIM_STATE            As Long = &H1&
Private Const MIIM_ID               As Long = &H2&
Private Const MFS_GRAYED            As Long = &H3&
Private Const WM_NCACTIVATE         As Long = &H86

Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long

Private Sub Command1_Click()
SetupChildForm frmSociety
End Sub

Private Sub cmdChat_Click()
SetupChildForm frmChat
End Sub

Private Sub cmdSendFile_Click()
SetupChildForm frmSendFile
End Sub

Private Sub cmdSociety_Click()
SetupChildForm frmSociety
End Sub

Private Sub cmdSwitch_Click()
SetupForm frmMain
End Sub

Private Sub MDIForm_Activate()
Me.Top = frmMain.Top
Me.Left = frmMain.Left
End Sub

Private Sub MDIForm_Load()
Me.Caption = pCaption
cmdChat.Caption = MDI_COMMAND_CHAT
cmdSendFile.Caption = MDI_COMMAND_SENDFILE
cmdSociety.Caption = MDI_COMMAND_SOCIETY
DisableFormResize Me
End Sub

Private Sub DisableFormResize(frm As Form)
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

SendMessage2 frm.hwnd, WM_NCACTIVATE, True, 0

frm.Width = frm.Width - 1
frm.Width = frm.Width + 1
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MSG As Long
    MSG = X / Screen.TwipsPerPixelX

Select Case MSG
    Case WM_LBUTTONDOWN
    Case WM_LBUTTONUP
        Vali = True
        frmContainer.Show
        frmContainer.WindowState = 0
        Shell_NotifyIcon NIM_DELETE, NID    'Del tray icon

    Case WM_LBUTTONDBLCLK
    Case WM_RBUTTONDOWN: frmMain.PopupMenu frmMain.myPOP
    Case WM_RBUTTONUP
    Case WM_RBUTTONDBLCLK
End Select
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
frmContainer.Hide
SetupForm frmMain
End Sub

Private Sub MDIForm_Resize()
If Me.WindowState = 1 Then
    If Vali = False Then
        If Setting.MIN_TICK Then
            MinimizeToTray frmContainer
        End If
    End If
    Vali = False
End If
End Sub
