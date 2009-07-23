VERSION 5.00
Object = "{E9260B1F-3B85-11D8-8579-000C7641C3F2}#40.0#0"; "menudesp.ocx"
Begin VB.Form frmDESP 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   Caption         =   "frmDESP"
   ClientHeight    =   2040
   ClientLeft      =   4140
   ClientTop       =   5130
   ClientWidth     =   3360
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "frmDESP"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   136
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   224
   ShowInTaskbar   =   0   'False
   Begin menudesp.menudes menudes1 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3625
   End
   Begin VB.Timer stopTimer 
      Enabled         =   0   'False
      Left            =   120
      Top             =   240
   End
End
Attribute VB_Name = "frmDESP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents fEvents As Form
Attribute fEvents.VB_VarHelpID = -1
Private bNoResize As Boolean, bNoActivate As Boolean

Private Sub Form_Load()
Me.Top = Screen.Height - 2100
Me.Left = Screen.Width - 2882
Set fEvents = frmBlank

' ShowInTaskBar
SetWindowLong Me.hwnd, GWL_STYLE, GetWindowLong(Me.hwnd, GWL_STYLE) Or WS_SYSMENU Or WS_MINIMIZEBOX

' Set Transparency
Me.BackColor = vbCyan
'Text2.BackColor = Me.BackColor
SetTrans Me, , Me.BackColor
SetTrans fEvents, 1

Me.Show
fEvents.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload fEvents
Set fEvents = Nothing
End Sub

Public Sub YouAreOnline()
menudes1.Activate "Hello " & frmConfig.txtNick.Text, 5, 2, False
With stopTimer
    .Interval = 3000
    .Enabled = True
End With
End Sub

Private Sub fEvents_Activate()
    If Not GetNextWindow(fEvents.hwnd) = Me.hwnd Then Me.ZOrder
End Sub

Private Sub Form_Activate()
    If Not GetNextWindow(Me.hwnd) = fEvents.hwnd Then fEvents.ZOrder
End Sub

Private Sub fEvents_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If IsOverCtl(Me, X \ Screen.TwipsPerPixelX, Y \ Screen.TwipsPerPixelY) Then Me.SetFocus
End Sub

Private Sub stopTimer_Timer()
Unload Me
End Sub
