VERSION 5.00
Object = "{E9260B1F-3B85-11D8-8579-000C7641C3F2}#40.0#0"; "menudesp.ocx"
Begin VB.Form frmDESP 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "frmDESP"
   ClientHeight    =   1650
   ClientLeft      =   4140
   ClientTop       =   5130
   ClientWidth     =   2775
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
   ScaleHeight     =   110
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   185
   ShowInTaskbar   =   0   'False
   Begin VB.Timer stopTimer 
      Enabled         =   0   'False
      Left            =   120
      Top             =   1200
   End
   Begin menudesp.menudes menudes1 
      Height          =   2055
      Left            =   -10
      TabIndex        =   0
      Top             =   -40
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   3625
   End
End
Attribute VB_Name = "frmDESP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
Me.Top = Screen.Height - 2100
Me.Left = Screen.Width - 2882
End Sub

Private Sub UserIsOnline(uzer As String)
menudes1.Activate uzer & " is online!", 3, 2, True
With stopTimer
    .Interval = 2500
    .Enabled = True
End With
End Sub

Private Sub stopTimer_Timer()
Unload Me
End Sub

Public Sub YouAreOnline()
menudes1.Activate "Hello, " & frmConfig.txtNick.Text, 1, 4, True
With stopTimer
    .Interval = 3000
    .Enabled = True
End With
End Sub
