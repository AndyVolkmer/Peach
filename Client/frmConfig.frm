VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4155
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "&Disconnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      Caption         =   "Configuration"
      ForeColor       =   &H8000000E&
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7215
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000C&
         Caption         =   "Server Information :"
         ForeColor       =   &H8000000E&
         Height          =   975
         Left            =   3600
         TabIndex        =   10
         Top             =   240
         Width           =   3015
         Begin VB.Label Label6 
            BackColor       =   &H8000000C&
            Caption         =   "Port : "
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label5 
            BackColor       =   &H8000000C&
            Caption         =   "IP : "
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000C&
         Caption         =   "Client Information :"
         ForeColor       =   &H8000000E&
         Height          =   975
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   3015
         Begin VB.Label Label4 
            BackColor       =   &H8000000C&
            Caption         =   "Name : "
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000C&
            Caption         =   "IP : "
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtNick 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         MaxLength       =   15
         TabIndex        =   0
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000C&
         Caption         =   "Port :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000C&
         Caption         =   "IP :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblNick 
         BackColor       =   &H8000000C&
         Caption         =   "Nickname :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   975
      End
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000C&
      Caption         =   "Version : 1.0.0.2"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000C&
      Caption         =   "Author : Notron"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
' Check if text's are empty
Select Case txtNick.Text
Case ""
    MsgBox "Nickname cant be empty.", vbCritical
    txtNick.SetFocus
Case Else
    Select Case txtIP.Text
    Case ""
        MsgBox "IP cant be empty.", vbCritical
        txtIP.SetFocus
    Case Else
        Select Case txtPort.Text
        Case ""
            MsgBox "Port cant be empty.", vbCritical
            txtPort.SetFocus
        Case Else
            ' If the nick is to short then no xP
            If Len(txtNick.Text) < 4 Then
                MsgBox "Your nickname is to short!    ", vbExclamation, " Error - Nickname"
                Exit Sub
            End If
            ' Do the enable stuff
            Command2.Enabled = True
            Command1.Enabled = False
            txtNick.Enabled = False
            txtIP.Enabled = False
            txtPort.Enabled = False
            
            With frmChat
                .cmdSend.Enabled = True
                .cmdClear.Enabled = True
                .txtToSend.Enabled = True
            End With
                        
            ' Connect winsocks
          With frmMain.Winsock1(0)
                .RemotePort = txtPort.Text
                .RemoteHost = txtIP.Text
                .Connect
            
            Label5.Caption = "IP: " & .RemoteHost
            Label6.Caption = "Port : " & .RemotePort
          End With
            
            frmMain.StatusBar1.Panels(1).Text = "Status: Connecting to '" & txtIP.Text & ":" & txtPort.Text & "'"
            frmConfig.Hide
            frmChat.Show
            frmChat.txtToSend.SetFocus
                   
        End Select
    End Select
End Select
End Sub

Private Sub Command2_Click()
' Do the buttons
    Command1.Enabled = True
    Command2.Enabled = False
    txtNick.Enabled = True
    txtIP.Enabled = True
    txtPort.Enabled = True
    Label5.Caption = "IP : "
    With frmChat
        .cmdSend.Enabled = False
        .cmdClear.Enabled = False
        .txtToSend.Enabled = False
    End With
        
' Close connection
frmMain.Winsock1(0).Close

frmMain.StatusBar1.Panels(1).Text = "Status: Disconnected"
End Sub

Private Sub Form_Load()

Me.Top = 0
Me.Left = 0

':::
Label3.Caption = "IP : " & frmMain.Winsock1(0).LocalIP
Label4.Caption = "Name : " & frmMain.Winsock1(0).LocalHostName

Dim TSSO As TypeSSO
TSSO = ReadConfigFile(App.Path & "\bin.conf")
With TSSO
    txtNick.Text = Trim(TSSO.Nickname)
    txtIP.Text = Trim(TSSO.ConnectIP)
    txtPort.Text = Trim(TSSO.Port)
End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
WriteConfigFile (App.Path & "\bin.conf")
End Sub

Private Sub txtIP_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub

Private Sub txtNick_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub

Private Sub txtNick_LostFocus()
txtNick.Text = StrConv(txtNick.Text, vbProperCase)
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub
