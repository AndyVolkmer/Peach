VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Server Configuration"
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7560
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "&Disconnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      Caption         =   "Configuration"
      ForeColor       =   &H8000000E&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "4728"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtNick 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "Server"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000C&
         Caption         =   "IP : "
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000C&
         Caption         =   "Name : "
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000C&
         Caption         =   "Port :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblNick 
         BackColor       =   &H8000000C&
         Caption         =   "Nickname :"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000C&
      Caption         =   "Author : Notron"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000C&
      Caption         =   "Version : 1.0.0.3"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   1335
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
' Do the buttons
    txtNick.Enabled = False
    txtPort.Enabled = False
    Command1.Enabled = False
    Command2.Enabled = True
        
' Activate the listen of sendfile form
    frmSendFile.cmdConnect_Click
    
' Connect sockets and start listening
    With frmMain
        .Winsock1(0).LocalPort = txtPort.Text
        .Winsock1(0).Listen
' frmMain.StatusBar1.Panels(1).Text = "Status: Waiting for connections .. "
        .StatusBar1.Panels(1).Text = "Status: Connected with  " & .Winsock1.Count - 1 & " Client(s)."
    End With
    
    frmConfig.Hide
    
    With frmChat
        .Show
        .txtToSend.SetFocus
    End With
    
End Sub

Private Sub Command2_Click()
Dim i As Integer

' Disconnect all sockets
    With frmMain
        For i = 1 To .Winsock1.Count - 1
            .Winsock1(i).Close
            Unload .Winsock1(i)
        Next i
        
        .Winsock1(0).Close
        
        .StatusBar1.Panels(1).Text = "Status: Disconnected"
        
    End With

' stop the listen of sendfile form
    frmSendFile.cmdConnect_Click
    
' Do the buttons
    txtNick.Enabled = True
    txtPort.Enabled = True
    Command1.Enabled = True
    Command2.Enabled = False
End Sub

Private Sub Form_Load()

Me.Top = 0
Me.Left = 0

With frmMain
    Label5.Caption = "IP : " & .Winsock1(0).LocalIP
    Label6.Caption = "Name : " & .Winsock1(0).LocalHostName
End With

End Sub

