VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmConfig 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   Caption         =   "frmConfig"
   ClientHeight    =   4140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7605
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
   ScaleHeight     =   4140
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox txt_log 
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   2143
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmConfig.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer connCounter 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1920
      Top             =   3480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Disconnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Connection Settings"
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         MaxLength       =   5
         TabIndex        =   2
         Text            =   "4728"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F4F4F4&
         Caption         =   "Offline"
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackColor       =   &H00F4F4F4&
         Caption         =   "IP : "
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F4F4F4&
         Caption         =   "Port :"
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label3 
      Caption         =   " Server Log:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Author : Notron"
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00F4F4F4&
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   120
      TabIndex        =   6
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
If HasError = True Then
    Command1.Enabled = False
    MsgBox "Database error occured, read the log for more information.", vbInformation
    Exit Sub
End If

connCounter.Enabled = True
'Do the buttons
txtPort.Enabled = False
Command1.Enabled = False
Command2.Enabled = True
    
On Error GoTo ErrListen
'Connect sockets and start listening
With frmMain
    .Winsock1(0).LocalPort = txtPort.Text
    .Winsock1(0).Listen
    .StatusBar1.Panels(1).Text = "Status: Connected with " & .Winsock1.Count - 1 & " Client(s)."
End With

With frmAccountPanel
    .RegSock(0).LocalPort = rPort
    .RegSock(0).Listen
End With

frmMain.SetupForms frmChat

frmChat.txtToSend.SetFocus

Exit Sub
ErrListen:
Select Case Err.Number
Case 10048
    MsgBox "This adress is already in use, please select another port.", vbInformation
    txtPort.Enabled = True
    Command1.Enabled = True
    Command2.Enabled = False
    txtPort.SetFocus
    txtPort.SelStart = Len(txtPort.Text)
    connCounter.Enabled = False
    Label2.Caption = "Offline"
Case Else
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & "Report the number above to developement.", vbInformation
    Unload Me
End Select
End Sub

Private Sub Command2_Click()
Dim WiSk As Winsock

connCounter.Enabled = False
Label2.Caption = "Offline"

With frmMain
    For Each WiSk In .Winsock1
        If WiSk.State = 7 Then
            WiSk.Close
            Unload WiSk
        End If
    Next
    .Winsock1(0).Close
    .StatusBar1.Panels(1).Text = "Status: Disconnected"
End With

With frmAccountPanel
    For Each WiSk In .RegSock
        If WiSk.State = 7 Then
            WiSk.Close
            Unload WiSk
        End If
    Next
    .RegSock(0).Close
End With
        
'Clear frmPanel ListView
frmPanel.ListView1.ListItems.Clear

'Do the buttons
txtPort.Enabled = True
Command1.Enabled = True
Command2.Enabled = False
End Sub

Private Function CheckTx(TB As TextBox, MB As String) As Boolean
If Len(TB.Text) = 0 Then
    MsgBox MB, vbInformation
    TB.SetFocus
    CheckTx = True
End If
End Function

Private Sub connCounter_Timer()
Static X As Long
X = X + 1
Label2.Caption = "Online Time : " & Format$(TimeSerial(0, 0, X), "hh:mm:ss")
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0

With frmMain
    Label5.Caption = "IP : " & .Winsock1(0).LocalIP
    Label8.Caption = "Version : " & Rev
End With
End Sub

Private Sub txt_log_Change()
With txt_log
    .SelStart = Len(.Text)
End With
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub
