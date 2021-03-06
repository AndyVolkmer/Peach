VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmConfig 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   ClientHeight    =   5220
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
   ScaleHeight     =   5220
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox txt_log 
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3201
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmConfig.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft JhengHei"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Disconnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Connection Settings"
      ForeColor       =   &H00000000&
      Height          =   975
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
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00F4F4F4&
         Caption         =   "IP:"
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F4F4F4&
         Caption         =   "Port :"
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label7 
      BackColor       =   &H00F4F4F4&
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00F4F4F4&
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   1335
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public START_TIME As String

Private Sub Command1_Click()
SetButtons True

On Error GoTo ErrListen
'Connect sockets and start listening
With frmMain
    .Winsock1(0).LocalPort = txtPort.Text
    .Winsock1(0).Listen
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
        SetButtons False
        txtPort.SetFocus
        txtPort.SelStart = Len(txtPort.Text)

    Case Else
        MsgBox "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & "Report the number above to developement.", vbInformation
        Unload Me

End Select
End Sub

Private Sub Command2_Click()
Dim WiSk As Winsock

SetButtons False

For Each WiSk In frmMain.Winsock1
    If WiSk.Index <> 0 Then
        Unload WiSk
    End If
Next

frmMain.Winsock1(0).Close

For Each WiSk In frmAccountPanel.RegSock
    If WiSk.Index <> 0 Then
        Unload WiSk
    End If
Next

frmAccountPanel.RegSock(0).Close
frmPanel.lvUsers.ListItems.Clear
frmChannel.lvChannels.ListItems.Clear
frmChannel.lvUsers.ListItems.Clear
End Sub

Private Function CheckTx(TB As TextBox, MB As String) As Boolean
If LenB(TB.Text) = 0 Then
    MsgBox MB, vbInformation
    TB.SetFocus
    CheckTx = True
End If
End Function

Private Sub Form_Load()
Top = 0: Left = 0

Label5.Caption = "IP: " & frmMain.Winsock1(0).LocalIP
Label7.Caption = "Author: " & pAuthor
Label8.Caption = "Version: " & pRev
End Sub

Private Sub txt_log_Change()
With txt_log
    .SelStart = Len(.Text)
End With
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub
