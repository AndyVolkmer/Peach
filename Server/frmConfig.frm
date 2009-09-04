VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   Caption         =   "frmConfig"
   ClientHeight    =   4140
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
   ScaleHeight     =   4140
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "&Connect Database"
      Height          =   375
      Left            =   5280
      TabIndex        =   22
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Database Settings"
      Height          =   1815
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   7215
      Begin VB.TextBox txtDBTable 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   21
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtDB 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtDBPassword 
         Alignment       =   2  'Center
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3840
         PasswordChar    =   "*"
         TabIndex        =   17
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtDBUser 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   15
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtDBIP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lbl_DB_Table 
         BackColor       =   &H00F4F4F4&
         Caption         =   "Table:"
         Height          =   255
         Left            =   2040
         TabIndex        =   20
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lbl_DB 
         BackColor       =   &H00F4F4F4&
         Caption         =   "Database:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lbl_DB_Password 
         BackColor       =   &H00F4F4F4&
         Caption         =   "Password:"
         Height          =   255
         Left            =   3840
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lbl_DB_User 
         BackColor       =   &H00F4F4F4&
         Caption         =   "User:"
         Height          =   255
         Left            =   2040
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lbl_DB_IP 
         BackColor       =   &H00F4F4F4&
         Caption         =   "IP:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
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
      TabIndex        =   5
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
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
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "4728"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtNick 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "Server"
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F4F4F4&
         Caption         =   "Offline"
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackColor       =   &H00F4F4F4&
         Caption         =   "IP : "
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F4F4F4&
         Caption         =   "Port :"
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblNick 
         BackColor       =   &H00F4F4F4&
         Caption         =   "Nickname :"
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label7 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Author : Notron"
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00F4F4F4&
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   120
      TabIndex        =   8
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
connCounter.Enabled = True
'Do the buttons
txtNick.Enabled = False
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
    .RegSock(0).LocalPort = RegPort
    .RegSock(0).Listen
End With

frmMain.SetupForms frmChat

frmChat.txtToSend.SetFocus

Exit Sub
ErrListen:
Select Case Err.Number
Case 10048
    MsgBox "This adress is already in use, please select another port.", vbInformation
    txtNick.Enabled = True
    txtPort.Enabled = True
    Command1.Enabled = True
    Command2.Enabled = False
    txtPort.SetFocus
    txtPort.SelStart = Len(txtPort.Text)
    connCounter.Enabled = False
    Label2.Caption = "Offline"
Case Else
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & "Report the number above to developer.", vbInformation
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
txtNick.Enabled = True
txtPort.Enabled = True
Command1.Enabled = True
Command2.Enabled = False
End Sub

Private Sub Command3_Click()
If CheckTx(txtDBIP, "No IP entered.") = True Then Exit Sub
If CheckTx(txtDBUser, "No User entered.") = True Then Exit Sub
If CheckTx(txtDBPassword, "No Password entered.") = True Then Exit Sub
If CheckTx(txtDB, "No Database entered.") = True Then Exit Sub
If CheckTx(txtDBTable, "No Table entered.") = True Then Exit Sub

frmMain.ConnectMySQL txtDB.Text, txtDBUser.Text, txtDBPassword.Text, txtDBIP.Text
frmMain.LoadCustomerListView txtDBTable.Text

End Sub

Public Sub DisableDatabaseField()
txtDBIP.Enabled = False
txtDB.Enabled = False
txtDBUser.Enabled = False
txtDBPassword.Enabled = False
txtDBTable.Enabled = False

Frame2.Enabled = False
Command3.Enabled = False

lbl_DB.Enabled = False
lbl_DB_IP.Enabled = False
lbl_DB_Password.Enabled = False
lbl_DB_Table.Enabled = False
lbl_DB_User.Enabled = False
End Sub

Private Function CheckTx(TB As TextBox, MB As String) As Boolean
If Len(TB.Text) = 0 Then
    MsgBox MB, vbInformation
    TB.SetFocus
    CheckTx = True
End If
End Function

Private Sub connCounter_Timer()
Static x As Long
x = x + 1
Label2.Caption = "Online Time : " & Format$(TimeSerial(0, 0, x), "hh:mm:ss")
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0

With frmMain
    Label5.Caption = "IP : " & .Winsock1(0).LocalIP
    Label8.Caption = "Version : " & Rev
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
WriteIniValue App.Path & "\Config.ini", "Database", "IP", txtDBIP.Text
WriteIniValue App.Path & "\Config.ini", "Database", "User", txtDBUser.Text
WriteIniValue App.Path & "\Config.ini", "Database", "Password", Encode(txtDBPassword.Text)
WriteIniValue App.Path & "\Config.ini", "Database", "Database", txtDB.Text
WriteIniValue App.Path & "\Config.ini", "Database", "Table", txtDBTable.Text
End Sub

Private Sub txtNick_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub
