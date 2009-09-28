VERSION 5.00
Begin VB.Form frmConfig 
   Appearance      =   0  'Flat
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   Caption         =   "frmConfig"
   ClientHeight    =   4185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
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
   ScaleHeight     =   4185
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Register Account"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   5400
      TabIndex        =   18
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Update"
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
      Left            =   1440
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Disconnect"
      Enabled         =   0   'False
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
      Left            =   5760
      TabIndex        =   8
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Language"
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
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Connect"
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
      Left            =   4080
      TabIndex        =   7
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Personal Settings"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   7335
      Begin VB.CheckBox SPT 
         BackColor       =   &H00F4F4F4&
         Caption         =   "Save Password"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   19
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtAccount 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         MaxLength       =   15
         TabIndex        =   0
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtNick 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblPassword 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Password:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblAccount 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Account:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblNickname 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Nickname:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Connection Settings"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   7335
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblPort 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Port:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblIP 
         BackColor       =   &H00F4F4F4&
         Caption         =   " IP:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00F4F4F4&
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label lblAuthor 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Author : Notron"
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3480
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
If Command1.Enabled = False Then
    Exit Sub
End If

'Account can't be empty
If CheckTx(txtAccount, CONFIG_MSG_ACCOUNT) = True Then Exit Sub
 
'Password can't be empty
If CheckTx(txtPassword, CONFIG_MSG_PASSWORD) = True Then Exit Sub
 
'Nick can't be empty
If CheckTx(txtNick, CONFIG_MSG_NAME) = True Then Exit Sub
 
'IP can't be empty
If CheckTx(txtIP, CONFIG_MSG_IP) = True Then Exit Sub
 
'Port can't be empty
If CheckTx(txtPort, CONFIG_MSG_PORT) = True Then Exit Sub

'Nick can't be numeric
If IsNumeric(txtNick.Text) = True Then
    txtNick.Text = vbNullString
    MsgBox CONFIG_MSG_NUMERIC, vbInformation
    txtNick.SetFocus
    Exit Sub
End If

'Nick can't be shorter then 4 characters
If Len(txtNick.Text) < 4 Then
    MsgBox "Your nickname is to short!", vbInformation, " Error - Nickname"
    txtNick.SelStart = Len(txtNick.Text)
    txtNick.SetFocus
    Exit Sub
End If

'Make the name proper case
txtNick.Text = StrConv(txtNick.Text, vbProperCase)

'Connect winsocks
With frmMain.Winsock1
    .RemotePort = txtPort.Text
    .RemoteHost = txtIP.Text
    .Connect
End With

'Set Recieve-Request-Winsock to listen
With frmMain
    .FSocket2(0).LocalPort = aPort
    .FSocket2(0).Listen
End With

'Set Recieve-File-Winsock to listen
With frmSendFile2
    .SckReceiveFile(0).LocalPort = bPort
    .SckReceiveFile(0).Listen
End With

Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False

frmMain.StatusBar1.Panels(1).Text = MDI_STAT_CONNECTING
End Sub

Public Sub Command2_Click()
Disconnect
With frmMain
    .Winsock1.Close
    .StatusBar1.Panels(1).Text = MDI_STAT_DISCONNECTED
End With
End Sub

Private Sub Command3_Click()
frmLanguage.Show 1
End Sub

Private Sub Command4_Click()
On Error Resume Next
If FileExists(App.Path & "\peachUpdater.exe") = True Then
    Shell App.Path & "\peachUpdater.exe", vbNormalFocus
Else
    MsgBox "You need the peach updater to be able to upgrade your peach." & vbCrLf & vbCrLf & "Download it here: http://riplegion.ri.funpic.de/Peach/peachUpdater.exe", vbInformation
End If
End Sub

Private Sub Command5_Click()
frmRegistration.Show 1
End Sub

Public Sub Form_Load()
Top = 0
Left = 0
lblVersion.Caption = "Version : " & Rev

LoadConfigForm

'Read 'IP' from .ini file
If Len(ReadIniValue(App.Path & "\Config.ini", "Connection", "IP")) <> 0 Then
    txtIP.Text = ReadIniValue(App.Path & "\Config.ini", "Connection", "IP")
End If

'Read 'Port' from .ini file
If Len(ReadIniValue(App.Path & "\Config.ini", "Connection", "Port")) <> 0 Then
    txtPort.Text = ReadIniValue(App.Path & "\Config.ini", "Connection", "Port")
End If

'Read 'Nickname' from .ini file
If Len(ReadIniValue(App.Path & "\Config.ini", "Private", "Nickname")) <> 0 Then
    txtNick.Text = ReadIniValue(App.Path & "\Config.ini", "Private", "Nickname")
End If

'Read 'Account' from .ini file
If Len(ReadIniValue(App.Path & "\Config.ini", "Private", "Account")) <> 0 Then
    txtAccount.Text = ReadIniValue(App.Path & "\Config.ini", "Private", "Account")
End If

'Read 'Save Password Tick'
If Len(ReadIniValue(App.Path & "\Config.ini", "Private", "SPT")) = 0 Then
    SPT.Value = 0
Else
    If ReadIniValue(App.Path & "\Config.ini", "Private", "SPT") = 0 Then
        SPT.Value = 0
    Else
        SPT.Value = 1
    End If
End If

'Read 'Password' from .ini file
If SPT.Value = 1 Then
    txtPassword.Text = DeCode(DeCode(ReadIniValue(App.Path & "\Config.ini", "Private", "Password")))
Else
    txtPassword.Text = vbNullString
End If
End Sub

Public Sub LoadConfigForm()
Command1.Caption = CONFIG_COMMAND_CONNECT
Command2.Caption = CONFIG_COMMAND_DISCONNECT
Command3.Caption = CONFIG_COMMAND_LANGUAGE
Command4.Caption = CONFIG_COMMAND_UPDATE
Command5.Caption = CONFIG_COMMAND_REGISTER
Frame1.Caption = CONFIG_FRAME_PERSONAL
Frame2.Caption = CONFIG_FRAME_CONNECTION
SPT.Caption = CONFIG_CHECK_SAVE_PASSWORD
lblNickname.Caption = CONFIG_LABEL_CI_NAME
End Sub

Private Function CheckTx(txtBox As TextBox, mBox As String) As Boolean
If Len(txtBox.Text) = 0 Then
    MsgBox mBox, vbInformation
    txtBox.SetFocus
    CheckTx = True
End If
End Function

Private Sub Form_Unload(Cancel As Integer)
'Write data entries to .ini file
WriteIniValue App.Path & "\Config.ini", "Private", "SPT", SPT.Value
WriteIniValue App.Path & "\Config.ini", "Private", "Password", Encode(Encode(txtPassword.Text))
WriteIniValue App.Path & "\Config.ini", "Private", "Account", txtAccount.Text
WriteIniValue App.Path & "\Config.ini", "Private", "Nickname", txtNick.Text
WriteIniValue App.Path & "\Config.ini", "Connection", "Port", txtPort.Text
WriteIniValue App.Path & "\Config.ini", "Connection", "IP", txtIP.Text

'Write position entries to .ini file
WriteIniValue App.Path & "\Config.ini", "Position", "Top", frmMain.Top
WriteIniValue App.Path & "\Config.ini", "Position", "Left", frmMain.Left
End Sub

Private Sub lblAuthor_Click()
If Command1.Enabled = False Then Exit Sub
frmAbout.Show 1
End Sub

Private Sub lblVersion_Click()
If Command1.Enabled = False Then Exit Sub
frmAbout.Show 1
End Sub

Private Sub SPT_Click()
If SPT.Value = 0 Then
    txtPassword.Text = vbNullString
End If
End Sub

Private Sub txtAccount_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
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

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub
