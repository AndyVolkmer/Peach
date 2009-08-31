VERSION 5.00
Begin VB.Form frmConfig 
   Appearance      =   0  'Flat
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   Caption         =   "frmConfig"
   ClientHeight    =   4185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4185
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Register Account"
      Height          =   320
      Left            =   5280
      TabIndex        =   18
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Update"
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
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Language"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Connect"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Personal Settings"
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   7215
      Begin VB.CheckBox SPT 
         BackColor       =   &H00F4F4F4&
         Caption         =   "Save Password"
         Height          =   255
         Left            =   5280
         TabIndex        =   19
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   240
         MaxLength       =   15
         TabIndex        =   0
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtNick 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblNickname 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Nickname:"
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
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   7215
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblPort 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Port:"
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
If CheckTx(txtAccount, CONFIGmsgbox_account) = True Then Exit Sub

'Password can't be empty
If CheckTx(txtPassword, CONFIGmsgbox_password) = True Then Exit Sub

'Nick can't be empty
If CheckTx(txtNick, CONFIGmsgbox_namenoempty) = True Then Exit Sub

'IP can't be empty
If CheckTx(txtIP, CONFIGmsgbox_ipnoempty) = True Then Exit Sub

'Port can't be empty
If CheckTx(txtPort, CONFIGmsgbox_portnoempty) = True Then Exit Sub

'Nick can't be numeric
If IsNumeric(txtNick.Text) = True Then
    txtNick.Text = ""
    MsgBox CONFIGmsgbox_nonumeric, vbInformation
    txtNick.SetFocus
    Exit Sub
End If

'Nick can't be shorter then 4 characters
If Len(txtNick.Text) < 4 Then
    MsgBox "Your nickname is to short!    ", vbInformation, " Error - Nickname"
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

frmMain.StatusBar1.Panels(1).Text = MDIstatusbar_connecting & txtIP.Text & ":" & txtPort.Text
End Sub

Public Sub Command2_Click()
Disconnect
With frmMain
    .Winsock1.Close
    .StatusBar1.Panels(1).Text = MDIstatusbar_disconnected
End With
End Sub

Private Sub Command3_Click()
frmMain.Hide
frmList.Hide
frmLanguage.Label1 = CONFIGlabel_selectlanguage
With frmLanguage.Combo1
    .Clear
    .AddItem CONFIGcombo_german
    .AddItem CONFIGcombo_english
    .AddItem CONFIGcombo_spanish
    .AddItem CONFIGcombo_swedish
    .AddItem CONFIGcombo_italian
    .AddItem CONFIGcombo_dutch
    .AddItem CONFIGcombo_serbian
    .AddItem CONFIGcombo_french
End With
frmLanguage.Show
End Sub

Private Sub Command4_Click()
On Error Resume Next
Shell App.Path & "\peachUpdater.exe", vbNormalFocus
End Sub

Private Sub Command5_Click()
frmRegistration.Show 1
End Sub

Public Sub Form_Load()
Top = 0
Left = 0
lblVersion.Caption = "Version : " & Rev

LoadConfigForm

'Read 'Nickname' from .ini file
If ReadIniValue(App.Path & "\Config.ini", "Data", "Nickname") = "" Then
    txtNick.Text = "Nickname"
Else
    txtNick.Text = ReadIniValue(App.Path & "\Config.ini", "Data", "Nickname")
End If

'Read 'IP' from .ini file
If ReadIniValue(App.Path & "\Config.ini", "Data", "IP") = "" Then
    txtIP.Text = "0.0.0.0"
Else
    txtIP.Text = ReadIniValue(App.Path & "\Config.ini", "Data", "IP")
End If

'Read 'Port' from .ini file
If ReadIniValue(App.Path & "\Config.ini", "Data", "Port") = "" Then
    txtPort.Text = "4728"
Else
    txtPort.Text = ReadIniValue(App.Path & "\Config.ini", "Data", "Port")
End If

'Read 'Account' from .ini file
If ReadIniValue(App.Path & "\Config.ini", "Data", "Account") = "" Then
    txtAccount.Text = "Your Account"
Else
    txtAccount.Text = ReadIniValue(App.Path & "\Config.ini", "Data", "Account")
End If

'Read 'Save Password Tick'
If ReadIniValue(App.Path & "\Config.ini", "Data", "SPT") = "" Then
    SPT.Value = 0
Else
    If ReadIniValue(App.Path & "\Config.ini", "Data", "SPT") = "0" Then
        SPT.Value = 0
    Else
        SPT.Value = 1
    End If
End If

'Read 'Password' from .ini file
If SPT.Value = 1 Then
    txtPassword.Text = DeCode(DeCode(ReadIniValue(App.Path & "\Config.ini", "Data", "Password")))
Else
    txtPassword.Text = ""
End If
End Sub

Public Sub LoadConfigForm()
Command1.Caption = CONFIGcommand_connect
Command2.Caption = CONFIGcommand_disconnect
Command3.Caption = CONFIGcommand_language
Frame1.Caption = CONFIGframe_personal
Frame2.Caption = CONFIGframe_connection
lblNickname.Caption = CONFIGlabel_CI_name
End Sub

Private Function CheckTx(txtBox As TextBox, mBox As String) As Boolean
If txtBox.Text = "" Then
    MsgBox mBox, vbInformation
    txtBox.SetFocus
    CheckTx = True
End If
End Function

Private Sub Form_Unload(Cancel As Integer)
'Write data entries to .ini file
WriteIniValue App.Path & "\Config.ini", "Data", "SPT", SPT.Value
WriteIniValue App.Path & "\Config.ini", "Data", "Password", Encode(Encode(txtPassword.Text))
WriteIniValue App.Path & "\Config.ini", "Data", "Account", txtAccount.Text
WriteIniValue App.Path & "\Config.ini", "Data", "Port", txtPort.Text
WriteIniValue App.Path & "\Config.ini", "Data", "IP", txtIP.Text
WriteIniValue App.Path & "\Config.ini", "Data", "Nickname", txtNick.Text

'Write position entries to .ini file
WriteIniValue App.Path & "\Config.ini", "Position", "Top", frmMain.Top
WriteIniValue App.Path & "\Config.ini", "Position", "Left", frmMain.Left
End Sub

Private Sub Label7_Click()
frmAbout.Show 1
End Sub

Private Sub Label8_Click()
frmAbout.Show 1
End Sub

Private Sub SPT_Click()
If SPT.Value = 0 Then
    txtPassword.Text = ""
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
