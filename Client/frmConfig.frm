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
   MDIChild        =   -1  'True
   ScaleHeight     =   4185
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Server Information:"
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   3720
      TabIndex        =   13
      Top             =   360
      Width           =   3015
      Begin VB.Label Label5 
         BackColor       =   &H00F4F4F4&
         Caption         =   "IP: "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackColor       =   &H00F4F4F4&
         Caption         =   "Port: "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Client Information:"
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   360
      TabIndex        =   12
      Top             =   360
      Width           =   3015
      Begin VB.Label Label3 
         BackColor       =   &H00F4F4F4&
         Caption         =   "IP: "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00F4F4F4&
         Caption         =   "Name: "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Disconnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Language"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Connect"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Configuration"
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7215
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
         BackColor       =   &H00F4F4F4&
         Caption         =   " Port:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F4F4F4&
         Caption         =   " IP:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblNick 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Nickname:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   975
      End
   End
   Begin VB.Label Label8 
      BackColor       =   &H00F4F4F4&
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   1455
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
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'Nick cant be empty
If txtNick.Text = "" Then
    MsgBox CONFIGmsgbox_namenoempty, vbInformation
    txtNick.SetFocus
    Exit Sub
End If

'IP cant be empty
If txtIP.Text = "" Then
    MsgBox CONFIGmsgbox_ipnoempty, vbInformation
    txtIP.SetFocus
    Exit Sub
End If

'Port cant be empty
If txtPort.Text = "" Then
    MsgBox CONFIGmsgbox_portnoempty, vbInformation
    txtPort.SetFocus
    Exit Sub
End If

'If the nick is numeric then no
If IsNumeric(txtNick.Text) = True Then
    txtNick.Text = ""
    MsgBox CONFIGmsgbox_nonumeric, vbInformation
    txtNick.SetFocus
    Exit Sub
End If

'If the nick is to short then no
If Len(txtNick.Text) < 4 Then
    MsgBox "Your nickname is to short!    ", vbInformation, " Error - Nickname"
    txtNick.SelStart = Len(txtNick.Text)
    txtNick.SetFocus
    Exit Sub
End If

'Make it proper case
txtNick.Text = StrConv(txtNick.Text, vbProperCase)

'Do the enable stuff
With Me
    .Command2.Enabled = True
    .Command1.Enabled = False
    .txtNick.Enabled = False
    .txtIP.Enabled = False
    .txtPort.Enabled = False
End With

With frmChat
    .cmdSend.Enabled = True
    .cmdClear.Enabled = True
    .txtToSend.Enabled = True
End With

'Connect winsocks
With frmMain.Winsock1
    .RemotePort = txtPort.Text
    .RemoteHost = txtIP.Text
    .Connect
    Label5.Caption = "IP: " & .RemoteHost
    Label6.Caption = "Port: " & .RemotePort
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

frmMain.StatusBar1.Panels(1).Text = MDIstatusbar_connecting & txtIP.Text & ":" & txtPort.Text
frmConfig.Hide
frmChat.Show
frmChat.txtToSend.SetFocus
Exit Sub
End Sub

Public Sub Command2_Click()
Dim WiSk As Winsock

' Do the buttons
Command1.Enabled = True
Command2.Enabled = False
txtNick.Enabled = True
txtIP.Enabled = True
txtPort.Enabled = True
Label5.Caption = "IP: "
Label6.Caption = "Port: "

With frmChat
    .cmdSend.Enabled = False
    .cmdClear.Enabled = False
    .txtToSend.Enabled = False
End With

'Clear the online user list
frmList.ListView1.ListItems.Clear
frmSendFile.Combo1.Clear

With frmMain
    For Each WiSk In .FSocket2
        If WiSk.State = 7 Then
            WiSk.Close
            Unload WiSk
        End If
    Next
    .FSocket2(0).Close
End With


'Close and unload all connected winsocks
With frmSendFile2
    For Each WiSk In .SckReceiveFile
        If WiSk.State = 7 Then
            WiSk.Close
            Unload WiSk
        End If
    Next
    .SckReceiveFile(0).Close
End With

'Close connection
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

Public Sub Form_Load()
Dim TSSO As TypeSSO
With Me
    .Top = 0
    .Left = 0
    .Label3.Caption = "IP: " & frmMain.Winsock1.LocalIP
    .Label4.Caption = "Name: " & frmMain.Winsock1.LocalHostName
    .Label8.Caption = "Version : " & Rev
End With
LoadConfigForm
TSSO = ReadConfigFile(App.Path & "\bin.conf")
With TSSO
    txtNick.Text = Trim(TSSO.Nickname)
    txtIP.Text = Trim(TSSO.ConnectIP)
    txtPort.Text = Trim(TSSO.Port)
End With
End Sub

Public Sub LoadConfigForm()
Command1.Caption = CONFIGcommand_connect
Command2.Caption = CONFIGcommand_disconnect
Command3.Caption = CONFIGcommand_language
Label4.Caption = CONFIGlabel_CI_name & frmMain.Winsock1.LocalHostName
Frame1.Caption = CONFIGframe_config
Frame2.Caption = CONFIGframe_client
Frame3.Caption = CONFIGframe_server
lblNick = CONFIGlabel_CI_name
End Sub

Private Sub Form_Unload(Cancel As Integer)
WriteConfigFile (App.Path & "\bin.conf")
End Sub

Private Sub Label7_Click()
frmAbout.Show 1
End Sub

Private Sub Label8_Click()
frmAbout.Show 1
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
