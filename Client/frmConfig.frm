VERSION 5.00
Begin VB.Form frmConfig 
   Appearance      =   0  'Flat
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   Caption         =   "frmConfig"
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4425
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdConnect 
      Caption         =   "cmdConnect"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   2640
      Width           =   1935
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
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Settings"
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
      TabIndex        =   7
      Top             =   2400
      Width           =   1335
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
      Left            =   2880
      MaxLength       =   15
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
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
      Left            =   2880
      MaxLength       =   15
      TabIndex        =   1
      Top             =   720
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
      Left            =   2880
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "&Forgot Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "&Register Account"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label lblNickname 
      BackColor       =   &H00F4F4F4&
      Caption         =   " Name"
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
      Left            =   2850
      TabIndex        =   12
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblAccount 
      BackColor       =   &H00F4F4F4&
      Caption         =   " Account"
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
      Left            =   2880
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblPassword 
      BackColor       =   &H00F4F4F4&
      Caption         =   " Password"
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
      Left            =   2880
      TabIndex        =   11
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00F4F4F4&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label lblAuthor 
      BackColor       =   &H00F4F4F4&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   1455
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub cmdConnect_Click()
If cmdConnect.Caption = CONFIG_COMMAND_CONNECT Then
    'Account can't be empty
    If CheckTx(txtAccount, CONFIG_MSG_ACCOUNT) Then Exit Sub
    
    'Password can't be empty
    If CheckTx(txtPassword, CONFIG_MSG_PASSWORD) Then Exit Sub
    
    'Nick can't be empty
    If CheckTx(txtNick, CONFIG_MSG_NAME) Then Exit Sub
    
    'Nick can't be numeric
    If IsNumeric(txtNick) Then
        MsgBox CONFIG_MSG_NUMERIC, vbInformation
        txtNick.SetFocus
        Exit Sub
    End If
    
    'Nick can't be shorter then 4 characters
    If Len(txtNick) < 4 Then
        MsgBox CONFIG_MSG_NAME_SHORT, vbInformation, "Error - Nickname"
        txtNick.SelStart = Len(txtNick)
        txtNick.SetFocus
        Exit Sub
    End If
    
    'Nick can't contain invalid characters
    If IsInvalid(txtNick) Then
        MsgBox CONFIG_MSG_NAME_INVALID, vbInformation
        txtNick.SelStart = Len(txtNick)
        txtNick.SetFocus
        Exit Sub
    End If
    
    'Make the name proper case
    txtNick = StrConv(txtNick, vbProperCase)
    
    'Connect winsocks
    With frmMain.Winsock1
        .RemotePort = Setting.SERVER_PORT
        .RemoteHost = Setting.SERVER_IP
        .Connect
    End With
    
    If Not Right$(App.EXEName, 5) = "DEBUG" Then
        'Set Recieve-Request-Winsock to listen
        With frmMain.FSocket2(0)
            .LocalPort = aPort
            .Listen
        End With
        
        'Set Recieve-File-Winsock to listen
        With frmSendFile2.SckReceiveFile(0)
            .LocalPort = bPort
            .Listen
        End With
    End If
    
    SwitchButtons False, True
    
    frmMain.StatusBar1.Panels(1).Text = MDI_STAT_CONNECTING
Else
    Disconnect
End If
End Sub

Private Sub Command3_Click()
frmSettings.Show 1
End Sub

Private Sub Command4_Click()
If FileExists(App.Path & "\peachUpdater.exe") Then
    Shell App.Path & "\peachUpdater.exe", vbNormalFocus
Else
    MsgBox CONFIG_MSG_UPDATE_FILE, vbInformation
End If
End Sub

Public Sub Form_Load()
Top = 0: Left = 0
lblAuthor.Caption = "Author : " & pAuthor
lblVersion.Caption = "Version : " & pRev

txtAccount = Setting.ACCOUNT
txtNick = Setting.NICKNAME
txtPassword = Setting.PASSWORD

Command3.Caption = CONFIG_COMMAND_SETTINGS
Command4.Caption = CONFIG_COMMAND_UPDATE
Label1.Caption = CONFIG_COMMAND_REGISTER
Label2.Caption = CONFIG_COMMAND_FORGOT_PASSWORD
cmdConnect.Caption = CONFIG_COMMAND_CONNECT
lblAccount.Caption = CONFIG_LABEL_ACCOUNT
lblPassword.Caption = CONFIG_LABEL_PASSWORD
lblNickname.Caption = CONFIG_LABEL_NAME
End Sub

Private Function CheckTx(txtBox As TextBox, mBox As String) As Boolean
If LenB(Trim$(txtBox)) = 0 Then
    MsgBox mBox, vbInformation
    txtBox.BackColor = vbYellow
    txtBox.SetFocus
    CheckTx = True
End If
End Function

Private Sub Form_Unload(Cancel As Integer)
'Write data entries into registry
InsertIntoRegistry "Client\Configuration", "Password", Encode(Encode(txtPassword))
InsertIntoRegistry "Client\Configuration", "Account", txtAccount
InsertIntoRegistry "Client\Configuration", "Nickname", txtNick

If frmMain.WindowState = 1 Then
    InsertIntoRegistry "Client\Configuration", "Top", Screen.Height / Screen.TwipsPerPixelY
    InsertIntoRegistry "Client\Configuration", "Left", Screen.Width / Screen.TwipsPerPixelX
Else
    InsertIntoRegistry "Client\Configuration", "Top", frmMain.Top
    InsertIntoRegistry "Client\Configuration", "Left", frmMain.Left
End If
End Sub

Private Sub Label1_Click()
frmRegistration.Show 1
End Sub

Private Sub Label2_Click()
frmForgotPassword.Show 1
End Sub

Private Sub lblAuthor_Click()
frmAbout.Show
End Sub

Private Sub lblVersion_Click()
frmAbout.Show
End Sub

Private Sub txtAccount_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmdConnect_Click
    KeyAscii = 0
End If
If txtAccount.BackColor <> vbWhite Then txtAccount.BackColor = vbWhite
End Sub

Private Sub txtNick_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmdConnect_Click
    KeyAscii = 0
End If
If txtNick.BackColor <> vbWhite Then txtNick.BackColor = vbWhite
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmdConnect_Click
    KeyAscii = 0
End If
If txtPassword.BackColor <> vbWhite Then txtPassword.BackColor = vbWhite
End Sub

Private Sub txtNick_LostFocus()
txtNick = StrConv(txtNick, vbProperCase)
End Sub
