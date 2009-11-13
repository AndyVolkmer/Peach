VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmRegistration 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Peach - Registration"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegistration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   3855
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbSecretQuestion 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox txtSecretAnswer 
      Height          =   285
      Left            =   240
      TabIndex        =   14
      Top             =   3120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin MSWinsockLib.Winsock RegSock 
      Left            =   1560
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Show Password"
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
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Submit"
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
      Left            =   2040
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Enter your details"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3615
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F4F4F4&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   2160
         ScaleHeight     =   855
         ScaleWidth      =   1215
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
         Begin VB.Label Label5 
            BackColor       =   &H00F4F4F4&
            Height          =   615
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.TextBox txtPassword2 
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
         Left            =   120
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtPassword1 
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
         Left            =   120
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtAccount 
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
         Left            =   120
         MaxLength       =   15
         TabIndex        =   2
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Secret question:"
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
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   3255
      End
      Begin VB.Label Label8 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Secret answer:"
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
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   3255
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Confirm the Password:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Password: "
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
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Account Name: "
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
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H00F4F4F4&
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1 = vbChecked Then
    txtPassword1.PasswordChar = vbNullString
    txtPassword2.PasswordChar = vbNullString
Else
    txtPassword1.PasswordChar = "*"
    txtPassword2.PasswordChar = "*"
End If
End Sub

Private Sub Command1_Click()
'Close form
If Command1.Caption = REG_COMMAND_CLOSE Then
    Unload Me
    Exit Sub
End If

'Can't register if there is no Account.
If Len(Trim$(txtAccount)) = 0 Then
    MsgBox REG_MSG_ACCOUNT_EMPTY, vbInformation
    txtAccount.SetFocus
    Exit Sub
End If

'Can't register if account is shorter then 4 chars.
If Len(txtAccount) < 4 Then
    MsgBox REG_MSG_ACCOUNT_SHORT, vbInformation
    txtAccount.SetFocus
    Exit Sub
End If

'Check for spaces and other signs in the account.
If CheckString(txtAccount) = True Then
    MsgBox REG_MSG_ACCOUNT_INVALID, vbInformation
    txtAccount.SetFocus
    Exit Sub
End If

'Can't register if the Account is made out of numbers.
If IsNumeric(txtAccount) = True Then
    MsgBox REG_MSG_ACCOUNT_NUMERIC, vbInformation
    txtAccount.SetFocus
    Exit Sub
End If

'Can't register if there is no Password.
If Len(txtPassword1) = 0 Then
    MsgBox REG_MSG_PASSWORD_EMPTY, vbInformation
    txtPassword1.SetFocus
    Exit Sub
End If

'Can't register if password is shorter then 6 chars.
If Len(txtPassword1) < 6 Then
    MsgBox REG_MSG_PASSWORD_SHORT, vbInformation
    txtPassword1.SetFocus
    Exit Sub
End If

'Can't register if passwords dont match.
If txtPassword1 <> txtPassword2 Then
    MsgBox REG_MSG_PASSWORD_MATCH, vbInformation
    txtPassword2.SetFocus
    Exit Sub
End If

'Can't register if email is empty.
If Len(Trim$(txtSecretAnswer)) = 0 Then
    MsgBox REG_MSG_SECRET_ANSWER_EMPTY, vbInformation
    txtSecretAnswer.SetFocus
    Exit Sub
End If

'Can't register if email is to shorter then 8 chars.
If Len(txtSecretAnswer) < 8 Then
    MsgBox REG_MSG_SECRET_ANSWER_SHORT, vbInformation
    txtSecretAnswer.SetFocus
    Exit Sub
End If

'Can't register if the winsock is not connected.
If RegSock.State <> 7 Then
    MsgBox REG_MSG_CONNECTION_BROKEN, vbInformation
    Exit Sub
End If

RegSock.SendData "!register#" & txtAccount & "#" & txtPassword1 & "#" & cmbSecretQuestion.ListIndex & "#" & txtSecretAnswer & "#"
End Sub

Private Sub Form_Activate()
Dim SC As String
SC = Setting.SCHEME_COLOR
BackColor = SC
Check1.BackColor = SC
Frame1.BackColor = SC
Label1.BackColor = SC
Label2.BackColor = SC
Label3.BackColor = SC
Label4.BackColor = SC
Label5.BackColor = SC
Label7.BackColor = SC
Label8.BackColor = SC
Picture1.BackColor = SC
End Sub

Private Sub Form_Load()
LoadRegistrationForm
With RegSock
    .RemoteHost = Setting.SERVER_IP
    .RemotePort = rPort
    .Connect
End With
Screen.MousePointer = vbArrowHourglass
Me.Caption = REG_MSG_LOADING
End Sub

Public Sub LoadRegistrationForm()
Frame1.Caption = REG_FRAME_DETAIL
Label1.Caption = REG_LABEL_ACCOUNT_NAME
Label2.Caption = REG_LABEL_PASSWORD
Label3.Caption = REG_LABEL_PASSWORD_CONFIRM
Check1.Caption = REG_CHECK_PASSWORD_SHOW
Command1.Caption = REG_COMMAND_CLOSE
With cmbSecretQuestion
    .List(0) = REG_CMB_SECRET_QUESTION_0
    .List(1) = REG_CMB_SECRET_QUESTION_1
    .List(2) = REG_CMB_SECRET_QUESTION_2
    .List(3) = REG_CMB_SECRET_QUESTION_3
    .List(4) = REG_CMB_SECRET_QUESTION_4
    .List(5) = REG_CMB_SECRET_QUESTION_5
    .ListIndex = 0
End With
Label7.Caption = REG_LABEL_SECRET_QUESTION
Label8.Caption = REG_LABEL_SECRET_ANSWER
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = vbDefault
End Sub

Private Sub RegSock_Close()
Me.Caption = REG_MSG_ERROR_OCCURED
Label4.Caption = REG_MSG_ERROR
Command1.Caption = REG_COMMAND_CLOSE
Frame1.Visible = False
Check1.Visible = False
cmbSecretQuestion.Visible = False
txtSecretAnswer.Visible = False
Screen.MousePointer = vbDefault
End Sub

Private Sub RegSock_Connect()
Me.Caption = " Peach - Registration"
Frame1.Visible = True
Check1.Visible = True
Command1.Visible = True
Command1.Caption = REG_COMMAND_SUBMIT
Screen.MousePointer = vbDefault
End Sub

Private Sub RegSock_DataArrival(ByVal bytesTotal As Long)
Dim GetMessage As String
Dim array1() As String

RegSock.GetData GetMessage

array1 = Split(GetMessage, "#")

Select Case array1(0)
Case "!nameexist"
    MsgBox REG_MSG_ACCOUNT_EXIST, vbInformation
    txtAccount = vbNullString
    txtAccount.SetFocus
Case "!done"
    MsgBox REG_MSG_SUCCESSFULLY, vbInformation
    Unload Me
End Select
End Sub

Private Sub RegSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Me.Caption = REG_MSG_ERROR_OCCURED
Label4.Caption = REG_MSG_ERROR
Command1.Caption = REG_COMMAND_CLOSE
Command1.Visible = True
Screen.MousePointer = vbDefault
End Sub

Private Sub txtAccount_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub

Private Sub txtPassword1_Change()
Select Case Len(txtPassword1)
Case Is < 7
    Picture1.BackColor = vbRed
    Picture1.BorderStyle = 1
    Label5.BackColor = vbRed
    Label5.ForeColor = vbWhite
    Label5.Caption = REG_LABEL_PASSWORD_WEAK
Case 7, 8, 9
    Picture1.BackColor = vbYellow
    Picture1.BorderStyle = 1
    Label5.BackColor = vbYellow
    Label5.ForeColor = vbBlack
    Label5.Caption = REG_LABEL_PASSWORD_NORMAL
Case Is > 10
    Picture1.BackColor = vbGreen
    Picture1.BorderStyle = 1
    Label5.BackColor = vbGreen
    Label5.ForeColor = vbBlack
    Label5.Caption = REG_LABEL_PASSWORD_STRONG
End Select
End Sub

Private Sub txtPassword1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub

Private Sub txtPassword2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub
