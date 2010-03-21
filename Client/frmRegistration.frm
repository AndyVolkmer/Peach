VERSION 5.00
Begin VB.Form frmCreateAccount 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Peach - Registration"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3870
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   3870
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbGender 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   3840
      Visible         =   0   'False
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
      Left            =   240
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   240
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
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
      Left            =   240
      MaxLength       =   15
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.ComboBox cmbSecretQuestion 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox txtSecretAnswer 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   3255
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
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4440
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
      TabIndex        =   7
      Top             =   4440
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
      Height          =   4215
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
         TabIndex        =   12
         Top             =   1200
         Width           =   1215
         Begin VB.Label Label5 
            BackColor       =   &H00F4F4F4&
            Height          =   615
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Label Label9 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Gender:"
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
         TabIndex        =   17
         Top             =   3480
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   2880
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H00F4F4F4&
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmCreateAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
If Check1.Value = 1 Then
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
If LenB(Trim$(txtAccount)) = 0 Then
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
If IsInvalid(txtAccount) Then
    MsgBox REG_MSG_ACCOUNT_INVALID, vbInformation
    txtAccount.SetFocus
    Exit Sub
End If

'Can't register if the Account is made out of numbers.
If IsNumeric(txtAccount) Then
    MsgBox REG_MSG_ACCOUNT_NUMERIC, vbInformation
    txtAccount.SetFocus
    Exit Sub
End If

txtAccount.Text = StrConv(txtAccount.Text, vbProperCase)

'Can't register if there is no Password.
If LenB(txtPassword1) = 0 Then
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
If LenB(Trim$(txtSecretAnswer)) = 0 Then
    MsgBox REG_MSG_SECRET_ANSWER_EMPTY, vbInformation
    txtSecretAnswer.SetFocus
    Exit Sub
End If

'Can't register if the winsock is not connected.
If frmMain.RegSock.State <> 7 Then
    MsgBox REG_MSG_CONNECTION_BROKEN, vbInformation
    Exit Sub
End If

frmMain.RegSock.SendData "!register#" & txtAccount & "#" & txtPassword1 & "#" & cmbSecretQuestion.ListIndex & "#" & txtSecretAnswer & "#" & cmbGender.ListIndex & "#"
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
Label9.BackColor = SC
Picture1.BackColor = SC
End Sub

Private Sub Form_Load()
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
Me.Caption = REG_MSG_LOADING
Label9.Caption = REG_LABEL_GENDER

With cmbGender
    .List(0) = REG_CMB_GENDER_MALE
    .List(1) = REG_CMB_GENDER_FEMALE
    .ListIndex = 0
End With

With frmMain.RegSock
    .Close
    .RemoteHost = Setting.SERVER_IP
    .RemotePort = rPort
    .Connect
End With

Screen.MousePointer = vbArrowHourglass
ACC_SWITCH = "REG"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = vbDefault
End Sub

Private Sub txtAccount_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub

Private Sub txtAccount_LostFocus()
txtAccount.Text = StrConv(txtAccount.Text, vbProperCase)
End Sub

Private Sub txtPassword1_Change()
Select Case Len(txtPassword1)
    Case Is < 5
        Picture1.BackColor = vbRed
        Picture1.BorderStyle = 1
        Label5.BackColor = vbRed
        Label5.ForeColor = vbWhite
        Label5.Caption = REG_LABEL_PASSWORD_WEAK
        
    Case 5, 6, 7
        Picture1.BackColor = vbYellow
        Picture1.BorderStyle = 1
        Label5.BackColor = vbYellow
        Label5.ForeColor = vbBlack
        Label5.Caption = REG_LABEL_PASSWORD_NORMAL
        
    Case Is > 8
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

Private Sub txtSecretAnswer_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub
