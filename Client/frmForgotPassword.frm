VERSION 5.00
Begin VB.Form frmForgotPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Peach"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3495
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmForgotPassword.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   3495
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRequest 
      Caption         =   "&Request"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtSecretAnswer 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox cmbSecretQuestion 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Forgot Password"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Label lblSecretAnswer 
         Caption         =   " Secret Answer"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Label lblSecretQuestion 
         Caption         =   " Secret Question"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label lblEmail 
         Caption         =   " Enter your email adress"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Label lblStatus 
      Height          =   2415
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmForgotPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRequest_Click()
If cmdRequest.Caption = FP_COMMAND_REQUEST Then
    'Can't request if email empty
    If LenB(txtEmail) = 0 Then
        MsgBox REG_MSG_EMAIL_EMPTY, vbInformation
        txtEmail.SetFocus
        Exit Sub
    End If
    
    'Can't request if secret answer empty
    If LenB(txtSecretAnswer) = 0 Then
        MsgBox REG_MSG_SECRET_ANSWER_EMPTY, vbInformation
        txtSecretAnswer.SetFocus
        Exit Sub
    End If
    
    frmMain.RegSock.SendData "!request_password" & pSplit & Trim$(txtEmail) & pSplit & cmbSecretQuestion.ListIndex & pSplit & txtSecretAnswer & pSplit
Else
    Unload Me
End If

Exit Sub
End Sub

Private Sub Form_Activate()
Dim SC As String
    SC = Setting.SCHEME_COLOR

BackColor = SC
Frame1.BackColor = SC
lblEmail.BackColor = SC
lblSecretQuestion.BackColor = SC
lblSecretAnswer.BackColor = SC
lblStatus.BackColor = SC
End Sub

Private Sub Form_Load()
Frame1.Caption = FP_FRAME_FORGOT_PASSWORD
lblEmail.Caption = FP_LABEL_EMAIL
lblSecretQuestion.Caption = FP_LABEL_SECRET_QUESTION
lblSecretAnswer.Caption = FP_LABEL_SECRET_ANSWER
cmdRequest.Caption = FP_COMMAND_REQUEST

With cmbSecretQuestion
    .List(0) = REG_CMB_SECRET_QUESTION_0
    .List(1) = REG_CMB_SECRET_QUESTION_1
    .List(2) = REG_CMB_SECRET_QUESTION_2
    .List(3) = REG_CMB_SECRET_QUESTION_3
    .List(4) = REG_CMB_SECRET_QUESTION_4
    .List(5) = REG_CMB_SECRET_QUESTION_5
    .ListIndex = 0
End With

With frmMain.RegSock
    .Close
    .RemoteHost = Setting.SERVER_IP
    .RemotePort = rPort
    .Connect
End With

Screen.MousePointer = vbArrowHourglass
Caption = REG_MSG_LOADING
ACC_SWITCH = "FP"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = vbDefault
End Sub

Private Sub txtAccount_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdRequest_Click
End Sub

Private Sub txtSecretAnswer_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdRequest_Click
End Sub
