VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmRegistration 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Peach - Registration"
   ClientHeight    =   3165
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
   MaxButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   3870
   StartUpPosition =   1  'CenterOwner
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
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock RegSock 
      Left            =   120
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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
      Top             =   2640
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
      Height          =   2415
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
      Height          =   255
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
If Command1.Caption = "&Close" Then
    Unload Me
    Exit Sub
End If

'Can't register if there is no Account
If Len(txtAccount.Text) = 0 Then
    MsgBox "No account entered.", vbInformation
    txtAccount.SetFocus
    Exit Sub
End If

'Can't register if account is shorter then 4 chars.
If Len(txtAccount.Text) < 4 Then
    MsgBox "Account name to short, it requieres at least 4 characters.", vbInformation
    txtAccount.SetFocus
    Exit Sub
End If

'Check for spaces and other signs in the account
With txtAccount
    If CheckString(.Text, " ") = True Then Exit Sub
    If CheckString(.Text, ".") = True Then Exit Sub
    If CheckString(.Text, "*") = True Then Exit Sub
    If CheckString(.Text, "#") = True Then Exit Sub
    If CheckString(.Text, "{") = True Then Exit Sub
    If CheckString(.Text, "}") = True Then Exit Sub
    If CheckString(.Text, ",") = True Then Exit Sub
    If CheckString(.Text, "(") = True Then Exit Sub
    If CheckString(.Text, ")") = True Then Exit Sub
    If CheckString(.Text, "&") = True Then Exit Sub
    If CheckString(.Text, "!") = True Then Exit Sub
    If CheckString(.Text, "@") = True Then Exit Sub
    If CheckString(.Text, "?") = True Then Exit Sub
    If CheckString(.Text, "/") = True Then Exit Sub
    If CheckString(.Text, "�") = True Then Exit Sub
    If CheckString(.Text, "=") = True Then Exit Sub
    If CheckString(.Text, "<") = True Then Exit Sub
    If CheckString(.Text, ">") = True Then Exit Sub
    If CheckString(.Text, "[") = True Then Exit Sub
    If CheckString(.Text, "]") = True Then Exit Sub
    If CheckString(.Text, "'") = True Then Exit Sub
    If CheckString(.Text, "�") = True Then Exit Sub
    If CheckString(.Text, "�") = True Then Exit Sub
    If CheckString(.Text, "�") = True Then Exit Sub
    If CheckString(.Text, "\") = True Then Exit Sub
    If CheckString(.Text, "|") = True Then Exit Sub
    If CheckString(.Text, "~") = True Then Exit Sub
    If CheckString(.Text, "�") = True Then Exit Sub
    If CheckString(.Text, "`") = True Then Exit Sub
    If CheckString(.Text, "+") = True Then Exit Sub
    If CheckString(.Text, "-") = True Then Exit Sub
    If CheckString(.Text, "^") = True Then Exit Sub
    If CheckString(.Text, "_") = True Then Exit Sub
    If CheckString(.Text, "�") = True Then Exit Sub
End With

'Can't register if the Account is made out of numbers
If IsNumeric(txtAccount.Text) = True Then
    MsgBox "Account can't be made of numeric characters.", vbInformation
    txtAccount.SetFocus
    Exit Sub
End If

'Can't register if there is no Password
If Len(txtPassword1.Text) = 0 Then
    MsgBox "No Password entered.", vbInformation
    txtPassword1.SetFocus
    Exit Sub
End If

'Can't register if password is shorter then 6 chars.
If Len(txtPassword1.Text) < 6 Then
    MsgBox "Password to short, it requieres at least 6 characters."
    txtPassword1.SetFocus
    Exit Sub
End If

'Can't register if passwords dont match
If txtPassword1.Text <> txtPassword2.Text Then
    MsgBox "Your Passwords don't match.", vbInformation
    txtPassword2.SetFocus
    Exit Sub
End If

'Can't register if the winsock is not connected
If RegSock.State <> 7 Then
    MsgBox "Connection is broken please try again later.", vbInformation
    Exit Sub
End If

RegSock.SendData "!register" & "#" & txtAccount.Text & "#" & txtPassword1.Text & "#"
End Sub

Private Function CheckString(pString As String, pChar As String) As Boolean
If InStr(1, pString, pChar) <> 0 Then
    MsgBox "Invalid account name.", vbInformation
    With txtAccount
        .SetFocus
        .Text = vbNullString
    End With
    CheckString = True
End If
End Function

Private Sub Form_Load()
With RegSock
    .RemoteHost = frmConfig.txtIP.Text
    .RemotePort = rPort
    .Connect
End With
Screen.MousePointer = vbArrowHourglass
Me.Caption = " Loading ..."
End Sub

Private Sub Form_Unload(Cancel As Integer)
RegSock.Close
End Sub

Private Sub RegSock_Close()
Me.Caption = "Error has occured ..."
Label4.Caption = "An error has occured please try later again."
Command1.Caption = "&Close"
Command1.Visible = False
Frame1.Visible = False
Check1.Visible = False
Screen.MousePointer = vbDefault
End Sub

Private Sub RegSock_Connect()
Me.Caption = " Peach - Registration"
Frame1.Visible = True
Check1.Visible = True
Command1.Visible = True
Screen.MousePointer = vbDefault
End Sub

Private Sub RegSock_DataArrival(ByVal bytesTotal As Long)
Dim GetMessage As String
Dim Array1() As String

RegSock.GetData GetMessage

Array1 = Split(GetMessage, "#")

Select Case Array1(0)
Case "!nameexist"
    MsgBox "The account name already exists.", vbInformation
    txtAccount.Text = vbNullString
    txtAccount.SetFocus
Case "!done"
    MsgBox "The account was successfully registered.", vbInformation
    Unload Me
End Select
End Sub

Private Sub RegSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Me.Caption = "Error has occured ..."
Label4.Caption = "An error has occured please try later again."
Command1.Caption = "&Close"
Command1.Visible = True
Screen.MousePointer = vbDefault
End Sub

Private Sub txtAccount_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub

Private Sub txtPassword1_Change()
Select Case Len(txtPassword1.Text)
Case Is < 6
    Picture1.BackColor = vbRed
    Picture1.BorderStyle = 1
    Label5.BackColor = vbRed
    Label5.ForeColor = vbWhite
    Label5.Caption = "The password is weak."
Case 6, 7, 8
    Picture1.BackColor = vbYellow
    Picture1.BorderStyle = 1
    Label5.BackColor = vbYellow
    Label5.ForeColor = vbBlack
    Label5.Caption = "The password is normal."
Case Is > 9
    Picture1.BackColor = vbGreen
    Picture1.BorderStyle = 1
    Label5.BackColor = vbGreen
    Label5.ForeColor = vbBlack
    Label5.Caption = "The password is strong."
End Select
End Sub

Private Sub txtPassword1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub

Private Sub txtPassword2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub
