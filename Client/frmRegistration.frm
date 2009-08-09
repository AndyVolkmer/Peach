VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmRegistration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Peach - Registration"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3885
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
   ScaleWidth      =   3885
   StartUpPosition =   1  'CenterOwner
   Begin MSWinsockLib.Winsock RegSock 
      Left            =   120
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Submit"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter your details."
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3615
      Begin VB.TextBox txtPassword2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtPassword1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtAccount 
         Height          =   285
         Left            =   120
         MaxLength       =   15
         TabIndex        =   2
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   " Confirm the Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   " Password: "
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   " Account Name: "
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label Label4 
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
Private Sub Command1_Click()
'Can't register if there is no Account
If txtAccount.Text = "" Then
    MsgBox "You have to enter an Account Name.", vbInformation
    txtAccount.SetFocus
    Exit Sub
End If

'Can't register if account is shorter then 5 chars.
If Len(txtAccount.Text) < 5 Then
    MsgBox "Your Account needs at least 5 characters.", vbInformation
    txtAccount.SetFocus
    Exit Sub
End If

'Can't register if the Account is made out of numbers
If IsNumeric(txtAccount.Text) = True Then
    MsgBox "Your Account can't be made of numeric characters.", vbInformation
    txtAccount.SetFocus
    Exit Sub
End If

'Can't register if there is no Password
If txtPassword1.Text = "" Then
    MsgBox "You have to enter a Password.", vbInformation
    txtPassword1.SetFocus
    Exit Sub
End If

'Can't register if password is shorter then 4 chars.
If Len(txtPassword1.Text) < 4 Then
    MsgBox "Your Password needs at least 4 characters."
    txtPassword1.SetFocus
    Exit Sub
End If

'Can't register if passwords dont match
If txtPassword1.Text <> txtPassword2.Text Then
    MsgBox "Your Passwords don't match.", vbInformation
    txtPassword1.Text = ""
    txtPassword2.Text = ""
    txtPassword1.SetFocus
    Exit Sub
End If

'Can't register if the winsock is not connected
If RegSock.State <> 7 Then
    MsgBox "Connection is broken please try again later.", vbInformation
    Exit Sub
End If

RegSock.SendData "!register" & "#" & txtAccount.Text & "#" & txtPassword1.Text & "#"
End Sub

Private Sub Form_Load()
With RegSock
    .RemoteHost = frmConfig.txtIP.Text
    .RemotePort = RegPort
    .Connect
End With
Me.Caption = " Loading ..."
End Sub

Private Sub Form_Unload(Cancel As Integer)
RegSock.Close
End Sub

Private Sub RegSock_Close()
Me.Caption = "Error has occured ..."
Label4.Caption = "An error has occured please try later again."
Frame1.Visible = False
Command1.Visible = False
End Sub

Private Sub RegSock_Connect()
Me.Caption = " Peach - Registration"
Frame1.Visible = True
Command1.Visible = True
End Sub

Private Sub RegSock_DataArrival(ByVal bytesTotal As Long)
Dim GetMessage As String
Dim Array1() As String

RegSock.GetData GetMessage

Array1 = Split(GetMessage, "#")

Select Case Array1(0)
Case "!nameexist"
    MsgBox "The account name already exists.", vbInformation
    txtAccount.Text = ""
    txtAccount.SetFocus
Case "!done"
    MsgBox "The account was successfully registered.", vbInformation
    Unload Me
End Select
End Sub

Private Sub RegSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Me.Caption = "Error has occured ..."
Label4.Caption = "An error has occured please try later again."
End Sub

Private Sub txtAccount_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub

Private Sub txtPassword1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub

Private Sub txtPassword2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub
