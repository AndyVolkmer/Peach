VERSION 5.00
Begin VB.Form frmChat 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   Caption         =   "frmChat"
   ClientHeight    =   5220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7635
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
   ScaleHeight     =   5220
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtToSend 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   5535
   End
   Begin VB.TextBox txtConver 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   7215
   End
   Begin VB.CommandButton txtClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSend_Click()
If Len(Trim$(txtToSend)) = 0 Then
    txtToSend = vbNullString
    Exit Sub
End If

SendMessage "Server Notice: " & txtToSend
CMSG txtToSend
txtToSend = vbNullString
End Sub

Private Sub Form_Load()
Me.Top = 0: Me.Left = 0
End Sub

Private Sub txtClear_Click()
txtConver = vbNullString
End Sub

Private Sub txtToSend_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmdSend_Click
    KeyAscii = 0
End If
End Sub
