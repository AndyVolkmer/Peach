VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.OCX"
Begin VB.Form frmChat 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "frmChat"
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7890
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
   ScaleHeight     =   4380
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox txtToSend 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1508
      _Version        =   393217
      Enabled         =   0   'False
      MultiLine       =   0   'False
      MaxLength       =   180
      TextRTF         =   $"frmChat.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtConver 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4471
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmChat.frx":007B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSend_Click()
frmMain.Prefix = "[" & Format(Time, "hh:nn:ss") & "]"
Select Case txtToSend.Text
Case ""
    MsgBox "Nothing inserted!", vbInformation
    txtToSend.Text = ""
    txtToSend.SetFocus
    Exit Sub
Case " "
    MsgBox "Nothing inserted!", vbInformation
    txtToSend.Text = ""
    txtToSend.SetFocus
    Exit Sub
Case "  "
    MsgBox "Nothing inserted!", vbInformation
    txtToSend.Text = ""
    txtToSend.SetFocus
    Exit Sub
Case "!time"
    txtConver.Text = txtConver.Text & vbCrLf & frmMain.Prefix & " [" & frmConfig.txtNick.Text & "] : " & txtToSend.Text & vbCrLf & frmMain.Prefix & " [System] : The time is " & Format(Time, "hh:nn:ss")
    txtToSend.Text = ""
Case "!online"
    txtConver.Text = txtConver.Text & vbCrLf & frmMain.Prefix & " [" & frmConfig.txtNick.Text & "] : " & txtToSend.Text
    
    With frmMain
        .ConverText = txtToSend.Text
        .NameText = frmConfig.txtNick.Text
        .Message = .NameText & "#" & .ConverText & "#"
        
    SendMessage .Message
    End With
    
    txtToSend.Text = ""
Case Else
    
    'Message = " [" & frmConfig.txtNick.Text & "] : " & txtToSend.Text
    
    With frmMain
        .ConverText = txtToSend.Text
        .NameText = frmConfig.txtNick.Text
        .Message = .NameText & "#" & .ConverText & "#"
        
    SendMessage .Message
    End With
            
    txtToSend.Text = ""
End Select
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
End Sub

Private Sub cmdClear_Click()
txtConver.Text = ""
txtToSend.Text = ""
End Sub

Private Sub txtConver_Change()
txtConver.SelStart = Len(txtConver)
If frmMain.WindowState = vbMinimized Then
    Call FlashTitle(frmMain.hwnd, True)
Else
End If
End Sub

Private Sub txtToSend_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdSend_Click
End Sub
