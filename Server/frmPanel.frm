VERSION 5.00
Begin VB.Form frmPanel 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7515
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   810
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get all Connections"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   6855
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   6855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Kick"
      Height          =   405
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   240
      TabIndex        =   0
      Text            =   "0"
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text1.Text = "" Then Exit Sub
    If Text1.Text = "0" Then Exit Sub
    
    ' Disconnect the socket
    frmMain.Winsock1(Text1.Text).Close
    Unload frmMain.Winsock1(Text1.Text)
    
    ' Clear entry
    Text1.Text = "0"
    
    ' Update Statusbar
    frmMain.StatusBar1.Panels(1).Text = "Status: Connected with  " & frmMain.Winsock1.Count - 1 & " Client(s)."
    Command2_Click
End Sub

Private Sub Command2_Click()
Dim Wsk As Winsock

List1.Clear
List2.Clear

For Each Wsk In frmMain.Winsock1
    If Wsk.State = sckConnected Then
        List1.AddItem "IP : " & Wsk.RemoteHostIP & " | " & "Number : " & Wsk.Index
        List2.AddItem Wsk.Index
    End If
Next

End Sub

Private Sub Form_Load()
Me.Top = "0"
Me.Left = "0"
End Sub

Private Sub List1_Click()
    List2.ListIndex = List1.ListIndex
End Sub

Private Sub List2_Click()
    Text1.Text = List2.List(List2.ListIndex)
End Sub
