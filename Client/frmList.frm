VERSION 5.00
Begin VB.Form frmList 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Online List"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   3765
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   3240
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "Online Users : "
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmMain.UpdateListPosition.Enabled = False
frmList.Hide
End Sub

Public Sub Form_Load()
Me.Caption = LISTcaption
Command1.Caption = LISTcommand_close
End Sub

Private Sub Form_Resize()
If Not frmMain.WindowState = vbMinimized Then
    List1.Height = Me.Height - 1300
    Command1.Top = List1.Height + 400
End If
End Sub

