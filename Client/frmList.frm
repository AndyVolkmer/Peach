VERSION 5.00
Begin VB.Form frmList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Online List"
   ClientHeight    =   5115
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   3765
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   4680
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Online Users : "
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
frmList.Hide
End Sub

Private Sub Form_Load()
'::::::::::::::::::
' Add item into list from an variable where we saved the getdata value from server..

End Sub

Private Sub Form_Resize()
List1.Height = Me.Height - 1400
Command1.Top = List1.Height + 500
End Sub

