VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Peach"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   5790
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Accept"
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
      Left            =   4080
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F4F4F4&
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Me.BackColor = Setting.SCHEME_COLOR
Label1.BackColor = Setting.SCHEME_COLOR
End Sub

Private Sub Form_Load()
Label1.Caption = "Author : Andy V." & vbCrLf _
            & "Version : " & Rev & vbCrLf & vbCrLf _
            & "Peach is beeing developed by Andy V. " & _
            "Do not publish this anywhere without " & _
            "permissions of the author."
End Sub
