VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Peach"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8760
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
   ScaleHeight     =   5010
   ScaleWidth      =   8760
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Accept"
      Height          =   375
      Left            =   7080
      TabIndex        =   0
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F4F4F4&
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   5535
   End
   Begin VB.Image Image1 
      Height          =   4110
      Left            =   -120
      Picture         =   "frmAbout.frx":08CA
      Top             =   -120
      Width           =   9000
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

Private Sub Form_Load()
Label1.Caption = "Author : Notron" & vbCrLf _
                    & "Version : " & Rev & vbCrLf & vbCrLf _
                    & "Peach is beeing developed by Notron. Do not publish this anywhere without permissions of the author."
End Sub
