VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Peach Server"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2925
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   2925
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Frame frameDatabase 
      Caption         =   "Database"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.TextBox txtHost 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtDatabase 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblHost 
         Caption         =   "Host"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblUser 
         Caption         =   "User"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblDatabase 
         Caption         =   "Database"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSave_Click()
With Database
    .Database = txtDatabase.Text
    .Host = txtHost.Text
    .Password = txtPassword.Text
    .User = txtUser.Text
    
InsertIntoRegistry "Server\Database", "Name", .Database
InsertIntoRegistry "Server\Database", "User", .User
InsertIntoRegistry "Server\Database", "Password", Encode(.Password)
InsertIntoRegistry "Server\Database", "Host", .Host
End With

Unload Me
End Sub

Private Sub Form_Load()
With Database
    txtDatabase.Text = .Database
    txtHost.Text = .Host
    txtUser.Text = .User
    txtPassword.Text = .Password
    txtHost.Text = .Host
End With
End Sub
