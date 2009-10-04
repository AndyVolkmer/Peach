VERSION 5.00
Begin VB.Form frmSettings 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Peach Settings"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4470
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
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Connection Setting"
      Height          =   1815
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   4215
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label3 
         Caption         =   " Server Port:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   " Server IP:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Options"
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   4215
      Begin VB.CheckBox SavePassword 
         Caption         =   "Save Password"
         Height          =   195
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox SaveAccount 
         Caption         =   "Save Account"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Language"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Color Scheme"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton cmdColor 
         Caption         =   "..."
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         ToolTipText     =   "Click here to change the color."
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtColor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Current Color:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   255
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdColor_Click()
With frmSendFile.CDialog
    .ShowColor
    txtColor.BackColor = .Color
End With
Command2.Enabled = True
End Sub

Private Sub Command1_Click()
frmLanguage.Show 1
End Sub

Private Sub Command2_Click()
Command2.Enabled = False
Dim Path As String
Path = App.Path & "\Config.ini"

WriteIniValue Path, "Connection", "IP", txtIP.Text
WriteIniValue Path, "Connection", "Port", txtPort.Text
WriteIniValue Path, "Private", "SchemeColor", txtColor.BackColor
WriteIniValue Path, "Private", "PasswordTick", CBool(SavePassword.Value)
WriteIniValue Path, "Private", "AccountTick", CBool(SaveAccount.Value)

With Setting
    .SERVER_IP = txtIP.Text
    .SERVER_PORT = txtPort.Text
    .SCHEME_COLOR = txtColor.BackColor
    .PASSWORD_TICK = SavePassword.Value
    .ACCOUNT_TICK = SaveAccount.Value
End With

SetScheme
End Sub

Private Sub Form_Load()
With Setting
    txtColor.BackColor = .SCHEME_COLOR
    If .ACCOUNT_TICK = True Then
        SaveAccount.Value = 1
    Else
        SaveAccount.Value = 0
    End If
    If .PASSWORD_TICK = True Then
        SavePassword.Value = 1
    Else
        SavePassword.Value = 0
    End If
    txtIP.Text = .SERVER_IP
    txtPort.Text = .SERVER_PORT
End With
End Sub

Private Sub SaveAccount_Click()
Command2.Enabled = True
End Sub

Private Sub SavePassword_Click()
Command2.Enabled = True
End Sub

Private Sub txtIP_Change()
Command2.Enabled = True
End Sub

Private Sub txtPort_Change()
Command2.Enabled = True
End Sub
