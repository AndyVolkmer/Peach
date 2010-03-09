VERSION 5.00
Begin VB.Form frmSettings 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Peach Settings"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4455
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
   ScaleHeight     =   6300
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   3735
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Connection Setting"
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   4215
      Begin VB.Label Label3 
         Caption         =   " Server Port:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   " Server IP:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Options"
      Height          =   2535
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   4215
      Begin VB.CheckBox chkAutoLogin 
         Caption         =   "Login automatically"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   3855
      End
      Begin VB.CheckBox chkMinimize 
         BackColor       =   &H00F4F4F4&
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   255
      End
      Begin VB.CheckBox chkAskClosing 
         BackColor       =   &H00F4F4F4&
         Caption         =   "Ask before closing"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   3975
      End
      Begin VB.CheckBox chkSavePassword 
         BackColor       =   &H00F4F4F4&
         Caption         =   "Save Password"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3855
      End
      Begin VB.CheckBox chkSaveAccount 
         BackColor       =   &H00F4F4F4&
         Caption         =   "Save Account"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblMinimizeTray 
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   1800
         Width           =   3615
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Language"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5760
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
      Begin VB.TextBox txtFont 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "Example"
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdFont 
         Caption         =   "..."
         Height          =   375
         Left            =   3360
         TabIndex        =   17
         ToolTipText     =   "Click here to change the font."
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "..."
         Height          =   375
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
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblFont 
         Caption         =   "Font:"
         Height          =   255
         Left            =   2160
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblColor 
         Caption         =   "Current Color:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   255
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheckAsk_Click()
Command2.Enabled = True
End Sub

Private Sub CheckMin_Click()
Command2.Enabled = True
End Sub

Private Sub chkAskClosing_Click()
Command2.Enabled = True
End Sub

Private Sub chkAutoLogin_Click()
Command2.Enabled = True
End Sub

Private Sub chkMinimize_Click()
Command2.Enabled = True
End Sub

Private Sub chkSaveAccount_Click()
Command2.Enabled = True
End Sub

Private Sub chkSavePassword_Click()
Command2.Enabled = True
End Sub

Private Sub cmdColor_Click()
On Error Resume Next
With frmMain.CDialog
    .ShowColor
    txtColor.BackColor = .Color
End With

Command2.Enabled = True
End Sub

Private Sub cmdFont_Click()
On Error Resume Next
With frmMain.CDialog
    .FontBold = Fonts.Bold
    .FontName = Fonts.Name
    .FontItalic = Fonts.Italic
    .FontSize = Fonts.Size
    .FontStrikethru = Fonts.Strike
    .FontUnderline = Fonts.Under
    
    .Min = 9
    .Max = 16
    .Flags = cdlCFEffects + cdlCFForceFontExist + cdlCFPrinterFonts + cdlCFLimitSize
    .ShowFont
    
    Fonts.Bold = .FontBold
    Fonts.Name = .FontName
    Fonts.Italic = .FontItalic
    Fonts.Size = .FontSize
    Fonts.Strike = .FontStrikethru
    Fonts.Under = .FontUnderline
End With

With Fonts
    txtFont.Font.Bold = .Bold
    txtFont.Font.Italic = .Italic
    txtFont.Font.Name = .Name
    txtFont.Font.Size = .Size
    txtFont.Font.Strikethrough = .Strike
    txtFont.Font.Underline = .Under
End With

Command2.Enabled = True
End Sub

Private Sub Command1_Click()
frmLanguage.Show 1
End Sub

Private Sub Command2_Click()
Command2.Enabled = False

InsertIntoRegistry "Client\Configuration", "IP", txtIP
InsertIntoRegistry "Client\Configuration", "Port", txtPort
InsertIntoRegistry "Client\Configuration", "SchemeColor", txtColor.BackColor
InsertIntoRegistry "Client\Configuration", "PasswordTick", chkSavePassword.Value
InsertIntoRegistry "Client\Configuration", "AutoLogin", chkAutoLogin.Value
InsertIntoRegistry "Client\Configuration", "AccountTick", chkSaveAccount.Value
InsertIntoRegistry "Client\Configuration", "AskTick", chkAskClosing.Value
InsertIntoRegistry "Client\Configuration", "MinimizeTray", chkMinimize.Value

With txtFont
    InsertIntoRegistry "Client\Font", "FontName", .Font
    InsertIntoRegistry "Client\Font", "FontBold", CStr(.FontBold)
    InsertIntoRegistry "Client\Font", "FontItalic", CStr(.FontItalic)
    InsertIntoRegistry "Client\Font", "FontSize", CStr(.FontSize)
    InsertIntoRegistry "Client\Font", "FontStrike", CStr(.FontStrikethru)
    InsertIntoRegistry "Client\Font", "FontUnder", CStr(.FontUnderline)
End With

With Setting
    .SERVER_IP = txtIP
    .SERVER_PORT = txtPort
    .SCHEME_COLOR = txtColor.BackColor
    .PASSWORD_TICK = chkSavePassword.Value
    .ACCOUNT_TICK = chkSaveAccount.Value
    .AUTO_LOGIN = chkAutoLogin.Value
    .ASK_TICK = chkAskClosing.Value
    .MIN_TICK = chkMinimize.Value
End With

SetScheme
End Sub

Private Sub Form_Activate()
Dim SC As String
    SC = Setting.SCHEME_COLOR

Me.BackColor = SC
Frame1.BackColor = SC
Frame2.BackColor = SC
Frame3.BackColor = SC
lblColor.BackColor = SC
lblFont.BackColor = SC
Label2.BackColor = SC
Label3.BackColor = SC
chkSaveAccount.BackColor = SC
chkSavePassword.BackColor = SC
chkAutoLogin.BackColor = SC
chkAskClosing.BackColor = SC
chkMinimize.BackColor = SC
lblMinimizeTray.BackColor = SC
End Sub

Private Sub Form_Load()
lblColor.Caption = SET_LABEL_COLOR
lblFont.Caption = SET_LABEL_FONT
lblMinimizeTray = SET_CHECK_MINIMIZE
Frame1.Caption = SET_FRAME_STYLE
Frame2.Caption = SET_FRAME_OPTIONS
Frame3.Caption = SET_FRAME_CONNECTION
chkSaveAccount.Caption = SET_CHECK_SAVE_ACCOUNT
chkSavePassword.Caption = SET_CHECK_SAVE_PASSWORD
chkAutoLogin.Caption = SET_CHECK_AUTO_LOGIN
chkAskClosing.Caption = SET_CHECK_ASK_CLOSING
chkMinimize.Caption = SET_CHECK_MINIMIZE
Command1.Caption = SET_COMMAND_LANGUAGE
Command2.Caption = SET_COMMAND_SAVE

With Setting
    txtColor.BackColor = .SCHEME_COLOR
    
    If .ACCOUNT_TICK Then
        chkSaveAccount.Value = 1
    Else
        chkSaveAccount.Value = 0
    End If
    
    If .PASSWORD_TICK Then
        chkSavePassword.Value = 1
    Else
        chkSavePassword.Value = 0
    End If
    
    If .AUTO_LOGIN Then
        chkAutoLogin.Value = 1
    Else
        chkAutoLogin.Value = 0
    End If
    
    If .ASK_TICK Then
        chkAskClosing.Value = 1
    Else
        chkAskClosing.Value = 0
    End If
    
    If .MIN_TICK Then
        chkMinimize.Value = 1
    Else
        chkMinimize.Value = 0
    End If
    
    txtIP = .SERVER_IP
    txtPort = .SERVER_PORT
End With

With Fonts
    txtFont.Font.Bold = .Bold
    txtFont.Font.Italic = .Italic
    txtFont.Font.Name = .Name
    txtFont.Font.Size = .Size
    txtFont.Font.Strikethrough = .Strike
    txtFont.Font.Underline = .Under
End With
End Sub

Private Sub lblMinimizeTray_Click()
With chkMinimize
    If .Value = 0 Then
        .Value = 1
    Else
        .Value = 0
    End If
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
