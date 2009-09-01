VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Begin VB.Form frmSettings 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Peach - Chat Settings"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4725
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
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4725
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CheckBox CheckMWindow 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Enable Message Windows"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Chat Settings"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtBackgroundColor 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdBackgroundColor 
         Caption         =   "..."
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox cmbFontSize 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdFontColor 
         Caption         =   "..."
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtFontColor 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox cmbFont 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Background Color:"
         Height          =   255
         Left            =   2280
         TabIndex        =   15
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblExample 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " This is an example text."
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   4215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Font Size:"
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Font Color:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Font:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label lblEMWindow 
      Alignment       =   2  'Center
      BackColor       =   &H00F4F4F4&
      Caption         =   "[&?]"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      ToolTipText     =   "Show example."
      Top             =   2520
      Width           =   375
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbFont_Click()
UpdateExample
End Sub

Private Sub cmbFontSize_Click()
UpdateExample
End Sub

Private Sub cmdBackgroundColor_Click()
With CommonDialog1
    .ShowColor
    txtBackgroundColor.BackColor = .Color
End With
UpdateExample
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdFontColor_Click()
With CommonDialog1
    .ShowColor
    txtFontColor.BackColor = .Color
End With
UpdateExample
End Sub

Private Sub cmdSave_Click()
WriteIniValue App.Path & "\Config.ini", "Chat", "Font", cmbFont.ListIndex
WriteIniValue App.Path & "\Config.ini", "Chat", "FontSize", cmbFontSize.ListIndex
WriteIniValue App.Path & "\Config.ini", "Chat", "FontCol", txtFontColor.BackColor
WriteIniValue App.Path & "\Config.ini", "Chat", "BackCol", txtBackgroundColor.BackColor
With frmChat.txtConver
    .Font = cmbFont.Text
    Select Case cmbFontSize.ListIndex
    Case 0
        .Font.Size = 8
    Case 1
        .Font.Size = 9
    Case 2
        .Font.Size = 10
    End Select
    .SelStart = 1
    .SelLength = Len(.Text)
    .SelColor = txtFontColor.BackColor
    .BackColor = txtBackgroundColor.BackColor
    frmChat.txtToSend.BackColor = txtBackgroundColor.BackColor
End With
Unload Me
End Sub

Private Sub Form_Load()
'Add fonts to combobox
With cmbFont
    .AddItem "Arial"
    .AddItem "Lucida Console"
    .AddItem "Tahoma"
    .AddItem "Times New Roman"
    .AddItem "Trebuchet MS"
    .AddItem "Verdana"
End With

'Add sizes to combobox
With cmbFontSize
    .AddItem 8
    .AddItem 9
    .AddItem 10
End With

'Load values from .ini file
LoadIniSettings

'Setup Example
UpdateExample
End Sub

Private Sub UpdateExample()
With lblExample
    .ForeColor = txtFontColor.BackColor
    .BackColor = txtBackgroundColor.BackColor
    .Font = cmbFont.Text
    If cmbFontSize.Text = "" Then
        .FontSize = 8
    Else
        .FontSize = cmbFontSize.Text
    End If
End With
End Sub

Private Sub LoadIniSettings()
'Read 'Font Index' from .ini file
If Len(ReadIniValue(App.Path & "\Config.ini", "Chat", "Font")) = 0 Then
    cmbFont.ListIndex = 0
Else
    cmbFont.ListIndex = ReadIniValue(App.Path & "\Config.ini", "Chat", "Font")
End If

'Read 'Font Size Index' from .ini file
If Len(ReadIniValue(App.Path & "\Config.ini", "Chat", "FontSize")) = 0 Then
    cmbFontSize.ListIndex = 0
Else
    cmbFontSize.ListIndex = ReadIniValue(App.Path & "\Config.ini", "Chat", "FontSize")
End If

'Read 'Font Color' from .ini file
If Len(ReadIniValue(App.Path & "\Config.ini", "Chat", "FontCol")) = 0 Then
    txtFontColor.BackColor = 0
Else
    txtFontColor.BackColor = ReadIniValue(App.Path & "\Config.ini", "Chat", "FontCol")
End If

'Read 'Background Color' from .ini file
If Len(ReadIniValue(App.Path & "\Config.ini", "Chat", "BackCol")) = 0 Then
    txtBackgroundColor.BackColor = 0
Else
    txtBackgroundColor.BackColor = ReadIniValue(App.Path & "\Config.ini", "Chat", "BackCol")
End If
End Sub
