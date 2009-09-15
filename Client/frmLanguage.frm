VERSION 5.00
Begin VB.Form frmLanguage 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Peach"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   2895
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLanguage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   2895
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Enter"
      Height          =   350
      Left            =   1440
      TabIndex        =   2
      Top             =   850
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Select your language :"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   165
      Width           =   2295
   End
End
Attribute VB_Name = "frmLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnter_Click()
Select Case Combo1.ListIndex
Case 0
    SetLangGerman
Case 1
    SetLangEnglish
Case 2
    SetLangSpanish
Case 3
    SetLangSwedish
Case 4
    SetLangItalian
Case 5
    SetLangDutch
Case 6
    SetLangSerbian
Case 7
    SetLangFrench
End Select

frmMain.LoadMDIForm
frmSendFile.LoadSendFileForm
frmChat.LoadChatForm
frmConfig.LoadConfigForm

WriteIniValue App.Path & "\Config.ini", "Language", "Validate", "0"
WriteIniValue App.Path & "\Config.ini", "Language", "Language", frmLanguage.Combo1.ListIndex
frmLanguage.Hide
End Sub

Private Sub Combo1_Click()
Select Case Combo1.ListIndex
Case 0 'German
    SetLangGerman
Case 1 'English
    SetLangEnglish
Case 2 'Spanish
    SetLangSpanish
Case 3 'Swedish
    SetLangSwedish
Case 4 'Italian
    SetLangItalian
Case 5 'Dutch
    SetLangDutch
Case 6 'Serbian
    SetLangSerbian
Case 7 'French
    SetLangFrench
End Select
SetLang
End Sub

Private Sub SetLang()
Label1.Caption = LANG_LABEL_SELLANG
cmdEnter.Caption = LANG_COMMAND_ENTER
Combo1.List(0) = LANG_GERMAN
Combo1.List(1) = LANG_ENGLISH
Combo1.List(2) = LANG_SPANISH
Combo1.List(3) = LANG_SWEDISH
Combo1.List(4) = LANG_ITALIAN
Combo1.List(5) = LANG_DUTCH
Combo1.List(6) = LANG_SERBIAN
Combo1.List(7) = LANG_FRENCH
End Sub

Private Sub Form_Activate()
SetLang
End Sub
