VERSION 5.00
Begin VB.Form frmLanguage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Peach (Client)"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   2895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLanguage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   2895
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdEnter 
      Caption         =   "&Enter"
      Height          =   350
      Left            =   1440
      TabIndex        =   2
      Top             =   850
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Select your language :"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   170
      Width           =   1695
   End
End
Attribute VB_Name = "frmLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim N As String
Dim M As String


Private Sub cmdEnter_Click()
Select Case Combo1.ListIndex
Case 0
    SetLangGerman
Case 1
    SetLangEnglish
Case 2
    SetLangSpanish
End Select
frmLanguage.Hide
frmMain.MDIForm_Load
frmList.Form_Load
frmConfig.Form_Load
frmSendFile.Form_Load
frmChat.Form_Load
WriteConfigFile2 (App.Path & "\validaten.conf")
End Sub

Private Sub Combo1_Click()
Select Case Combo1.ListIndex
Case 0
    Label1.Caption = "Wähle deine Sprache aus:"
    cmdEnter.Caption = "&Öffnen"
    Combo1.List(0) = "Deutsch"
    Combo1.List(1) = "Englisch"
    Combo1.List(2) = "Spanisch"
Case 1
    Label1.Caption = "Select your language:"
    cmdEnter.Caption = "&Open"
    Combo1.List(0) = "German"
    Combo1.List(1) = "English"
    Combo1.List(2) = "Spanish"
Case 2
    Label1.Caption = "Elige tu idioma:"
    cmdEnter.Caption = "&Abrir"
    Combo1.List(0) = "Aleman"
    Combo1.List(1) = "Inglés"
    Combo1.List(2) = "Español"
End Select
End Sub

Public Sub Form_Load()
Dim TSSO2 As TypeSSO2
TSSO2 = ReadConfigFile2(App.Path & "\validaten.conf")

N = Trim(TSSO2.ValidateN)
M = Trim(TSSO2.Language)
If N = "0" Then
    Select Case M
    Case 0
        SetLangGerman
    Case 1
        SetLangEnglish
    Case 2
        SetLangSpanish
    End Select
    frmLanguage.Hide
    frmMain.Show
Else
    Combo1.Clear
    With Combo1
        .AddItem "German"
        .AddItem "English"
        .AddItem "Spanish"
    End With
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmMain
Unload frmList
End Sub
