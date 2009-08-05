VERSION 5.00
Begin VB.Form frmLanguage 
   BackColor       =   &H00F4F4F4&
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
Dim N       As String
Dim M       As String
Dim TSSO2   As TypeSSO2

Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub cmdEnter_Click()
Select Case Combo1.ListIndex
Case 0
    SetLangGerman
    M = "0"
Case 1
    SetLangEnglish
    M = "1"
Case 2
    SetLangSpanish
    M = "2"
Case 3
    SetLangSwedish
    M = "3"
Case 4
    SetLangItalian
    M = "4"
Case 5
    SetLangDutch
    M = "5"
Case 6
    SetLangSerbian
    M = "6"
Case 7
    SetLangFrench
    M = "7"
End Select
frmLanguage.Hide
frmMain.LoadMDIForm
frmList.LoadListForm
frmSendFile.LoadSendFileForm
frmChat.LoadChatForm
frmConfig.LoadConfigForm
frmMain.Show
WriteConfigFile2 (App.Path & "\validaten.conf")
End Sub

Private Sub Combo1_Click()
Select Case Combo1.ListIndex
Case 0 ' German
    Label1.Caption = "Wähle deine Sprache aus:"
    cmdEnter.Caption = "&Öffnen"
    Combo1.List(0) = "Deutsch"
    Combo1.List(1) = "Englisch"
    Combo1.List(2) = "Spanisch"
    Combo1.List(3) = "Swedisch"
    Combo1.List(4) = "Italienisch"
    Combo1.List(5) = "Holländisch"
    Combo1.List(6) = "Serbisch"
    Combo1.List(7) = "Französisch"
Case 1 ' English
    Label1.Caption = "Select your language:"
    cmdEnter.Caption = "&Open"
    Combo1.List(0) = "German"
    Combo1.List(1) = "English"
    Combo1.List(2) = "Spanish"
    Combo1.List(3) = "Swedish"
    Combo1.List(4) = "Italian"
    Combo1.List(5) = "Dutch"
    Combo1.List(6) = "Serbian"
    Combo1.List(7) = "French"
Case 2 ' Spanish
    Label1.Caption = "Elige tu idioma:"
    cmdEnter.Caption = "&Abrir"
    Combo1.List(0) = "Aleman"
    Combo1.List(1) = "Inglés"
    Combo1.List(2) = "Español"
    Combo1.List(3) = "Sueco"
    Combo1.List(4) = "Italiano"
    Combo1.List(5) = "Holandés"
    Combo1.List(6) = "Serbio"
    Combo1.List(7) = "Francés"
Case 3 ' Swedish
    Label1.Caption = "Välj språk:"
    cmdEnter.Caption = "&Öppna"
    Combo1.List(O) = "Tyska"
    Combo1.List(1) = "Engelska"
    Combo1.List(2) = "Spanska"
    Combo1.List(3) = "Svenska"
    Combo1.List(4) = "Italienska"
    Combo1.List(5) = "Holländska"
    Combo1.List(6) = "Serbiska"
    Combo1.List(7) = "Franska"
Case 4 ' Italian
    Label1.Caption = "Seleziona la tua lingua:"
    cmdEnter.Caption = "&Apri"
    Combo1.List(0) = "Tedesco"
    Combo1.List(1) = "Inglese"
    Combo1.List(2) = "Spagnolo"
    Combo1.List(3) = "Svedese"
    Combo1.List(4) = "Italiano"
    Combo1.List(5) = "Olandese"
    Combo1.List(6) = "Serbo"
    Combo1.List(7) = "Francese"
Case 5 ' Dutch
    Label1.Caption = "Selecteer jou taal:"
    cmdEnter.Caption = "Openen"
    Combo1.List(O) = "Duits"
    Combo1.List(1) = "Engels"
    Combo1.List(2) = "Spaans"
    Combo1.List(3) = "Zweeds"
    Combo1.List(4) = "Italiaans"
    Combo1.List(5) = "Nederlands"
    Combo1.List(6) = "Serbisch"
    Combo1.List(7) = "Frans"
Case 6 ' Serbian
    Label1.Caption = "Dodaj svoj jezik:"
    cmdEnter.Caption = "Otvori"
    Combo1.List(O) = "Nemacki"
    Combo1.List(1) = "Engleski"
    Combo1.List(2) = "Spanski"
    Combo1.List(3) = "Svedski"
    Combo1.List(4) = "Italijanski"
    Combo1.List(5) = "Holandski"
    Combo1.List(6) = "Srpski"
    Combo1.List(7) = "Francuski"
Case 7 ' French
    Label1.Caption = "Choisissez votre langue:"
    cmdEnter.Caption = "Ouvrir"
    Combo1.List(O) = "Allemand"
    Combo1.List(1) = "Anglais"
    Combo1.List(2) = "Espagnol"
    Combo1.List(3) = "Suédois"
    Combo1.List(4) = "Italien"
    Combo1.List(5) = "Hollandais"
    Combo1.List(6) = "Serbisch"
    Combo1.List(7) = "Français"
End Select
End Sub

Private Sub Form_Activate()
Select Case M
Case 0
    Label1.Caption = CONFIGlabel_selectlanguage
    cmdEnter.Caption = "&Öffnen"
    Combo1.List(0) = CONFIGcombo_german
    Combo1.List(1) = CONFIGcombo_english
    Combo1.List(2) = CONFIGcombo_spanish
    Combo1.List(3) = CONFIGcombo_swedish
    Combo1.List(4) = CONFIGcombo_italian
    Combo1.List(5) = CONFIGcombo_dutch
    Combo1.List(6) = CONFIGcombo_serbian
    Combo1.List(7) = CONFIGcombo_french
Case 1 ' English
    cmdEnter.Caption = "&Open"
    Label1.Caption = CONFIGlabel_selectlanguage
    Combo1.List(0) = CONFIGcombo_german
    Combo1.List(1) = CONFIGcombo_english
    Combo1.List(2) = CONFIGcombo_spanish
    Combo1.List(3) = CONFIGcombo_swedish
    Combo1.List(4) = CONFIGcombo_italian
    Combo1.List(5) = CONFIGcombo_dutch
    Combo1.List(6) = CONFIGcombo_serbian
    Combo1.List(7) = CONFIGcombo_french
Case 2 ' Spanish
    cmdEnter.Caption = "&Abrir"
    Label1.Caption = CONFIGlabel_selectlanguage
    Combo1.List(0) = CONFIGcombo_german
    Combo1.List(1) = CONFIGcombo_english
    Combo1.List(2) = CONFIGcombo_spanish
    Combo1.List(3) = CONFIGcombo_swedish
    Combo1.List(4) = CONFIGcombo_italian
    Combo1.List(5) = CONFIGcombo_dutch
    Combo1.List(6) = CONFIGcombo_serbian
    Combo1.List(7) = CONFIGcombo_french
Case 3 ' Swedish
    cmdEnter.Caption = "&Öppna"
    Label1.Caption = CONFIGlabel_selectlanguage
    Combo1.List(0) = CONFIGcombo_german
    Combo1.List(1) = CONFIGcombo_english
    Combo1.List(2) = CONFIGcombo_spanish
    Combo1.List(3) = CONFIGcombo_swedish
    Combo1.List(4) = CONFIGcombo_italian
    Combo1.List(5) = CONFIGcombo_dutch
    Combo1.List(6) = CONFIGcombo_serbian
    Combo1.List(7) = CONFIGcombo_french
Case 4 ' Italian
    cmdEnter.Caption = "&Apri"
    Label1.Caption = CONFIGlabel_selectlanguage
    Combo1.List(0) = CONFIGcombo_german
    Combo1.List(1) = CONFIGcombo_english
    Combo1.List(2) = CONFIGcombo_spanish
    Combo1.List(3) = CONFIGcombo_swedish
    Combo1.List(4) = CONFIGcombo_italian
    Combo1.List(5) = CONFIGcombo_dutch
    Combo1.List(6) = CONFIGcombo_serbian
    Combo1.List(7) = CONFIGcombo_french
Case 5 ' Dutch
    cmdEnter.Caption = "Openen"
    Label1.Caption = CONFIGlabel_selectlanguage
    Combo1.List(0) = CONFIGcombo_german
    Combo1.List(1) = CONFIGcombo_english
    Combo1.List(2) = CONFIGcombo_spanish
    Combo1.List(3) = CONFIGcombo_swedish
    Combo1.List(4) = CONFIGcombo_italian
    Combo1.List(5) = CONFIGcombo_dutch
    Combo1.List(6) = CONFIGcombo_serbian
    Combo1.List(7) = CONFIGcombo_french
Case 6 ' Serbian
    cmdEnter.Caption = "Otvori"
    Label1.Caption = CONFIGlabel_selectlanguage
    Combo1.List(0) = CONFIGcombo_german
    Combo1.List(1) = CONFIGcombo_english
    Combo1.List(2) = CONFIGcombo_spanish
    Combo1.List(3) = CONFIGcombo_swedish
    Combo1.List(4) = CONFIGcombo_italian
    Combo1.List(5) = CONFIGcombo_dutch
    Combo1.List(6) = CONFIGcombo_serbian
    Combo1.List(7) = CONFIGcombo_french
Case 7 ' French
    cmdEnter.Caption = "Ouvrir"
    Label1.Caption = CONFIGlabel_selectlanguage
    Combo1.List(0) = CONFIGcombo_german
    Combo1.List(1) = CONFIGcombo_english
    Combo1.List(2) = CONFIGcombo_spanish
    Combo1.List(3) = CONFIGcombo_swedish
    Combo1.List(4) = CONFIGcombo_italian
    Combo1.List(5) = CONFIGcombo_dutch
    Combo1.List(6) = CONFIGcombo_serbian
    Combo1.List(7) = CONFIGcombo_french
End Select
End Sub

Private Sub Form_Initialize()
If App.PrevInstance = True Then
    MsgBox "Peach is already running!", vbInformation
     'you can add code to maximize already open instance
    Exit Sub
End If
Call InitCommonControls
End Sub

Public Sub Form_Load()
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
    frmLanguage.Hide
    frmMain.Show
Else
    Combo1.Clear
    With Combo1
        .AddItem "German"
        .AddItem "English"
        .AddItem "Spanish"
        .AddItem "Swedish"
        .AddItem "Italian"
        .AddItem "Dutch"
        .AddItem "Serbian"
        .AddItem "French"
        .ListIndex = 1
        SetLangEnglish
        M = .ListIndex
    End With
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmMain
Unload frmList
End Sub
