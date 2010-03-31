VERSION 5.00
Begin VB.Form frmLanguage 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Peach"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   2880
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
   ScaleHeight     =   1410
   ScaleWidth      =   2880
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Enter"
      Enabled         =   0   'False
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
Option Explicit

Dim FIRST_LANG As Long

Private Sub cmdEnter_Click()
SET_LANG_1 Combo1.ListIndex

InsertIntoRegistry "Client\Configuration", "Validate", "0"
InsertIntoRegistry "Client\Configuration", "Language", frmLanguage.Combo1.ListIndex

If Setting.VALIDATE = 0 Then
    If MsgBox(LANG_QUIT, vbQuestion + vbYesNo) = vbYes Then
        CloseThis
    Else
        Unload Me
    End If
Else
    SetupForm frmMain
End If
End Sub

Private Sub Combo1_Click()
cmdEnter.Enabled = True
SET_LANG_1 Combo1.ListIndex
SET_LANG
End Sub

Private Sub SET_LANG()
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
SET_LANG
FIRST_LANG = CURRENT_LANG
Me.BackColor = Setting.SCHEME_COLOR
Label1.BackColor = Setting.SCHEME_COLOR
End Sub

Private Sub Form_Unload(Cancel As Integer)
SET_LANG_1 FIRST_LANG
End Sub

Private Sub SET_LANG_1(Lang As Long)
Select Case Lang
    Case 0: SET_LANG_GERMAN
    Case 1: SET_LANG_ENGLISH
    Case 2: SET_LANG_SPANISH
    Case 3: SET_LANG_SWEDISH
    Case 4: SET_LANG_ITALIAN
    Case 5: SET_LANG_DUTCH
    Case 6: SET_LANG_SERBIAN
    Case 7: SET_LANG_FRENCH
End Select
End Sub
