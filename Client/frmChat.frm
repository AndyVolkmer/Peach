VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.OCX"
Begin VB.Form frmChat 
   Appearance      =   0  'Flat
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   Caption         =   "frmChat"
   ClientHeight    =   4185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4185
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Settings"
      Height          =   300
      Left            =   5280
      TabIndex        =   5
      Top             =   3720
      Width           =   2055
   End
   Begin RichTextLib.RichTextBox txtToSend 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1508
      _Version        =   393217
      Enabled         =   0   'False
      MultiLine       =   0   'False
      MaxLength       =   180
      TextRTF         =   $"frmChat.frx":0000
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   6840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Clear"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Send"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox txtConver 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4471
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmChat.frx":007B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WM_PASTE = &H302

Private Sub cmdSend_Click()
Dim Array1() As String: Array1 = Split(txtToSend.Text, " ")

'Display the time
If LCase$(RTrim$(txtToSend.Text)) = "/time" Then
    With txtConver
        .SelStart = Len(.Text)
        .SelRTF = vbCrLf & "[" & Format$(Time, "hh:nn:ss") & "]" & CHATtimetext & Format$(Time, "hh:nn")
    End With
    GoTo Next1
End If

'Show online list
If LCase$(RTrim$(txtToSend.Text)) = "/online" Then
    frmMain.UpdateListPosition.Enabled = True
    With frmList
        .Left = frmMain.Left + .Width * 2 + 20
        .Top = frmMain.Top
        .Height = frmMain.Height - 400
        .Show
    End With
    GoTo Next1
End If

'No whitespaces
If Len(Trim$(txtToSend.Text)) = 0 Then
    GoTo Next1
End If

'If any checkbox is checked then send it private to that client
With frmList.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i).Checked = True Then
            SendMsg "!w" & "#" & frmConfig.txtNick.Text & "|" & .Item(i) & "#" & txtToSend.Text & "#"
            GoTo Next1
        End If
    Next i
End With

'Send public message
SendMsg "!msg" & "#" & frmConfig.txtNick.Text & "#" & txtToSend.Text & "#"

Next1:
txtToSend.Text = vbNullString
txtToSend.SetFocus

End Sub

Private Sub Command1_Click()
frmSettings.Show 1
End Sub

Public Sub Form_Load()
Top = 0
Left = 0
LoadChatForm
LoadChatSettings
End Sub

Public Sub LoadChatSettings()
Dim FONT_BACK_COL   As Long

'Read 'BackColor' from .ini file
If Len(ReadIniValue(App.Path & "\Config.ini", "Chat", "BackCol")) = 0 Then
    txtConver.BackColor = 16777215
Else
    FONT_BACK_COL = ReadIniValue(App.Path & "\Config.ini", "Chat", "BackCol")
    txtConver.BackColor = FONT_BACK_COL
    txtToSend.BackColor = FONT_BACK_COL
End If

'Read 'Font' from .ini file
If Len(ReadIniValue(App.Path & "\Config.ini", "Chat", "Font")) = 0 Then
    txtConver.Font = "Tahoma"
Else
    With txtConver
        Select Case Font
        Case 0 'Arial
            .Font = "Arial"
        Case 1 'Lucida Console
            .Font = "Lucida Console"
        Case 2 'Tahoma
            .Font = "Tahoma"
        Case 3 'Time New Roman
            .Font = "Times New Roman"
        Case 4 'Trebuchet MS
            .Font = "Trebuchet MS"
        Case 5 'Verdana
            .Font = "Verdana"
        End Select
    End With
End If

'Read 'FontSize' from .ini file
If Len(ReadIniValue(App.Path & "\Config.ini", "Chat", "FontSize")) = 0 Then
    txtConver.Font.Size = 8
Else
    With txtConver.Font
        Select Case ReadIniValue(App.Path & "\Config.ini", "Chat", "FontSize")
        Case 0
            .Size = 8
        Case 1
            .Size = 9
        Case 2
            .Size = 10
        End Select
    End With
End If
End Sub

Public Sub LoadChatForm()
cmdSend.Caption = CHATcommand_send
cmdClear.Caption = CHATcommand_clear
End Sub

Private Sub cmdClear_Click()
txtConver.Text = vbNullString
txtToSend.Text = vbNullString
End Sub

Private Sub txtConver_Change()
Dim hWnd1 As Long: hWnd1 = GetActiveWindow

'Unlock so we can convert smiñeys
txtConver.Locked = False

'If window doenst have focus then flash
With frmMain
    If Not hWnd1 = .hwnd Then
        Call FlashTitle(.hwnd, True)
    End If
End With

'Create smileys
Call Create_Smileys(txtConver)

'Set cursor to last position
txtConver.SelStart = Len(txtConver.Text)

'Lock again
txtConver.Locked = True
End Sub

Private Sub txtToSend_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdSend_Click
End Sub

Public Sub Create_Smileys(RTF As Control)
Dim Smileys() As String
Dim SmileysFile() As String
Dim Smilestring As String
Dim SmileFileString As String
Dim Pos As Long, Start As Long
Dim IconPath As String

Screen.MousePointer = vbHourglass

Pos = RTF.SelStart

Start = 1

IconPath = App.Path & "\smileys\"

Smilestring = _
    ":) " & _
    ":-) " & _
    ":( " & _
    ":-( " & _
    ";) " & _
    ";-) " & _
    ":O " & _
    ":o " & _
    ":D " & _
    ":P " & _
    ":p " & _
    ":cool: " & _
    ":rolleyes: " & _
    ">:("

SmileFileString = _
    "smiley1.gif," & _
    "smiley1.gif," & _
    "smiley2.gif," & _
    "smiley2.gif," & _
    "smiley3.gif," & _
    "smiley3.gif," & _
    "smiley4.gif," & _
    "smiley4.gif," & _
    "smiley5.gif," & _
    "smiley6.gif," & _
    "smiley6.gif," & _
    "smiley7.gif," & _
    "smiley8.gif," & _
    "smiley9.gif"

Smileys = Split(Smilestring, " ")
SmileysFile = Split(SmileFileString, ",")

If UBound(Smileys) <> UBound(SmileysFile) Then
    Debug.Print "Failure in array."
    Exit Sub
End If

For i = LBound(Smileys) To UBound(Smileys)
    While RTF.Find(Smileys(i), Start - 1) >= 0
        Picture1.Picture = LoadPicture(Trim$(IconPath & SmileysFile(i)))
        RTF.SelStart = RTF.Find(Smileys(i), Start - 1)
        RTF.SelLength = Len(Smileys(i))
        Start = RTF.SelStart + RTF.SelLength + 1
        RTF.SelText = vbNullString
        CopyPictureToRTF RTF, Picture1.Picture
    Wend

    Start = 1
Next i

RTF.SelStart = Pos

Screen.MousePointer = vbNormal
End Sub

Private Sub CopyPictureToRTF(RTF As Control, Bild As Picture)
Dim Buf As Variant
Dim Text As String

If Clipboard.GetFormat(vbCFText) = True Then
  Text = Clipboard.GetText
Else
  Buf = Clipboard.GetData
End If

Clipboard.Clear
Clipboard.SetData Picture1.Picture
DoEvents

SendMessage RTF.hwnd, WM_PASTE, 0, 0
DoEvents
Sleep 30

Clipboard.Clear
On Error Resume Next
If Text <> "" Then
  Clipboard.SetText Text
Else
  If Buf <> 0 Then
    Clipboard.SetData Buf
  End If
End If
End Sub
