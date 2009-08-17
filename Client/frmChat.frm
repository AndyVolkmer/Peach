VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.OCX"
Begin VB.Form frmChat 
   Appearance      =   0  'Flat
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   Caption         =   "frmChat"
   ClientHeight    =   3810
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3810
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   6840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Clear"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Send"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox txtToSend 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1508
      _Version        =   393217
      Enabled         =   0   'False
      MultiLine       =   0   'False
      MaxLength       =   180
      TextRTF         =   $"frmChat.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtConver 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4471
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
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
Dim i As Integer
Dim Array1() As String

'Assign variables
Prefix = "[" & Format(Time, "hh:nn:ss") & "]"
Array1 = Split(txtToSend.Text, " ")
GetName = StrConv(frmConfig.txtNick, vbProperCase)

'No whitespaces 0-5
Select Case txtToSend.Text
Case "", " ", "  ", "   ", "    ", "     ", "      "
    txtToSend.Text = ""
    Exit Sub
End Select

'Check if there is an emote or command used
On Error GoTo Error1
Select Case Array1(0)
'All avaible emotes
Case _
    "/lol", "/LOL", "/Lol", "/Laugh", "/laugh", _
    "/rofl", "/ROFL", "/Rofl", _
    "/beer", "/Beer", _
    "/fart", "/Fart", _
    "/lmao", "/LMAO", _
    "/insult", "/Insult", _
    "/facepalm", "/Facepalm", _
    "/violin", "/Violin"
    
    SendMsg "!emote" & "#" & GetName & "#" & Array1(0) & "#"
    GoTo NextI
End Select

Select Case txtToSend.Text
        
'Check if its the same Message as before
Case LastMsg
    VisualizeMessage False, "System", CHATflood_protection
    txtToSend.Text = ""
    Exit Sub
    
'Display the time
Case Trim("/time"), Trim("/Time"), Trim("/TIME")
    txtConver.Text = txtConver.Text & vbCrLf & Prefix & CHATtimetext & Format(Time, "hh:nn")
    
'Open online user list
Case Trim("/online"), Trim("/Online"), Trim("/ONLINE")
    frmMain.UpdateListPosition.Enabled = True
    With frmList
        .Left = frmMain.Left + .Width * 2 + 20
        .Top = frmMain.Top
        .Height = frmMain.Height - 400
        .Show
    End With

'Send Message
Case Else
    'If any checkbox is checked then send it private to that client
    For i = 1 To frmList.ListView1.ListItems.Count
        If frmList.ListView1.ListItems.Item(i).Checked = True Then
            'If the the selected name is yours then no
            If frmList.ListView1.ListItems.Item(i).Text = StrConv(frmConfig.txtNick, vbProperCase) Then
                MsgBox "You cant whisper yourself.", vbInformation
                txtToSend.Text = ""
                txtToSend.SetFocus
                Exit Sub
            End If
            
            GetConver = txtToSend.Text
            ForWho = StrConv(frmList.ListView1.ListItems.Item(i), vbProperCase)
            Message = "!w" & "#" & GetName & "|" & ForWho & "#" & GetConver & "#"
            SendMsg Message
            VisualizeMessage True, ForWho, GetConver

        End If
    Next i
    
    'Send public message
    GetConver = txtToSend.Text
    GetName = frmConfig.txtNick.Text
    Message = "!msg" & "#" & GetName & "#" & GetConver & "#"
    SendMsg Message
    
End Select

NextI:
    LastMsg = txtToSend.Text
    txtToSend.Text = ""
    txtToSend.SetFocus

Exit Sub
Error1:
MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Public Sub Form_Load()
Me.Top = 0
Me.Left = 0
LoadChatForm
End Sub

Public Sub LoadChatForm()
cmdSend.Caption = CHATcommand_send
cmdClear.Caption = CHATcommand_clear
End Sub

Private Sub cmdClear_Click()
txtConver.Text = ""
txtToSend.Text = ""
End Sub

Private Sub txtConver_Change()
Dim hWnd1 As Long
'txtConver.Locked = False
hWnd1 = GetActiveWindow
With frmMain
    If hWnd1 = .hwnd Then
    Else
        Call FlashTitle(.hwnd, True)
    End If
    .Hyperlink1.URLFormat txtConver
End With
'Create_Smileys txtConver
txtConver.SelStart = Len(txtConver.Text)
'txtConver.Locked = True
End Sub

Private Sub txtConver_Click()
frmMain.Hyperlink1.URLLaunch
End Sub

Private Sub txtConver_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.Hyperlink1.RichWordOver Me, txtConver, X, Y
End Sub

Private Sub txtToSend_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdSend_Click
End Sub

Public Sub Create_Smileys(RTF As Control)
Dim Smileys() As String
Dim SmileysFile() As String
Dim Smilestring As String
Dim SmileFileString As String
Dim i As Integer
Dim Pos As Long, Start As Long
Dim IconPath As String

Screen.MousePointer = vbHourglass

Pos = RTF.SelStart

Start = 1

IconPath = App.Path & "\smileys\"

Smilestring = ":) :-) :( :-( ;) ;-) " & _
  ":o :D :p :cool: :rolleyes: :mad:"

SmileFileString = "smiley1.gif,smiley1.gif," & _
  "smiley2.gif,smiley2.gif," & _
  "smiley3.gif,smiley3.gif," & _
  "smiley4.gif,smiley5.gif," & _
  "smiley6.gif,smiley7.gif," & _
  "smiley8.gif,smiley9.gif"

Smileys = Split(Smilestring, " ")
SmileysFile = Split(SmileFileString, ",")

If UBound(Smileys) <> UBound(SmileysFile) Then
  Debug.Print "Arrays are not same!"
  Exit Sub
End If

For i = LBound(Smileys) To UBound(Smileys)
  While RTF.Find(Smileys(i), Start - 1) >= 0
    Picture1.Picture = LoadPicture(Trim$(IconPath & SmileysFile(i)))
    RTF.SelStart = RTF.Find(Smileys(i), Start - 1)
    RTF.SelLength = Len(Smileys(i))
    Start = RTF.SelStart + RTF.SelLength + 1
    RTF.SelText = ""
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

If Text <> "" Then
  Clipboard.SetText Text
Else
  If Buf <> 0 Then
    Clipboard.SetData Buf
  End If
End If
End Sub
