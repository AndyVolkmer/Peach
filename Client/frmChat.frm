VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmChat 
   Appearance      =   0  'Flat
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   Caption         =   "frmChat"
   ClientHeight    =   3870
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
   ScaleHeight     =   3870
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox txtToSend 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1508
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   0   'False
      MultiLine       =   0   'False
      MaxLength       =   180
      TextRTF         =   $"frmChat.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      TextRTF         =   $"frmChat.frx":007D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
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

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const EM_CHARFROMPOS As Long = &HD7&
Private Const WM_PASTE = &H302

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Sign(255) As Integer

Private Sub cmdSend_Click()

'No whitespaces
If Len(Trim$(txtToSend.Text)) = 0 Then
    Call Clear
    Exit Sub
End If

'Display the time
If LCase$(RTrim$(txtToSend.Text)) = "/time" Then
    With txtConver
        .SelStart = Len(.Text)
        .SelRTF = vbCrLf & CHAT_TIME_TEXT & Format$(Time, "hh:nn:ss") & "."
    End With
    Call Clear
    Exit Sub
End If

'Send public message
SendMsg "!msg#" & frmConfig.txtNick & "#" & Trim$(txtToSend.Text) & "#"
Call Clear

End Sub

Private Sub Clear()
With txtToSend
    .Text = vbNullString
    .SetFocus
End With
End Sub

Private Sub Form_Load()
Top = 0: Left = 0
Call LoadChatForm
End Sub

Public Sub LoadChatForm()
cmdSend.Caption = CHAT_COMMAND_SEND
cmdClear.Caption = CHAT_COMMAND_CLEAR
End Sub

Private Sub cmdClear_Click()
txtConver.Text = vbNullString
txtToSend.Text = vbNullString
End Sub

Private Sub txtConver_Change()
Dim hWnd1 As Long: hWnd1 = GetActiveWindow

'Unlock so we can convert smileys
txtConver.Locked = False

'Create smileys
Call Create_Smileys(txtConver)

Call Highlight(txtConver)

'If window doenst have focus then flash
With frmMain
    If Not hWnd1 = .hwnd Then
        Call FlashTitle(.hwnd, True)
    End If
End With

'Set cursor to last position
txtConver.SelStart = Len(txtConver.Text)

'Lock again
txtConver.Locked = True
End Sub

Private Sub txtConver_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Text As String
Dim lnk As Long
Dim ret As Long

Text = GetWord(txtConver, X, Y)

lnk = IsUrlOrMail(Text)

If lnk > 0 Then
    ret = RemoveSign(Text)
    ret = RemoveBrackets(Text)
    
    If lnk > 100 Then
        Text = "mailto:" + Text
    End If
    
    'MsgBox Text$
    Call SendLink(Text)
End If

End Sub

Private Sub txtConver_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Text As String
    
Text = GetWord(txtConver, X, Y)

If IsUrlOrMail(Text) Then
    txtConver.MousePointer = 99
Else
    txtConver.MousePointer = 0
End If
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
    ":["
    
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

Private Sub SendLink(ByVal Link As String)
Dim Success As Long
Success = ShellExecute(0&, vbNullString, Link, vbNullString, "C:\", 1)
End Sub

Private Function GetWord(Rich As RichTextBox, ByVal X&, ByVal Y&) As String
Dim Pos As Long, P1 As Long, P2 As Long
Dim CHAR As Long
Dim MousePointer As POINTAPI

MousePointer.X = X \ Screen.TwipsPerPixelX
MousePointer.Y = Y \ Screen.TwipsPerPixelY
Pos = SendMessage(Rich.hwnd, EM_CHARFROMPOS, 0&, MousePointer)
If Pos <= 0 Then Exit Function

For P1 = Pos To 1 Step -1
    CHAR = Asc(Mid$(Rich.Text, P1, 1))
    If Sign(CHAR) = 2 Then
        Exit For
    End If
Next P1
P1 = P1 + 1

' Wortende finden.
For P2 = Pos To Len(Rich.Text)
    CHAR = Asc(Mid$(Rich.Text, P2, 1))
    If Sign(CHAR) = 2 Then
        Exit For
    End If
Next P2
P2 = P2 - 1

If P1 < P2 Then GetWord = Mid$(Rich.Text, P1, P2 - P1 + 1)
End Function

Private Function RemoveSign(ByRef Test As String) As Long
Dim Last As Long
Last = Asc(Right$(Test, 1))
If Sign(Last) = 1 Then
    Test = Left$(Test, Len(Test) - 1)
    RemoveSign = 1
End If
End Function

Private Function RemoveBrackets(Test As String) As Long
If Left$(Test, 1) = Chr$(40) Then
    If Right(Test, 1) = Chr$(41) Then
        Test = Mid$(Test, 2, Len(Test) - 2)
        RemoveBrackets = 1
    End If
End If

If Left$(Test, 1) = Chr$(34) Then
    If Right(Test, 1) = Chr$(34) Then
        Test = Mid$(Test$, 2, Len(Test) - 2)
        RemoveBrackets = 1
    End If
End If

If Left$(Test, 1) = Chr$(39) Then
    If Right(Test, 1) = Chr$(39) Then
        Test = Mid$(Test, 2, Len(Test) - 2)
        RemoveBrackets = 1
    End If
End If
End Function

Private Function IsUrlOrMail(Test As String) As Long
Dim ok As Long
Dim Pos As Long

Pos = InStr(1, Test$, "://", 1)
If Pos > 0 Then
    Pos = InStr(1, Test$, "http", 1)
    If Pos > 0 Then
        ok = 1
    Else
        Pos = InStr(1, Test$, "ftp", 1)
        If Pos > 0 Then
            ok = 11
        End If
    End If
    If ok > 0 Then
        Pos = InStr(1, Test$, ".", 1)
        If Pos = 0 Then
            ok = 0
        End If
    End If
     
Else
    If LCase(Left$(Test$, 4)) = "www." Then
        Pos = InStr(5, Test$, ".", 1)
        If Pos > 0 Then
            ok = 5
        End If
    End If
          
End If

If ok > 0 Then
    IsUrlOrMail = ok
    Exit Function
End If

Pos = InStr(1, Test$, "@", 1)
If Pos > 1 Then
    Pos = InStr(Pos + 1, Test$, ".", 1)
    If Pos > 0 Then
        ok = 101
    End If
End If

IsUrlOrMail = ok
End Function

Private Sub Highlight(Rtb As RichTextBox)
Dim Pos2 As Long
Dim pos1 As Long
Dim br As Long
Dim lnk As Long
Dim ret As Long
Dim l As Long
Dim Text As String
Dim Test As String

Text = Rtb.Text
l = Len(Text$)
pos1 = 1

Do
    Pos2 = InStr(pos1, Text$, " ", 1)
    If Pos2 > pos1 Then
        Test = Mid$(Text, pos1, (Pos2 - pos1))
        br = RemoveBrackets(Test)
        ret = RemoveSign(Test)
        lnk = IsUrlOrMail(Test)
        
        If lnk > 0 Then
            Rtb.SelStart = pos1 - 1 + br
            Rtb.SelLength = Len(Test)
            
            Select Case lnk
                Case 1 To 10
                    Rtb.SelColor = vbBlue
                Case 11 To 20
                    Rtb.SelColor = RGB(0, 127, 0)
                Case Is > 100
                    Rtb.SelColor = vbRed
            End Select
            
            Rtb.SelBold = True
        Else
            Rtb.SelStart = pos1 - 1
            Rtb.SelLength = Len(Test$)
            Rtb.SelColor = 0
            Rtb.SelBold = False
        End If
        
        pos1 = Pos2 + 1
    Else
        If Pos2 = pos1 Then
            pos1 = Pos2 + 1
        End If
        
    End If
Loop Until Pos2 = 0 Or Pos2 >= l
End Sub
