VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   BorderStyle     =   0  'None
   Caption         =   "frmChat"
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7905
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtToSend 
      Height          =   855
      Left            =   120
      MaxLength       =   180
      TabIndex        =   5
      Top             =   2760
      Width           =   5535
   End
   Begin RichTextLib.RichTextBox NRTB 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   873
      _Version        =   393217
      BorderStyle     =   0
      TextRTF         =   $"frmChat.frx":0000
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Send"
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
      Enabled         =   -1  'True
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
   Begin VB.Menu UserPop 
      Caption         =   "UserPop"
      Visible         =   0   'False
      Begin VB.Menu pWhisper 
         Caption         =   "Whisper"
      End
      Begin VB.Menu pAddToFriendlist 
         Caption         =   "Add to Friendlist"
      End
      Begin VB.Menu pIgnoreUser 
         Caption         =   "Ignore User"
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const EM_CHARFROMPOS    As Long = &HD7&
Private Const WM_PASTE          As Long = &H302

Private Sign(255)               As Integer
Private menuUser                As String

Private Sub cmdSend_Click()
'No whitespaces
If LenB(Trim$(txtToSend.Text)) = 0 Then Exit Sub

'Send public message
SendMessage "!message" & pSplit & RTrim$(txtToSend.Text) & pSplit

'Wipeout textbox
txtToSend.Text = vbNullString
End Sub

Private Sub Form_Activate()
frmContainer.cmdChat.Caption = MDI_COMMAND_CHAT
frmMain.ChatNotifyTimer.Enabled = False
End Sub

Private Sub Form_Load()
Top = 0: Left = 0

cmdSend.Caption = CHAT_COMMAND_SEND
cmdClear.Caption = CHAT_COMMAND_CLEAR
pAddToFriendlist.Caption = SOC_COMMAND_FRIEND
pIgnoreUser.Caption = SOC_COMMAND_IGNORE
pWhisper.Caption = SOC_COMMAND_WHISPER

Call InitSigns
End Sub

Private Sub cmdClear_Click()
txtConver.Text = vbNullString
txtToSend.Text = vbNullString
End Sub

Private Sub pAddToFriendlist_Click()
SendMessage "!friend" & pSplit & "-add" & pSplit & menuUser & pSplit
End Sub

Private Sub pIgnoreUser_Click()
SendMessage "!ignore" & pSplit & "-add" & pSplit & menuUser & pSplit
End Sub

Private Sub pWhisper_Click()
With txtToSend
    .Text = "/whisper " & menuUser & " "
    .SelStart = Len(.Text)
    .SetFocus
End With
End Sub

Private Sub txtConver_Change()
Dim hWnd1 As Long
    hWnd1 = GetActiveWindow

'Unlock so we can convert smileys
txtConver.Locked = False

'Create smileys
Call Create_Smileys(txtConver)

'Highlight links, emails and ftp links
Call Highlight(txtConver)

'If window doenst have focus then flash
With frmContainer
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
Dim Text    As String
Dim lnk     As Long
Dim ret     As Long
Dim i       As Long
Dim j       As Long
Dim pTemp   As String

Text = GetWord(txtConver, X, Y)
menuUser = GetRichWordOver(txtConver, X, Y)

lnk = IsUrlOrMail(Text)

If lnk > 0 Then
    ret = RemoveSign(Text)
    ret = RemoveBrackets(Text)

    If lnk > 100 Then
        Text = "mailto:" + Text
    End If

    Call SendLink(Text)
Else
    'Proceed only if the button pressed is right button
    If Button = 2 Then
        With frmSociety.lvOnlineList.ListItems
            For i = 1 To .Count
                'If the it is the user and not your self then proceed
                If LCase$(.Item(i)) = LCase$(menuUser) And Not LCase$(menuUser) = LCase$(frmMain.txtAccount.Text) Then
                    'Check if the user is already added in friend list ( to disable control )
                    With frmSociety.lvFriendList.ListItems
                        For j = 1 To .Count
                            If LCase$(.Item(j)) = LCase$(menuUser) Then
                                pAddToFriendlist.Enabled = False
                                Exit For
                            Else
                                If j = .Count Then
                                    pAddToFriendlist.Enabled = True
                                End If
                            End If
                        Next j
                    End With

                    'Check if user is already beeing ignored ( to disable control )
                    With frmSociety.lvIgnoreList.ListItems
                        For j = 1 To .Count
                            If LCase$(.Item(j)) = LCase$(menuUser) Then
                                pIgnoreUser.Enabled = False
                                Exit For
                            Else
                                If j = .Count Then
                                    pIgnoreUser.Enabled = True
                                End If
                            End If
                        Next j
                    End With

                    PopupMenu UserPop
                    Exit For
                End If
            Next i
        End With
    End If
End If

pAddToFriendlist.Enabled = True
pIgnoreUser.Enabled = True
pWhisper.Enabled = True
End Sub

Private Sub txtToSend_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmdSend_Click
    KeyAscii = 0
End If
End Sub

Public Sub Create_Smileys(RTF As Control)
Dim Smileys()       As String
Dim SmileyResID(13) As Long
Dim Smilestring     As String
Dim SmileFileString As String
Dim Start           As Long
Dim i               As Long

Screen.MousePointer = vbHourglass

Start = 1

Smilestring = _
    ":)," & _
    ":-)," & _
    ":(," & _
    ":-(," & _
    ";)," & _
    ";-)," & _
    ":O," & _
    ":o," & _
    ":D," & _
    ":P," & _
    ":p," & _
    "*cool*," & _
    "*roll*," & _
    "*mad*"

SmileyResID(0) = 101
SmileyResID(1) = 101
SmileyResID(2) = 102
SmileyResID(3) = 102
SmileyResID(4) = 103
SmileyResID(5) = 103
SmileyResID(6) = 104
SmileyResID(7) = 104
SmileyResID(8) = 105
SmileyResID(9) = 106
SmileyResID(10) = 106
SmileyResID(11) = 107
SmileyResID(12) = 108
SmileyResID(13) = 109

Smileys = Split(Smilestring, ",")

For i = LBound(Smileys) To UBound(Smileys)
    While RTF.Find(Smileys(i), Start - 1) >= 0
        Picture1.Picture = LoadResPicture(SmileyResID(i), vbResBitmap)
        RTF.SelStart = RTF.Find(Smileys(i), Start - 1)
        RTF.SelLength = Len(Smileys(i))
        Start = RTF.SelStart + RTF.SelLength + 1
        RTF.SelText = vbNullString
        CopyPictureToRTF RTF, Picture1.Picture
    Wend
    Start = 1
Next i

Screen.MousePointer = vbNormal
End Sub

Private Sub CopyPictureToRTF(RTF As Control, Bild As Picture)
Dim Buf As Variant
Dim Text As String

If Clipboard.GetFormat(vbCFText) Then
    Text = Clipboard.GetText
Else
    Buf = Clipboard.GetData
End If

Clipboard.Clear
Clipboard.SetData Picture1.Picture
DoEvents

SendMessage2 RTF.hwnd, WM_PASTE, 0, 0
DoEvents

Clipboard.Clear
On Error Resume Next
If LenB(Text) <> 0 Then
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

Public Function GetRichWordOver(rch As RichTextBox, X As Single, Y As Single) As String
Dim pAPI            As POINTAPI
Dim pPosition       As Integer
Dim p_START_POS     As Integer
Dim p_END_POS       As Integer
Dim ch              As String
Dim pText           As String
Dim txtlen          As Integer

'Convert the position to pixels.
pAPI.X = X \ Screen.TwipsPerPixelX
pAPI.Y = Y \ Screen.TwipsPerPixelY

pPosition = SendMessage2(rch.hwnd, EM_CHARFROMPOS, 0&, pAPI)
If pPosition <= 0 Then Exit Function
        
pText = rch.Text

For p_START_POS = pPosition To 1 Step -1
    ch = Mid$(rch.Text, p_START_POS, 1)

    If Not ( _
        (ch >= "0" And ch <= "9") Or _
        (ch >= "a" And ch <= "z") Or _
        (ch >= "A" And ch <= "Z") Or _
        ch = "_" _
    ) Then Exit For
Next p_START_POS

p_START_POS = p_START_POS + 1

txtlen = Len(pText)

For p_END_POS = pPosition To txtlen
    ch = Mid$(pText, p_END_POS, 1)
    If Not ( _
        (ch >= "0" And ch <= "9") Or _
        (ch >= "a" And ch <= "z") Or _
        (ch >= "A" And ch <= "Z") Or _
        ch = "_" _
    ) Then Exit For
Next p_END_POS

p_END_POS = p_END_POS - 1

If p_START_POS <= p_END_POS Then
    GetRichWordOver = Mid$(pText, p_START_POS, p_END_POS - p_START_POS + 1)
End If
End Function

Private Function GetWord(Rich As RichTextBox, ByVal X&, ByVal Y&) As String
Dim Pos As Long, P1 As Long, P2 As Long
Dim CHAR As Long
Dim MousePointer As POINTAPI

MousePointer.X = X \ Screen.TwipsPerPixelX
MousePointer.Y = Y \ Screen.TwipsPerPixelY
Pos = SendMessage2(Rich.hwnd, EM_CHARFROMPOS, 0&, MousePointer)
If Pos <= 0 Then Exit Function

For P1 = Pos To 1 Step -1
    CHAR = Asc(Mid$(Rich.Text, P1, 1))
    If Sign(CHAR) = 2 Then
        Exit For
    End If
Next P1
P1 = P1 + 1

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
On Error GoTo leaveUNEXPECTED
Dim Last As Long
    Last = Asc(Right$(Test, 1))

If Sign(Last) = 1 Then
    Test = Left$(Test, Len(Test) - 1)
    RemoveSign = 1
End If
Exit Function
leaveUNEXPECTED:
End Function

Private Function RemoveBrackets(Test As String) As Long
On Error GoTo leaveUNEXP
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
Exit Function
leaveUNEXP:
End Function

Private Function IsUrlOrMail(Test As String) As Long
Dim ok  As Long
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

Private Sub InitSigns()
Dim i    As Long
Dim K    As Long
Dim Test As String

Test = ".,;:?!"
For i = 1 To 6
    K = Asc(Mid$(Test, i, 1))
    Sign(K) = 1
Next i

Test = " " + vbCrLf + Chr$(160)
For i = 1 To Len(Test)
    K = Asc(Mid$(Test, i, 1))
    Sign(K) = 2
Next i
End Sub

Private Sub Highlight(Rtb As RichTextBox)
Dim Pos1    As Long
Dim Pos2    As Long
Dim br      As Long
Dim lnk     As Long
Dim ret     As Long
Dim L       As Long
Dim Text    As String
Dim Test    As String

NRTB.TextRTF = Rtb.TextRTF

Text = Rtb.Text
L = Len(Text$)
Pos1 = 1

Do
    Pos2 = InStr(Pos1, Text$ & " ", " ", 1)
    If Pos2 > Pos1 Then
        Test = Mid$(Text, Pos1, (Pos2 - Pos1))
        br = RemoveBrackets(Test)
        ret = RemoveSign(Test)
        lnk = IsUrlOrMail(Test)

        If lnk > 0 Then
            NRTB.SelStart = Pos1 - 1 + br
            NRTB.SelLength = Len(Test)

            Select Case lnk
                Case 1 To 10: NRTB.SelColor = RGB(34, 0, 204)
                Case 11 To 20: NRTB.SelColor = RGB(0, 127, 0)
                Case Is > 100: NRTB.SelColor = vbRed
            End Select

            NRTB.SelBold = True
        Else
            With NRTB
                .SelStart = Pos1 - 1
                .SelLength = Len(Test$)
                .SelColor = 0
                .SelBold = False
            End With
        End If

        Pos1 = Pos2 + 1
    Else
        If Pos2 = Pos1 Then
            Pos1 = Pos2 + 1
        End If
    End If
Loop Until Pos2 = 0 Or Pos2 >= L

Rtb.TextRTF = NRTB.TextRTF
End Sub

Public Sub WriteText(Text As String)
With txtConver
    .SelStart = Len(.Text)
    .SelRTF = vbCrLf & " [" & Format$(Time, "hh:nn:ss") & "] " & Text
End With
End Sub
