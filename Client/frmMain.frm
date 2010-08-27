VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Peach"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   5955
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   5955
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "^"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5490
      TabIndex        =   9
      Top             =   50
      Width           =   375
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1890
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox txtAccount 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1890
      MaxLength       =   15
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "cmdConnect"
      Height          =   495
      Left            =   1890
      TabIndex        =   0
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Timer STimer 
      Enabled         =   0   'False
      Left            =   0
      Top             =   480
   End
   Begin VB.Timer ChatNotifyTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   480
   End
   Begin MSWinsockLib.Winsock FSocket2 
      Index           =   0
      Left            =   960
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock FSocket 
      Left            =   480
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock RegSock 
      Left            =   1440
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   1440
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label lblAuthor 
      BackColor       =   &H00F4F4F4&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00F4F4F4&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label lblPassword 
      BackColor       =   &H00F4F4F4&
      Caption         =   "lblPassword"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1890
      TabIndex        =   6
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label lblAccount 
      BackColor       =   &H00F4F4F4&
      Caption         =   "lblAccount"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1890
      TabIndex        =   5
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label lblCreateAccount 
      Alignment       =   2  'Center
      Caption         =   "lblCreateAccount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2010
      TabIndex        =   4
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label lblForgotPassword 
      Alignment       =   2  'Center
      Caption         =   "lblForgotPassword"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2010
      TabIndex        =   3
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Menu mainMenu 
      Caption         =   "mainMenu"
      Begin VB.Menu menuConfig 
         Caption         =   "menuConfig"
      End
      Begin VB.Menu menuUpdate 
         Caption         =   "menuUpdate"
      End
   End
   Begin VB.Menu myPOP 
      Caption         =   "myPOP"
      Visible         =   0   'False
      Begin VB.Menu ExitC 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Vali As Boolean

Private Sub ChatNotifyTimer_Timer()
Static onORoff As Boolean

If ChatNotifyTimer.Interval <> 500 Then ChatNotifyTimer.Interval = 500

With frmContainer.cmdChat
    If onORoff Then
        .Caption = MDI_COMMAND_CHAT
        onORoff = False
    Else
        .Caption = "! - " & MDI_COMMAND_CHAT & " - !"
        onORoff = True
    End If
End With
End Sub

Private Sub cmdSwitch_Click()
SetupForm frmContainer
SetupChildForm frmChat
End Sub

Private Sub ExitC_Click()
CloseThis
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Setting.ASK_TICK Then
    If MsgBox(MDI_MSG_UNLOAD, vbQuestion + vbYesNo) = vbNo Then
        Cancel = 1
    End If
End If
End Sub

Private Sub FSocket2_Close(Index As Integer)
Unload FSocket2(Index)
End Sub

Private Sub FSocket2_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Unload FSocket2(Index)
End Sub

Private Sub txtAccount_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmdConnect_Click
    KeyAscii = 0
End If
If txtAccount.BackColor <> vbWhite Then txtAccount.BackColor = vbWhite
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmdConnect_Click
    KeyAscii = 0
End If
If txtPassword.BackColor <> vbWhite Then txtPassword.BackColor = vbWhite
End Sub

Public Sub cmdConnect_Click()
If cmdConnect.Caption = CONFIG_COMMAND_CONNECT Then
    If CheckTx(txtAccount, CONFIG_MSG_ACCOUNT) Then Exit Sub
    If CheckTx(txtPassword, CONFIG_MSG_PASSWORD) Then Exit Sub

    'Connect winsocks
    With Winsock1
        .RemotePort = Setting.SERVER_PORT
        .RemoteHost = Setting.SERVER_IP
        .Connect
    End With

    If Not Right$(App.EXEName, 5) = "DEBUG" Then
        'Set Recieve-Request-Winsock to listen
        With FSocket2(0)
            .LocalPort = aPort
            .Listen
        End With

        'Set Recieve-File-Winsock to listen
        With frmSendFile2.SckReceiveFile(0)
            .LocalPort = bPort
            .Listen
        End With
    End If

    SwitchButtons False, True
Else
    Disconnect
End If
End Sub

Private Sub Exit_Click()
CloseThis
End Sub

Private Sub FSocket_Close()
Disconnect
End Sub

Private Sub FSocket_DataArrival(ByVal bytesTotal As Long)
Dim Message As String

FSocket.GetData Message

Select Case Message
    Case "1": frmSendFile.SendF FSocket.RemoteHost
    Case "0": MsgBox SF_MSG_DECILINED, vbInformation
End Select
End Sub

Private Sub Form_Load()
mainMenu.Caption = MDI_MENU
menuConfig.Caption = CONFIG_COMMAND_SETTINGS
menuUpdate.Caption = CONFIG_COMMAND_UPDATE

ExitC.Caption = REG_COMMAND_CLOSE

lblCreateAccount.Caption = CONFIG_COMMAND_REGISTER
lblForgotPassword.Caption = CONFIG_COMMAND_FORGOT_PASSWORD
lblAccount.Caption = CONFIG_LABEL_ACCOUNT
lblPassword.Caption = CONFIG_LABEL_PASSWORD
lblAuthor.Caption = "Author: " & pAuthor
lblVersion.Caption = "Version: " & pRev

cmdConnect.Caption = CONFIG_COMMAND_CONNECT

Caption = pCaption
Top = Setting.MAIN_TOP
Left = Setting.MAIN_LEFT

txtAccount.Text = Setting.ACCOUNT
txtPassword.Text = Setting.PASSWORD
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MSG As Long
    MSG = X / Screen.TwipsPerPixelX

Select Case MSG
    Case WM_LBUTTONDOWN
    Case WM_LBUTTONUP
        Vali = True
        Show
        WindowState = 0
        Shell_NotifyIcon NIM_DELETE, NID    'Del tray icon

    Case WM_LBUTTONDBLCLK
    Case WM_RBUTTONDOWN: PopupMenu myPOP
    Case WM_RBUTTONUP
    Case WM_RBUTTONDBLCLK
End Select
End Sub

Private Sub Form_Resize()
If WindowState = 1 Then
    If Vali = False Then
        If Setting.MIN_TICK Then
            MinimizeToTray Me
        End If
    End If
    Vali = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
CloseThis
End Sub

Private Sub lblCreateAccount_Click()
frmCreateAccount.Show 1
End Sub

Private Sub lblForgotPassword_Click()
frmForgotPassword.Show 1
End Sub

Private Sub menuConfig_Click()
frmSettings.Show 1
End Sub

Private Sub lblAuthor_Click()
showAbout
End Sub

Private Sub lblVersion_Click()
showAbout
End Sub

Private Sub showAbout()
MsgBox "Author: " & pAuthor & vbCrLf _
     & "Version: " & pRev & vbCrLf & vbCrLf _
     & "Peach is beeing developed by " & pAuthor _
     & " Do not publish this anywhere without " _
     & "permissions of the author.", vbInformation
End Sub

Private Sub menuUpdate_Click()
If FileExists(App.Path & "\peachUpdater.exe") Then
    Shell App.Path & "\peachUpdater.exe", vbNormalFocus
Else
    MsgBox CONFIG_MSG_UPDATE_FILE, vbInformation
End If
End Sub

Private Sub RegSock_Close()
If ACC_SWITCH = "REG" Then
    With frmCreateAccount
        .Caption = REG_MSG_ERROR_OCCURED
        .Label4.Caption = REG_MSG_ERROR
        .Command1.Caption = REG_COMMAND_CLOSE
        .Command1.Visible = True
        .Frame1.Visible = False
        .Check1.Visible = False
        .cmbSecretQuestion.Visible = False
        .cmbGender.Visible = False
        .txtSecretAnswer.Visible = False
        .txtAccount.Visible = False
        .txtPassword1.Visible = False
        .txtPassword2.Visible = False
        .txtEmail.Visible = False
    End With
Else
    With frmForgotPassword
        .Caption = REG_MSG_ERROR_OCCURED
        .lblStatus.Caption = REG_MSG_ERROR
        .cmdRequest.Caption = REG_COMMAND_CLOSE
        .cmdRequest.Visible = True
        .Frame1.Visible = False
        .txtEmail.Visible = False
        .cmbSecretQuestion.Visible = False
        .txtSecretAnswer.Visible = False
    End With
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub RegSock_Connect()
If ACC_SWITCH = "REG" Then
    With frmCreateAccount
        .Caption = REG_CAPTION
        .Frame1.Visible = True
        .Check1.Visible = True
        .txtAccount.Visible = True
        .txtPassword1.Visible = True
        .txtPassword2.Visible = True
        .cmbSecretQuestion.Visible = True
        .cmbGender.Visible = True
        .txtSecretAnswer.Visible = True
        .txtEmail.Visible = True
        .Command1.Visible = True
        .Command1.Caption = REG_COMMAND_SUBMIT
    End With
Else
    With frmForgotPassword
        .Caption = FP_CAPTION
        .Frame1.Visible = True
        .txtEmail.Visible = True
        .txtSecretAnswer.Visible = True
        .cmbSecretQuestion.Visible = True
        .cmdRequest.Visible = True
        .cmdRequest.Caption = FP_COMMAND_REQUEST
    End With
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub RegSock_DataArrival(ByVal bytesTotal As Long)
Dim pArray() As String
Dim Message  As String
Dim Text     As String

RegSock.GetData Message

pArray = Split(Message, "#")

Select Case pArray(0)
    Case "!nameexist"
        MsgBox REG_MSG_ACCOUNT_EXIST, vbInformation

        With frmCreateAccount
            .txtAccount = vbNullString
            .txtAccount.SetFocus
        End With

    Case "!emailtaken"
        MsgBox REG_MSG_EMAIL_TAKEN, vbInformation

        With frmCreateAccount
            .txtEmail = vbNullString
            .txtAccount.SetFocus
        End With

    Case "!done"
        MsgBox REG_MSG_SUCCESSFULLY, vbInformation
        Unload frmCreateAccount

    Case "!error_fp"
        MsgBox FP_MSG_WRONG_ANSWER, vbInformation

        With frmForgotPassword
            .txtSecretAnswer = vbNullString
            .txtSecretAnswer.SetFocus
        End With

    Case "!successfull"
        Text = Replace$(FP_MSG_SUCCESSFULL, "%p", pArray(1))
        Text = Replace$(Text, "%u", pArray(2))

        MsgBox Text, vbInformation

        txtPassword = pArray(1)
        txtAccount = pArray(2)

        Unload frmForgotPassword

    Case "!email_not_exist"
        MsgBox FP_MSG_WRONG_EMAIL, vbInformation

        With frmForgotPassword
            .txtEmail = vbNullString
            .txtEmail.SetFocus
        End With

End Select
End Sub

Private Sub RegSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If ACC_SWITCH = "REG" Then
    With frmCreateAccount
        .Caption = REG_MSG_ERROR_OCCURED
        .Label4.Caption = REG_MSG_ERROR & vbCrLf & vbCrLf & "Error: " & Description
        .Command1.Caption = REG_COMMAND_CLOSE
        .Command1.Visible = True
        .Frame1.Visible = False
        .Check1.Visible = False
        .cmbSecretQuestion.Visible = False
        .cmbGender.Visible = False
        .txtSecretAnswer.Visible = False
        .txtEmail.Visible = False
    End With
Else
    With frmForgotPassword
        .Caption = REG_MSG_ERROR_OCCURED
        .lblStatus.Caption = REG_MSG_ERROR & vbCrLf & vbCrLf & "Error: " & Description
        .cmdRequest.Caption = REG_COMMAND_CLOSE
        .cmdRequest.Visible = True
        .Frame1.Visible = False
        .txtEmail.Visible = False
        .cmbSecretQuestion.Visible = False
        .txtSecretAnswer.Visible = False
    End With
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Winsock1_Close()
Disconnect
End Sub

Private Sub Winsock1_Connect()
SendMessage "!login#" & txtAccount.Text & "#" & GetMD5(txtPassword.Text) & "#"
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim k          As Long
Dim i          As Long
Dim Message    As String
Dim PreArray() As String
Dim Command    As String
Dim StrArr()   As String
Dim StrArr2()  As String
Dim Buffer     As String

'Get data from socket
Winsock1.GetData Message
DoEvents

'Do first array to avoid incorrect package reading
'Chr(24) and Chr(25) are the delimeters for each packet
PreArray = Split(Message, Chr(24) & Chr(25))

'Start looping through
For k = 0 To UBound(PreArray) - 1

'We split the message into an array
StrArr = Split(PreArray(k), "#")

'Assign the variables to the array
If UBound(StrArr) > -1 Then
    Command = StrArr(0)
End If

Select Case Command
    Case "!clear"
        frmChat.txtConver.Text = vbNullString

    Case "!split_text"
        For i = 1 To UBound(StrArr)
            Buffer = Buffer & vbCrLf & " " & StrArr(i)
        Next i

        With frmChat.txtConver
            .SelStart = Len(.Text)
            .SelRTF = Buffer
        End With

    Case "!namechange"
        txtAccount.Text = StrArr(1)

    'We can't login
    Case "!decilined"
        Disconnect
        MsgBox MDI_MSG_NAME_TAKEN, vbInformation

    'We can login
    Case "!accepted"
        cmdSwitch.Enabled = True

        With frmContainer
            .Top = Top
            .Left = Left
        End With

        txtAccount.Text = StrArr(1)
        SetupChildForm frmChat
        SendMessage "!connected#"
        SendMessage "!friend#-get#"
        SendMessage "!ignore#-get#"

    'Wipe out current ignore list and insert new values
    Case "!update_ignore"
        With frmSociety.lvIgnoreList.ListItems
            .Clear
            For i = 1 To UBound(StrArr) - 1
                .Add , , StrArr(i)
            Next i
        End With

    'Wipe out current friend list and insert new values
    Case "!update_friends"
        Dim f_array()   As String
        Dim j           As Long

        With frmSociety.lvFriendList.ListItems
            .Clear
            For i = LBound(StrArr) + 1 To UBound(StrArr) - 1
                f_array = Split(StrArr(i), "$")

                'Add account name of friend
                .Add , , f_array(0)
                j = .Count

                'Write down status
                .Item(j).SubItems(1) = f_array(1)
                .Item(j).ListSubItems(1).Bold = True

                'Color the listview with apropiate color
                If Len(f_array(1)) = 6 Then
                    .Item(j).ListSubItems(1).ForeColor = RGB(0, 132, 43)
                Else
                    .Item(j).ListSubItems(1).ForeColor = RGB(132, 0, 0)
                End If
            Next i
        End With

    'Wipe out current list and insert new values
    Case "!update_online"
        'Clear the current list so we don't get multiply entries
        frmSociety.lvOnlineList.ListItems.Clear
        frmSendFile.Combo1.Clear

        'Go through array and add users
        For i = LBound(StrArr) + 1 To UBound(StrArr) - 1
            frmSociety.lvOnlineList.ListItems.Add , , StrArr(i)
            frmSendFile.Combo1.AddItem StrArr(i)
        Next i

    'We get login answer here
    Case "!login"
        Select Case StrArr(1)
            Case "Password"
                MsgBox MDI_MSG_WRONG_PASSWORD, vbInformation
                Disconnect
                txtPassword.Text = vbNullString
                txtPassword.SetFocus
                Exit Sub

            Case "Account"
                MsgBox MDI_MSG_WRONG_ACCOUNT, vbInformation
                Disconnect
                txtAccount.Text = vbNullString
                txtAccount.SetFocus
                Exit Sub

            Case "Banned"
                MsgBox MDI_MSG_BANNED, vbInformation
                Disconnect
                Exit Sub

        End Select

    'We get ip here
    Case "!iprequest"
        'Connect new winsock to client
        FSocket.Close
        With FSocket
            .RemoteHost = StrArr(1)
            .RemotePort = aPort
            .Connect
        End With

        'Start timer to send file
        With STimer
            .Interval = 5
            .Enabled = True
        End With

    'Function to display information messageboxes
    Case "!msgbox"
        Select Case StrArr(1)
            Case "MSG_CANT_ADD_YOU"
                MsgBox MDI_MSG_CANT_ADD_YOU, vbInformation

            Case "MSG_ALREADY_IN_IGNORE_LIST"
                MsgBox Replace$(MDI_MSG_ALREADY_IN_IGNORE_LIST, "%u", StrArr(2)), vbInformation

            Case "MSG_ALREADY_IN_FRIEND_LIST"
                MsgBox Replace$(MDI_MSG_ALREADY_IN_FRIEND_LIST, "%u", StrArr(2)), vbInformation

            Case "MSG_ACCOUNT_NOT_EXIST"
                MsgBox Replace$(MDI_MSG_ACCOUNT_NOT_EXIST, "%u", StrArr(2)), vbInformation

            Case Else
                MsgBox StrArr(1), vbInformation

        End Select

    Case "!channel_password"
        SendMessage "!channel_password#" & InputBox(Replace(CH_MSG_PASSWORD, "%c", StrArr(1))) & "#" & StrArr(1)

    Case "!pmessage"
        Dim temp As String

        Select Case StrArr(1)
            Case "online"
                frmChat.WriteText Replace(MSG_USER_ONLINE, "%u", StrArr(2))

            Case "offline"
                frmChat.WriteText Replace(MSG_USER_OFFLINE, "%u", StrArr(2))

            Case "announce"
                temp = Replace(MSG_ANNOUNCE, "%f", StrArr(2))
                temp = Replace(temp, "%u", StrArr(3))
                temp = Replace(temp, "%m", StrArr(4))

                frmChat.WriteText temp

            Case "table_reload"
                temp = Replace(MSG_TABLE_RELOAD, "%u", StrArr(2))
                temp = Replace(temp, "%t", StrArr(3))
                temp = Replace(temp, "%i", StrArr(4))

                frmChat.WriteText temp

            Case "config_reload"
                temp = Replace(MSG_CONFIG_RELOAD, "%u", StrArr(2))
                temp = Replace(temp, "%t", StrArr(3))

                frmChat.WriteText temp

            Case "incorrect_syntax"
                frmChat.WriteText Replace(MSG_INCORRECT_SYNTAX, "%s", StrArr(2))

            Case "table_not_exist"
                frmChat.WriteText MSG_TABLE_NOT_EXIST

            Case "table_cant_reload"
                frmChat.WriteText MSG_TABLE_CANT_RELOAD

            Case "user_not_found"
                frmChat.WriteText Replace(MSG_USER_NOT_FOUND, "%u", StrArr(2))

            Case "deleted_account"
                temp = Replace(MSG_DELETED_ACCOUNT, "%u", StrArr(2))
                temp = Replace(temp, "%id", StrArr(3))

                frmChat.WriteText temp

            Case "gm_flag_enable"
                frmChat.WriteText MSG_GM_FLAG_ENABLE

            Case "gm_flag_disable"
                frmChat.WriteText MSG_GM_FLAG_DISABLE

            Case "unknown_command"
                frmChat.WriteText MSG_UNKNOWN_COMMAND

            Case "muted"
                frmChat.WriteText MSG_MUTED

            Case "flood_protection"
                frmChat.WriteText MSG_FLOOD_PROTECTION

            Case "roll"
                temp = Replace(MSG_ROLL, "%u", StrArr(2))
                temp = Replace(temp, "%r", StrArr(3))
                temp = Replace(temp, "%minR", StrArr(4))
                temp = Replace(temp, "%maxR", StrArr(5))

                frmChat.WriteText temp

            Case "not_afk"
                frmChat.WriteText MSG_NOT_AFK

            Case "afk"
                frmChat.WriteText MSG_AFK

            Case "online_time"
                frmChat.WriteText Replace(MSG_ONLINE_TIME, "%t", StrArr(2))

            Case "valid_channel"
                frmChat.WriteText MSG_VALID_CHANNEL

            Case "already_in_channel"
                frmChat.WriteText Replace(MSG_ALREADY_IN_CHANNEL, "%c", StrArr(2))

            Case "not_in_channel"
                frmChat.WriteText Replace(MSG_NOT_IN_CHANNEL, "%c", StrArr(2))

            Case "channel_announcements"
                temp = Replace(MSG_CHANNEL_ANNOUNCEMENTS, "%c", StrArr(2))
                temp = Replace(temp, "%u", StrArr(3))

                frmChat.WriteText temp

            Case "channel_password"
                temp = Replace(MSG_CHANNEL_PASSWORD, "%c", StrArr(2))
                temp = Replace(temp, "%p", StrArr(3))

                frmChat.WriteText temp

            Case "channel_wrong_password"
                frmChat.WriteText Replace(MSG_CHANNEL_WRONG_PASSWORD, "%c", StrArr(2))

            Case "not_channel_leader"
                frmChat.WriteText MSG_NOT_CHANNEL_LEADER

            Case "message_blocked"
                frmChat.WriteText MSG_MESSAGE_BLOCKED

            Case "cant_whisper_self"
                frmChat.WriteText MSG_CANT_WHISPER_SELF

            Case "is_ignoring_you"
                frmChat.WriteText Replace(MSG_IS_IGNORING_YOU, "%t", StrArr(2))

            Case "you_whisper_to"
                temp = Replace(MSG_YOU_WHISPER_TO, "%t", StrArr(2))
                temp = Replace(temp, "%m", StrArr(3))

                frmChat.WriteText temp

            Case "target_is_afk"
                frmChat.WriteText Replace(MSG_TARGET_IS_AFK, "%t", StrArr(2))

            Case "whisper"
                temp = Replace(MSG_WHISPER, "%f", StrArr(2))
                temp = Replace(temp, "%u", StrArr(3))
                temp = Replace(temp, "%m", StrArr(4))

                frmChat.WriteText temp

            Case "user_already_muted"
                frmChat.WriteText Replace(MSG_USER_ALREADY_MUTED, "%u", StrArr(2))

            Case "user_is_not_muted"
                frmChat.WriteText Replace(MSG_IS_NOT_MUTED, "%u", StrArr(2))

            Case "muted_by"
                temp = Replace(MSG_MUTED_BY, "%t", StrArr(2))
                temp = Replace(temp, "%u", StrArr(3))

                frmChat.WriteText temp

            Case "unmuted_by"
                temp = Replace(MSG_UNMUTED_BY, "%t", StrArr(2))
                temp = Replace(temp, "%u", StrArr(3))

                frmChat.WriteText temp

            Case "muted_by_reason"
                temp = Replace(MSG_MUTED_BY_REASON, "%t", StrArr(2))
                temp = Replace(temp, "%u", StrArr(3))
                temp = Replace(temp, "%r", StrArr(4))

                frmChat.WriteText temp
            
            Case "unmuted_by_reason"
                temp = Replace(MSG_UNMUTED_BY_REASON, "%t", StrArr(2))
                temp = Replace(temp, "%u", StrArr(3))
                temp = Replace(temp, "%r", StrArr(4))

                frmChat.WriteText temp

            Case "already_banned"
                frmChat.WriteText Replace(MSG_ALREADY_BANNED, "%u", StrArr(2))

            Case "already_unbanned"
                frmChat.WriteText Replace(MSG_ALREADY_UNBANNED, "%u", StrArr(2))

            Case "banned_by"
                temp = Replace(MSG_BANNED_BY, "%t", StrArr(2))
                temp = Replace(temp, "%u", StrArr(3))

                frmChat.WriteText temp

            Case "unbanned_by"
                temp = Replace(MSG_UNBANNED_BY, "%t", StrArr(2))
                temp = Replace(temp, "%u", StrArr(3))

                frmChat.WriteText temp

            Case "banned_by_reason"
                temp = Replace(MSG_BANNED_BY_REASON, "%t", StrArr(2))
                temp = Replace(temp, "%u", StrArr(3))
                temp = Replace(temp, "%r", StrArr(4))

                frmChat.WriteText temp

            Case "unbanned_by_reason"
                temp = Replace(MSG_UNBANNED_BY_REASON, "%t", StrArr(2))
                temp = Replace(temp, "%u", StrArr(3))
                temp = Replace(temp, "%r", StrArr(4))

                frmChat.WriteText temp

            Case "successfull_rename"
                temp = Replace(MSG_SUCCESSFULL_RENAME, "%u", StrArr(2))
                temp = Replace(temp, "%t", StrArr(3))

                frmChat.WriteText temp

            Case "renamed_you_to"
                temp = Replace(MSG_RENAMED_YOU_TO, "%u", StrArr(2))
                temp = Replace(temp, "%t", StrArr(3))

                frmChat.WriteText temp

            Case "user_already_used"
                frmChat.WriteText Replace(MSG_USER_ALREADY_USED, "%u", StrArr(2))

            Case "level_incorrect_value"
                frmChat.WriteText MSG_LEVEL_INCORRECT_VALUE

            Case "successfull_level"
                temp = Replace(MSG_SUCCESSFULL_LEVEL, "%u", StrArr(2))
                temp = Replace(temp, "%l", StrArr(3))
                
                frmChat.WriteText temp

            Case "changed_your_level"
                temp = Replace(MSG_CHANGED_YOUR_LEVEL, "%u", StrArr(2))
                temp = Replace(temp, "%l", StrArr(3))

                frmChat.WriteText temp

            Case "gender_incorrect_value"
                frmChat.WriteText MSG_GENDER_INCORRECT_VALUE

            Case "successfull_gender"
                temp = Replace(MSG_SUCCESSFULL_GENDER, "%u", StrArr(2))
                temp = Replace(temp, "%g", StrArr(3))

                frmChat.WriteText temp

            Case "changed_your_gender"
                temp = Replace(MSG_CHANGED_YOUR_GENDER, "%u", StrArr(2))
                temp = Replace(temp, "%g", StrArr(3))

                frmChat.WriteText temp

            Case "successfull_password"
                temp = Replace(MSG_SUCCESSFULL_PASSWORD, "%u", StrArr(2))
                temp = Replace(temp, "%p", StrArr(3))

                frmChat.WriteText temp

            Case "changed_your_password"
                temp = Replace(MSG_CHANGED_YOUR_PASSWORD, "%u", StrArr(2))
                temp = Replace(temp, "%p", StrArr(3))

                frmChat.WriteText temp

            Case "successfull_email"
                temp = Replace(MSG_SUCCESSFULL_EMAIL, "%u", StrArr(2))
                temp = Replace(temp, "%e", StrArr(3))

                frmChat.WriteText temp

            Case "changed_your_email"
                temp = Replace(MSG_CHANGED_YOUR_EMAIL, "%u", StrArr(2))
                temp = Replace(temp, "%e", StrArr(3))

                frmChat.WriteText temp

            Case "joined_channel"
                temp = Replace(MSG_JOINED_CHANNEL, "%c", StrArr(2))

                frmChat.WriteText temp

            Case "channel_user_join"
                temp = Replace(MSG_CHANNEL_USER_JOIN, "%c", StrArr(2))
                temp = Replace(temp, "%u", StrArr(3))

                frmChat.WriteText temp

            Case "channel_user_leave"
                temp = Replace(MSG_CHANNEL_USER_LEAVE, "%c", StrArr(2))
                temp = Replace(temp, "%u", StrArr(3))

                frmChat.WriteText temp

            Case "message_sent_offline"
                temp = Replace(MSG_MESSAGE_SENT_OFFLINE, "%t", StrArr(2))

                frmChat.WriteText temp

            Case "offline_message"
                temp = Replace(MSG_OFFLINE_MESSAGE, "%from", StrArr(2))
                temp = Replace(temp, "%message", StrArr(3))
                temp = Replace(temp, "%time", StrArr(4))

                frmChat.WriteText temp

            Case "incorrect_password"
                frmChat.WriteText MDI_MSG_WRONG_PASSWORD

            Case "successfull_sudo"
                frmChat.WriteText MSG_SUCCESSFULL_SUDO_LOGIN

        End Select

    Case Else
        If Not Screen.ActiveForm.Name = frmChat.Name Then
            ChatNotifyTimer.Enabled = True
        End If
        frmChat.WriteText PreArray(k)

End Select
Next k
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Disconnect
End Sub

Private Sub STimer_Timer()
With FSocket
    If .State = 7 Then
        STimer.Enabled = False
        .SendData "!filerequest#" & CDialog.FileTitle & "#" & txtAccount & "#"
    End If
End With
End Sub

Private Sub FSocket2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim pArray() As String
Dim Message  As String
Dim Text     As String

FSocket2(Index).GetData Message

pArray = Split(Message, "#")

If pArray(0) = "!filerequest" Then
    Text = Replace$(SF_MSG_INCOMMING_FILE, "%f", pArray(1))
    Text = Replace$(Text, "%u", pArray(2))

    If MsgBox(Text, vbYesNo + vbQuestion) = vbYes Then
        FSocket2(Index).SendData "1"
        frmSendFile2.Show
    Else
        FSocket2(Index).SendData "0"
    End If
End If
End Sub

Private Sub FSocket2_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim i As Long
    i = LoadSocket

FSocket2(i).LocalPort = aPort
FSocket2(i).Accept requestID
End Sub

Private Function GetFreeSocket() As Long
Dim i As Long
Dim j As Long

On Error GoTo HandleErrorFreeSocket
For i = FSocket2.LBound + 1 To FSocket2.UBound
    j = FSocket2(i).LocalIP
Next i

GetFreeSocket = FSocket2.UBound + 1

Exit Function
HandleErrorFreeSocket:
    GetFreeSocket = i
End Function

Private Function LoadSocket() As Integer
Dim i As Long
    i = GetFreeSocket

Load FSocket2(i)
LoadSocket = i
End Function

Private Function CheckTx(txtBox As TextBox, mBox As String) As Boolean
If LenB(Trim$(txtBox)) = 0 Then
    MsgBox mBox, vbInformation
    txtBox.BackColor = vbYellow
    txtBox.SetFocus
    CheckTx = True
End If
End Function
