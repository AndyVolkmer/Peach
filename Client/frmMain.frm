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

Private Vali   As Boolean
Public RunOnce As Boolean

Private Sub ChatNotifyTimer_Timer()
Static onORoff As Boolean

If ChatNotifyTimer.Interval <> 500 Then
    ChatNotifyTimer.Interval = 500
End If

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
    With frmMain.Winsock1
        .RemotePort = Setting.SERVER_PORT
        .RemoteHost = Setting.SERVER_IP
        .Connect
    End With
    
    If Not Right$(App.EXEName, 5) = "DEBUG" Then
        'Set Recieve-Request-Winsock to listen
        With frmMain.FSocket2(0)
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
Dim GetMessage  As String
Dim array1()    As String

FSocket.GetData GetMessage

array1 = Split(GetMessage, "#")

Select Case array1(0)
    Case "!acceptfile": frmSendFile.SendF FSocket.RemoteHost
    Case "!denyfile": MsgBox SF_MSG_DECILINED, vbInformation
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

Me.Caption = pCaption

Me.Top = Setting.MAIN_TOP
Me.Left = Setting.MAIN_LEFT
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
        frmMain.Show
        frmMain.WindowState = 0
        Shell_NotifyIcon NIM_DELETE, NID    'Del tray icon
        
    Case WM_LBUTTONDBLCLK
    Case WM_RBUTTONDOWN: frmMain.PopupMenu myPOP
    Case WM_RBUTTONUP
    Case WM_RBUTTONDBLCLK
End Select
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then
    If Vali = False Then
        If Setting.MIN_TICK Then
            MinimizeToTray frmMain
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
    & "Peach is beeing developed by " & pAuthor & _
    " Do not publish this anywhere without " & _
    "permissions of the author.", vbInformation
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
    End With
Else
    With frmForgotPassword
        .Caption = REG_MSG_ERROR_OCCURED
        .lblStatus.Caption = REG_MSG_ERROR
        .cmdRequest.Caption = REG_COMMAND_CLOSE
        .cmdRequest.Visible = True
        .Frame1.Visible = False
        .txtAccount.Visible = False
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
        .Command1.Visible = True
        .Command1.Caption = REG_COMMAND_SUBMIT
    End With
Else
    With frmForgotPassword
        .Caption = FP_CAPTION
        .Frame1.Visible = True
        .txtAccount.Visible = True
        .txtSecretAnswer.Visible = True
        .cmbSecretQuestion.Visible = True
        .cmdRequest.Visible = True
        .cmdRequest.Caption = FP_COMMAND_REQUEST
    End With
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub RegSock_DataArrival(ByVal bytesTotal As Long)
Dim GetMessage  As String
Dim array1()    As String

RegSock.GetData GetMessage

array1 = Split(GetMessage, "#")

Select Case array1(0)
    Case "!nameexist"
        MsgBox REG_MSG_ACCOUNT_EXIST, vbInformation
        With frmCreateAccount
            .txtAccount = vbNullString
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
        MsgBox FP_MSG_SUCCESSFULL & array1(1), vbInformation
        Unload frmForgotPassword
        
    Case "!account_not_exist"
        MsgBox MDI_MSG_WRONG_ACCOUNT, vbInformation
        With frmForgotPassword
            .txtAccount = vbNullString
            .txtAccount.SetFocus
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
    End With
Else
    With frmForgotPassword
        .Caption = REG_MSG_ERROR_OCCURED
        .lblStatus.Caption = REG_MSG_ERROR & vbCrLf & vbCrLf & "Error: " & Description
        .cmdRequest.Caption = REG_COMMAND_CLOSE
        .cmdRequest.Visible = True
        .Frame1.Visible = False
        .txtAccount.Visible = False
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
With frmMain
    SendMessage "!login#" & .txtAccount.Text & "#" & .txtPassword.Text & "#" & CURRENT_LANG & "#"
End With
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim p_PreArray()    As String
Dim k               As Long
Dim i               As Long

Dim GetCommand      As String
Dim StrArr()        As String
Dim StrArr2()       As String
Dim GetMessage      As String
Dim Buffer          As String

'We get the message
Winsock1.GetData GetMessage
DoEvents

'Do first array to avoid spam
p_PreArray = Split(GetMessage, Chr(24) & Chr(25))

'Start looping through
For k = 0 To UBound(p_PreArray) - 1
    'We split the message into an array
    StrArr = Split(p_PreArray(k), "#")
        
    'Assign the variables to the array
    GetCommand = StrArr(0)
    
    Select Case GetCommand
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
            
        'Show server information
        Case "!server_info"
            For i = 1 To UBound(StrArr)
                Buffer = Buffer & vbCrLf & " " & StrArr(i)
            Next i
            With frmChat.txtConver
                .SelStart = Len(.Text)
                .SelRTF = Buffer
            End With
            SendMessage "!ignore#-get#"
            
        'We can't login
        Case "!decilined"
            Disconnect
            MsgBox MDI_MSG_NAME_TAKEN, vbInformation
            
        'We can login
        Case "!accepted"
            frmMain.cmdSwitch.Enabled = True
            
            With frmContainer
                .Top = frmMain.Top
                .Left = frmMain.Left
            End With
            
            SetupChildForm frmChat
            txtAccount.Text = StrArr(1)
            SendMessage "!connected#"
            
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
            Dim f_array() As String
            Dim j As Long
            
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
                
                'Ask for server information
                If RunOnce = False Then
                    RunOnce = True
                    SendMessage "!server_info#"
                End If
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
            
            'Ask for friendlist
            SendMessage "!friend#-get#"
            
        'We get login answer here
        Case "!login"
            Select Case StrArr(1)
                Case "Password"
                    MsgBox MDI_MSG_WRONG_PASSWORD, vbInformation
                    Disconnect
                    With frmMain
                        .txtPassword.Text = vbNullString
                        .txtPassword.SetFocus
                    End With
                    Exit Sub
                    
                Case "Account"
                    MsgBox MDI_MSG_WRONG_ACCOUNT, vbInformation
                    Disconnect
                    With frmMain
                        .txtAccount.Text = vbNullString
                        .txtAccount.SetFocus
                    End With
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
                    MsgBox Replace(MDI_MSG_ALREADY_IN_IGNORE_LIST, "%u", StrArr(2)), vbInformation
                    
                Case "MSG_ALREADY_IN_FRIEND_LIST"
                    MsgBox Replace(MDI_MSG_ALREADY_IN_FRIEND_LIST, "%u", StrArr(2)), vbInformation
                    
                Case "MSG_ACCOUNT_NOT_EXIST"
                    MsgBox Replace(MDI_MSG_ACCOUNT_NOT_EXIST, "%u", StrArr(2)), vbInformation
                    
                Case Else
                    MsgBox StrArr(1), vbInformation
                    
            End Select
            
        Case Else
            If Not Screen.ActiveForm.Name = frmChat.Name Then
                ChatNotifyTimer.Enabled = True
            End If
            With frmChat.txtConver
                .SelStart = Len(.Text)
                .SelRTF = vbCrLf & Space(1) & "[" & Format$(Time, "hh:nn:ss") & "]" & Space(1) & p_PreArray(k)
            End With
            
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
        .SendData "!filerequest#" & frmMain.CDialog.FileTitle & "#" & frmMain.txtAccount & "#"
    End If
End With
End Sub

Private Sub FSocket2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim GetMessage  As String
Dim array1()    As String

FSocket2(Index).GetData GetMessage

array1 = Split(GetMessage, "#")

If array1(0) = "!filerequest" Then
    SF_MSG_INCOMMING_FILE = Replace$(SF_MSG_INCOMMING_FILE, "%f", array1(1))
    SF_MSG_INCOMMING_FILE = Replace$(SF_MSG_INCOMMING_FILE, "%u", array1(2))
    
    If MsgBox(SF_MSG_INCOMMING_FILE, vbYesNo + vbQuestion) = vbYes Then
        FSocket2(Index).SendData "!acceptfile#"
        frmSendFile2.Show
    Else
        FSocket2(Index).SendData "!denyfile#"
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
