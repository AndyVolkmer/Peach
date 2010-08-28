VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00F4F4F4&
   Caption         =   " Peach Server"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7395
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   5
      Top             =   4740
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   7938
            MinWidth        =   7938
         EndProperty
      EndProperty
      Enabled         =   0   'False
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
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   7395
      TabIndex        =   0
      Top             =   0
      Width           =   7395
      Begin VB.CommandButton Command7 
         Caption         =   "OFF-M"
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
         Left            =   6090
         TabIndex        =   8
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "CHAN"
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
         Left            =   5130
         TabIndex        =   7
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "FL / IL"
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
         Left            =   4170
         TabIndex        =   6
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "USER"
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
         Left            =   3210
         TabIndex        =   4
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ACC"
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
         Left            =   2250
         TabIndex        =   3
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "CHAT"
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
         Left            =   1290
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CONF"
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
         Left            =   330
         TabIndex        =   1
         Top             =   120
         Width           =   975
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

Private Sub Command1_Click()
SetupForms frmConfig
End Sub

Private Sub Command2_Click()
SetupForms frmChat
End Sub

Private Sub Command3_Click()
SetupForms frmAccountPanel
End Sub

Private Sub Command4_Click()
SetupForms frmPanel
End Sub

Private Sub Command5_Click()
SetupForms frmFriendIgnoreList
End Sub

Private Sub Command6_Click()
SetupForms frmChannel
End Sub

Private Sub Command7_Click()
SetupForms frmOfflineMessages
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MSG As Long
    MSG = X / Screen.TwipsPerPixelX

Select Case MSG
    Case WM_LBUTTONDOWN
    Case WM_LBUTTONUP
        Vali = True
        Show
        WindowState = 0

    Case WM_LBUTTONDBLCLK
    Case WM_RBUTTONDOWN
    Case WM_RBUTTONUP
    Case WM_RBUTTONDBLCLK
End Select
End Sub

Private Sub MDIForm_Resize()
If WindowState = 1 Then
    If Vali = False Then
        MinimizeToTray
    End If
    Vali = False
End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'Remove tray icon
Shell_NotifyIcon NIM_DELETE, NID

'Disconnect the database
pDB.CloseDatabase

'== Position ==
InsertIntoRegistry "Server\Configuration", "Top", Top
InsertIntoRegistry "Server\Configuration", "Left", Left
End Sub

Private Sub Winsock1_Close(Index As Integer)
KickUser GetAccountByIndex(Index)
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim i As Long
Dim j As Long
    j = LoadSocket()

With Winsock1(j)
    .LocalPort = frmConfig.txtPort.Text
    .Accept requestID
End With

'Add new user to panel without account and name
With frmPanel.lvUsers.ListItems
    .Add
    i = .Count
    If Winsock1(j).RemoteHostIP = "127.0.0.1" Then
        .Item(i).SubItems(INDEX_IP) = Winsock1(0).LocalIP
    Else
        .Item(i).SubItems(INDEX_IP) = Winsock1(j).RemoteHostIP
    End If
    .Item(i).SubItems(INDEX_WINSOCK_ID) = j
    .Item(i).SubItems(INDEX_MUTED) = "0"
    .Item(i).SubItems(INDEX_LOGIN_TIME) = Format$(Time, "hh:mm:ss")
    .Item(i).SubItems(INDEX_GM_FLAG) = "0"
    .Item(i).SubItems(INDEX_AFK_FLAG) = "0"
    .Item(i).SubItems(INDEX_IS_ROOT) = "0"
End With

UPDATE_STATUS_BAR
End Sub

Private Function GetFreeSocket() As Long
Dim i As Long

On Error GoTo HandleErrorFreeSocket
With Winsock1
    For i = .LBound + 1 To .UBound
        .Item (i)
    Next i

    GetFreeSocket = .UBound + 1
End With

Exit Function
HandleErrorFreeSocket:
    GetFreeSocket = i
End Function

Private Function LoadSocket() As Long
Dim i As Long
    i = GetFreeSocket()

Load Winsock1(i)
LoadSocket = i
End Function

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim k               As Long
Dim i               As Long
Dim p_Message       As String   'All data send
Dim p_PreArray()    As String   'All data send split by chr(24) & chr(25) to not allow stack of wrong information
Dim p_MainArray()   As String   'Whole message string is saved here and split up by Chr(3) & Chr(4) sign
Dim p_Command       As String   'First part of main array ( always the command )
Dim IsMuted         As Boolean  'Mute explains itself
Dim properAccount   As String   'Used to save account to avoid multiply loading in loops

'Get data from socket
Winsock1(Index).GetData p_Message
DoEvents

'Do first array to avoid incorrect package reading
'Chr(24) and Chr(25) are the delimeters for each packet
p_PreArray = Split(p_Message, Chr(24) & Chr(25))

'No more use for p_Message variable so we clear
p_Message = vbNullString

'Start looping through
For k = 0 To UBound(p_PreArray) - 1
    
'Print the message
WriteText p_PreArray(k) & " | Index: " & Index

'We decode (split) the message into an array
p_MainArray = Split(p_PreArray(k), pSplit)

'Assign the variable to the array
If UBound(p_MainArray) > -1 Then
    p_Command = p_MainArray(0)
End If

Select Case p_Command
    'Select action and execute command
    Case "!friend"
        Select Case p_MainArray(1)
            'Update Friend list
            Case "-get"
                UPDATE_FRIEND GetAccountByIndex(Index), Index

            'Add friend to list
            Case "-add"
                frmFriendIgnoreList.AddFriend GetAccountByIndex(Index), GetProperAccountName(p_MainArray(2)), Index

            'Remove friend from list
            Case "-remove"
                frmFriendIgnoreList.RemoveFriend GetAccountByIndex(Index), GetProperAccountName(p_MainArray(2)), Index

        End Select

    Case "!ignore"
        Select Case p_MainArray(1)
            'Update Ignore list
            Case "-get"
                UPDATE_IGNORE GetAccountByIndex(Index), Index

            'Add ignore to list
            Case "-add"
                frmFriendIgnoreList.AddIgnore GetAccountByIndex(Index), GetProperAccountName(p_MainArray(2)), Index

            'Remove ignore from list
            Case "-remove"
                frmFriendIgnoreList.RemoveIgnore GetAccountByIndex(Index), GetProperAccountName(p_MainArray(2)), Index

        End Select

    Case "!connected"
        UPDATE_ONLINE

        properAccount = GetAccountByIndex(Index)

        With frmPanel.lvUsers.ListItems
            For i = 1 To .Count
                If Not .Item(i) = properAccount Then
                    SendSingle "!pmessage" & pSplit & "online" & pSplit & properAccount & pSplit, .Item(i).SubItems(INDEX_WINSOCK_ID)
                End If
            Next i
        End With

        'Send Server information
        SendSingle vbCrLf & _
                   " Welcome to Peach Servers." & vbCrLf & _
                   " Server: Peach r" & pRev & "/" & GetOS & vbCrLf & _
                   " Online User: " & Winsock1.Count - 1 & vbCrLf & _
                   " Server Uptime: " & TimeSerial(0, 0, DateDiff("s", frmConfig.START_TIME, Time)), Index

        frmOfflineMessages.SendAllOfflineMessages properAccount, Index

        properAccount = vbNullString

    Case "!login"
        With frmAccountPanel.lvAccounts.ListItems
            If .Count = 0 Then
                SendSingle "!login" & pSplit & "Account" & pSplit, Index
            Else
                'Authentication; password, is banned?, account etc ..
                For i = 1 To .Count
                    If LCase$(.Item(i).SubItems(INDEX_NAME)) = LCase$(p_MainArray(1)) Then
                        'Ban Check
                        If .Item(i).SubItems(INDEX_BANNED) = "1" Then
                            SendSingle "!login" & pSplit & "Banned" & pSplit, Index
                            Exit Sub
                        End If

                        'Password Check
                        If Not GetMD5(.Item(i).SubItems(INDEX_PASSWORD)) = p_MainArray(2) Then
                            SendSingle "!login" & pSplit & "Password" & pSplit, Index
                            Exit Sub
                        End If
                        Exit For
                    Else
                        If i = .Count Then
                            SendSingle "!login" & pSplit & "Account" & pSplit, Index
                            Exit Sub
                        End If
                    End If
                Next i

                With frmPanel.lvUsers.ListItems
                    'If the account is already beeing used kick first instance
                    For i = 1 To .Count
                        If LCase$(.Item(i)) = LCase$(p_MainArray(1)) Then
                            KickUser .Item(i)
                            Exit For
                        End If
                    Next i

                    .Item(.Count).Text = GetProperAccountName(p_MainArray(1))

                    UPDATE_STATUS_BAR
                    SendSingle "!accepted" & pSplit & .Item(.Count).Text & pSplit, Index
                End With
            End If
        End With

    'We get ip request and send ip back
    Case "!iprequest"
        With frmPanel.lvUsers.ListItems
            For i = 1 To .Count
                If .Item(i) = p_MainArray(1) Then
                    SendSingle "!iprequest" & pSplit & .Item(i).SubItems(INDEX_IP) & pSplit, Index
                    Exit For
                End If
            Next i
        End With

    Case "!channel_password"
        With frmChannel.lvChannels.ListItems
            For i = 1 To .Count
                If LCase$(.Item(i)) = LCase$(p_MainArray(2)) And .Item(i).SubItems(CHANNEL_PASSWORD) = p_MainArray(1) Then
                    frmChannel.JoinChannelReal .Item(i), GetAccountByIndex(Index)
                    Exit For
                Else
                    If i = .Count Then SendSingle "!pmessage" & pSplit & "channel_wrong_password" & pSplit & p_MainArray(2) & pSplit, Index
                End If
            Next i
        End With

    Case "!message"
        Dim p_CHAT_ARRAY() As String
        Dim p_TEXT_FIRST   As String
        Dim p_TEXT_SECOND  As String
        Dim IsCommand      As Boolean
        Dim IsSlash        As Boolean

        'Split the conversation text by spaces
        p_CHAT_ARRAY = Split(p_MainArray(1), " ")

        'Check first position of the text for a point indicating command or emote
        If UBound(p_CHAT_ARRAY) > -1 Then
            Select Case Left$(p_CHAT_ARRAY(0), 1)
                Case Chr(46)
                    If GetLevel(GetAccountByIndex(Index)) > 0 Then IsCommand = True
    
                Case Chr(47)
                    IsSlash = True
    
            End Select
        End If

        'Capture first part of the text
        If UBound(p_CHAT_ARRAY) > 0 Then
            p_TEXT_FIRST = p_CHAT_ARRAY(1)
        End If

        'Capture second part
        If UBound(p_CHAT_ARRAY) > 1 Then
            p_TEXT_SECOND = p_CHAT_ARRAY(2)
        End If

        'Check if user is muted
        With frmPanel.lvUsers.ListItems
            properAccount = GetAccountByIndex(Index)

            For i = 1 To .Count
                If .Item(i) = properAccount Then
                    If .Item(i).SubItems(INDEX_MUTED) = "1" Then
                        IsMuted = True
                        Exit For
                    End If
                End If
            Next i

            properAccount = vbNullString
        End With

        'If a command is used check out which
        If IsCommand Then
            Dim m       As Long
            Dim n       As Long
            Dim Reason  As String

            'Save the reason
            For i = 2 To UBound(p_CHAT_ARRAY)
                Reason = Reason & p_CHAT_ARRAY(i) & " "
            Next i

            Select Case LCase$(p_CHAT_ARRAY(0))
                Case ".show"
                    If IsPartOf(p_TEXT_FIRST, "accounts") Then
                        SendSingle "!split_text" & pSplit & GetAccountList, Index

                    ElseIf IsPartOf(p_TEXT_FIRST, "onliners") Then
                        SendSingle "!split_text" & pSplit & GetOnlineList, Index

                    Else
                        SendSingle "!pmessage" & pSplit & "incorrect_syntax" & pSplit & ".show [accounts, onliners]", Index

                    End If

                Case ".userinfo", ".uinfo"
                    GetUserInfo GetProperAccountName(p_TEXT_FIRST), p_CHAT_ARRAY(0), Index

                Case ".accountinfo", ".accinfo", ".ainfo"
                    GetAccountInfo GetProperAccountName(p_TEXT_FIRST), p_CHAT_ARRAY(0), Index

                Case ".kick"
                    With frmPanel.lvUsers.ListItems
                        properAccount = GetProperAccountName(p_TEXT_FIRST)

                        For i = 1 To .Count
                            If .Item(i) = properAccount Then
                                KickUser .Item(i)
                                Exit For
                            Else
                                If i = .Count Then
                                    If LenB(p_TEXT_FIRST) = 0 Then
                                        SendSingle "!pmessage" & pSplit & "incorrect_syntax" & pSplit & ".kick [User]" & pSplit, Index
                                    Else
                                        SendSingle "!pmessage" & pSplit & "user_not_found" & pSplit & p_TEXT_FIRST & pSplit, Index
                                    End If
                                End If
                            End If
                        Next i

                        properAccount = vbNullString
                    End With

                Case ".ban"
                    Ban GetProperAccountName(p_TEXT_FIRST), 1, Trim$(Reason), Index

                Case ".unban"
                    Ban GetProperAccountName(p_TEXT_FIRST), 0, Trim$(Reason), Index

                Case ".mute"
                    MuteUser GetProperAccountName(p_TEXT_FIRST), 1, Trim$(Reason), Index

                Case ".unmute"
                    MuteUser GetProperAccountName(p_TEXT_FIRST), 0, Trim$(Reason), Index

                Case ".announce", ".ann", ".notify"
                    Dim p_ANN_MSG As String

                    'Capture announce message
                    If UBound(p_CHAT_ARRAY) > 0 Then
                        p_ANN_MSG = Mid$(p_MainArray(1), Len(p_CHAT_ARRAY(0)) + 2, Len(p_MainArray(1)))
                    End If

                    If LenB(p_ANN_MSG) = 0 Then
                        SendSingle "!pmessage" & pSplit & "incorrect_syntax" & pSplit & p_CHAT_ARRAY(0) & " [Text]" & pSplit, Index
                    Else
                        Select Case LCase$(p_CHAT_ARRAY(0))
                            Case ".announce", ".ann"
                                Reason = GetAccountByIndex(Index)
                                SendMessage "!pmessage" & pSplit & "announce" & pSplit & GetGMFlag(Reason) & GetAFKFlag(Reason) & pSplit & Reason & pSplit & p_ANN_MSG & pSplit

                            Case ".notify"
                                SendMessage "!msgbox" & pSplit & GetAccountByIndex(Index) & " - " & p_ANN_MSG & pSplit

                        End Select
                    End If

                Case ".help", ".command", ".commands"
                    SendSingle GetCommands, Index

                Case ".change", ".modify", ".mod"
                    Select Case LCase$(p_TEXT_FIRST)
                        Case "name"
                            properAccount = GetProperAccountName(p_TEXT_SECOND)

                            'Check if there are enough parameters
                            If UBound(p_CHAT_ARRAY) > 2 Then
                                'Check if the account you want to modify exist
                                With frmAccountPanel.lvAccounts.ListItems
                                    For i = 1 To .Count
                                        If .Item(i).SubItems(INDEX_NAME) = properAccount Then
                                            'Check if change name is already given out
                                            For n = 1 To .Count
                                                If LCase$(.Item(n).SubItems(INDEX_NAME)) = LCase$(p_CHAT_ARRAY(3)) Then
                                                    SendSingle "!pmessage" & pSplit & "user_already_used" & pSplit & p_CHAT_ARRAY(3) & pSplit, Index
                                                    Exit For
                                                Else
                                                    If n = .Count Then
                                                        'Modify account in database and panel
                                                        frmAccountPanel.ModifyAccount p_CHAT_ARRAY(3), .Item(i).SubItems(INDEX_PASSWORD), .Item(i).SubItems(INDEX_BANNED), .Item(i).SubItems(INDEX_LEVEL), .Item(i), i, .Item(i).SubItems(INDEX_GENDER), .Item(i).SubItems(INDEX_EMAIL)

                                                        'Send feedback to person who changed name
                                                        SendSingle "!pmessage" & pSplit & "successfull_rename" & pSplit & properAccount & pSplit & p_CHAT_ARRAY(3) & pSplit, Index

                                                        'Check if player is online and directly rename it, also tell user that he got renamed
                                                        With frmPanel.lvUsers.ListItems
                                                            For m = 1 To .Count
                                                                If .Item(m) = properAccount Then
                                                                    .Item(m) = p_CHAT_ARRAY(3)

                                                                    SendSingle "!pmessage" & pSplit & "renamed_you_to" & pSplit & GetAccountByIndex(Index) & pSplit & p_CHAT_ARRAY(3) & pSplit, .Item(m).SubItems(INDEX_WINSOCK_ID)
                                                                    SendSingle "!namechange" & pSplit & p_CHAT_ARRAY(3) & pSplit, .Item(m).SubItems(INDEX_WINSOCK_ID)
                                                                    Exit For
                                                                End If
                                                            Next m
                                                        End With

                                                        'Rename it in friend list
                                                        With frmFriendIgnoreList.ListView1.ListItems
                                                            'First column
                                                            For m = 1 To .Count
                                                                If .Item(m).SubItems(1) = properAccount Then
                                                                    .Item(m).SubItems(1) = p_CHAT_ARRAY(3)
                                                                End If
                                                            Next m

                                                            'Second column
                                                            For m = 1 To .Count
                                                                If .Item(m).SubItems(2) = properAccount Then
                                                                    .Item(m).SubItems(2) = p_CHAT_ARRAY(3)
                                                                End If
                                                            Next m
                                                        End With

                                                        'Rename it in ignore list
                                                        With frmFriendIgnoreList.ListView2.ListItems
                                                            'First column
                                                            For m = 1 To .Count
                                                                If .Item(m).SubItems(1) = properAccount Then
                                                                    .Item(m).SubItems(1) = p_CHAT_ARRAY(3)
                                                                End If
                                                            Next m

                                                            'Second column
                                                            For m = 1 To .Count
                                                                If .Item(m).SubItems(2) = properAccount Then
                                                                    .Item(m).SubItems(2) = p_CHAT_ARRAY(3)
                                                                End If
                                                            Next m
                                                        End With

                                                        'Rename it in offline message list
                                                        With frmOfflineMessages.lvOfflineMessages.ListItems
                                                            'First column
                                                            For m = 1 To .Count
                                                                If .Item(m) = properAccount Then
                                                                    .Item(m) = p_CHAT_ARRAY(3)
                                                                End If
                                                            Next m

                                                            For m = 1 To .Count
                                                                If .Item(m).SubItems(INDEX_TO) = properAccount Then
                                                                    .Item(m).SubItems(INDEX_TO) = p_CHAT_ARRAY(3)
                                                                End If
                                                            Next m
                                                        End With

                                                        pDB.ExecuteCommand "UPDATE " & DATABASE_TABLE_FRIENDS & " SET Name = '" & p_CHAT_ARRAY(3) & "' WHERE Name = '" & properAccount & "'"
                                                        pDB.ExecuteCommand "UPDATE " & DATABASE_TABLE_FRIENDS & " SET Friend = '" & p_CHAT_ARRAY(3) & "' WHERE Friend = '" & properAccount & "'"

                                                        UPDATE_ONLINE
                                                    End If
                                                End If
                                            Next n
                                            Exit For
                                        Else
                                            If i = .Count Then SendSingle "!pmessage" & pSplit & "user_not_found" & pSplit & properAccount & pSplit, Index
                                        End If
                                    Next i
                                End With

                                properAccount = vbNullString
                            Else
                                SendSingle "!pmessage" & pSplit & "incorrect_syntax" & pSplit & p_CHAT_ARRAY(0) & Space(1) & p_TEXT_FIRST & " [Oldname] [Newname]" & pSplit, Index
                            End If

                        Case "level"
                            If UBound(p_CHAT_ARRAY) > 2 Then
                                'Level must be numeric and in range of 0-2
                                If (Not IsNumeric(p_CHAT_ARRAY(3))) Or (p_CHAT_ARRAY(3) > 2 Or p_CHAT_ARRAY(3) < 0) Then
                                    SendSingle "!pmessage" & pSplit & "level_incorrect_value" & pSplit, Index
                                    Exit Sub
                                End If

                                With frmAccountPanel.lvAccounts.ListItems
                                    properAccount = GetProperAccountName(p_TEXT_SECOND)

                                    For i = 1 To .Count
                                        If .Item(i).SubItems(INDEX_NAME) = properAccount Then
                                            'Modify level for account in database
                                            frmAccountPanel.ModifyAccount .Item(i).SubItems(INDEX_NAME), .Item(i).SubItems(INDEX_PASSWORD), .Item(i).SubItems(INDEX_BANNED), p_CHAT_ARRAY(3), .Item(i), i, .Item(i).SubItems(INDEX_GENDER), .Item(i).SubItems(INDEX_EMAIL)

                                            'Feedback to the person who modified the level
                                            SendSingle "!pmessage" & pSplit & "successfull_level" & pSplit & .Item(i).SubItems(INDEX_NAME) & pSplit & p_CHAT_ARRAY(3) & pSplit, Index

                                            With frmPanel.lvUsers.ListItems
                                                For n = 1 To .Count
                                                    If .Item(n) = properAccount Then
                                                        SendSingle "!pmessage" & pSplit & "changed_your_level" & pSplit & GetAccountByIndex(Index) & pSplit & p_CHAT_ARRAY(3) & pSplit, .Item(n).SubItems(INDEX_WINSOCK_ID)
                                                        Exit For
                                                    End If
                                                Next n
                                            End With
                                            Exit For
                                        Else
                                            If i = .Count Then SendSingle "!pmessage" & pSplit & "user_not_found" & pSplit & p_TEXT_SECOND & pSplit, Index
                                        End If
                                    Next i

                                    properAccount = vbNullString
                                End With
                            Else
                                SendSingle "!pmessage" & pSplit & "incorrect_syntax" & pSplit & p_CHAT_ARRAY(0) & Space(1) & p_TEXT_FIRST & " [Name] [Level]" & pSplit, Index
                            End If

                        Case "gender"
                            Dim GenderInNumeric As String

                            If UBound(p_CHAT_ARRAY) > 2 Then
                                p_CHAT_ARRAY(3) = LCase$(p_CHAT_ARRAY(3))

                                If p_CHAT_ARRAY(3) = "male" Then
                                    GenderInNumeric = "0"

                                ElseIf p_CHAT_ARRAY(3) = "female" Then
                                    GenderInNumeric = "1"

                                End If

                                If LenB(GenderInNumeric) = 0 Then
                                    SendSingle "!pmessage" & pSplit & "gender_incorrect_value" & pSplit, Index
                                Else
                                    With frmAccountPanel.lvAccounts.ListItems
                                        properAccount = GetProperAccountName(p_TEXT_SECOND)

                                        For i = 1 To .Count
                                            If .Item(i).SubItems(INDEX_NAME) = properAccount Then
                                                'Modify gender in database and panel
                                                frmAccountPanel.ModifyAccount .Item(i).SubItems(INDEX_NAME), .Item(i).SubItems(INDEX_PASSWORD), .Item(i).SubItems(INDEX_BANNED), .Item(i).SubItems(INDEX_LEVEL), .Item(i), i, GenderInNumeric, .Item(i).SubItems(INDEX_EMAIL)

                                                'Feedback to the person who modified the gender
                                                SendSingle "!pmessage" & pSplit & "successfull_gender" & pSplit & properAccount & pSplit & GenderInNumeric & pSplit, Index

                                                'Notify user that gender got changed
                                                With frmPanel.lvUsers.ListItems
                                                    For n = 1 To .Count
                                                        If .Item(n) = properAccount Then
                                                            SendSingle "!pmessage" & pSplit & "changed_your_gender" & pSplit & GetAccountByIndex(Index) & pSplit & GenderInNumeric & pSplit, .Item(n).SubItems(INDEX_WINSOCK_ID)
                                                            Exit For
                                                        End If
                                                    Next n
                                                End With
                                                Exit For
                                            Else
                                                If i = .Count Then SendSingle "!pmessage" & pSplit & "user_not_found" & pSplit & p_TEXT_SECOND & pSplit, Index
                                            End If
                                        Next i

                                        properAccount = vbNullString
                                    End With
                                End If
                            Else
                                SendSingle "!pmessage" & pSplit & "incorrect_syntax" & pSplit & p_CHAT_ARRAY(0) & Space(1) & p_TEXT_FIRST & " [Name] [Gender]" & pSplit, Index
                            End If

                        Case "password"
                            If UBound(p_CHAT_ARRAY) > 2 Then
                                With frmAccountPanel.lvAccounts.ListItems
                                    properAccount = GetProperAccountName(p_TEXT_SECOND)

                                    For i = 1 To .Count
                                        If .Item(i).SubItems(INDEX_NAME) = properAccount Then
                                            'Modify password in database and panel
                                            frmAccountPanel.ModifyAccount .Item(i).SubItems(INDEX_NAME), p_CHAT_ARRAY(3), .Item(i).SubItems(INDEX_BANNED), .Item(i).SubItems(INDEX_LEVEL), .Item(i), i, .Item(i).SubItems(INDEX_GENDER), .Item(i).SubItems(INDEX_EMAIL)

                                            'Feedback to the person who modified the password
                                            SendSingle "!pmessage" & pSplit & "successfull_password" & pSplit & properAccount & pSplit & p_CHAT_ARRAY(3) & pSplit, Index

                                            'Notify user that password got changed
                                            With frmPanel.lvUsers.ListItems
                                                For n = 1 To .Count
                                                    If .Item(n) = properAccount Then
                                                        SendSingle "!pmessage" & pSplit & "changed_your_password" & pSplit & GetAccountByIndex(Index) & pSplit & p_CHAT_ARRAY(3) & pSplit, .Item(n).SubItems(INDEX_WINSOCK_ID)
                                                        Exit For
                                                    End If
                                                Next n
                                            End With
                                            Exit For
                                        Else
                                            If i = .Count Then SendSingle "!pmessage" & pSplit & "user_not_found" & pSplit & p_TEXT_SECOND & pSplit, Index
                                        End If
                                    Next i

                                    properAccount = vbNullString
                                End With
                            Else
                                SendSingle "!pmessage" & pSplit & "incorrect_syntax" & pSplit & p_CHAT_ARRAY(0) & Space(1) & p_TEXT_FIRST & " [Name] [Password]" & pSplit, Index
                            End If

                        Case "email"
                            If UBound(p_CHAT_ARRAY) > 2 Then
                                With frmAccountPanel.lvAccounts.ListItems
                                    properAccount = GetProperAccountName(p_TEXT_SECOND)

                                    For i = 1 To .Count
                                        If .Item(i).SubItems(INDEX_NAME) = properAccount Then
                                            'Modify email in database and panel
                                            frmAccountPanel.ModifyAccount .Item(i).SubItems(INDEX_NAME), .Item(i).SubItems(INDEX_PASSWORD), .Item(i).SubItems(INDEX_BANNED), .Item(i).SubItems(INDEX_LEVEL), .Item(i), i, .Item(i).SubItems(INDEX_GENDER), p_CHAT_ARRAY(3)

                                            'Feedback to the person who modified the email
                                            SendSingle "!pmessage" & pSplit & "successfull_email" & pSplit & properAccount & pSplit & p_CHAT_ARRAY(3) & pSplit, Index

                                            'Notify user that password got changed
                                            With frmPanel.lvUsers.ListItems
                                                For n = 1 To .Count
                                                    If .Item(n) = properAccount Then
                                                        SendSingle "!pmessage" & pSplit & "changed_your_email" & pSplit & GetAccountByIndex(Index) & pSplit & p_CHAT_ARRAY(3) & pSplit, .Item(n).SubItems(INDEX_WINSOCK_ID)
                                                        Exit For
                                                    End If
                                                Next n
                                            End With
                                            Exit For
                                        Else
                                            If i = .Count Then SendSingle "!pmessage" & pSplit & "user_not_found" & pSplit & p_TEXT_SECOND & pSplit, Index
                                        End If
                                    Next i

                                    properAccount = vbNullString
                                End With
                            Else
                                SendSingle "!pmessage" & pSplit & "incorrect_syntax" & pSplit & p_CHAT_ARRAY(0) & Space(1) & p_TEXT_FIRST & " [Name] [Email]" & pSplit, Index
                            End If

                        Case Else
                            SendSingle _
                            "!pmessage" & pSplit & "incorrect_syntax" & pSplit & _
                                vbCrLf & Space(2) & _
                            p_CHAT_ARRAY(0) & " name [Oldname] [Newname]" & _
                                vbCrLf & Space(2) & _
                            p_CHAT_ARRAY(0) & " level [Name] [Level]" & _
                                vbCrLf & Space(2) & _
                            p_CHAT_ARRAY(0) & " gender [Name] [Gender]" & _
                                vbCrLf & Space(2) & _
                            p_CHAT_ARRAY(0) & " password [Name] [Password]" & _
                                vbCrLf & Space(2) & _
                            p_CHAT_ARRAY(0) & " email [Name] [Email]" & pSplit, Index

                    End Select

                Case ".reload"
                    Dim loadTime As Long

                    With Database
                        Select Case LCase$(p_TEXT_FIRST)
                            Case LCase$(DATABASE_TABLE_ACCOUNTS), LCase$(DATABASE_TABLE_FRIENDS), LCase$(DATABASE_TABLE_IGNORES)
                                SendSingle "!pmessage" & pSplit & "table_cant_reload" & pSplit, Index

                            Case LCase$(DATABASE_TABLE_COMMANDS)
                                loadTime = timeGetTime
                                Erase Commands
                                LoadCommands
                                SendMessage "!pmessage" & pSplit & "table_reload" & pSplit & GetAccountByIndex(Index) & pSplit & DATABASE_TABLE_COMMANDS & pSplit & timeGetTime - loadTime & pSplit

                            Case LCase$(DATABASE_TABLE_DECLINED_NAMES)
                                loadTime = timeGetTime
                                Erase DeclinedNames
                                LoadDeclinedNames
                                SendMessage "!pmessage" & pSplit & "table_reload" & pSplit & GetAccountByIndex(Index) & pSplit & DATABASE_TABLE_DECLINED_NAMES & pSplit & timeGetTime - loadTime & pSplit

                            Case LCase$(DATABASE_TABLE_EMOTES)
                                loadTime = timeGetTime
                                Erase Emotes
                                LoadEmotes
                                SendMessage "!pmessage" & pSplit & "table_reload" & pSplit & GetAccountByIndex(Index) & pSplit & DATABASE_TABLE_EMOTES & pSplit & timeGetTime - loadTime & pSplit

                            Case LCase$("config"), LCase$("c")
                                loadTime = timeGetTime
                                LoadConfigValue
                                SendMessage "!pmessage" & pSplit & "config_reload" & pSplit & GetAccountByIndex(Index) & pSplit & timeGetTime - loadTime & pSplit

                            Case Else
                                If LenB(p_TEXT_FIRST) = 0 Then
                                    SendSingle "!pmessage" & pSplit & "incorrect_syntax" & pSplit & ".reload Table" & pSplit, Index
                                Else
                                    SendSingle "!pmessage" & pSplit & "table_not_exist" & pSplit, Index
                                End If

                        End Select
                    End With

                Case ".clear"
                    If LenB(p_TEXT_FIRST) = 0 Then
                        SendSingle "!pmessage" & pSplit & "incorrect_syntax" & pSplit & ".clear [User]" & pSplit, Index
                    Else
                        If LCase$(p_TEXT_FIRST) = "this" Or LCase$(p_TEXT_FIRST) = "me" Then
                            SendSingle "!clear" & pSplit, Index
                            Exit Sub

                        ElseIf LCase$(p_TEXT_FIRST) = "all" Then
                            SendMessage "!clear" & pSplit
                            Exit Sub

                        End If

                        With frmPanel.lvUsers.ListItems
                            properAccount = GetProperAccountName(p_TEXT_FIRST)

                            For i = 1 To .Count
                                If .Item(i) = properAccount Then
                                    SendSingle "!clear" & pSplit, .Item(i).SubItems(INDEX_WINSOCK_ID)
                                    Exit For
                                Else
                                    If i = .Count Then SendSingle "!pmessage" & pSplit & "user_not_found" & pSplit & p_TEXT_FIRST & pSplit, Index
                                End If
                            Next i

                            properAccount = vbNullString
                        End With
                    End If

                Case ".delete", ".del"
                    With frmAccountPanel.lvAccounts.ListItems
                        properAccount = GetProperAccountName(p_TEXT_FIRST)

                        For i = 1 To .Count
                            If .Item(i).SubItems(INDEX_NAME) = properAccount Then
                                With frmPanel.lvUsers.ListItems
                                    For m = 1 To .Count
                                        If .Item(m) = properAccount Then
                                            KickUser properAccount
                                            Exit For
                                        End If
                                    Next m
                                End With

                                'If account exist delete it
                                pDB.ExecuteCommand "DELETE FROM " & DATABASE_TABLE_ACCOUNTS & " WHERE Name1 = '" & properAccount & "'"
                                frmFriendIgnoreList.RemoveAllFriendsFromUser properAccount
                                frmFriendIgnoreList.RemoveAllIgnoresFromUser properAccount
                                frmOfflineMessages.RemoveAllOfflineMessagesFromAndToUser properAccount

                                SendSingle "!pmessage" & pSplit & "deleted_account" & pSplit & properAccount & pSplit & .Item(i) & pSplit, Index
                                .Remove i
                                Exit For
                            Else
                                If i = .Count Then
                                    If LenB(properAccount) = 0 Then
                                        SendSingle "!pmessage" & pSplit & "incorrect_syntax" & pSplit & ".delete [Account]" & pSplit, Index
                                    Else
                                        SendSingle "!pmessage" & pSplit & "user_not_found" & pSplit & p_TEXT_FIRST & pSplit, Index
                                    End If
                                End If
                            End If
                        Next i

                        properAccount = vbNullString
                    End With

                Case ".gm"
                    With frmPanel.lvUsers.ListItems
                        Select Case LCase$(p_TEXT_FIRST)
                            Case "on"
                                For i = 1 To .Count
                                    If .Item(i).SubItems(INDEX_WINSOCK_ID) = Index Then
                                        .Item(i).SubItems(INDEX_GM_FLAG) = "1"
                                        SendSingle "!pmessage" & pSplit & "gm_flag_enable" & pSplit, Index
                                        Exit For
                                    End If
                                Next i

                            Case "off"
                                For i = 1 To .Count
                                    If .Item(i).SubItems(INDEX_WINSOCK_ID) = Index Then
                                        .Item(i).SubItems(INDEX_GM_FLAG) = "0"
                                        SendSingle "!pmessage" & pSplit & "gm_flag_disable" & pSplit, Index
                                        Exit For
                                    End If
                                Next i

                            Case Else
                                SendSingle "!pmessage" & pSplit & "incorrect_syntax" & pSplit & ".gm [on / off]" & pSplit, Index

                        End Select
                    End With

                Case Else
                    SendSingle "!pmessage" & pSplit & "unknown_command" & pSplit, Index

            End Select
            Exit Sub 'Always exit after any command used or else it will be written for everyone
        End If

        'Check if user is muted for emotes and normal text
        If IsMuted Then
            SendSingle "!pmessage" & "muted" & pSplit, Index
            Exit Sub
        End If

        If Options.REPEAT_CHECK = 1 Then
            If IsRepeating(GetAccountByIndex(Index), p_MainArray(1)) Then
                SendSingle "!pmessage" & pSplit & "flood_protection" & pSplit, Index
                Exit Sub
            End If
        End If

        If IsSlash Then
            Dim f       As Long
            Dim IsUser  As Boolean

            'IsUser variable determines if the first part of the text is an user and exist or not
            With frmPanel.lvUsers.ListItems
                properAccount = GetProperAccountName(p_TEXT_FIRST)

                For i = 1 To .Count
                    If .Item(i) = properAccount Then
                        IsUser = True
                        Exit For
                    End If
                Next i

                properAccount = vbNullString
            End With

            Select Case LCase(p_CHAT_ARRAY(0))
                'Roll function
                Case "/roll"
                    Dim Roll    As Long
                    Dim MinRoll As Long
                    Dim MaxRoll As Long

                    If UBound(p_CHAT_ARRAY) > 1 Then
                        If IsNumeric(p_CHAT_ARRAY(1)) Then
                            If p_CHAT_ARRAY(1) > MAX_INT_VALUE Or p_CHAT_ARRAY(1) < MIN_INT_VALUE Then
                                MinRoll = 1
                            Else
                                MinRoll = p_CHAT_ARRAY(1)
                            End If
                        Else
                            MinRoll = 1
                        End If

                        If IsNumeric(p_CHAT_ARRAY(2)) Then
                            If p_CHAT_ARRAY(2) > MAX_INT_VALUE Or p_CHAT_ARRAY(2) < MIN_INT_VALUE Then
                                MaxRoll = 100
                            Else
                                MaxRoll = p_CHAT_ARRAY(2)
                            End If
                        Else
                            MaxRoll = 100
                        End If

                    ElseIf UBound(p_CHAT_ARRAY) > 0 Then
                        MinRoll = 1

                        If IsNumeric(p_CHAT_ARRAY(1)) Then
                            If p_CHAT_ARRAY(1) > MAX_INT_VALUE Or p_CHAT_ARRAY(1) < MIN_INT_VALUE Then
                                MaxRoll = 100
                            Else
                                MaxRoll = p_CHAT_ARRAY(1)
                            End If
                        Else
                            MaxRoll = 100
                        End If

                    Else
                        MinRoll = 1
                        MaxRoll = 100

                    End If

                    If MinRoll > MaxRoll Then MaxRoll = MinRoll
                    If MinRoll = 0 Then MinRoll = 1
                    If MaxRoll = 0 Then MaxRoll = 100

                    Roll = GetRandomNumber(MinRoll, MaxRoll)

                    properAccount = GetAccountByIndex(Index)

                    SendProtectedMessage properAccount, "!pmessage" & pSplit & "roll" & pSplit & properAccount & pSplit & Roll & pSplit & MinRoll & pSplit & MaxRoll & pSplit

                    properAccount = vbNullString

                'Whisper X to Z from Y
                Case "/w", "/whisper"
                    Dim Message As String

                    If UBound(p_CHAT_ARRAY) > 1 Then
                        For i = 2 To UBound(p_CHAT_ARRAY)
                            Message = Message & p_CHAT_ARRAY(i) & " "
                        Next i

                        Whisper GetAccountByIndex(Index), GetProperAccountName(p_TEXT_FIRST), Trim$(Message), Index
                    Else
                        SendSingle "!pmessage" & pSplit & "incorrect_syntax" & pSplit & "/whisper [Name] [Text]" & pSplit, Index
                    End If

                Case "/afk"
                    With frmPanel.lvUsers.ListItems
                        For i = 1 To .Count
                            If .Item(i).SubItems(INDEX_WINSOCK_ID) = Index Then
                                If .Item(i).SubItems(INDEX_AFK_FLAG) = "1" Then
                                    .Item(i).SubItems(INDEX_AFK_FLAG) = "0"
                                    SendSingle "!pmessage" & pSplit & "not_afk" & pSplit, Index
                                Else
                                    .Item(i).SubItems(INDEX_AFK_FLAG) = "1"
                                    SendSingle "!pmessage" & pSplit & "afk" & pSplit, Index
                                End If
                                Exit For
                            End If
                        Next i
                    End With

                Case "/online"
                    SendSingle "!pmessage" & pSplit & "online_time" & pSplit & Trim$(GetOnlineTime(GetAccountByIndex(Index))) & pSplit, Index

                Case "/logout"
                    KickUser GetAccountByIndex(Index)

                Case "/sudo"
                    If LenB(p_TEXT_FIRST) = 0 Then
                        SendSingle "!pmessage" & pSplit & "incorrect_syntax" & pSplit & "/sudo [Password]" & pSplit, Index
                    Else
                        If p_TEXT_FIRST = Options.ROOT_PASSWORD Then
                            With frmPanel.lvUsers.ListItems
                                For i = 1 To .Count
                                    If .Item(i).SubItems(INDEX_WINSOCK_ID) = Index Then
                                        .Item(i).SubItems(INDEX_IS_ROOT) = 1
                                        SendSingle "!pmessage" & pSplit & "successfull_sudo" & pSplit, Index
                                        Exit For
                                    End If
                                Next i
                            End With
                        Else
                            SendSingle "!pmessage" & pSplit & "incorrect_password" & pSplit, Index
                        End If
                    End If

                Case "/join"
                    If LenB(p_TEXT_FIRST) = 0 Then
                        SendSingle "!pmessage" & pSplit & "valid_channel" & pSplit, Index
                    Else
                        With frmChannel.lvUsers.ListItems
                            If .Count = 0 Then
                                frmChannel.JoinChannel p_TEXT_FIRST, GetAccountByIndex(Index), Index
                            Else
                                properAccount = GetAccountByIndex(Index)

                                For i = 1 To .Count
                                    If .Item(i) = properAccount And LCase$(.Item(i).SubItems(CHANNEL_USER_CHANNEL)) = LCase$(p_TEXT_FIRST) Then
                                        SendSingle "!pmessage" & pSplit & "already_in_channel" & pSplit & .Item(i).SubItems(CHANNEL_USER_CHANNEL) & pSplit, Index
                                        Exit For
                                    Else
                                        If i = .Count Then frmChannel.JoinChannel p_TEXT_FIRST, properAccount, Index
                                    End If
                                Next i

                                properAccount = vbNullString
                            End If
                        End With
                    End If

                Case "/leave"
                    If LenB(p_TEXT_FIRST) = 0 Then
                        SendSingle "!pmessage" & pSplit & "valid_channel" & pSplit, Index
                    Else
                        frmChannel.LeaveChannel p_TEXT_FIRST, GetAccountByIndex(Index)
                    End If

                Case "/announce"
                    If LenB(p_TEXT_FIRST) = 0 Then
                        SendSingle "!pmessage" & pSplit & "valid_channel" & pSplit, Index
                    Else
                        If frmChannel.lvChannels.ListItems.Count = 0 Then
                            SendSingle "!pmessage" & pSplit & "not_in_channel" & pSplit & p_TEXT_FIRST & pSplit, Index
                        Else
                            With frmChannel.lvUsers.ListItems
                                properAccount = GetAccountByIndex(Index)

                                For i = 1 To .Count
                                    If .Item(i) = properAccount And LCase$(.Item(i).SubItems(CHANNEL_USER_CHANNEL)) = LCase$(p_TEXT_FIRST) Then
                                        If .Item(i).SubItems(CHANNEL_USER_IS_OWNER) = "1" Then
                                            With frmChannel.lvChannels.ListItems
                                                For f = 1 To .Count
                                                    If LCase$(.Item(f)) = LCase$(p_TEXT_FIRST) Then
                                                        If .Item(f).SubItems(CHANNEL_JOIN_ANNOUNCE) = "1" Then
                                                            .Item(f).SubItems(CHANNEL_JOIN_ANNOUNCE) = "0"
                                                        Else
                                                            .Item(f).SubItems(CHANNEL_JOIN_ANNOUNCE) = "1"
                                                        End If

                                                        SendMessageToChannel .Item(f), properAccount, "!pmessage" & pSplit & "channel_announcements" & pSplit & .Item(f) & pSplit & properAccount & pSplit
                                                        Exit For
                                                    End If
                                                Next f
                                            End With
                                        Else
                                            SendSingle "!pmessage" & pSplit & "not_channel_leader" & pSplit, Index
                                        End If
                                        Exit For
                                    Else
                                        If i = .Count Then SendSingle "!pmessage" & pSplit & "not_in_channel" & pSplit & p_TEXT_FIRST & pSplit, Index
                                    End If
                                Next i

                                properAccount = vbNullString
                            End With
                        End If
                    End If

                Case "/setpassword"
                    If LenB(p_TEXT_FIRST) = 0 Then
                        SendSingle "!pmessage" & pSplit & "valid_channel" & pSplit, Index
                    Else
                        If LenB(p_TEXT_SECOND) = 0 Then
                            SendSingle "!pmessage" & pSplit & "incorrect_syntax" & pSplit & "/setpassword [Channel] [Password]" & pSplit, Index
                        Else
                            If frmChannel.lvChannels.ListItems.Count = 0 Then
                                SendSingle "!pmessage" & pSplit & "not_in_channel" & pSplit & p_TEXT_FIRST & pSplit, Index
                            Else
                                With frmChannel.lvUsers.ListItems
                                    properAccount = GetAccountByIndex(Index)

                                    For i = 1 To .Count
                                        If .Item(i) = properAccount And LCase$(.Item(i).SubItems(CHANNEL_USER_CHANNEL)) = LCase$(p_TEXT_FIRST) Then
                                            If .Item(i).SubItems(CHANNEL_USER_IS_OWNER) = "1" Then
                                                With frmChannel.lvChannels.ListItems
                                                    For f = 1 To .Count
                                                        If LCase$(.Item(f)) = LCase$(p_TEXT_FIRST) Then
                                                            .Item(f).SubItems(CHANNEL_PASSWORD) = p_TEXT_SECOND
                                                            SendSingle "!pmessage" & pSplit & "channel_password" & .Item(f) & pSplit & p_TEXT_SECOND & pSplit, Index
                                                            Exit For
                                                        End If
                                                    Next f
                                                End With
                                            Else
                                                SendSingle "!pmessage" & pSplit & "not_channel_leader" & pSplit, Index
                                            End If
                                            Exit For
                                        Else
                                            If i = .Count Then SendSingle "!pmessage" & pSplit & "not_in_channel" & pSplit & p_TEXT_FIRST & pSplit, Index
                                        End If
                                    Next i

                                    properAccount = vbNullString
                                End With
                            End If
                        End If
                    End If

                Case Else
                    'Emotes
                    properAccount = GetAccountByIndex(Index)

                    For i = LBound(Emotes) To UBound(Emotes)
                        If LCase$(Emotes(i).Command) = LCase$(p_CHAT_ARRAY(0)) Then
                            If properAccount = GetProperAccountName(p_TEXT_FIRST) Then
                                IsUser = False
                            End If

                            Dim Gender As String
                            Dim Temp   As String
                            Dim j      As Long

                            With frmAccountPanel.lvAccounts.ListItems
                                For j = 1 To .Count
                                    If .Item(j).SubItems(INDEX_NAME) = properAccount Then
                                        Select Case .Item(j).SubItems(INDEX_GENDER)
                                            Case "0": Gender = "his"
                                            Case "1": Gender = "her"
                                        End Select
                                        Exit For
                                    End If
                                Next j
                            End With

                            If IsUser Then
                                Temp = Replace(Emotes(i).TargetEmote, "%u", properAccount)
                                Temp = Replace(Temp, "%g", Gender)
                                Temp = Replace(Temp, "%t", GetProperAccountName(p_TEXT_FIRST))
                            Else
                                Temp = Replace(Emotes(i).SingleEmote, "%u", properAccount)
                                Temp = Replace(Temp, "%g", Gender)
                            End If

                            SendProtectedMessage properAccount, Temp
                            SetLastMessage properAccount, p_MainArray(1)
                            Exit Sub
                        End If
                    Next i

                    properAccount = vbNullString

                    With frmChannel.lvChannels.ListItems
                        If .Count = 0 Then
                            If LenB(p_TEXT_FIRST) = 0 Then
                                SendSingle "!pmessage" & pSplit & "valid_channel" & pSplit, Index
                            Else
                                SendSingle "!pmessage" & pSplit & "not_in_channel" & pSplit & Right$(p_CHAT_ARRAY(0), Len(p_CHAT_ARRAY(0)) - 1) & pSplit, Index
                            End If
                        Else
                            If LenB(p_TEXT_FIRST) = 0 Then
                                SendSingle "!pmessage" & pSplit & "valid_channel" & pSplit, Index
                            Else
                                properAccount = Right$(p_CHAT_ARRAY(0), Len(p_CHAT_ARRAY(0)) - 1)

                                For i = 1 To .Count
                                    If LCase$(.Item(i)) = LCase$(properAccount) Then
                                        With frmChannel.lvUsers.ListItems
                                            Temp = GetAccountByIndex(Index)

                                            For f = 1 To .Count
                                                If .Item(f) = Temp And LCase$(.Item(f).SubItems(CHANNEL_USER_CHANNEL)) = LCase$(properAccount) Then
                                                    SendMessageToChannel properAccount, Temp, "[" & .Item(f).SubItems(CHANNEL_USER_CHANNEL) & "]" & GetGMFlag(Temp) & GetAFKFlag(Temp) & "[" & Temp & "]: " & Right$(p_MainArray(1), Len(p_MainArray(1)) - Len(properAccount) - 1)
                                                    Exit For
                                                Else
                                                    If f = .Count Then SendSingle "!pmessage" & pSplit & "not_in_channel" & pSplit & properAccount & pSplit, Index
                                                End If
                                            Next f

                                            Temp = vbNullString
                                        End With
                                        Exit Sub
                                    End If
                                Next i

                                properAccount = vbNullString
                            End If
                        End If
                    End With

            End Select
            SetLastMessage GetAccountByIndex(Index), p_MainArray(1)
            Exit Sub
        End If

        Dim S1 As Long
        Dim E  As String

        'We just bother checking if the text is longer then 5 characters
        If Len(p_MainArray(1)) > 5 Then
            'If the text is just made of numbers we don't check
            If IsNumeric(p_MainArray(2)) = False Then
                'If there are letters and not just signs then check
                If IsAlphaCharacter(p_MainArray(1)) Then

                    For i = 1 To Len(p_MainArray(1))
                        E = Mid$(p_MainArray(1), i, 1)
                        If UCase$(E) = E Then S1 = S1 + 1
                    Next i

                    'Exit if there are more then 75% of caps
                    If Format$(100 * S1 / Len(p_MainArray(1)), "0") > 75 And Options.CAPS_CHECK = 1 Then
                        SendSingle "!pmessage" & pSplit & "message_blocked" & pSplit, Index
                        Exit Sub
                    End If
                End If
            End If
        End If

        'Send Message and print in chat
        E = GetAccountByIndex(Index) '//Temp variable used to save account to not load 4 times
        SendProtectedMessage E, GetGMFlag(E) & GetAFKFlag(E) & "[" & E & "]: " & p_MainArray(1)

        'Set last message
        SetLastMessage E, p_MainArray(1)

    Case Else
        SendSingle "ERROR", Index

End Select
Next k
End Sub

Private Sub SetLastMessage(User As String, Message As String)
Dim i As Long

With frmPanel.lvUsers.ListItems
    For i = 1 To .Count
        If .Item(i) = User Then
            .Item(i).SubItems(INDEX_LAST_MESSAGE) = Message
            Exit For
        End If
    Next i
End With
End Sub

Private Function IsRepeating(User As String, Message As String) As Boolean
Dim i As Long

With frmPanel.lvUsers.ListItems
    For i = 1 To .Count
        If .Item(i) = User Then
            If .Item(i).SubItems(INDEX_LAST_MESSAGE) = Message Then IsRepeating = True
            Exit For
        End If
    Next i
End With
End Function

Private Function IsPartOf(Part As String, Command As String) As Boolean
Dim i As Long

For i = 1 To Len(Part)
    If LCase$(Mid(Part, 1, i)) = LCase$(Left$(Command, Len(Mid(Part, 1, i)))) Then
        IsPartOf = True
    Else
        IsPartOf = False
        Exit Function
    End If
Next i
End Function

Private Function GetOnlineTime(User As String) As String
Dim TD    As String
Dim TD1() As String
Dim i     As Long
Dim j     As Long

With frmPanel.lvUsers.ListItems
    For i = 1 To .Count
        If .Item(i) = User Then
            TD = TimeSerial(0, 0, DateDiff("s", .Item(i).SubItems(INDEX_LOGIN_TIME), Time))

            TD1 = Split(TD, ":")
            TD = vbNullString

            For j = LBound(TD1) To UBound(TD1)
                Select Case j
                    Case 0
                        If TD1(j) <> 0 Then
                            If Left$(TD1(j), 1) = 0 Then
                                TD = TD & Right$(TD1(j), 1) & " hours "
                            Else
                                TD = TD & TD1(j) & " hours "
                            End If
                        End If

                    Case 1
                        If TD1(j) <> 0 Then
                            If Left$(TD1(j), 1) = 0 Then
                                TD = TD & Right$(TD1(j), 1) & " minutes "
                            Else
                                TD = TD & TD1(j) & " minutes "
                            End If
                        End If

                    Case 2
                        If TD1(j) <> 0 Then
                            If Left$(TD1(j), 1) = 0 Then
                                TD = TD & Right$(TD1(j), 1) & " seconds"
                            Else
                                TD = TD & TD1(j) & " seconds"
                            End If
                        End If

                End Select
            Next j

            GetOnlineTime = TD
        End If
    Next i
End With
End Function

Private Function GetProperAccountName(Account As String) As String
Dim i As Long

With frmAccountPanel.lvAccounts.ListItems
    For i = 1 To .Count
        If LCase$(.Item(i).SubItems(INDEX_NAME)) = LCase$(Account) Then
            GetProperAccountName = .Item(i).SubItems(INDEX_NAME)
            Exit For
        Else
            If i = .Count Then GetProperAccountName = Account
        End If
    Next i
End With
End Function

Private Sub Whisper(User As String, Target As String, Message As String, Index As Integer)
Dim i As Long
Dim j As Long

'Check if user is whispering itself
If User = Target Then
    SendSingle "!pmessage" & pSplit & "cant_whisper_self" & pSplit, Index
    Exit Sub
End If

With frmAccountPanel.lvAccounts.ListItems
    For i = 1 To .Count
        'Check if user exist
        If .Item(i).SubItems(INDEX_NAME) = Target Then
            'Check if target is ignoring user
            If IsIgnoring(Target, User) Then
                SendSingle "!pmessage" & pSplit & "is_ignoring_you" & pSplit & Target & pSplit, Index
            Else
                With frmPanel.lvUsers.ListItems
                    For j = 1 To .Count
                        'Check if target is online
                        If .Item(j) = Target Then
                            SendSingle "!pmessage" & pSplit & "you_whisper_to" & pSplit & Target & pSplit & Message & pSplit, Index
                            If LenB(GetAFKFlag(Target)) <> 0 Then
                                SendSingle "!pmessage" & pSplit & "target_is_afk" & pSplit & Target & pSplit, Index
                            End If
                            SendSingle "!pmessage" & pSplit & "whisper" & pSplit & GetGMFlag(User) & GetAFKFlag(User) & pSplit & User & pSplit & Message & pSplit, .Item(j).SubItems(INDEX_WINSOCK_ID)
                            Exit For
                        Else
                            If j = .Count Then
                                SendSingle "!pmessage" & pSplit & "message_sent_offline" & pSplit & Target & pSplit, Index
                                frmOfflineMessages.AddOfflineMessage User, Target, Message, Time
                            End If
                        End If
                    Next j
                End With
            End If
            Exit For
        Else
            If i = .Count Then SendSingle "!pmessage" & pSplit & "user_not_found" & pSplit & Target & pSplit, Index
        End If
    Next i
End With
End Sub

Private Sub MuteUser(User As String, IsMuted As Long, Reason As String, Index As Integer)
Dim i As Long

With frmPanel.lvUsers.ListItems
    For i = 1 To .Count
        If .Item(i) = User Then
            If IsMuted = 1 And .Item(i).SubItems(INDEX_MUTED) = "1" Then
                SendSingle "!pmessage" & pSplit & "user_already_muted" & pSplit & User & pSplit, Index
                Exit Sub

            ElseIf IsMuted = 0 And .Item(i).SubItems(INDEX_MUTED) = "0" Then
                SendSingle "!pmessage" & pSplit & "user_is_not_muted" & pSplit & User & pSplit, Index
                Exit Sub

            End If

            'Set flag in userlist
            .Item(i).SubItems(INDEX_MUTED) = IsMuted

            'Announce the action
            If LenB(Reason) = 0 Then
                If IsMuted = 1 Then
                    SendMessage "!pmessage" & pSplit & "muted_by" & pSplit & User & pSplit & GetAccountByIndex(Index) & pSplit
                Else
                    SendMessage "!pmessage" & pSplit & "unmuted_by" & pSplit & User & pSplit & GetAccountByIndex(Index) & pSplit
                End If
            Else
                If IsMuted = 1 Then
                    SendMessage "!pmessage" & pSplit & "muted_by_reason" & pSplit & User & pSplit & GetAccountByIndex(Index) & pSplit & Reason & pSplit
                Else
                    SendMessage "!pmessage" & pSplit & "unmuted_by_reason" & pSplit & User & pSplit & GetAccountByIndex(Index) & pSplit & Reason & pSplit
                End If
            End If
            Exit For
        Else
            If i = .Count Then
                If LenB(Trim$(User)) = 0 Then
                    If IsMuted = 1 Then
                        SendSingle "!pmessage" & pSplit & "incorrect_syntax" & pSplit & ".mute [User] [Reason]" & pSplit, Index
                    Else
                        SendSingle "!pmessage" & pSplit & "incorrect_syntax" & pSplit & ".unmute [User] [Reason]" & pSplit, Index
                    End If
                Else
                    SendSingle "!pmessage" & pSplit & "user_not_found" & pSplit & User & pSplit, Index
                End If
            End If
        End If
    Next i
End With
End Sub

Private Function GetAccountList() As String
Dim i As Long

With frmAccountPanel.lvAccounts.ListItems
    GetAccountList = "Account List:" & pSplit
    For i = 1 To .Count
        GetAccountList = GetAccountList & .Item(i).SubItems(INDEX_NAME) & pSplit
    Next i
End With
End Function

Private Function GetOnlineList() As String
Dim i As Long

With frmPanel.lvUsers.ListItems
    GetOnlineList = "User List:" & pSplit
    For i = 1 To .Count
        GetOnlineList = GetOnlineList & .Item(i) & pSplit
    Next i
End With
End Function

Private Sub Ban(Account As String, Ban As Long, Reason As String, Index As Integer)
Dim i As Long

With frmAccountPanel.lvAccounts.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(INDEX_NAME) = Account Then

            'If the account is already un/banned send feedback
            If Ban = 1 Then
                If .Item(i).SubItems(INDEX_BANNED) = "1" Then
                    SendSingle "!pmessage" & pSplit & "already_banned" & pSplit & Account & pSplit, Index
                    Exit Sub
                End If
            Else
                If .Item(i).SubItems(INDEX_BANNED) = "0" Then
                    SendSingle "!pmesssage" & pSplit & "already_unbanned" & pSplit & Account & pSplit, Index
                    Exit Sub
                End If
            End If

            'Ban account in database
            frmAccountPanel.ModifyAccount Account, .Item(i).SubItems(INDEX_PASSWORD), Ban, .Item(i).SubItems(INDEX_LEVEL), .Item(i), .Item(i).Index, .Item(i).SubItems(INDEX_GENDER), .Item(i).SubItems(INDEX_EMAIL)

            'Announce the action
            If LenB(Reason) = 0 Then
                If Ban = 1 Then
                    SendMessage "!pmessage" & pSplit & "banned_by" & pSplit & Account & pSplit & GetAccountByIndex(Index) & pSplit
                Else
                    SendMessage "!pmessage" & pSplit & "unbanned_by" & pSplit & Account & pSplit & GetAccountByIndex(Index) & pSplit
                End If
            Else
                If Ban = 1 Then
                    SendMessage "!pmessage" & pSplit & "banned_by_reason" & pSplit & Account & pSplit & GetAccountByIndex(Index) & pSplit & Reason & pSplit
                Else
                    SendMessage "!pmessage" & pSplit & "unbanned_by_reason" & pSplit & Account & pSplit & GetAccountByIndex(Index) & pSplit & Reason & pSplit
                End If
            End If
            Exit For
        Else
            If i = .Count Then
                If LenB(Trim$(Account)) = 0 Then
                    If Ban = 1 Then
                        SendSingle "!pmessage" & pSplit & "incorrect_syntax" & pSplit & ".ban [Name] [Reason]" & pSplit, Index
                    Else
                        SendSingle "!pmessage" & pSplit & "incorrect_syntax" & pSplit & ".unban [Name] [Reason]" & pSplit, Index
                    End If
                Else
                    SendSingle "!pmessage" & pSplit & "user_not_found" & pSplit & Account & pSplit, Index
                End If
            End If
        End If
    Next i
End With
End Sub

Private Sub KickUser(User As String)
Dim i As Long
Dim j As Long

With frmPanel.lvUsers.ListItems
    For i = 1 To .Count
        If .Item(i) = User Then
            Unload Winsock1(.Item(i).SubItems(INDEX_WINSOCK_ID))

            pDB.ExecuteCommand "UPDATE " & DATABASE_TABLE_ACCOUNTS & " SET LastIP1 = '" & .Item(i).SubItems(INDEX_IP) & "' WHERE Name1 = '" & User & "'"

            With frmAccountPanel.lvAccounts.ListItems
                For j = 1 To .Count
                    If .Item(j).SubItems(INDEX_NAME) = User Then
                        .Item(j).SubItems(INDEX_LAST_IP) = frmPanel.lvUsers.ListItems.Item(i).SubItems(INDEX_IP)
                        Exit For
                    End If
                Next j
            End With

            .Remove (i)

            UPDATE_ONLINE
            UPDATE_STATUS_BAR

            SendMessage "!pmessage" & pSplit & "offline" & pSplit & User & pSplit
            frmChannel.LeaveAllChannels User
            Exit For
        End If
    Next i
End With
End Sub

Private Sub GetAccountInfo(Account As String, UsedSyntax As String, Index As Integer)
Dim i As Long

With frmAccountPanel.lvAccounts.ListItems
    For i = 1 To .Count
        If LCase(.Item(i).SubItems(INDEX_NAME)) = LCase(Account) Then
            SendSingle vbCrLf & " Account information about '" & Account & "'" & vbCrLf & " ID: " & .Item(i) & vbCrLf & " Password: " & .Item(i).SubItems(INDEX_PASSWORD) & vbCrLf & " Registration Time: " & .Item(i).SubItems(INDEX_TIME) & vbCrLf & " Registration Date: " & .Item(i).SubItems(INDEX_DATE) & vbCrLf & " Banned: " & .Item(i).SubItems(INDEX_BANNED) & vbCrLf & " Level: " & .Item(i).SubItems(INDEX_LEVEL) & vbCrLf & " Gender: " & .Item(i).SubItems(INDEX_GENDER) & vbCrLf & " Email: " & .Item(i).SubItems(INDEX_EMAIL) & vbCrLf & " Last IP:  " & .Item(i).SubItems(INDEX_LAST_IP), Index
            Exit For
        Else
            If i = .Count Then
                If LenB(Trim$(Account)) = 0 Then
                    SendSingle "!pmessage" & pSplit & "incorrect_syntax" & pSplit & UsedSyntax & " [User]" & pSplit, Index
                Else
                    SendSingle "!pmessage" & pSplit & "user_not_found" & pSplit & Account & pSplit, Index
                End If
            End If
        End If
    Next i
End With
End Sub

Private Sub GetUserInfo(User As String, UsedSyntax As String, Index As Integer)
Dim i As Long

With frmPanel.lvUsers.ListItems
    For i = 1 To .Count
        If .Item(i) = User Then
            SendSingle vbCrLf & "User information about '" & .Item(i) & "'" & vbCrLf & " IP : " & .Item(i).SubItems(INDEX_IP) & vbCrLf & " Winsock ID: " & .Item(i).SubItems(INDEX_WINSOCK_ID) & vbCrLf & " Last Message: " & .Item(i).SubItems(INDEX_LAST_MESSAGE) & vbCrLf & " Muted: " & .Item(i).SubItems(INDEX_MUTED) & vbCrLf & " Login Time: " & .Item(i).SubItems(INDEX_LOGIN_TIME) & vbCrLf & " GM Flag: " & .Item(i).SubItems(INDEX_GM_FLAG) & vbCrLf & " AFK Flag: " & .Item(i).SubItems(INDEX_AFK_FLAG), Index
            Exit For
        Else
            If i = .Count Then
                If LenB(User) = 0 Then
                    SendSingle "!pmessage" & pSplit & "incorrect_syntax" & pSplit & UsedSyntax & " [User]" & pSplit, Index
                Else
                    SendSingle "!pmessage" & pSplit & "user_not_found" & pSplit & User & pSplit, Index
                End If
            End If
        End If
    Next i
End With
End Sub

Private Function GetCommands() As String
Dim i As Long

GetCommands = vbCrLf & "*********************************************" & vbCrLf & "* List of all avaible commands:" & vbCrLf
For i = 0 To UBound(Commands) - 1
    GetCommands = GetCommands & "* " & Commands(i).Syntax & " (" & Commands(i).Description & ")" & vbCrLf
Next i
GetCommands = GetCommands & "*********************************************"
End Function

Private Function GetLevel(User As String) As Long
Dim i As Long
Dim j As Long

With frmAccountPanel.lvAccounts.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(INDEX_NAME) = User Then
            If .Item(i).SubItems(INDEX_LEVEL) = "0" Then
                With frmPanel.lvUsers.ListItems
                    For j = 1 To .Count
                        If .Item(j) = User Then
                            If .Item(j).SubItems(INDEX_IS_ROOT) = "1" Then
                                GetLevel = 1
                            End If
                            Exit For
                        End If
                    Next j
                End With
            Else
                GetLevel = .Item(i).SubItems(INDEX_LEVEL)
            End If
            Exit For
        End If
    Next i
End With
End Function

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
KickUser GetAccountByIndex(Index)
End Sub

Public Sub SetupForms(NewForm As Form)
Dim pForm As Form

For Each pForm In Forms
    If Not pForm.Name = Name Then pForm.Hide
Next
NewForm.Show
End Sub

Private Function GetAccountByIndex(Index As Integer) As String
Dim i As Long

With frmPanel.lvUsers.ListItems
    For i = 1 To .Count
        If CInt(.Item(i).SubItems(INDEX_WINSOCK_ID)) = Index Then
            GetAccountByIndex = .Item(i)
            Exit For
        End If
    Next i
End With
End Function

Private Function GetGMFlag(User As String) As String
Dim i As Long

With frmPanel.lvUsers.ListItems
    For i = 1 To .Count
        If .Item(i) = User Then
            If CInt(.Item(i).SubItems(INDEX_GM_FLAG)) = 1 Then GetGMFlag = "<GM>"
            Exit For
        End If
    Next i
End With
End Function

Private Function GetAFKFlag(User As String) As String
Dim i As Long

With frmPanel.lvUsers.ListItems
    For i = 1 To .Count
        If .Item(i) = User Then
            If CInt(.Item(i).SubItems(INDEX_AFK_FLAG)) = 1 Then GetAFKFlag = "<AFK>"
            Exit For
        End If
    Next i
End With
End Function
