VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
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
      Begin VB.CommandButton Command6 
         Caption         =   "C&HANNELS"
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
         Left            =   5850
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&FL / IL"
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
         Left            =   4770
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&USERS"
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
         Left            =   3690
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "A&CCOUNTS"
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
         Left            =   2610
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "CH&AT"
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
         Left            =   1530
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&CONF"
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
         Left            =   450
         TabIndex        =   1
         Top             =   120
         Width           =   1095
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
Dim i    As Long
Dim j    As Long
Dim User As String

User = GetAccountByIndex(Index)
Unload Winsock1(Index)

If LenB(User) <> 0 Then SendMessage "!pmessage#offline#" & User & "#"

With frmPanel.lvUsers.ListItems
    For i = 1 To .Count
        If .Item(i) = User Then
            If Len(User) <> 0 Then
                Call pDB.ExecuteCommand("UPDATE " & DATABASE_TABLE_ACCOUNTS & " SET LastIP1 = '" & .Item(i).SubItems(INDEX_IP) & "' WHERE Name1 = '" & User & "'")

                With frmAccountPanel.lvAccounts.ListItems
                    For j = 1 To .Count
                        If .Item(j).SubItems(INDEX_NAME) = User Then
                            .Item(j).SubItems(INDEX_LAST_IP) = frmPanel.lvUsers.ListItems.Item(i).SubItems(INDEX_IP)
                            Exit For
                        End If
                    Next j
                End With
            End If

            .Remove i
            Exit For
        End If
    Next i
End With

frmChannel.LeaveAllChannels User

UPDATE_ONLINE
UPDATE_STATUS_BAR
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
Dim p_MainArray()   As String   'Whole message string is saved here and split up by # sign
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
p_MainArray = Split(p_PreArray(k), "#")

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
                    SendSingle "!pmessage#online#" & properAccount & "#", .Item(i).SubItems(INDEX_WINSOCK_ID)
                End If
            Next i
        End With

        properAccount = vbNullString

        'Send Server information
        SendSingle vbCrLf & _
                   " Welcome to Peach Servers." & vbCrLf & _
                   " Server: Peach r" & pRev & "/" & GetOS & vbCrLf & _
                   " Online User: " & Winsock1.Count - 1 & vbCrLf & _
                   " Server Uptime: " & TimeSerial(0, 0, DateDiff("s", frmConfig.START_TIME, Time)), Index

    Case "!login"
        With frmAccountPanel.lvAccounts.ListItems
            If .Count = 0 Then
                SendSingle "!login#Account#", Index
            Else
                'Authentication; password, is banned?, account etc ..
                For i = 1 To .Count
                    If LCase$(.Item(i).SubItems(INDEX_NAME)) = LCase$(p_MainArray(1)) Then
                        'Ban Check
                        If .Item(i).SubItems(INDEX_BANNED) = "1" Then
                            SendSingle "!login#Banned#", Index
                            Exit Sub
                        End If

                        'Password Check
                        If Not GetMD5(.Item(i).SubItems(INDEX_PASSWORD)) = p_MainArray(2) Then
                            SendSingle "!login#Password#", Index
                            Exit Sub
                        End If
                        Exit For
                    Else
                        If i = .Count Then
                            SendSingle "!login#Account#", Index
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
                    SendSingle "!accepted#" & .Item(.Count).Text & "#", Index
                End With
            End If
        End With

    'We get ip request and send ip back
    Case "!iprequest"
        With frmPanel.lvUsers.ListItems
            For i = 1 To .Count
                If .Item(i) = p_MainArray(1) Then
                    SendSingle "!iprequest#" & .Item(i).SubItems(INDEX_IP) & "#", Index
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
                    If i = .Count Then SendSingle "!pmessage#channel_wrong_password#" & p_MainArray(2) & "#", Index
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
        Select Case Left$(p_CHAT_ARRAY(0), 1)
            Case Chr(46)
                If GetLevel(GetAccountByIndex(Index)) > 0 Then IsCommand = True

            Case Chr(47)
                IsSlash = True

        End Select

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
                        SendSingle "!split_text#" & GetAccountList, Index

                    ElseIf IsPartOf(p_TEXT_FIRST, "onliners") Then
                        SendSingle "!split_text#" & GetOnlineList, Index

                    Else
                        SendSingle "!pmessage#incorrect_syntax#.show [accounts, onliners]", Index

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
                                        SendSingle "!pmessage#incorrect_syntax#.kick [User]#", Index
                                    Else
                                        SendSingle "!pmessage#user_not_found#" & p_TEXT_FIRST & "#", Index
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
                        SendSingle "!pmessage#incorrect_syntax#" & p_CHAT_ARRAY(0) & " [Text]#", Index
                    Else
                        Select Case LCase$(p_CHAT_ARRAY(0))
                            Case ".announce", ".ann"
                                Reason = GetAccountByIndex(Index)
                                SendMessage "!pmessage#announce#" & GetGMFlag(Reason) & GetAFKFlag(Reason) & "#" & Reason & "#" & p_ANN_MSG & "#"

                            Case ".notify"
                                SendMessage "!msgbox#" & GetAccountByIndex(Index) & ": " & p_ANN_MSG & "#"

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
                                                    SendSingle "!pmessage#user_already_used#" & p_CHAT_ARRAY(3) & "#", Index
                                                    Exit For
                                                Else
                                                    If n = .Count Then
                                                        'Modify account in database and panel
                                                        frmAccountPanel.ModifyAccount p_CHAT_ARRAY(3), .Item(i).SubItems(INDEX_PASSWORD), .Item(i).SubItems(INDEX_BANNED), .Item(i).SubItems(INDEX_LEVEL), .Item(i), i, .Item(i).SubItems(INDEX_GENDER), .Item(i).SubItems(INDEX_EMAIL)

                                                        'Send feedback to person who changed name
                                                        SendSingle "!pmessage#successfull_rename#" & properAccount & "#" & p_CHAT_ARRAY(3) & "#", Index

                                                        'Check if player is online and directly rename it, also tell user that he got renamed
                                                        With frmPanel.lvUsers.ListItems
                                                            For m = 1 To .Count
                                                                If .Item(m) = properAccount Then
                                                                    .Item(m) = p_CHAT_ARRAY(3)

                                                                    SendSingle "!pmessage#renamed_you_to#" & GetAccountByIndex(Index) & "#" & p_CHAT_ARRAY(3) & "#", .Item(m).SubItems(INDEX_WINSOCK_ID)
                                                                    SendSingle "!namechange#" & p_CHAT_ARRAY(3) & "#", .Item(m).SubItems(INDEX_WINSOCK_ID)
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

                                                        pDB.ExecuteCommand "UPDATE " & DATABASE_TABLE_FRIENDS & " SET Name = '" & p_CHAT_ARRAY(3) & "' WHERE Name = '" & properAccount & "'"
                                                        pDB.ExecuteCommand "UPDATE " & DATABASE_TABLE_FRIENDS & " SET Friend = '" & p_CHAT_ARRAY(3) & "' WHERE Friend = '" & properAccount & "'"

                                                        UPDATE_ONLINE
                                                    End If
                                                End If
                                            Next n
                                            Exit For
                                        Else
                                            If i = .Count Then SendSingle "!pmessage#user_not_found#" & properAccount & "#", Index
                                        End If
                                    Next i
                                End With

                                properAccount = vbNullString
                            Else
                                SendSingle "!pmessage#incorrect_syntax#" & p_CHAT_ARRAY(0) & Space(1) & p_TEXT_FIRST & " [Oldname] [Newname]#", Index
                            End If

                        Case "level"
                            If UBound(p_CHAT_ARRAY) > 2 Then
                                'Level must be numeric and in range of 0-2
                                If (Not IsNumeric(p_CHAT_ARRAY(3))) Or (p_CHAT_ARRAY(3) > 2 Or p_CHAT_ARRAY(3) < 0) Then
                                    SendSingle "!pmessage#level_incorrect_value#", Index
                                    Exit Sub
                                End If

                                With frmAccountPanel.lvAccounts.ListItems
                                    properAccount = GetProperAccountName(p_TEXT_SECOND)

                                    For i = 1 To .Count
                                        If .Item(i).SubItems(INDEX_NAME) = properAccount Then
                                            'Modify level for account in database
                                            frmAccountPanel.ModifyAccount .Item(i).SubItems(INDEX_NAME), .Item(i).SubItems(INDEX_PASSWORD), .Item(i).SubItems(INDEX_BANNED), p_CHAT_ARRAY(3), .Item(i), i, .Item(i).SubItems(INDEX_GENDER), .Item(i).SubItems(INDEX_EMAIL)

                                            'Feedback to the person who modified the level
                                            SendSingle "!pmessage#successfull_level#" & .Item(i).SubItems(INDEX_NAME) & "#" & p_CHAT_ARRAY(3) & "#", Index

                                            With frmPanel.lvUsers.ListItems
                                                For n = 1 To .Count
                                                    If .Item(n) = properAccount Then
                                                        SendSingle "!pmessage#changed_your_level#" & GetAccountByIndex(Index) & "#" & p_CHAT_ARRAY(3) & "#", .Item(n).SubItems(INDEX_WINSOCK_ID)
                                                        Exit For
                                                    End If
                                                Next n
                                            End With
                                            Exit For
                                        Else
                                            If i = .Count Then SendSingle "!pmessage#user_not_found#" & p_TEXT_SECOND & "#", Index
                                        End If
                                    Next i

                                    properAccount = vbNullString
                                End With
                            Else
                                SendSingle "!pmessage#incorrect_syntax#" & p_CHAT_ARRAY(0) & Space(1) & p_TEXT_FIRST & " [Name] [Level]#", Index
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
                                    SendSingle "!pmessage#gender_incorrect_value#", Index
                                Else
                                    With frmAccountPanel.lvAccounts.ListItems
                                        properAccount = GetProperAccountName(p_TEXT_SECOND)

                                        For i = 1 To .Count
                                            If .Item(i).SubItems(INDEX_NAME) = properAccount Then
                                                'Modify gender in database and panel
                                                frmAccountPanel.ModifyAccount .Item(i).SubItems(INDEX_NAME), .Item(i).SubItems(INDEX_PASSWORD), .Item(i).SubItems(INDEX_BANNED), .Item(i).SubItems(INDEX_LEVEL), .Item(i), i, GenderInNumeric, .Item(i).SubItems(INDEX_EMAIL)

                                                'Feedback to the person who modified the gender
                                                SendSingle "!pmessage#successfull_gender#" & properAccount & "#" & GenderInNumeric & "#", Index

                                                'Notify user that gender got changed
                                                With frmPanel.lvUsers.ListItems
                                                    For n = 1 To .Count
                                                        If .Item(n) = properAccount Then
                                                            SendSingle "!pmessage#changed_your_gender#" & GetAccountByIndex(Index) & "#" & GenderInNumeric & "#", .Item(n).SubItems(INDEX_WINSOCK_ID)
                                                            Exit For
                                                        End If
                                                    Next n
                                                End With
                                                Exit For
                                            Else
                                                If i = .Count Then SendSingle "!pmessage#user_not_found#" & p_TEXT_SECOND & "#", Index
                                            End If
                                        Next i

                                        properAccount = vbNullString
                                    End With
                                End If
                            Else
                                SendSingle "!pmessage#incorrect_syntax#" & p_CHAT_ARRAY(0) & Space(1) & p_TEXT_FIRST & " [Name] [Gender]#", Index
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
                                            SendSingle "!pmessage#successfull_password#" & properAccount & "#" & p_CHAT_ARRAY(3) & "#", Index

                                            'Notify user that password got changed
                                            With frmPanel.lvUsers.ListItems
                                                For n = 1 To .Count
                                                    If .Item(n) = properAccount Then
                                                        SendSingle "!pmessage#changed_your_password#" & GetAccountByIndex(Index) & "#" & p_CHAT_ARRAY(3) & "#", .Item(n).SubItems(INDEX_WINSOCK_ID)
                                                        Exit For
                                                    End If
                                                Next n
                                            End With
                                            Exit For
                                        Else
                                            If i = .Count Then SendSingle "!pmessage#user_not_found#" & p_TEXT_SECOND & "#", Index
                                        End If
                                    Next i

                                    properAccount = vbNullString
                                End With
                            Else
                                SendSingle "!pmessage#incorrect_syntax#" & p_CHAT_ARRAY(0) & Space(1) & p_TEXT_FIRST & " [Name] [Password]#", Index
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
                                            SendSingle "!pmessage#successfull_email#" & properAccount & "#" & p_CHAT_ARRAY(3) & "#", Index

                                            'Notify user that password got changed
                                            With frmPanel.lvUsers.ListItems
                                                For n = 1 To .Count
                                                    If .Item(n) = properAccount Then
                                                        SendSingle "!pmessage#changed_your_email#" & GetAccountByIndex(Index) & "#" & p_CHAT_ARRAY(3) & "#", .Item(n).SubItems(INDEX_WINSOCK_ID)
                                                        Exit For
                                                    End If
                                                Next n
                                            End With
                                            Exit For
                                        Else
                                            If i = .Count Then SendSingle "!pmessage#user_not_found#" & p_TEXT_SECOND & "#", Index
                                        End If
                                    Next i

                                    properAccount = vbNullString
                                End With
                            Else
                                SendSingle "!pmessage#incorrect_syntax#" & p_CHAT_ARRAY(0) & Space(1) & p_TEXT_FIRST & " [Name] [Email]#", Index
                            End If

                        Case Else
                            SendSingle _
                            "!pmessage#incorrect_syntax#" & _
                                vbCrLf & Space(2) & _
                            p_CHAT_ARRAY(0) & " name [Oldname] [Newname]" & _
                                vbCrLf & Space(2) & _
                            p_CHAT_ARRAY(0) & " level [Name] [Level]" & _
                                vbCrLf & Space(2) & _
                            p_CHAT_ARRAY(0) & " gender [Name] [Gender]" & _
                                vbCrLf & Space(2) & _
                            p_CHAT_ARRAY(0) & " password [Name] [Password]" & _
                                vbCrLf & Space(2) & _
                            p_CHAT_ARRAY(0) & " email [Name] [Email]#", Index

                    End Select

                Case ".reload"
                    Dim loadTime As Long

                    With Database
                        Select Case LCase$(p_TEXT_FIRST)
                            Case LCase$(DATABASE_TABLE_ACCOUNTS), LCase$(DATABASE_TABLE_FRIENDS), LCase$(DATABASE_TABLE_IGNORES)
                                SendSingle "!pmessage#table_cant_reload#", Index

                            Case LCase$(DATABASE_TABLE_COMMANDS)
                                loadTime = timeGetTime
                                Erase Commands
                                LoadCommands
                                SendMessage "!pmessage#table_reload#" & GetAccountByIndex(Index) & "#" & DATABASE_TABLE_COMMANDS & "#" & timeGetTime - loadTime & "#"

                            Case LCase$(DATABASE_TABLE_DECLINED_NAMES)
                                loadTime = timeGetTime
                                Erase DeclinedNames
                                LoadDeclinedNames
                                SendMessage "!pmessage#table_reload#" & GetAccountByIndex(Index) & "#" & DATABASE_TABLE_DECLINED_NAMES & "#" & timeGetTime - loadTime & "#"

                            Case LCase$(DATABASE_TABLE_EMOTES)
                                loadTime = timeGetTime
                                Erase Emotes
                                LoadEmotes
                                SendMessage "!pmessage#table_reload#" & GetAccountByIndex(Index) & "#" & DATABASE_TABLE_EMOTES & "#" & timeGetTime - loadTime & "#"

                            Case LCase$("config"), LCase$("c")
                                loadTime = timeGetTime
                                LoadConfigValue
                                SendMessage "!pmessage#config_reload#" & GetAccountByIndex(Index) & "#" & timeGetTime - loadTime & "#"

                            Case Else
                                If LenB(p_TEXT_FIRST) = 0 Then
                                    SendSingle "!pmessage#incorrect_syntax#.reload Table#", Index
                                Else
                                    SendSingle "!pmessage#table_not_exist#.", Index
                                End If

                        End Select
                    End With

                Case ".clear"
                    If LenB(p_TEXT_FIRST) = 0 Then
                        SendSingle "!pmessage#incorrect_syntax#.clear [User]#", Index
                    Else
                        If LCase$(p_TEXT_FIRST) = "this" Or LCase$(p_TEXT_FIRST) = "me" Then
                            SendSingle "!clear#", Index
                            Exit Sub

                        ElseIf LCase$(p_TEXT_FIRST) = "all" Then
                            SendMessage "!clear#"
                            Exit Sub

                        End If

                        With frmPanel.lvUsers.ListItems
                            properAccount = GetProperAccountName(p_TEXT_FIRST)

                            For i = 1 To .Count
                                If .Item(i) = properAccount Then
                                    SendSingle "!clear#", .Item(i).SubItems(INDEX_WINSOCK_ID)
                                    Exit For
                                Else
                                    If i = .Count Then SendSingle "!pmessage#user_not_found#" & p_TEXT_FIRST & "#", Index
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
                                pDB.ExecuteCommand "DELETE FROM " & DATABASE_TABLE_ACCOUNTS & " WHERE ID = " & .Item(i)
                                frmFriendIgnoreList.RemoveAllFriendsFromUser .Item(i).SubItems(INDEX_NAME)
                                frmFriendIgnoreList.RemoveAllIgnoresFromUser .Item(i).SubItems(INDEX_NAME)

                                SendSingle "!pmessage#deleted_account#" & properAccount & "#" & .Item(i) & "#", Index
                                .Remove i
                                Exit For
                            Else
                                If i = .Count Then
                                    If LenB(properAccount) = 0 Then
                                        SendSingle "!pmessage#incorrect_syntax#.delete [Account]#", Index
                                    Else
                                        SendSingle "!pmessage#user_not_found#" & p_TEXT_FIRST & "#", Index
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
                                        SendSingle "!pmessage#gm_flag_enable#", Index
                                        Exit For
                                    End If
                                Next i

                            Case "off"
                                For i = 1 To .Count
                                    If .Item(i).SubItems(INDEX_WINSOCK_ID) = Index Then
                                        .Item(i).SubItems(INDEX_GM_FLAG) = "0"
                                        SendSingle "!pmessage#gm_flag_disable#", Index
                                        Exit For
                                    End If
                                Next i

                            Case Else
                                SendSingle "!pmessage#incorrect_syntax#.gm [on / off]#", Index

                        End Select
                    End With

                Case Else
                    SendSingle "!pmessage#unknown_command#", Index

            End Select
            Exit Sub 'Always exit after any command used or else it will be written for everyone
        End If

        'Check if user is muted for emotes and normal text
        If IsMuted Then
            SendSingle "!pmessage#muted#", Index
            Exit Sub
        End If

        If Options.REPEAT_CHECK = 1 Then
            If IsRepeating(GetAccountByIndex(Index), p_MainArray(1)) Then
                SendSingle "!pmessage#flood_protection#", Index
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

                    SendProtectedMessage properAccount, "!pmessage#roll#" & properAccount & "#" & Roll & "#" & MinRoll & "#" & MaxRoll & "#"

                    properAccount = vbNullString

                'Whisper X to Z from Y
                Case "/w", "/whisper"
                    If IsUser Then
                        Dim Message As String

                        If UBound(p_CHAT_ARRAY) > 1 Then
                            For i = 2 To UBound(p_CHAT_ARRAY)
                                Message = Message & p_CHAT_ARRAY(i) & " "
                            Next i

                            Whisper GetAccountByIndex(Index), GetProperAccountName(p_TEXT_FIRST), Trim$(Message), Index
                        End If
                    Else
                        If LenB(p_TEXT_FIRST) = 0 Then
                            SendSingle "!pmessage#incorrect_syntax#/whisper [Name] [Text]#", Index
                            Exit Sub
                        Else
                            SendSingle "!pmessage#user_not_found#" & p_TEXT_FIRST & "#", Index
                        End If
                    End If

                Case "/afk"
                    With frmPanel.lvUsers.ListItems
                        For i = 1 To .Count
                            If .Item(i).SubItems(INDEX_WINSOCK_ID) = Index Then
                                If .Item(i).SubItems(INDEX_AFK_FLAG) = "1" Then
                                    .Item(i).SubItems(INDEX_AFK_FLAG) = "0"
                                    SendSingle "!pmessage#not_afk#", Index
                                Else
                                    .Item(i).SubItems(INDEX_AFK_FLAG) = "1"
                                    SendSingle "!pmessage#afk#", Index
                                End If
                                Exit For
                            End If
                        Next i
                    End With

                Case "/online"
                    SendSingle "!pmessage#online_time#" & Trim$(GetOnlineTime(GetAccountByIndex(Index))) & "#", Index

                Case "/logout"
                    KickUser GetAccountByIndex(Index)

                Case "/join"
                    If LenB(p_TEXT_FIRST) = 0 Then
                        SendSingle "!pmessage#valid_channel#", Index
                    Else
                        With frmChannel.lvUsers.ListItems
                            If .Count = 0 Then
                                frmChannel.JoinChannel p_TEXT_FIRST, GetAccountByIndex(Index), Index
                            Else
                                properAccount = GetAccountByIndex(Index)

                                For i = 1 To .Count
                                    If .Item(i) = properAccount And LCase$(.Item(i).SubItems(CHANNEL_USER_CHANNEL)) = LCase$(p_TEXT_FIRST) Then
                                        SendSingle "!pmessage#already_in_channel#" & .Item(i).SubItems(CHANNEL_USER_CHANNEL) & "#", Index
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
                        SendSingle "!pmessage#valid_channel#", Index
                    Else
                        frmChannel.LeaveChannel p_TEXT_FIRST, GetAccountByIndex(Index)
                    End If

                Case "/announce"
                    If LenB(p_TEXT_FIRST) = 0 Then
                        SendSingle "!pmessage#valid_channel#", Index
                    Else
                        If frmChannel.lvChannels.ListItems.Count = 0 Then
                            SendSingle "!pmessage#not_in_channel" & p_TEXT_FIRST & "#", Index
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

                                                        SendMessageToChannel .Item(f), properAccount, "!pmessage#channel_announcements#" & .Item(f) & "#" & properAccount & "#"
                                                        Exit For
                                                    End If
                                                Next f
                                            End With
                                        Else
                                            SendSingle "!pmessage#not_channel_leader#", Index
                                        End If
                                        Exit For
                                    Else
                                        If i = .Count Then SendSingle "!pmessage#not_in_channel#" & p_TEXT_FIRST & "#", Index
                                    End If
                                Next i

                                properAccount = vbNullString
                            End With
                        End If
                    End If

                Case "/setpassword"
                    If LenB(p_TEXT_FIRST) = 0 Then
                        SendSingle "!pmessage#valid_channel#", Index
                    Else
                        If LenB(p_TEXT_SECOND) = 0 Then
                            SendSingle "!pmessage#incorrect_syntax#/setpassword [Channel] [Password]#", Index
                        Else
                            If frmChannel.lvChannels.ListItems.Count = 0 Then
                                SendSingle "!pmessage#not_in_channel#" & p_TEXT_FIRST & "#", Index
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
                                                            SendSingle "!pmessage#channel_password" & .Item(f) & "#" & p_TEXT_SECOND & "#", Index
                                                            Exit For
                                                        End If
                                                    Next f
                                                End With
                                            Else
                                                SendSingle "!pmessage#not_channel_leader#", Index
                                            End If
                                            Exit For
                                        Else
                                            If i = .Count Then SendSingle "!pmessage#not_in_channel#" & p_TEXT_FIRST & "#", Index
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
                                SendSingle "!pmessage#valid_channel#", Index
                            Else
                                SendSingle "!pmessage#not_in_channel#" & Right$(p_CHAT_ARRAY(0), Len(p_CHAT_ARRAY(0)) - 1) & "#", Index
                            End If
                        Else
                            If LenB(p_TEXT_FIRST) = 0 Then
                                SendSingle "!pmessage#valid_channel#", Index
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
                                                    If f = .Count Then SendSingle "!pmessage#not_in_channel#" & properAccount & "#", Index
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
                        SendSingle "!pmessage#message_blocked#", Index
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
                                TD = TD & Left$(TD1(j), 1) & " hours"
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

'Check if user is whispering itself
If User = Target Then
    SendSingle "!pmessage#cant_whisper_self#", Index
    Exit Sub
End If

With frmPanel.lvUsers.ListItems
    'Search target in list and send message
    For i = 1 To .Count
        If .Item(i) = Target Then
            If IsIgnoring(.Item(i), User) Then
                SendSingle "!pmessage#is_ignoring_you#" & Target & "#", Index
            Else
                SendSingle "!pmessage#you_whisper_to#" & Target & "#" & Message & "#", Index
                If LenB(GetAFKFlag(Target)) <> 0 Then SendSingle "!pmessage#target_is_afk#" & Target & "#", Index
                SendSingle "!pmessage#whisper#" & GetGMFlag(User) & GetAFKFlag(User) & "#" & User & "#" & Message & "#", .Item(i).SubItems(INDEX_WINSOCK_ID)
            End If
            Exit For
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
                SendSingle "!pmessage#user_already_muted#" & User & "#", Index
                Exit Sub

            ElseIf IsMuted = 0 And .Item(i).SubItems(INDEX_MUTED) = "0" Then
                SendSingle "!pmessage#user_is_not_muted#" & User & "#", Index
                Exit Sub

            End If

            'Set flag in userlist
            .Item(i).SubItems(INDEX_MUTED) = IsMuted

            'Announce the action
            If LenB(Reason) = 0 Then
                If IsMuted = 1 Then
                    SendMessage "!pmessage#muted_by#" & User & "#" & GetAccountByIndex(Index) & "#"
                Else
                    SendMessage "!pmessage#unmuted_by#" & User & "#" & GetAccountByIndex(Index) & "#"
                End If
            Else
                If IsMuted = 1 Then
                    SendMessage "!pmessage#muted_by_reason#" & User & "#" & GetAccountByIndex(Index) & "#" & Reason & "#"
                Else
                    SendMessage "!pmessage#unmuted_by_reason#" & User & "#" & GetAccountByIndex(Index) & "#" & Reason & "#"
                End If
            End If
            Exit For
        Else
            If i = .Count Then
                If LenB(Trim$(User)) = 0 Then
                    If IsMuted = 1 Then
                        SendSingle "!pmessage#incorrect_syntax#.mute [User] [Reason]#", Index
                    Else
                        SendSingle "!pmessage#incorrect_syntax#.unmute [User] [Reason]#", Index
                    End If
                Else
                    SendSingle "!pmessage#user_not_found#" & User & "#", Index
                End If
            End If
        End If
    Next i
End With
End Sub

Private Function GetAccountList() As String
Dim i As Long

With frmAccountPanel.lvAccounts.ListItems
    GetAccountList = "Account List:#"
    For i = 1 To .Count
        GetAccountList = GetAccountList & .Item(i).SubItems(INDEX_NAME) & "#"
    Next i
End With
End Function

Private Function GetOnlineList() As String
Dim i As Long

With frmPanel.lvUsers.ListItems
    GetOnlineList = "User List:#"
    For i = 1 To .Count
        GetOnlineList = GetOnlineList & .Item(i) & "#"
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
                    SendSingle "!pmessage#already_banned#" & Account & "#", Index
                    Exit Sub
                End If
            Else
                If .Item(i).SubItems(INDEX_BANNED) = "0" Then
                    SendSingle "!pmesssage#already_unbanned#" & Account & "#", Index
                    Exit Sub
                End If
            End If

            'Ban account in database
            frmAccountPanel.ModifyAccount Account, .Item(i).SubItems(INDEX_PASSWORD), Ban, .Item(i).SubItems(INDEX_LEVEL), .Item(i), .Item(i).Index, .Item(i).SubItems(INDEX_GENDER), .Item(i).SubItems(INDEX_EMAIL)

            'Announce the action
            If LenB(Reason) = 0 Then
                If Ban = 1 Then
                    SendMessage "!pmessage#banned_by#" & Account & "#" & GetAccountByIndex(Index) & "#"
                Else
                    SendMessage "!pmessage#unbanned_by#" & Account & "#" & GetAccountByIndex(Index) & "#"
                End If
            Else
                If Ban = 1 Then
                    SendMessage "!pmessage#banned_by_reason#" & Account & "#" & GetAccountByIndex(Index) & "#" & Reason & "#"
                Else
                    SendMessage "!pmessage#unbanned_by_reason#" & Account & "#" & GetAccountByIndex(Index) & "#" & Reason & "#"
                End If
            End If
            Exit For
        Else
            If i = .Count Then
                If LenB(Trim$(Account)) = 0 Then
                    If Ban = 1 Then
                        SendSingle "!pmessage#incorrect_syntax#.ban [Name] [Reason]", Index
                    Else
                        SendSingle "!pmessage#incorrect_syntax#.unban [Name] [Reason]", Index
                    End If
                Else
                    SendSingle "!pmessage#user_not_found" & Account & "#", Index
                End If
            End If
        End If
    Next i
End With
End Sub

Private Sub KickUser(User As String)
Dim i As Long

With frmPanel.lvUsers.ListItems
    For i = 1 To .Count
        If .Item(i) = User Then
            Unload Winsock1(.Item(i).SubItems(INDEX_WINSOCK_ID))

            .Remove (i)

            UPDATE_ONLINE
            UPDATE_STATUS_BAR

            SendMessage "!pmessage#offline#" & User & "#"
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
                    SendSingle "!pmessage#incorrect_syntax#" & UsedSyntax & " [User]", Index
                Else
                    SendSingle "!pmessage#user_not_found#" & Account & "#", Index
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
            SendSingle vbCrLf & "User information about '" & .Item(i) & "'" & vbCrLf & " IP : " & .Item(i).SubItems(INDEX_IP) & vbCrLf & " Winsock ID: " & .Item(i).SubItems(INDEX_WINSOCK_ID) & vbCrLf & " Last Message: " & .Item(i).SubItems(INDEX_LAST_MESSAGE) & vbCrLf & " Muted: " & .Item(i).SubItems(INDEX_MUTED) & vbCrLf & " Login Time: " & .Item(i).SubItems(INDEX_LOGIN_TIME) & " GM Flag: " & .Item(i).SubItems(INDEX_GM_FLAG) & " AFK Flag: " & .Item(i).SubItems(INDEX_AFK_FLAG), Index
            Exit For
        Else
            If i = .Count Then
                If LenB(User) = 0 Then
                    SendSingle "!pmessage#incorrect_syntax#" & UsedSyntax & " [User]", Index
                Else
                    SendSingle "!pmessage#user_not_found#" & User & "#", Index
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

With frmAccountPanel.lvAccounts.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(INDEX_NAME) = User Then
            GetLevel = .Item(i).SubItems(INDEX_LEVEL)
            Exit For
        End If
    Next i
End With
End Function

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim i    As Long
Dim j    As Long
Dim User As String

User = GetAccountByIndex(Index)
Unload Winsock1(Index)

If LenB(User) <> 0 Then SendMessage "!pmessage#offline#" & User & "#"

With frmPanel.lvUsers.ListItems
    For i = 1 To .Count
        If .Item(i) = User Then
            If Len(User) <> 0 Then
                pDB.ExecuteCommand "UPDATE " & DATABASE_TABLE_ACCOUNTS & " SET LastIP1 = '" & .Item(i).SubItems(INDEX_IP) & "' WHERE Name1 = '" & User & "'"

                With frmAccountPanel.lvAccounts.ListItems
                    For j = 1 To .Count
                        If .Item(j).SubItems(INDEX_NAME) = User Then
                            .Item(j).SubItems(INDEX_LAST_IP) = frmPanel.lvUsers.ListItems.Item(i).SubItems(INDEX_IP)
                            Exit For
                        End If
                    Next j
                End With
            End If

            .Remove i
            Exit For
        End If
    Next i
End With

frmChannel.LeaveAllChannels User

UPDATE_ONLINE
UPDATE_STATUS_BAR
End Sub

Public Sub SetupForms(NewForm As Form)
Dim pForm As Form

For Each pForm In Forms
    If Not pForm.Name = Name Then
        pForm.Hide
    End If
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
