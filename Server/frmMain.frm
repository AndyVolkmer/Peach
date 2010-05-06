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
Dim MSG         As Long
Dim sFilter     As String

MSG = X / Screen.TwipsPerPixelX

Select Case MSG
    Case WM_LBUTTONDOWN
    Case WM_LBUTTONUP
        Vali = True
        frmMain.Show
        frmMain.WindowState = 0

    Case WM_LBUTTONDBLCLK
    Case WM_RBUTTONDOWN
    Case WM_RBUTTONUP
    Case WM_RBUTTONDBLCLK
End Select
End Sub

Private Sub MDIForm_Resize()
If Me.WindowState = 1 Then
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
InsertIntoRegistry "Server\Configuration", "Top", Me.Top
InsertIntoRegistry "Server\Configuration", "Left", Me.Left
End Sub

Private Sub Winsock1_Close(Index As Integer)
Dim i As Long
Dim j As Long
Dim User As String

User = GetAccountByIndex(Index)
Unload Winsock1(Index)

If LenB(User) <> 0 Then SendMessage User & " has gone offline."

With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = User Then
            If Len(User) <> 0 Then
                Call pDB.ExecuteCommand("UPDATE " & DATABASE_TABLE_ACCOUNTS & " SET LastIP1 = '" & .Item(i).SubItems(INDEX_IP) & "' WHERE Name1 = '" & User & "'")

                With frmAccountPanel.ListView1.ListItems
                    For j = 1 To .Count
                        If .Item(j).SubItems(INDEX_NAME) = User Then
                            .Item(j).SubItems(INDEX_LAST_IP) = frmPanel.ListView1.ListItems.Item(i).SubItems(INDEX_IP)
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
With frmPanel.ListView1.ListItems
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
Dim p_PreArray()    As String
Dim k               As Long
Dim i               As Long

Dim p_Message       As String
Dim p_MainArray()   As String   'Whole message string is saved here and split up by # sign
Dim p_Command       As String   'First part of main array ( always the command )
Dim IsMuted         As Boolean  'Mute explains itself

'Get Message
frmMain.Winsock1(Index).GetData p_Message
DoEvents

'Do first array to avoid incorrect package reading
'Chr(24) and Chr(25) are the seperators for each packet
p_PreArray = Split(p_Message, Chr(24) & Chr(25))

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

            Dim pSocket As Winsock

            For Each pSocket In frmMain.Winsock1
                With pSocket
                    If .State = 7 And Not .Index = Index Then
                        .SendData GetAccountByIndex(Index) & " has come online." & Chr(24) & Chr(25)
                        DoEvents
                    End If
                End With
            Next

        'Send Server information
        Case "!server_info"
            SendSingle "!server_info#" & GetServerInformation, Index

        Case "!login"
            With frmAccountPanel.ListView1.ListItems
                If .Count = 0 Then
                    SendSingle "!login#Account#", Index
                    Exit Sub
                End If

                For i = 1 To .Count
                    If LCase$(.Item(i).SubItems(INDEX_NAME)) = LCase$(p_MainArray(1)) Then
                        'Ban Check
                        If .Item(i).SubItems(INDEX_BANNED) = "1" Then
                            SendSingle "!login#Banned#", Index
                            Exit Sub
                        End If

                        'Password Check
                        If Not .Item(i).SubItems(INDEX_PASSWORD) = p_MainArray(2) Then
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
            End With

            With frmPanel.ListView1.ListItems
                'If the account is already beeing used kick first instance
                For i = 1 To .Count
                    If .Item(i) = p_MainArray(1) Then
                        Unload Winsock1(.Item(i).SubItems(INDEX_WINSOCK_ID))
                        .Remove (i)
                        Exit For
                    End If
                Next i

                .Item(.Count).Text = GetProperAccountName(p_MainArray(1))
            End With

            UPDATE_STATUS_BAR
            SendSingle "!accepted#" & GetProperAccountName(p_MainArray(1)) & "#", Index

        'We get ip request and send ip back
        Case "!iprequest"
            With frmPanel.ListView1.ListItems
                For i = 1 To .Count
                    If .Item(i) = p_MainArray(1) Then
                        SendSingle "!iprequest#" & .Item(i).SubItems(INDEX_IP) & "#", Index
                        Exit For
                    End If
                Next i
            End With

        Case "!message"
            Dim p_CHAT_ARRAY()      As String
            Dim p_TEXT_FIRST        As String
            Dim p_TEXT_SECOND       As String
            Dim IsCommand           As Boolean
            Dim IsSlash             As Boolean
            Dim properAccount       As String

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
            With frmPanel.ListView1.ListItems
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
                Dim Reason As String
                Dim m As Long
                Dim n As Long

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
                            SendSingle "Incorrect syntax, use the following format .show [accounts, onliners].", Index

                        End If

                    Case ".userinfo", ".uinfo"
                        GetUserInfo GetProperAccountName(p_TEXT_FIRST), p_CHAT_ARRAY(0), Index

                    Case ".accountinfo", ".accinfo", ".ainfo"
                        GetAccountInfo GetProperAccountName(p_TEXT_FIRST), p_CHAT_ARRAY(0), Index

                    Case ".kick"
                        With frmPanel.ListView1.ListItems
                            properAccount = GetProperAccountName(p_TEXT_FIRST)

                            For i = 1 To .Count
                                If .Item(i) = properAccount Then
                                    KickUser .Item(i)
                                    Exit For
                                Else
                                    If i = .Count Then
                                        If LenB(p_TEXT_FIRST) = 0 Then
                                            SendSingle "Incorrect syntax, use following format .kick [User].", Index
                                        Else
                                            SendSingle "User '" & p_TEXT_FIRST & "' not found.", Index
                                        End If
                                    End If
                                End If
                            Next i

                            properAccount = vbNullString
                        End With

                    Case ".ban"
                        Ban GetProperAccountName(p_TEXT_FIRST), GetAccountByIndex(Index), 1, Index, Trim$(Reason)

                    Case ".unban"
                        Ban GetProperAccountName(p_TEXT_FIRST), GetAccountByIndex(Index), 0, Index, Trim$(Reason)

                    Case ".mute"
                        MuteUser GetProperAccountName(p_TEXT_FIRST), GetAccountByIndex(Index), 1, Index, Trim$(Reason)

                    Case ".unmute"
                        MuteUser GetProperAccountName(p_TEXT_FIRST), GetAccountByIndex(Index), 0, Index, Trim$(Reason)

                    Case ".announce", ".ann", ".notify"
                        Dim p_ANN_MSG As String

                        'Capture announce message
                        If UBound(p_CHAT_ARRAY) > 0 Then
                            p_ANN_MSG = Mid$(p_MainArray(1), Len(p_CHAT_ARRAY(0)) + 2, Len(p_MainArray(1)))
                        End If

                        If LenB(p_ANN_MSG) = 0 Then
                            SendSingle "Incorrect syntax, use the following format " & p_CHAT_ARRAY(0) & " [Text].", Index
                        Else
                            Select Case LCase$(p_CHAT_ARRAY(0))
                                Case ".announce", ".ann"
                                    Reason = GetAccountByIndex(Index) '//Temp variable used to not load account twice
                                    SendMessage GetGMFlag(Reason) & GetAFKFlag(Reason) & "[" & Reason & " announces]: " & p_ANN_MSG

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
                                    With frmAccountPanel.ListView1.ListItems
                                        For i = 1 To .Count
                                            If .Item(i).SubItems(INDEX_NAME) = properAccount Then
                                                'Check if change name is already given out
                                                For n = 1 To .Count
                                                    If LCase$(.Item(n).SubItems(INDEX_NAME)) = LCase$(p_CHAT_ARRAY(3)) Then
                                                        SendSingle "Account '" & p_CHAT_ARRAY(3) & "' is already taken.", Index
                                                        Exit For
                                                    Else
                                                        If n = .Count Then
                                                            'Modify account in database and panel
                                                            frmAccountPanel.ModifyAccount p_CHAT_ARRAY(3), .Item(i).SubItems(INDEX_PASSWORD), .Item(i).SubItems(INDEX_BANNED), .Item(i).SubItems(INDEX_LEVEL), .Item(i), i, .Item(i).SubItems(INDEX_GENDER), .Item(i).SubItems(INDEX_EMAIL)

                                                            'Send feedback to person who changed name
                                                            SendSingle "Successfully renamed '" & properAccount & "' to '" & p_CHAT_ARRAY(3) & "'.", Index

                                                            'Check if player is online and directly rename it, also tell user that he got renamed
                                                            With frmPanel.ListView1.ListItems
                                                                For m = 1 To .Count
                                                                    If .Item(m) = properAccount Then
                                                                        .Item(m) = p_CHAT_ARRAY(3)

                                                                        SendSingle GetAccountByIndex(Index) & " renamed you to '" & p_CHAT_ARRAY(3) & "'.", .Item(m).SubItems(INDEX_WINSOCK_ID)
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

                                                            Call pDB.ExecuteCommand("UPDATE " & DATABASE_TABLE_FRIENDS & " SET Name = '" & p_CHAT_ARRAY(3) & "' WHERE Name = '" & properAccount & "'")
                                                            Call pDB.ExecuteCommand("UPDATE " & DATABASE_TABLE_FRIENDS & " SET Friend = '" & p_CHAT_ARRAY(3) & "' WHERE Friend = '" & properAccount & "'")

                                                            UPDATE_ONLINE
                                                        End If
                                                    End If
                                                Next n
                                                Exit For
                                            Else
                                                If i = .Count Then SendSingle "Account '" & properAccount & "' not found.", Index
                                            End If
                                        Next i
                                    End With

                                    properAccount = vbNullString
                                Else
                                    SendSingle "Incorrect syntax, use the following format " & p_CHAT_ARRAY(0) & Space(1) & p_TEXT_FIRST & " [Oldname] [Newname].", Index
                                End If

                            Case "level"
                                If UBound(p_CHAT_ARRAY) > 2 Then
                                    If Not IsNumeric(p_CHAT_ARRAY(3)) Then
                                        SendSingle "Level contains incorrect values, must be in range of 0-2.", Index
                                        Exit Sub
                                    End If

                                    If p_CHAT_ARRAY(3) > 2 Or p_CHAT_ARRAY(3) < 0 Then
                                        SendSingle "Level contains incorrect values, must be in range of 0-2", Index
                                        Exit Sub
                                    End If

                                    With frmAccountPanel.ListView1.ListItems
                                        properAccount = GetProperAccountName(p_TEXT_SECOND)

                                        For i = 1 To .Count
                                            If .Item(i).SubItems(INDEX_NAME) = properAccount Then
                                                'Modify level for account in database
                                                frmAccountPanel.ModifyAccount .Item(i).SubItems(INDEX_NAME), .Item(i).SubItems(INDEX_PASSWORD), .Item(i).SubItems(INDEX_BANNED), p_CHAT_ARRAY(3), .Item(i), i, .Item(i).SubItems(INDEX_GENDER), .Item(i).SubItems(INDEX_EMAIL)

                                                'Feedback to the person who modified the level
                                                SendSingle "Successfully changed level of '" & .Item(i).SubItems(INDEX_NAME) & "' to '" & p_CHAT_ARRAY(3) & "'.", Index

                                                With frmPanel.ListView1.ListItems
                                                    For n = 1 To .Count
                                                        If .Item(n) = properAccount Then
                                                            SendSingle GetAccountByIndex(Index) & " changed your level to '" & p_CHAT_ARRAY(3) & "'.", .Item(n).SubItems(INDEX_WINSOCK_ID)
                                                            Exit For
                                                        End If
                                                    Next n
                                                End With
                                                Exit For
                                            Else
                                                If i = .Count Then SendSingle "Account '" & p_TEXT_SECOND & "' not found.", Index
                                            End If
                                        Next i

                                        properAccount = vbNullString
                                    End With
                                Else
                                    SendSingle "Incorrect syntax, use the following format " & p_CHAT_ARRAY(0) & Space(1) & p_TEXT_FIRST & " [Name] [Level].", Index
                                End If

                            Case "gender"
                                Dim pGender As String
                                If UBound(p_CHAT_ARRAY) > 2 Then
                                    p_CHAT_ARRAY(3) = LCase$(p_CHAT_ARRAY(3))

                                    If p_CHAT_ARRAY(3) = "male" Then
                                        pGender = "0"

                                    ElseIf p_CHAT_ARRAY(3) = "female" Then
                                        pGender = "1"

                                    End If

                                    If LenB(pGender) = 0 Then
                                        SendSingle "Incorrect gender format use 'male' or 'female'.", Index
                                    Else
                                        With frmAccountPanel.ListView1.ListItems
                                            properAccount = GetProperAccountName(p_TEXT_SECOND)

                                            For i = 1 To .Count
                                                If .Item(i).SubItems(INDEX_NAME) = properAccount Then
                                                    'Modify gender in database and panel
                                                    frmAccountPanel.ModifyAccount .Item(i).SubItems(INDEX_NAME), .Item(i).SubItems(INDEX_PASSWORD), .Item(i).SubItems(INDEX_BANNED), .Item(i).SubItems(INDEX_LEVEL), .Item(i), i, pGender, .Item(i).SubItems(INDEX_EMAIL)

                                                    'Feedback to the person who modified the gender
                                                    SendSingle "Successfully changed gender of '" & properAccount & "' to '" & p_CHAT_ARRAY(3) & "'.", Index

                                                    'Notify user that gender got changed
                                                    With frmPanel.ListView1.ListItems
                                                        For n = 1 To .Count
                                                            If .Item(n) = properAccount Then
                                                                SendSingle GetAccountByIndex(Index) & " changed your gender to '" & p_CHAT_ARRAY(3) & "'.", .Item(n).SubItems(INDEX_WINSOCK_ID)
                                                                Exit For
                                                            End If
                                                        Next n
                                                    End With
                                                    Exit For
                                                Else
                                                    If i = .Count Then SendSingle "Account '" & p_TEXT_SECOND & "' not found.", Index
                                                End If
                                            Next i

                                            properAccount = vbNullString
                                        End With
                                    End If
                                Else
                                    SendSingle "Incorrect syntax, use the following format " & p_CHAT_ARRAY(0) & Space(1) & p_TEXT_FIRST & " [Name] [Gender].", Index
                                End If

                            Case "password"
                                If UBound(p_CHAT_ARRAY) > 2 Then
                                    With frmAccountPanel.ListView1.ListItems
                                        properAccount = GetProperAccountName(p_TEXT_SECOND)

                                        For i = 1 To .Count
                                            If .Item(i).SubItems(INDEX_NAME) = properAccount Then
                                                'Modify password in database and panel
                                                frmAccountPanel.ModifyAccount .Item(i).SubItems(INDEX_NAME), p_CHAT_ARRAY(3), .Item(i).SubItems(INDEX_BANNED), .Item(i).SubItems(INDEX_LEVEL), .Item(i), i, .Item(i).SubItems(INDEX_GENDER), .Item(i).SubItems(INDEX_EMAIL)

                                                'Feedback to the person who modified the password
                                                SendSingle "Successfully changed password of '" & properAccount & "' to '" & p_CHAT_ARRAY(3) & "'.", Index

                                                'Notify user that password got changed
                                                With frmPanel.ListView1.ListItems
                                                    For n = 1 To .Count
                                                        If .Item(n) = properAccount Then
                                                            SendSingle GetAccountByIndex(Index) & " changed your password to '" & p_CHAT_ARRAY(3) & "'.", .Item(n).SubItems(INDEX_WINSOCK_ID)
                                                            Exit For
                                                        End If
                                                    Next n
                                                End With
                                                Exit For
                                            Else
                                                If i = .Count Then SendSingle "Account '" & p_TEXT_SECOND & "' not found.", Index
                                            End If
                                        Next i

                                        properAccount = vbNullString
                                    End With
                                Else
                                    SendSingle "Incorrect syntax, use the following format " & p_CHAT_ARRAY(0) & Space(1) & p_TEXT_FIRST & " [Name] [Password].", Index
                                End If

                            Case "email"
                                If UBound(p_CHAT_ARRAY) > 2 Then
                                    With frmAccountPanel.ListView1.ListItems
                                        properAccount = GetProperAccountName(p_TEXT_SECOND)

                                        For i = 1 To .Count
                                            If .Item(i).SubItems(INDEX_NAME) = properAccount Then
                                                'Modify email in database and panel
                                                frmAccountPanel.ModifyAccount .Item(i).SubItems(INDEX_NAME), .Item(i).SubItems(INDEX_PASSWORD), .Item(i).SubItems(INDEX_BANNED), .Item(i).SubItems(INDEX_LEVEL), .Item(i), i, .Item(i).SubItems(INDEX_GENDER), p_CHAT_ARRAY(3)

                                                'Feedback to the person who modified the email
                                                SendSingle "Successfully changed email of '" & properAccount & "' to ' " & p_CHAT_ARRAY(3) & " '.", Index

                                                'Notify user that password got changed
                                                With frmPanel.ListView1.ListItems
                                                    For n = 1 To .Count
                                                        If .Item(n) = properAccount Then
                                                            SendSingle GetAccountByIndex(Index) & " changed your email to ' " & p_CHAT_ARRAY(3) & " '.", .Item(n).SubItems(INDEX_WINSOCK_ID)
                                                            Exit For
                                                        End If
                                                    Next n
                                                End With
                                                Exit For
                                            Else
                                                If i = .Count Then SendSingle "Account '" & p_TEXT_SECOND & "' not found.", Index
                                            End If
                                        Next i

                                        properAccount = vbNullString
                                    End With
                                Else
                                    SendSingle "Incorrect syntax, use the following format " & p_CHAT_ARRAY(0) & Space(1) & p_TEXT_FIRST & " [Name] [Email].", Index
                                End If

                            Case Else
                                SendSingle _
                                    vbCrLf & Space(2) & _
                                "Incorrect syntax, use the following format:" & _
                                    vbCrLf & Space(2) & _
                                p_CHAT_ARRAY(0) & " name [Oldname] [Newname]" & _
                                    vbCrLf & Space(2) & _
                                p_CHAT_ARRAY(0) & " level [Name] [Level]" & _
                                    vbCrLf & Space(2) & _
                                p_CHAT_ARRAY(0) & " gender [Name] [Gender]" & _
                                    vbCrLf & Space(2) & _
                                p_CHAT_ARRAY(0) & " password [Name] [Password]" & _
                                    vbCrLf & Space(2) & _
                                p_CHAT_ARRAY(0) & " email [Name] [Email]", Index

                        End Select

                    Case ".reload"
                        Dim loadTime As Long
                        With Database
                            Select Case LCase$(p_TEXT_FIRST)
                                Case LCase$(DATABASE_TABLE_ACCOUNTS), LCase$(DATABASE_TABLE_FRIENDS), LCase$(DATABASE_TABLE_IGNORES)
                                    SendSingle "This table can't be reloaded.", Index

                                Case LCase$(DATABASE_TABLE_COMMANDS)
                                    loadTime = timeGetTime
                                    Erase Commands
                                    LoadCommands
                                    SendMessage GetAccountByIndex(Index) & " initiated the reload of '" & DATABASE_TABLE_COMMANDS & "' table. ( " & timeGetTime - loadTime & "ms )"

                                Case LCase$(DATABASE_TABLE_DECLINED_NAMES)
                                    loadTime = timeGetTime
                                    Erase DeclinedNames
                                    LoadDeclinedNames
                                    SendMessage GetAccountByIndex(Index) & " initiated the reload of '" & DATABASE_TABLE_DECLINED_NAMES & "' table. ( " & timeGetTime - loadTime & "ms )"

                                Case LCase$(DATABASE_TABLE_EMOTES)
                                    loadTime = timeGetTime
                                    Erase Emotes
                                    LoadEmotes
                                    SendMessage GetAccountByIndex(Index) & " initiated the reload of '" & DATABASE_TABLE_EMOTES & "' table. ( " & timeGetTime - loadTime & "ms )"

                                Case LCase$("config"), LCase$("c")
                                    loadTime = timeGetTime
                                    LoadConfigValue
                                    SendMessage GetAccountByIndex(Index) & " iniated the reload of configuration files. ( " & timeGetTime - loadTime & "ms )"

                                Case Else
                                    If LenB(p_TEXT_FIRST) = 0 Then
                                        SendSingle "Incorrect syntax. Use the following format .reload Table.", Index
                                    Else
                                        SendSingle "This table does not exist.", Index
                                    End If

                            End Select
                        End With

                    Case ".clear"
                        If LenB(p_TEXT_FIRST) = 0 Then
                            SendSingle "Incorrect syntax. Use the following format .clear [User]", Index
                        Else
                            If LCase$(p_TEXT_FIRST) = "this" Or LCase$(p_TEXT_FIRST) = "me" Then
                                SendSingle "!clear#", Index
                                Exit Sub

                            ElseIf LCase$(p_TEXT_FIRST) = "all" Then
                                SendMessage "!clear#"
                                Exit Sub

                            End If

                            With frmPanel.ListView1.ListItems
                                properAccount = GetProperAccountName(p_TEXT_FIRST)

                                For i = 1 To .Count
                                    If .Item(i) = properAccount Then
                                        SendSingle "!clear#", .Item(i).SubItems(INDEX_WINSOCK_ID)
                                        Exit For
                                    Else
                                        If i = .Count Then SendSingle "User '" & p_TEXT_FIRST & "' was not found.", Index
                                    End If
                                Next i

                                properAccount = vbNullString
                            End With
                        End If

                    Case ".delete", ".del"
                        With frmAccountPanel.ListView1.ListItems
                            properAccount = GetProperAccountName(p_TEXT_FIRST)

                            For i = 1 To .Count
                                If .Item(i).SubItems(INDEX_NAME) = properAccount Then
                                    With frmPanel.ListView1.ListItems
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

                                    SendSingle "Successfully deleted account '" & properAccount & "' ID: " & .Item(i) & ".", Index
                                    .Remove i
                                    Exit For
                                Else
                                    If i = .Count Then
                                        If LenB(properAccount) = 0 Then
                                            SendSingle "Incorrect syntax. Use the following format .delete [Account]", Index
                                        Else
                                            SendSingle "User '" & p_TEXT_FIRST & "' was not found.", Index
                                        End If
                                    End If
                                End If
                            Next i

                            properAccount = vbNullString
                        End With

                    Case ".gm"
                        With frmPanel.ListView1.ListItems
                            Select Case LCase$(p_TEXT_FIRST)
                                Case "on"
                                    For i = 1 To .Count
                                        If .Item(i).SubItems(INDEX_WINSOCK_ID) = Index Then
                                            .Item(i).SubItems(INDEX_GM_FLAG) = "1"
                                            SendSingle "Enabled [GM] flag. Use .gm off to disable.", Index
                                            Exit For
                                        End If
                                    Next i

                                Case "off"
                                    For i = 1 To .Count
                                        If .Item(i).SubItems(INDEX_WINSOCK_ID) = Index Then
                                            .Item(i).SubItems(INDEX_GM_FLAG) = "0"
                                            SendSingle "Disabled [GM] flag. Use .gm on to enable.", Index
                                            Exit For
                                        End If
                                    Next i

                                Case Else
                                    SendSingle "Incorrect syntax. Use the following format .gm [on / off].", Index

                            End Select
                        End With

                    Case Else
                        SendSingle "Unknown command used. Check .help for more information about commands.", Index

                End Select
                Exit Sub
            End If

            'Check if user is muted
            If IsMuted Then
                SendSingle "You are muted.", Index
                Exit Sub
            End If

            If IsSlash Then
                Dim f       As Long
                Dim IsUser  As Boolean

                If Options.REPEAT_CHECK = 1 Then
                    If IsRepeating(GetAccountByIndex(Index), p_MainArray(1)) Then
                        SendSingle "Your message has triggered serverside flood protection. Please don't repeat yourself.", Index
                        Exit Sub
                    End If
                End If

                With frmPanel.ListView1.ListItems
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
                        Dim pRoll       As Long
                        Dim pMinRoll    As Long
                        Dim pMaxRoll    As Long

                        If UBound(p_CHAT_ARRAY) > 1 Then
                            If IsNumeric(p_CHAT_ARRAY(1)) Then
                                If p_CHAT_ARRAY(1) > MAX_INT_VALUE Or p_CHAT_ARRAY(1) < MIN_INT_VALUE Then
                                    pMinRoll = 1
                                Else
                                    pMinRoll = p_CHAT_ARRAY(1)
                                End If
                            Else
                                pMinRoll = 1
                            End If

                            If IsNumeric(p_CHAT_ARRAY(2)) Then
                                If p_CHAT_ARRAY(2) > MAX_INT_VALUE Or p_CHAT_ARRAY(2) < MIN_INT_VALUE Then
                                    pMaxRoll = 100
                                Else
                                    pMaxRoll = p_CHAT_ARRAY(2)
                                End If
                            Else
                                pMaxRoll = 100
                            End If

                        ElseIf UBound(p_CHAT_ARRAY) > 0 Then
                            pMinRoll = 1

                            If IsNumeric(p_CHAT_ARRAY(1)) Then
                                If p_CHAT_ARRAY(1) > MAX_INT_VALUE Or p_CHAT_ARRAY(1) < MIN_INT_VALUE Then
                                    pMaxRoll = 100
                                Else
                                    pMaxRoll = p_CHAT_ARRAY(1)
                                End If
                            Else
                                pMaxRoll = 100
                            End If

                        Else
                            pMinRoll = 1
                            pMaxRoll = 100

                        End If

                        If pMinRoll > pMaxRoll Then pMaxRoll = pMinRoll
                        If pMinRoll = 0 Then pMinRoll = 1
                        If pMaxRoll = 0 Then pMaxRoll = 100

                        pRoll = GetRandomNumber(pMinRoll, pMaxRoll)

                        SendProtectedMessage GetAccountByIndex(Index), GetAccountByIndex(Index) & " rolls " & pRoll & ". (" & pMinRoll & " - " & pMaxRoll & ")"

                    'Whisper X to Z from Y
                    Case "/w", "/whisper"
                        If IsUser Then
                            Dim pMes As String

                            If UBound(p_CHAT_ARRAY) > 1 Then
                                For i = 2 To UBound(p_CHAT_ARRAY)
                                    pMes = pMes & p_CHAT_ARRAY(i) & " "
                                Next i

                                Whisper GetAccountByIndex(Index), GetProperAccountName(p_TEXT_FIRST), Trim$(pMes), Index
                            End If
                        Else
                            If LenB(p_TEXT_FIRST) = 0 Then
                                SendSingle "Please use the following format /whisper [Name] [Text]", Index
                                Exit Sub
                            Else
                                SendSingle "No user named '" & p_TEXT_FIRST & "' is currently online.", Index
                            End If
                        End If

                    Case "/afk"
                        With frmPanel.ListView1.ListItems
                            For i = 1 To .Count
                                If .Item(i).SubItems(INDEX_WINSOCK_ID) = Index Then
                                    If .Item(i).SubItems(INDEX_AFK_FLAG) = "1" Then
                                        .Item(i).SubItems(INDEX_AFK_FLAG) = "0"
                                        SendSingle "You are not away from keyboard anymore.", Index
                                    Else
                                        .Item(i).SubItems(INDEX_AFK_FLAG) = "1"
                                        SendSingle "You are away from keyboard now.", Index
                                    End If
                                    Exit For
                                End If
                            Next i
                        End With

                    Case "/online"
                        SendSingle "You are online for " & Trim$(GetOnlineTime(GetAccountByIndex(Index))) & ".", Index

                    Case "/logout"
                        KickUser GetAccountByIndex(Index)

                    Case "/join"
                        If LenB(p_TEXT_FIRST) = 0 Then
                            SendSingle "Please enter a valid channel name.", Index
                        Else
                            With frmChannel.lvUsers.ListItems
                                If .Count = 0 Then
                                    frmChannel.JoinChannel p_TEXT_FIRST, GetAccountByIndex(Index)
                                Else
                                    For i = 1 To .Count
                                        If .Item(i) = GetAccountByIndex(Index) And LCase$(.Item(i).SubItems(CHANNEL_USER_CHANNEL)) = LCase$(p_TEXT_FIRST) Then
                                            SendSingle "You are already in '" & .Item(i).SubItems(CHANNEL_USER_CHANNEL) & "'.", Index
                                            Exit For
                                        Else
                                            If i = .Count Then frmChannel.JoinChannel p_TEXT_FIRST, GetAccountByIndex(Index)
                                        End If
                                    Next i
                                End If
                            End With
                        End If

                    Case "/leave"
                        If LenB(p_TEXT_FIRST) = 0 Then
                            SendSingle "Please enter a valid channel name.", Index
                        Else
                            frmChannel.LeaveChannel p_TEXT_FIRST, GetAccountByIndex(Index)
                        End If

                    Case Else
                        'Emotes
                        For i = LBound(Emotes) To UBound(Emotes)
                            If LCase$(Emotes(i).Command) = LCase$(p_CHAT_ARRAY(0)) Then
                                If GetAccountByIndex(Index) = GetProperAccountName(p_TEXT_FIRST) Then
                                    IsUser = False
                                End If

                                Dim pTemp As String
                                Dim Gen   As String
                                Dim j     As Long

                                With frmAccountPanel.ListView1.ListItems
                                    properAccount = GetAccountByIndex(Index)

                                    For j = 1 To .Count
                                        If .Item(j).SubItems(INDEX_NAME) = properAccount Then
                                            Select Case .Item(j).SubItems(INDEX_GENDER)
                                                Case "0": Gen = "his"
                                                Case "1": Gen = "her"
                                            End Select
                                            Exit For
                                        End If
                                    Next j

                                    properAccount = vbNullString
                                End With

                                If IsUser Then
                                    pTemp = Replace(Emotes(i).TargetEmote, "%u", GetAccountByIndex(Index))
                                    pTemp = Replace(pTemp, "%g", Gen)
                                    pTemp = Replace(pTemp, "%t", GetProperAccountName(p_TEXT_FIRST))
                                Else
                                    pTemp = Replace(Emotes(i).SingleEmote, "%u", GetAccountByIndex(Index))
                                    pTemp = Replace(pTemp, "%g", Gen)
                                End If

                                SendProtectedMessage GetAccountByIndex(Index), pTemp
                                SetLastMessage GetAccountByIndex(Index), p_MainArray(1)
                                Exit Sub
                            End If
                        Next i

                        Dim tempA   As String
                        Dim Channel As String

                        With frmChannel.lvChannels.ListItems
                            If .Count = 0 Then
                                If LenB(p_TEXT_FIRST) = 0 Then
                                    SendSingle "Please enter a valid channel name.", Index
                                Else
                                    SendSingle "You are not in channel '" & Right$(p_CHAT_ARRAY(0), Len(p_CHAT_ARRAY(0)) - 1) & "'.", Index
                                End If
                            Else
                                If LenB(p_TEXT_FIRST) = 0 Then
                                    SendSingle "Please enter a valid channel name.", Index
                                Else
                                    Channel = Right$(p_CHAT_ARRAY(0), Len(p_CHAT_ARRAY(0)) - 1)

                                    For i = 1 To .Count
                                        If LCase$(.Item(i)) = LCase$(Channel) Then
                                            With frmChannel.lvUsers.ListItems
                                                tempA = GetAccountByIndex(Index)

                                                For f = 1 To .Count
                                                    If .Item(f) = tempA And LCase$(.Item(f).SubItems(CHANNEL_USER_CHANNEL)) = LCase$(Channel) Then
                                                        SendMessageToChannel Channel, tempA, "[" & .Item(f).SubItems(CHANNEL_USER_CHANNEL) & "][" & tempA & "]: " & Right$(p_MainArray(1), Len(p_MainArray(1)) - Len(Channel) - 1)
                                                        Exit For
                                                    Else
                                                        If f = .Count Then SendSingle "You are not in channel '" & Channel & "'.", Index
                                                    End If
                                                Next f
                                            End With
                                            Exit Sub
                                        End If
                                    Next i
                                End If
                            End If
                        End With

                End Select
                SetLastMessage GetAccountByIndex(Index), p_MainArray(1)

                Exit Sub
            End If

            Dim S1  As Long
            Dim E   As String

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
                            SendSingle "Message blocked. Please do not write more then 75% in caps.", Index
                            Exit Sub
                        End If
                    End If
                End If
            End If

            If Options.REPEAT_CHECK = 1 Then
                If IsRepeating(GetAccountByIndex(Index), p_MainArray(1)) Then
                    SendSingle "Your message has triggered serverside flood protection. Please don't repeat yourself.", Index
                    Exit Sub
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

Private Sub SetLastMessage(pUser As String, pMessage As String)
Dim i As Long

With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = pUser Then
            .Item(i).SubItems(INDEX_LAST_MESSAGE) = pMessage
            Exit For
        End If
    Next i
End With
End Sub

Private Function IsRepeating(pUser As String, pMessage As String) As Boolean
Dim i As Long

With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = pUser Then
            If .Item(i).SubItems(INDEX_LAST_MESSAGE) = pMessage Then
                 IsRepeating = True
            End If
            Exit For
        End If
    Next i
End With
End Function

Private Function IsPartOf(pPart As String, pCommand As String) As Boolean
Dim pTemp  As String
Dim i       As Long

For i = 1 To Len(pPart)
    pTemp = Mid(pPart, 1, i)
    If LCase$(pTemp) = LCase$(Left$(pCommand, Len(pTemp))) Then
        IsPartOf = True
    Else
        IsPartOf = False
        Exit Function
    End If
Next i
End Function

Private Function GetOnlineTime(pUser As String) As String
Dim TD      As String
Dim TD1()   As String
Dim i       As Long
Dim j       As Long

With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = pUser Then
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

Private Function GetProperAccountName(pAccount As String) As String
Dim i As Long

With frmAccountPanel.ListView1.ListItems
    For i = 1 To .Count
        If LCase$(.Item(i).SubItems(INDEX_NAME)) = LCase$(pAccount) Then
            GetProperAccountName = .Item(i).SubItems(INDEX_NAME)
            Exit For
        Else
            If i = .Count Then GetProperAccountName = pAccount
        End If
    Next i
End With
End Function

Private Function GetServerInformation() As String
GetServerInformation = _
"Welcome to Peach Servers." & "#" & _
"Server: Peach r" & pRev & "/" & GetOS & "#" & _
"Online User: " & frmMain.Winsock1.Count - 1 & "#" & _
"Server Uptime: " & TimeSerial(0, 0, DateDiff("s", frmConfig.START_TIME, Time))
End Function

Private Sub Whisper(pUser As String, pTarget As String, pConv As String, Index As Integer)
Dim i As Long

'Check if user is whispering itself
If pUser = pTarget Then
    SendSingle "You can't whisper yourself.", Index
    Exit Sub
End If

With frmPanel.ListView1.ListItems
    'Search target in list and send message
    For i = 1 To .Count
        If .Item(i) = pTarget Then
            If IsIgnoring(.Item(i), pUser) Then
                SendSingle pTarget & " is ignoring you.", Index
            Else
                SendSingle "[You whisper to " & pTarget & "]: " & pConv, Index
                If LenB(GetAFKFlag(pTarget)) <> 0 Then SendSingle pTarget & " is away from keyboard.", Index
                SendSingle GetGMFlag(pUser) & "[" & pUser & " whispers]: " & pConv, .Item(i).SubItems(INDEX_WINSOCK_ID)
            End If
            Exit For
        End If
    Next i
End With
End Sub

Private Sub MuteUser(User As String, AdminName As String, IsMuted As Long, pIndex As Integer, Reason As String)
Dim i As Long

With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = User Then
            If IsMuted = 1 And .Item(i).SubItems(INDEX_MUTED) = "1" Then
                SendSingle User & " is already muted.", pIndex
                Exit Sub

            ElseIf IsMuted = 0 And .Item(i).SubItems(INDEX_MUTED) = "0" Then
                SendSingle User & " is not muted.", pIndex
                Exit Sub

            End If

            'Set flag in userlist
            .Item(i).SubItems(INDEX_MUTED) = IsMuted

            'Announce the action
            If LenB(Reason) = 0 Then
                If IsMuted = 1 Then
                    SendMessage User & " got muted by " & AdminName & "."
                Else
                    SendMessage User & " got unmuted by " & AdminName & "."
                End If
            Else
                If IsMuted = 1 Then
                    SendMessage User & " got muted by " & AdminName & ". (" & Reason & ")"
                Else
                    SendMessage User & " got unmuted by " & AdminName & ". (" & Reason & ")"
                End If
            End If
            Exit For
        Else
            If i = .Count Then
                If LenB(Trim$(User)) = 0 Then
                    If IsMuted = 1 Then
                        SendSingle "Incorrect syntax, use the following format .mute [User] [Reason].", pIndex
                    Else
                        SendSingle "Incorrect syntax, use the following format .unmute [User] [Reason].", pIndex
                    End If
                Else
                    SendSingle "User '" & User & "' was not found.", pIndex
                End If
            End If
        End If
    Next i
End With
End Sub

Private Function GetAccountList() As String
Dim i As Long

With frmAccountPanel.ListView1.ListItems
    GetAccountList = "Account List:#"
    For i = 1 To .Count
        GetAccountList = GetAccountList & .Item(i).SubItems(INDEX_NAME) & "#"
    Next i
End With
End Function

Private Function GetOnlineList() As String
Dim i As Long

With frmPanel.ListView1.ListItems
    GetOnlineList = "User List:#"
    For i = 1 To .Count
        GetOnlineList = GetOnlineList & .Item(i) & "#"
    Next i
End With
End Function

Private Sub Ban(Account As String, AdminName As String, Ban As Long, pIndex As Integer, Reason As String)
Dim i As Long

With frmAccountPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(INDEX_NAME) = Account Then

            'If the account is already banned send feedback
            If Ban = 1 Then
                If .Item(i).SubItems(INDEX_BANNED) = "1" Then
                    SendSingle "Account '" & Account & "' is already banned.", pIndex
                    Exit Sub
                End If
            Else
                If .Item(i).SubItems(INDEX_BANNED) = "0" Then
                    SendSingle "Account '" & Account & "' is not banned.", pIndex
                    Exit Sub
                End If
            End If

            'Ban account in database
            frmAccountPanel.ModifyAccount Account, .Item(i).SubItems(INDEX_PASSWORD), Ban, .Item(i).SubItems(INDEX_LEVEL), .Item(i), .Item(i).Index, .Item(i).SubItems(INDEX_GENDER), .Item(i).SubItems(INDEX_EMAIL)

            'Announce the action
            If LenB(Reason) = 0 Then
                If Ban = 1 Then
                    SendMessage Account & " was account banned by " & AdminName & "."
                Else
                    SendMessage Account & " was account unbanned by " & AdminName & "."
                End If
            Else
                If Ban = 1 Then
                    SendMessage Account & " was account banned by " & AdminName & ". (" & Reason & ")"
                Else
                    SendMessage Account & " was account unbanned by " & AdminName & ". (" & Reason & ")"
                End If
            End If
            Exit For
        Else
            If i = .Count Then
                If LenB(Trim$(Account)) = 0 Then
                    If Ban = 1 Then
                        SendSingle "Incorrect syntax, use the following format .ban [Name] [Reason].", pIndex
                    Else
                        SendSingle "Incorrect syntax, use the following format .unban [Name] [Reason].", pIndex
                    End If
                Else
                    SendSingle "User '" & Account & "' not found.", pIndex
                End If
            End If
        End If
    Next i
End With
End Sub

Private Sub KickUser(pUser As String)
Dim i As Long

With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = pUser Then
            Unload frmMain.Winsock1(.Item(i).SubItems(INDEX_WINSOCK_ID))

            SendMessage .Item(i) & " has gone offline."

            frmChannel.LeaveAllChannels pUser

            .Remove (i)

            UPDATE_ONLINE
            UPDATE_STATUS_BAR
            Exit For
        End If
    Next i
End With
End Sub

Private Sub GetAccountInfo(Account As String, pUsedSyntax As String, pIndex As Integer)
Dim i As Long

With frmAccountPanel.ListView1.ListItems
    For i = 1 To .Count
        If LCase(.Item(i).SubItems(INDEX_NAME)) = LCase(Account) Then
            SendSingle vbCrLf & " Account information about '" & Account & "'" & vbCrLf & " ID: " & .Item(i) & vbCrLf & " Password: " & .Item(i).SubItems(INDEX_PASSWORD) & vbCrLf & " Registration Time: " & .Item(i).SubItems(INDEX_TIME) & vbCrLf & " Registration Date: " & .Item(i).SubItems(INDEX_DATE) & vbCrLf & " Banned: " & .Item(i).SubItems(INDEX_BANNED) & vbCrLf & " Level: " & .Item(i).SubItems(INDEX_LEVEL) & vbCrLf & " Gender: " & .Item(i).SubItems(INDEX_GENDER) & vbCrLf & " Email: " & .Item(i).SubItems(INDEX_EMAIL) & vbCrLf & " Last IP:  " & .Item(i).SubItems(INDEX_LAST_IP), pIndex
            Exit For
        Else
            If i = .Count Then
                If LenB(Trim$(Account)) = 0 Then
                    SendSingle "Incorrect syntax, use following format " & pUsedSyntax & " [Account].", pIndex
                Else
                    SendSingle "Account '" & Account & "' not found.", pIndex
                End If
            End If
        End If
    Next i
End With
End Sub

Private Sub GetUserInfo(pUser As String, pUsedSyntax As String, pIndex As Integer)
Dim i As Long

With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = pUser Then
            SendSingle vbCrLf & "User information about '" & pUser & "'" & vbCrLf & " IP : " & .Item(i).SubItems(INDEX_IP) & vbCrLf & " Winsock ID: " & .Item(i).SubItems(INDEX_WINSOCK_ID) & vbCrLf & " Last Message: " & .Item(i).SubItems(INDEX_LAST_MESSAGE) & vbCrLf & " Muted: " & .Item(i).SubItems(INDEX_MUTED) & vbCrLf & " Login Time: " & .Item(i).SubItems(INDEX_LOGIN_TIME) & " GM Flag: " & .Item(i).SubItems(INDEX_GM_FLAG) & " AFK Flag: " & .Item(i).SubItems(INDEX_AFK_FLAG), pIndex
            Exit For
        Else
            If i = .Count Then
                If LenB(pUser) = 0 Then
                    SendSingle "Incorrect syntax, use following format " & pUsedSyntax & " [User].", pIndex
                Else
                    SendSingle "User '" & pUser & "' was not found.", pIndex
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

Private Function GetLevel(pName As String) As Long
Dim i As Long

With frmAccountPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(INDEX_NAME) = pName Then
            GetLevel = .Item(i).SubItems(INDEX_LEVEL)
            Exit For
        End If
    Next i
End With
End Function

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim i As Long
Dim j As Long
Dim User As String

User = GetAccountByIndex(Index)
Unload Winsock1(Index)

If LenB(User) <> 0 Then SendMessage User & " has gone offline."

With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = User Then
            If Len(User) <> 0 Then
                Call pDB.ExecuteCommand("UPDATE " & DATABASE_TABLE_ACCOUNTS & " SET LastIP1 = '" & .Item(i).SubItems(INDEX_IP) & "' WHERE Name1 = '" & User & "'")

                With frmAccountPanel.ListView1.ListItems
                    For j = 1 To .Count
                        If .Item(j).SubItems(INDEX_NAME) = User Then
                            .Item(j).SubItems(INDEX_LAST_IP) = frmPanel.ListView1.ListItems.Item(i).SubItems(INDEX_IP)
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

Public Sub SetupForms(pNewForm As Form)
Dim pForm As Form

For Each pForm In Forms
    If Not pForm.Name = frmMain.Name Then
        pForm.Hide
    End If
Next
pNewForm.Show
End Sub

Private Function GetAccountByIndex(pIndex As Integer) As String
Dim i As Long

With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(INDEX_WINSOCK_ID) = pIndex Then
            GetAccountByIndex = .Item(i)
            Exit For
        End If
    Next i
End With
End Function

Private Function GetGMFlag(uName As String) As String
Dim i As Long

With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = uName Then
            If .Item(i).SubItems(INDEX_GM_FLAG) = "1" Then GetGMFlag = "<GM>"
            Exit For
        End If
    Next i
End With
End Function

Private Function GetAFKFlag(uName As String) As String
Dim i As Long

With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = uName Then
            If .Item(i).SubItems(INDEX_AFK_FLAG) = "1" Then GetAFKFlag = "<AFK>"
            Exit For
        End If
    Next i
End With
End Function
