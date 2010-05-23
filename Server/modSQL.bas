Attribute VB_Name = "modSQL"
Option Explicit

Public Function LoadCommands() As Boolean
Dim SQL     As String
Dim Counter As Long

With pDB

SQL = "SELECT * FROM " & DATABASE_TABLE_COMMANDS

On Error GoTo HandleErrorEmotes

.pCommand.CommandText = SQL
Set .pRecordSet = .pCommand.Execute

ReDim Commands(0)

With .pRecordSet
    Do Until .EOF
        ReDim Preserve Commands(Counter + 1)
        Commands(Counter).Syntax = !Syntax
        Commands(Counter).Description = !Description
        Counter = Counter + 1
        .MoveNext
    Loop
End With

Set .pRecordSet = Nothing

End With

WriteLog "Loaded " & Counter & " command(s)."

LoadCommands = True

Exit Function
HandleErrorEmotes:
    WriteLog Err.Description & "."
End Function

Public Function LoadDeclinedNames() As Boolean
Dim SQL     As String
Dim Counter As Long

With pDB

SQL = "SELECT * FROM " & DATABASE_TABLE_DECLINED_NAMES

On Error GoTo HandleErrorEmotes

.pCommand.CommandText = SQL
Set .pRecordSet = .pCommand.Execute

ReDim DeclinedNames(0)

With .pRecordSet
    Do Until .EOF
        ReDim Preserve DeclinedNames(Counter + 1)
        DeclinedNames(Counter) = !Name
        Counter = Counter + 1
        .MoveNext
    Loop
End With

Set .pRecordSet = Nothing

End With

WriteLog "Loaded " & Counter & " declined name(s)."

LoadDeclinedNames = True

Exit Function
HandleErrorEmotes:
    WriteLog Err.Description & "."
End Function

Public Function LoadEmotes() As Boolean
Dim SQL     As String
Dim Counter As Long

With pDB

SQL = "SELECT * FROM " & DATABASE_TABLE_EMOTES

On Error GoTo HandleErrorEmotes

.pCommand.CommandText = SQL
Set .pRecordSet = .pCommand.Execute

ReDim Emotes(0)

With .pRecordSet
    Do Until .EOF
        ReDim Preserve Emotes(Counter + 1)
        Emotes(Counter).Command = !Command
        Emotes(Counter).SingleEmote = !single_emote
        Emotes(Counter).TargetEmote = !target_emote
        .MoveNext
        Counter = Counter + 1
    Loop
End With

Set .pRecordSet = Nothing

End With

WriteLog "Loaded " & Counter & " emote(s)."

LoadEmotes = True

Exit Function
HandleErrorEmotes:
    WriteLog Err.Description
End Function

Public Function LoadFriends() As Boolean
Dim SQL     As String
Dim LItem   As ListItem
Dim Counter As Long

With pDB

SQL = "SELECT * FROM " & DATABASE_TABLE_FRIENDS

On Error GoTo HandleErrorFriends

.pCommand.CommandText = SQL
Set .pRecordSet = .pCommand.Execute

With .pRecordSet
    Do Until .EOF
        Set LItem = frmFriendIgnoreList.ListView1.ListItems.Add(, , !ID)
        LItem.SubItems(1) = !Name
        LItem.SubItems(2) = !Friend
        .MoveNext
        Counter = Counter + 1
    Loop
End With

Set LItem = Nothing
Set .pRecordSet = Nothing

End With

WriteLog "Loaded " & Counter & " relation(s)."

LoadFriends = True

Exit Function
HandleErrorFriends:
    WriteLog Err.Description
End Function

Public Function LoadIgnores() As Boolean
Dim SQL     As String
Dim LItem   As ListItem
Dim Counter As Long

With pDB

SQL = "SELECT * FROM " & DATABASE_TABLE_IGNORES

On Error GoTo HandleErrorFriends

.pCommand.CommandText = SQL
Set .pRecordSet = .pCommand.Execute

With .pRecordSet
    Do Until .EOF
        Set LItem = frmFriendIgnoreList.ListView2.ListItems.Add(, , !ID)
        LItem.SubItems(1) = !Name
        LItem.SubItems(2) = !IgnoredName
        .MoveNext
        Counter = Counter + 1
    Loop
End With

Set LItem = Nothing
Set .pRecordSet = Nothing

End With

WriteLog "Loaded " & Counter & " ignore(s)."

LoadIgnores = True

Exit Function
HandleErrorFriends:
    WriteLog Err.Description
End Function

Public Function LoadAccounts() As Boolean
Dim SQL     As String
Dim LItem   As ListItem
Dim Counter As Long

With pDB

SQL = "SELECT * FROM " & DATABASE_TABLE_ACCOUNTS

On Error GoTo HandleErrorTable

.pCommand.CommandText = SQL
Set .pRecordSet = .pCommand.Execute

With frmAccountPanel
    With .cmbBanned
        .AddItem "0"
        .AddItem "1"
    End With
    With .cmbLevel
        .AddItem "0"
        .AddItem "1"
        .AddItem "2"
    End With
    With .cmbGender
        .AddItem "0"
        .AddItem "1"
    End With
End With

With .pRecordSet
    Do Until .EOF
        Set LItem = frmAccountPanel.lvAccounts.ListItems.Add(, , !ID)
        LItem.SubItems(INDEX_NAME) = !Name1
        LItem.SubItems(INDEX_PASSWORD) = !Password1
        LItem.SubItems(INDEX_TIME) = Format$(!Time1, "hh:nn:ss")
        LItem.SubItems(INDEX_DATE) = !Date1
        LItem.SubItems(INDEX_BANNED) = !Banned1
        LItem.SubItems(INDEX_LEVEL) = !Level1
        LItem.SubItems(INDEX_SECRET_QUESTION) = !SecretQuestion1
        LItem.SubItems(INDEX_SECRET_ANSWER) = !SecretAnswer1
        LItem.SubItems(INDEX_GENDER) = !Gender1
        LItem.SubItems(INDEX_EMAIL) = !Email1
        LItem.SubItems(INDEX_LAST_IP) = !LastIP1
        .MoveNext
        Counter = Counter + 1
    Loop
End With

Set LItem = Nothing
Set .pRecordSet = Nothing

End With

WriteLog "Loaded " & Counter & " account(s)."

LoadAccounts = True

Exit Function
HandleErrorTable:
    WriteLog Err.Description
End Function
