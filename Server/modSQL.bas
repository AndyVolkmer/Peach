Attribute VB_Name = "modSQL"
Option Explicit

Public pConnection      As New ADODB.Connection
Public pCommand         As New ADODB.Command
Public pRecordSet       As New ADODB.Recordset

Public Function ConnectMySQL(pDatabase As String, pUser As String, pPassword As String, pIP As String) As Boolean
On Error GoTo HandleErrorConnection

Set pConnection = New ADODB.Connection
pConnection.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
    & "SERVER=" & pIP & ";" _
    & "DATABASE=" & pDatabase & ";" _
    & "UID=" & pUser & ";" _
    & "PWD=" & pPassword & ";" _
    & "OPTION=18475" '-> 1 + 2 + 8 + 32 + 2048 + 16384

pConnection.Open

Set pCommand = New ADODB.Command
Set pCommand.ActiveConnection = pConnection

pCommand.CommandType = adCmdText
WriteLog "Connected with Database"

ConnectMySQL = True

Exit Function
HandleErrorConnection:
'Reset database variable
WriteIniValue App.Path & "\peachConfig.conf", "Database", "Name", vbNullString

'Print error
WriteLog Err.Description

'Set error flag
ConnectMySQL = False
End Function

Public Function LoadCommands() As Boolean
Dim SQL     As String
Dim Counter As Long

SQL = "SELECT * FROM " & DATABASE_TABLE_COMMANDS

On Error GoTo HandleErrorEmotes

pCommand.CommandText = SQL
Set pRecordSet = pCommand.Execute

ReDim Commands(0)

With pRecordSet
    Do Until .EOF
        ReDim Preserve Commands(Counter + 1)
        Commands(Counter).Syntax = !Syntax
        Commands(Counter).Description = !Description
        Counter = Counter + 1
        .MoveNext
    Loop
End With

Set pRecordSet = Nothing

WriteLog "Loaded " & Counter & " command(s)."

LoadCommands = True

Exit Function
HandleErrorEmotes:
'Print error
WriteLog Err.Description & "."

LoadCommands = False
End Function

Public Function LoadDeclinedNames() As Boolean
Dim SQL     As String
Dim Counter As Long

SQL = "SELECT * FROM " & DATABASE_TABLE_DECLINED_NAMES
Counter = 0

On Error GoTo HandleErrorEmotes

pCommand.CommandText = SQL
Set pRecordSet = pCommand.Execute

ReDim DeclinedNames(0)

With pRecordSet
    Do Until .EOF
        ReDim Preserve DeclinedNames(Counter + 1)
        DeclinedNames(Counter) = !Name
        Counter = Counter + 1
        .MoveNext
    Loop
End With

Set pRecordSet = Nothing

WriteLog "Loaded " & Counter & " declined name(s)."

LoadDeclinedNames = True
Exit Function
HandleErrorEmotes:
'Print error
WriteLog Err.Description & "."

LoadDeclinedNames = False
End Function

Public Function LoadEmotes() As Boolean
Dim SQL     As String
Dim Counter As Long

SQL = "SELECT * FROM " & DATABASE_TABLE_EMOTES
Counter = 0

On Error GoTo HandleErrorEmotes

pCommand.CommandText = SQL
Set pRecordSet = pCommand.Execute

ReDim Emotes(0)

With pRecordSet
    Do Until .EOF
        ReDim Preserve Emotes(Counter + 1)
        Emotes(Counter).Command = !Command
        Emotes(Counter).SingleEmote = !single_emote
        Emotes(Counter).TargetEmote = !target_emote
        .MoveNext
        Counter = Counter + 1
    Loop
End With

Set pRecordSet = Nothing

WriteLog "Loaded " & Counter & " emote(s)."

LoadEmotes = True
Exit Function
HandleErrorEmotes:
'Print error
WriteLog Err.Description

LoadEmotes = False
End Function

Public Function LoadFriends() As Boolean
Dim SQL     As String
Dim LItem   As ListItem
Dim Counter As Long

SQL = "SELECT * FROM " & DATABASE_TABLE_FRIENDS
Counter = 0

On Error GoTo HandleErrorFriends

pCommand.CommandText = SQL
Set pRecordSet = pCommand.Execute

With pRecordSet
    Do Until .EOF
        Set LItem = frmFriendIgnoreList.ListView1.ListItems.Add(, , !ID)
        LItem.SubItems(1) = !Name
        LItem.SubItems(2) = !Friend
        .MoveNext
        Counter = Counter + 1
    Loop
End With

Set LItem = Nothing
Set pRecordSet = Nothing

WriteLog "Loaded " & Counter & " relation(s)."

LoadFriends = True
Exit Function
HandleErrorFriends:
'Print error
WriteLog Err.Description

LoadFriends = False
End Function

Public Function LoadIgnores() As Boolean
Dim SQL     As String
Dim LItem   As ListItem
Dim Counter As Long

SQL = "SELECT * FROM " & DATABASE_TABLE_IGNORES
Counter = 0

On Error GoTo HandleErrorFriends
pCommand.CommandText = SQL
Set pRecordSet = pCommand.Execute

With pRecordSet
    Do Until .EOF
        Set LItem = frmFriendIgnoreList.ListView2.ListItems.Add(, , !ID)
        LItem.SubItems(1) = !Name
        LItem.SubItems(2) = !IgnoredName
        .MoveNext
        Counter = Counter + 1
    Loop
End With

Set LItem = Nothing
Set pRecordSet = Nothing

WriteLog "Loaded " & Counter & " ignore(s)."

LoadIgnores = True
Exit Function
HandleErrorFriends:
'Print error
WriteLog Err.Description

LoadIgnores = False
End Function

Public Function LoadAccounts() As Boolean
Dim SQL     As String
Dim LItem   As ListItem
Dim Counter As Long

SQL = "SELECT * FROM " & DATABASE_TABLE_ACCOUNTS
Counter = 0

On Error GoTo HandleErrorTable
pCommand.CommandText = SQL
Set pRecordSet = pCommand.Execute

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
        .AddItem "Male"
        .AddItem "Female"
    End With
End With

With pRecordSet
    Do Until .EOF
        Set LItem = frmAccountPanel.ListView1.ListItems.Add(, , !ID)
        LItem.SubItems(1) = !Name1
        LItem.SubItems(2) = !Password1
        LItem.SubItems(3) = Format$(!Time1, "hh:nn:ss")
        LItem.SubItems(4) = !Date1
        LItem.SubItems(5) = !Banned1
        LItem.SubItems(6) = !Level1
        LItem.SubItems(7) = !SecretQuestion1
        LItem.SubItems(8) = !SecretAnswer1
        LItem.SubItems(9) = !Gender1
        .MoveNext
        Counter = Counter + 1
    Loop
End With

Set LItem = Nothing
Set pRecordSet = Nothing

WriteLog "Loaded " & Counter & " account(s)."

LoadAccounts = True
Exit Function
HandleErrorTable:
'Print error
WriteLog Err.Description

LoadAccounts = False
End Function

Public Function ExecuteCommand(pShell As String) As Boolean
On Error GoTo hHE

With pCommand
    .CommandText = pShell
    .Execute
End With

ExecuteCommand = True

Exit Function

hHE:
MsgBox Err.Description, vbInformation, Err.Number
ExecuteCommand = False
End Function
