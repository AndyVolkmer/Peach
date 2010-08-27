VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Documentation:
'-------------
'1. Use of recordset in order to load tables
'
'   pCommand.CommandText = "SELECT * / UPDATE / INSERT", the command you want to execute here
'   Set pRecordSet = pCommand.Execute, apply all the data which gets returned from the .Execute command ( each row in .EOF )
'
'   Example use: // Loop through rows to get data, when EOF finishes there are no more rows
'       Do Until pRecordSet.EOF
'           Variable(1) = pRecordSet!ColumnName1
'           Variable(2) = pRecordSet!ColumnName2
'       Loop
'
'   Now free up memory with:
'   Set pRecordSet = Nothing

Option Explicit

Public pConnection    As New ADODB.Connection
Public pCommand       As New ADODB.Command
Public pRecordSet     As New ADODB.Recordset

Public Function ConnectDatabase(pType As String, pHost As String, pDatabase As String, pUser As String, pPassword As String, pDatabaseFile As String) As String
On Error GoTo hCD
Select Case pType
    Case "0" 'Access Database
        If InStr(1, pDatabaseFile, ".") = 0 Then
            pDatabaseFile = pDatabaseFile & ".mdb"
        End If

        pConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" _
            & "Data Source=" _
            & App.Path & "\" _
            & pDatabaseFile

    Case "1" 'MySQL Database 3.51
        pConnection.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
            & "SERVER=" & pHost & ";" _
            & "DATABASE=" & pDatabase & ";" _
            & "UID=" & pUser & ";" _
            & "PWD=" & pPassword & ";" _
            & "OPTION=18475" '-> 1 + 2 + 8 + 32 + 2048 + 16384

    Case Else 'Other types database types
        '
End Select

pConnection.Open

Set pCommand = New ADODB.Command
Set pCommand.ActiveConnection = pConnection

pCommand.CommandType = adCmdText

ConnectDatabase = vbNullString
Exit Function
hCD:
ConnectDatabase = Err.Description
End Function

Public Sub CloseDatabase()
Set pCommand = Nothing
Set pRecordSet = Nothing

pConnection.Close

Set pConnection = Nothing
End Sub

Public Sub ExecuteCommand(pShell As String)
With pCommand
    .CommandText = pShell
    .Execute
End With

Exit Sub
End Sub

Public Function GetMaxID(RowName As String, TableName As String) As Long
pCommand.CommandText = "SELECT MAX(" & RowName & ") AS MaxID FROM " & TableName
Set pRecordSet = pCommand.Execute
DoEvents

If pRecordSet.EOF Or IsNull(pRecordSet!MaxID) Then
    GetMaxID = 1
Else
    GetMaxID = pRecordSet!MaxID + 1
End If

Set pRecordSet = Nothing
End Function
