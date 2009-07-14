Attribute VB_Name = "Module1"
Option Explicit

Public Sub SendMessage(Message As String)
Dim WinSk As Winsock

For Each WinSk In frmMain.Winsock1
    If WinSk.State = sckConnected Then
        WinSk.SendData Message
    End If
Next

End Sub

Public Sub SendRequest(Message As String, Wsk As Winsock)
If Wsk.State = 7 Then
    Wsk.SendData Message
End If
End Sub

Public Function UpdateUsersList() As Integer
Dim i As Integer
Dim tMsg As String
    frmMain.iMatch = True
    With frmPanel.ListView1.ListItems
        tMsg = "!listupdate#"
        For i = 1 To .Count
            tMsg = tMsg & .Item(i) & "#"
        Next i
    End With
    If tMsg <> "!listupdate#" Then
        SendMessage tMsg
    End If
End Function
