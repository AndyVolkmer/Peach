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
