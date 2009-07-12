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

'// Old way
'Dim i As Integer
'' Send to clients
'    For i = 1 To frmMain.Winsock1.ubound - 1
'        ' Check if connected
'        If frmMain.Winsock1(i).State = sckConnected Then
'            frmMain.Winsock1(i).SendData Message
'        End If
'    Next i
