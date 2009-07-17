Attribute VB_Name = "CodeModule"
Option Explicit
Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal binvert As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long

Public Sub SendMessage(iMessage As String)

If frmMain.Winsock1(0).State = 7 Then
    frmMain.Winsock1(0).SendData iMessage
Else
    MsgBox "Not connected!", vbInformation
End If

End Sub

Public Sub FlashTitle(Handle As Long, ReturnOrig As Boolean)
    Call FlashWindow(Handle, ReturnOrig)
End Sub
