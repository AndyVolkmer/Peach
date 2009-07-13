Attribute VB_Name = "FileModule"
Option Explicit

'//File Module
Public Type TypeSSO
    TopPos          As String * 6
    LeftPos         As String * 6
    Nickname        As String * 15
    ConnectIP       As String * 20
    Port            As String * 5
End Type

Public Function WriteConfigFile(PathToFile As String) As Boolean

Dim lnManip As Integer
Dim theOath As TypeSSO

If Trim(PathToFile) = "" Then
    WriteConfigFile = False
    Exit Function
End If

lnManip = FreeFile

Open PathToFile For Random As lnManip Len = Len(theOath)

With theOath
    .LeftPos = frmMain.Left
    .TopPos = frmMain.Top
    .Nickname = frmConfig.txtNick.Text
    .ConnectIP = frmConfig.txtIP.Text
    .Port = frmConfig.txtPort.Text
End With

Put #lnManip, 1, theOath

Close #lnManip

WriteConfigFile = True
Debug.Print "Succesfully Saved!" & " " & Format(Time, "hh:nn:ss") & " - " & Date

End Function

Public Function ReadConfigFile(PathToFile As String) As TypeSSO
    Dim lnManip As Integer
    Dim theOath As TypeSSO
    Dim lnRegistro As Integer

On Error GoTo Error_Function:

lnManip = FreeFile

Open PathToFile For Random As lnManip Len = Len(theOath)

lnRegistro = 1

Get #lnManip, lnRegistro, theOath
Close #lnManip
ReadConfigFile = theOath
'**********************************************************
Exit Function
Error_Function:
    On Local Error Resume Next
    Close #lnManip
End Function
