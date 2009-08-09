Attribute VB_Name = "FileModule"
Option Explicit

'//File Module
Public Type TypeSSO
    TopPos          As String * 6
    LeftPos         As String * 6
    Nickname        As String * 15
    ConnectIP       As String * 20
    Port            As String * 5
    Account         As String * 15
    Password        As String * 15
End Type

Public Type TypeSSO2
    ValidateN       As String * 1
    Language        As String * 1
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
    .Account = frmConfig.txtAccount.Text
    .Password = frmConfig.txtPassword.Text
End With

Put #lnManip, 1, theOath

Close #lnManip

WriteConfigFile = True

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

Public Function WriteConfigFile2(PathToFile2 As String) As Boolean

Dim lnManip2 As Integer
Dim theOath2 As TypeSSO2

If Trim(PathToFile2) = "" Then
    WriteConfigFile2 = False
    Exit Function
End If

lnManip2 = FreeFile

Open PathToFile2 For Random As lnManip2 Len = Len(theOath2)

With theOath2
    .ValidateN = "0"
    .Language = frmLanguage.Combo1.ListIndex
End With

Put #lnManip2, 1, theOath2

Close #lnManip2

WriteConfigFile2 = True

End Function

Public Function ReadConfigFile2(PathToFile2 As String) As TypeSSO2
    Dim lnManip2 As Integer
    Dim theOath2 As TypeSSO2
    Dim lnRegistro2 As Integer

On Error GoTo Error_Function:

lnManip2 = FreeFile

Open PathToFile2 For Random As lnManip2 Len = Len(theOath2)

lnRegistro2 = 1

Get #lnManip2, lnRegistro2, theOath2
Close #lnManip2
ReadConfigFile2 = theOath2
'**********************************************************
Exit Function
Error_Function:
    On Local Error Resume Next
    Close #lnManip2
End Function
