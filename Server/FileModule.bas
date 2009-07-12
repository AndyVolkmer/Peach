Attribute VB_Name = "FileModule"
Option Explicit

'//File Module
Public Type TypeSSO
    TopPos          As String * 6
    LeftPos         As String * 6
    Form1Realm      As String * 25
    Language        As Integer
    Expansion       As Integer
    B2Top           As String * 6
    B2Left          As String * 6
    B2Height        As String * 6
    B2Width         As String * 6
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
    .LeftPos = Form1.Left
    .TopPos = Form1.Top
    .Form1Realm = Form1.Text1.Text
    .Language = Form1.Combo1.ListIndex
    .Expansion = Form1.Combo2.ListIndex
    .B2Height = Form1.txtHeight
    .B2Width = Form1.txtWidth
    .B2Top = Form1.txtTop
    .B2Left = Form1.txtLeft
End With

Put #lnManip, 1, theOath

Close #lnManip

WriteConfigFile = True
Debug.Print "Succesfully Saved!" & " " & Time & " - " & Date

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
