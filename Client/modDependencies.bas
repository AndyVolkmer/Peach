Attribute VB_Name = "modDependencies"
Option Explicit

Public Function RegisterDependency() As String
Dim Path As String
    Path = GetWindowsDir & "\system32\"

'Unregister
If FileExists(Path & "COMDLG32.OCX") Then
    Shell "regsvr32 /u /s COMDLG32.OCX"
Else
    RegisterDependency = "Dependency COMDLG32.OCX missing."
    Exit Function
End If

If FileExists(Path & "RICHTX32.OCX") Then
    Shell "regsvr32 /u /s RICHTX32.OCX"
Else
    RegisterDependency = "Dependency RICHTX32.OCX missing."
    Exit Function
End If

If FileExists(Path & "RICHTX32.OCX") Then
    Shell "regsvr32 /u /s RICHTX32.OCX"
Else
    RegisterDependency = "Dependency RICHTX32.OCX missing."
    Exit Function
End If

If FileExists(Path & "TABCTL32.OCX") Then
    Shell "regsvr32 /u /s TABCTL32.OCX"
Else
    RegisterDependency = "Dependency TABCTL32.OCX missing."
    Exit Function
End If

If FileExists(Path & "MSCOMCTL.OCX") Then
    Shell "regsvr32 /u /s MSCOMCTL.OCX"
Else
    RegisterDependency = "Dependency MSCOMCTL.OCX missing."
    Exit Function
End If

If FileExists(Path & "MSWINSCK.OCX") Then
    Shell "regsvr32 /u /s MSWINSCK.OCX"
Else
    RegisterDependency = "Dependency MSWINSCK.OCX missing."
    Exit Function
End If

'Register
Shell "regsvr32 /s COMDLG32.OCX"
Shell "regsvr32 /s RICHTX32.OCX"
Shell "regsvr32 /s TABCTL32.OCX"
Shell "regsvr32 /s MSCOMCTL.OCX"
Shell "regsvr32 /s MSWINSCK.OCX"
End Function

