Attribute VB_Name = "modFunction"
Option Explicit

Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer

Private Type PROCESSENTRY32
   dwSize               As Long
   cntUsage             As Long
   th32ProcessID        As Long
   th32DefaultHeapID    As Long
   th32ModuleID         As Long
   cntThreads           As Long
   th32ParentProcessID  As Long
   pcPriClassBase       As Long
   dwFlags              As Long
   szExeFile            As String * 260
End Type

Private Type OSVERSIONINFO
   dwOSVersionInfoSize  As Long
   dwMajorVersion       As Long
   dwMinorVersion       As Long
   dwBuildNumber        As Long
   dwPlatformId         As Long
   szCSDVersion         As String * 128
End Type

Private Const PROCESS_ALL_ACCESS = 0
Private Const TH32CS_SNAPPROCESS As Long = 2&

Private Const WINNT     As Integer = 2
Private Const WIN98     As Integer = 1

Public KillAppReturn    As Boolean
Public i                As Long

Public Function getVersion() As Integer
Dim udtOSInfo As OSVERSIONINFO
Dim intRetVal As Integer
        
With udtOSInfo
    .dwOSVersionInfoSize = 148
    .szCSDVersion = Space$(128)
End With
   
intRetVal = GetVersionExA(udtOSInfo)
 
getVersion = udtOSInfo.dwPlatformId
End Function

Public Function Killapp(myName As String)
Select Case getVersion()
    Case WIN98
        Killapp9X (myName)
        
    Case WINNT
        KillappNT (myName)
End Select
End Function

Private Function KillappNT(myName As String)
Dim uProcess        As PROCESSENTRY32
Dim rProcessFound   As Long
Dim hSnapshot       As Long
Dim szExename       As String
Dim exitCode        As Long
Dim myProcess       As Long
Dim AppKill         As Boolean
Dim appCount        As Integer
Dim i               As Integer

On Local Error GoTo Finish
appCount = 0
uProcess.dwSize = Len(uProcess)
hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
rProcessFound = ProcessFirst(hSnapshot, uProcess)

Do While rProcessFound
    i = InStr(1, uProcess.szExeFile, Chr(0))
    szExename = LCase$(Left$(uProcess.szExeFile, i - 1))
       
    If Right$(szExename, Len(myName)) = LCase$(myName) Then
        KillAppReturn = True
        appCount = appCount + 1
        myProcess = OpenProcess(1&, -1&, uProcess.th32ProcessID)
        AppKill = TerminateProcess(myProcess, 0&)
        Call CloseHandle(myProcess)
    End If
    
    rProcessFound = ProcessNext(hSnapshot, uProcess)
Loop
   
Call CloseHandle(hSnapshot)

Finish:
KillAppReturn = False
End Function

Private Function Killapp9X(myName As String)
Dim uProcess        As PROCESSENTRY32
Dim rProcessFound   As Long
Dim hSnapshot       As Long
Dim szExename       As String
Dim exitCode        As Long
Dim myProcess       As Long
Dim AppKill         As Boolean
Dim appCount        As Integer
Dim i               As Integer
   
On Local Error GoTo Finish
   
appCount = 0
uProcess.dwSize = Len(uProcess)
hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
rProcessFound = ProcessFirst(hSnapshot, uProcess)
   
Do While rProcessFound
    i = InStr(1, uProcess.szExeFile, Chr(0))
    szExename = LCase$(Left$(uProcess.szExeFile, i - 1))
    
    If Right$(szExename, Len(myName)) = LCase$(myName) Then
        KillAppReturn = True
        appCount = appCount + 1
        myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
        AppKill = TerminateProcess(myProcess, exitCode)
        Call CloseHandle(myProcess)
    End If
    
    rProcessFound = ProcessNext(hSnapshot, uProcess)
Loop

Call CloseHandle(hSnapshot)

Finish:
KillAppReturn = False
End Function

Public Function DownloadFile(ByVal source_file As String, ByVal dest_file As String) As Boolean
Dim bytes() As Byte
Dim fnum As Integer

' Get the file's contents.
On Error GoTo DownloadError
bytes() = frmMain.inetftp.OpenURL(source_file, icByteArray)

' Remove the file if it exists.
On Error Resume Next
Kill dest_file
On Error GoTo DownloadError

' Write the contents into the destination file.
fnum = FreeFile
Open dest_file For Binary Access Write As #fnum
Put #fnum, , bytes()
Close #fnum

DownloadFile = True
Exit Function

DownloadError:
MsgBox "Error " & Err.Number & _
    " downloading file '" & _
    source_file & "' to '" & _
    dest_file & "'." & vbCrLf & Err.Description, _
    vbExclamation Or vbOKOnly, _
    "Download Error"
DownloadFile = False
Exit Function
End Function

Public Sub StartDownload(SourceF As String, DestF As String)
DoEvents
If DownloadFile(SourceF, DestF) Then
    frmMain.Label2.Caption = "Status : Downloaded Patch files."
End If
End Sub
