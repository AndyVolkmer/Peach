Attribute VB_Name = "modRegistry"
Option Explicit

Const KEY_ALL_ACCESS = &H2003F

Public Const RegistryDataType   As Long = 1                          '1 = String, 2 = Binary, 4 = DWORD
Public Const RegistryKeyROOT    As Long = &H80000002
'&H80000000 = HKEY_CLASSES_ROOT
'&H80000001 = HKEY_CURRENT_USER
'&H80000002 = HKEY_LOCAL_MACHINE <- Using this one
'&H80000003 = HKEY_USERS
'&H80000005 = HKEY_CURRENT_CONFIG
Public Const RegistryPath       As String = "SOFTWARE\Peach\"        'Path inside the default root
Public Const RegistryROOT       As String = "HKEY_LOCAL_MACHINE"     'Name again?

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Public Sub InsertIntoRegistry(pFolder As String, pTitle As String, pValue As String)
SetKeyDataValue RegistryKeyROOT, RegistryPath & pFolder, RegistryDataType, pTitle, pValue
End Sub

Public Function ReadFromRegistry(pFolder As String, pTitle As String) As String
ReadFromRegistry = GetKeyDataValue(RegistryKeyROOT, RegistryPath & pFolder, RegistryDataType, pTitle)
End Function

Public Sub DeleteFromRegistry(pFolder As String, pTitle As String)
DeleteRegValue RegistryKeyROOT, RegistryPath & pFolder, pTitle
End Sub

Private Sub SetKeyDataValue(RegKeyRoot As Long, RegKeyName As String, KeyDataType As Long, KeyValueName As String, KeyValueDate As Variant)
Dim OpenKey     As Long
Dim SetValue    As Long
Dim hKey        As Long
    
OpenKey = RegOpenKeyEx(RegKeyRoot, RegKeyName, 0, KEY_ALL_ACCESS, hKey)
    
If (OpenKey <> 0) Then
    Call RegCreateKey(RegKeyRoot, RegKeyName, hKey)
End If

Select Case KeyDataType
    Case 1:
        SetValue = RegSetValueEx(hKey, KeyValueName, 0&, KeyDataType, ByVal CStr(KeyValueDate & Chr$(0)), Len(KeyValueDate))
        
    Case 3:
        SetValue = RegSetValueEx(hKey, KeyValueName, 0&, KeyDataType, ByVal CStr(KeyValueDate & Chr$(0)), Len(KeyValueDate))
        
    Case 4:
        SetValue = RegSetValueEx(hKey, KeyValueName, 0&, KeyDataType, CLng(KeyValueDate), 4)
End Select
    
SetValue = RegCloseKey(hKey)
End Sub

Private Function GetKeyDataValue(RegKeyRoot As Long, RegKeyName As String, KeyDateType As Long, KeyValueName As String)
Dim OpenKey     As Long
Dim hKey        As Long
Dim strTempVal  As String
Dim KeyValSize  As Long
       
OpenKey = RegOpenKeyEx(RegKeyRoot, RegKeyName, 0, KEY_ALL_ACCESS, hKey)
        
strTempVal = String$(1024, 0)
KeyValSize = 1024
    
OpenKey = RegQueryValueEx(hKey, KeyValueName, 0, KeyDateType, strTempVal, KeyValSize)

On Error Resume Next
If Asc(Mid(strTempVal, KeyValSize, 1) = 0) Then
    strTempVal = Left(strTempVal, KeyValSize - 1)
Else
    strTempVal = Left(strTempVal, KeyValSize)
End If

Select Case KeyDateType
    Case 1, 3:
        If Not strTempVal = String$(1023, 0) Then
            GetKeyDataValue = strTempVal
        End If
        
    Case 4:
        For i = Len(strTempVal) To 1 Step -1
            GetKeyDataValue = Format(Hex(Asc(Mid(strTempVal, i, 1))), "00")
        Next i
        
        GetKeyDataValue = Format$("&h" + GetKeyDataValue)
        
End Select

OpenKey = RegCloseKey(hKey)
End Function

Private Sub DeleteRegValue(RegKeyRoot As Long, RegKeyName As String, KeyValueName As String)
Dim DeleteKeyValue  As Long
Dim hKey            As Long

DeleteKeyValue = RegOpenKeyEx(RegKeyRoot, RegKeyName, 0, KEY_ALL_ACCESS, hKey)
DeleteKeyValue = RegDeleteValue(hKey, KeyValueName)

End Sub

Private Sub DeleteRegKey(RegKeyRoot As Long, RegKeyName As String)
Dim DeleteKey As Long
DeleteKey = RegDeleteKey(RegKeyRoot, RegKeyName)
End Sub
