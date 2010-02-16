Attribute VB_Name = "modMain"
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32" ()

Sub Main()
Dim pStartTime As Long

InitCommonControls
pStartTime = timeGetTime

'Load registry values
LoadRegistry

'Load ini values
LoadIniValue

'Connect to MySQL Database
 If ConnectMySQL(Database.Database, Database.User, Database.Password, Database.Host) Then
    If Not LoadAccounts Then
        MsgBox frmConfig.txt_log.Text, vbInformation
        End
    End If
    
    If Not LoadCommands Then
        MsgBox frmConfig.txt_log.Text, vbInformation
        End
    End If
    
    If Not LoadDeclinedNames Then
        MsgBox frmConfig.txt_log.Text, vbInformation
        End
    End If
    
    If Not LoadEmotes Then
        MsgBox frmConfig.txt_log.Text, vbInformation
        End
    End If
    
    If Not LoadFriends Then
        MsgBox frmConfig.txt_log.Text, vbInformation
        End
    End If
    
    If Not LoadIgnores Then
        MsgBox frmConfig.txt_log.Text, vbInformation
        End
    End If
Else
    MsgBox frmConfig.txt_log.Text, vbInformation
    End
End If

frmMain.StatusBar1.Panels(1) = "Status: Disconnected"
frmMain.SetupForms frmConfig

WriteLog "Correctly loaded in " & timeGetTime - pStartTime & " ms."
pStartTime = vbNull

DisableFormResize frmMain

Dim L As Long
    L = GetWindowLong(frmMain.hwnd, GWL_STYLE)
'   L = L And Not (WS_MINIMIZEBOX)
    L = L And Not (WS_MAXIMIZEBOX)
    L = SetWindowLong(frmMain.hwnd, GWL_STYLE, L)
End Sub

Private Sub LoadRegistry()
With frmMain
    '== Position ==
    If LenB(ReadFromRegistry("Server\Configuration", "Top")) <> 0 Then
        .Top = ReadFromRegistry("Server\Configuration", "Top")
    Else
        .Top = 1200
    End If
    
    If LenB(ReadFromRegistry("Server\Configuration", "Left")) <> 0 Then
        .Left = ReadFromRegistry("Server\Configuration", "Left")
    Else
        .Left = 1200
    End If
End With
End Sub

Private Sub LoadIniValue()
Dim p_Path As String
    p_Path = App.Path & "\peachConfig.conf"

With Database
    '== Database ==
    If LenB(ReadIniValue(p_Path, "Database", "Name")) = 0 Then
        WriteIniValue p_Path, "Database", "Name", vbNullString
        If MsgBox("Database name not found, please check configuration file. Do you want to enter a value?", vbQuestion + vbYesNo) = vbYes Then
            WriteIniValue p_Path, "Database", "Name", InputBox("Enter a database name.", "Database name ..")
        End If
        End
    Else
        .Database = ReadIniValue(p_Path, "Database", "Name")
    End If
    
    If LenB(ReadIniValue(p_Path, "Database", "User")) = 0 Then
        WriteIniValue p_Path, "Database", "User", vbNullString
        If MsgBox("Database user not found, please check configuration file. Do you want to enter a value?", vbQuestion + vbYesNo) = vbYes Then
            WriteIniValue p_Path, "Database", "User", InputBox("Enter a database user.", "Database user ..")
        End If
        End
    Else
        .User = ReadIniValue(p_Path, "Database", "User")
    End If
    
    If LenB(ReadIniValue(p_Path, "Database", "Password")) = 0 Then
        WriteIniValue p_Path, "Database", "Password", vbNullString
        If MsgBox("Database password not found, please check configuration file. Do you want to enter a value?", vbQuestion + vbYesNo) = vbYes Then
            WriteIniValue p_Path, "Database", "Password", InputBox("Enter a database password.", "Database password ..")
        End If
        End
    Else
        .Password = DeCode(ReadIniValue(p_Path, "Database", "Password"))
    End If
    
    If LenB(ReadIniValue(p_Path, "Database", "Host")) = 0 Then
        WriteIniValue p_Path, "Database", "Host", vbNullString
        If MsgBox("Database host not found, please check configuration file. Do you want to enter a value?", vbQuestion + vbYesNo) = vbYes Then
            WriteIniValue p_Path, "Database", "Host", InputBox("Enter a database host.", "Database host ..")
        End If
        End
    Else
        .Host = ReadIniValue(p_Path, "Database", "Host")
    End If
    
    If LenB(ReadIniValue(p_Path, "Options", "CapsCheck")) = 0 Then
        WriteIniValue p_Path, "Options", "CapsCheck", "1"
        MsgBox "Option caps check value not found, default set.", vbInformation
        End
    Else
        Options.CAPS_CHECK = ReadIniValue(p_Path, "Options", "CapsCheck")
    End If
    
    If LenB(ReadIniValue(p_Path, "Options", "RepeatCheck")) = 0 Then
        WriteIniValue p_Path, "Options", "RepeatCheck", "1"
        MsgBox "Option repeat check vlaue not found, default set.", vbInformation
        End
    Else
        Options.REPEAT_CHECK = ReadIniValue(p_Path, "Options", "RepeatCheck")
    End If
End With
End Sub
