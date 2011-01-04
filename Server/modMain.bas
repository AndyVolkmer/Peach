Attribute VB_Name = "modMain"
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32" ()

Sub Main()
Dim pStartTime      As Long
Dim pConnectResult  As String

'Capture start time
pStartTime = timeGetTime

'Load windows own style
InitCommonControls

'Load registry values
LoadRegistry

'Load ini values
LoadIniValue

'Load config values
LoadConfigValue

'Connect to Database
pConnectResult = pDB.ConnectDatabase(Database.Type, Database.Host, Database.Database, Database.User, Database.Password, Database.File)

If LenB(pConnectResult) = 0 Then
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

    If Not LoadOfflineMessages Then
        MsgBox frmConfig.txt_log.Text, vbInformation
        End
    End If
Else
    MsgBox pConnectResult, vbInformation
    End
End If

frmMain.StatusBar1.Panels(1) = "Status: Disconnected"

DisableFormResize frmMain

Dim L As Long
    L = GetWindowLong(frmMain.hwnd, GWL_STYLE)
'   L = L And Not (WS_MINIMIZEBOX)
    L = L And Not (WS_MAXIMIZEBOX)
    L = SetWindowLong(frmMain.hwnd, GWL_STYLE, L)

frmMain.SetupForms frmConfig
WriteLog "Finished loading in " & timeGetTime - pStartTime & " ms."
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
Dim eCheck As String
Dim p_Path As String
    p_Path = App.Path & "\peachConfig.conf"

With Database
    If LenB(ReadIniValue(p_Path, "Database", "Type")) = 0 Then
        WriteIniValue p_Path, "Database", "Type", vbNullString
        If MsgBox("Database type not found, please check configuration file. Do you want to enter a value?", vbQuestion + vbYesNo) = vbYes Then
            eCheck = InputBox("Enter a database type." & vbCrLf & "0 - Access Database" & vbCrLf & "1 - MySQL Database", "Database type ..")
            If IsNumeric(eCheck) And eCheck < 2 Then
                .Type = eCheck
                WriteIniValue p_Path, "Database", "Type", eCheck
            Else
                MsgBox "Database type is incorrect.", vbError
                End
            End If
        Else
            End
        End If
    Else
        .Type = ReadIniValue(p_Path, "Database", "Type")
    End If

    If .Type = "0" Then
        If LenB(ReadIniValue(p_Path, "Database", "File")) = 0 Then
            WriteIniValue p_Path, "Database", "File", vbNullString
            If MsgBox("Database file name not found, please check configuration file. Do you want to enter a value?", vbQuestion + vbYesNo) = vbYes Then
                WriteIniValue p_Path, "Database", "File", InputBox("Enter a database file name.", "Database file name ..")
            Else
                End
            End If
        Else
            .File = ReadIniValue(p_Path, "Database", "File")
        End If

    ElseIf .Type = "1" Then
        If LenB(ReadIniValue(p_Path, "Database", "Name")) = 0 Then
            WriteIniValue p_Path, "Database", "Name", vbNullString
            If MsgBox("Database name not found, please check configuration file. Do you want to enter a value?", vbQuestion + vbYesNo) = vbYes Then
                eCheck = InputBox("Enter a database name.", "Database name ..")
                .Database = eCheck
                WriteIniValue p_Path, "Database", "Name", eCheck
            Else
                End
            End If
        Else
            .Database = ReadIniValue(p_Path, "Database", "Name")
        End If

        If LenB(ReadIniValue(p_Path, "Database", "User")) = 0 Then
            WriteIniValue p_Path, "Database", "User", vbNullString
            If MsgBox("Database user not found, please check configuration file. Do you want to enter a value?", vbQuestion + vbYesNo) = vbYes Then
                eCheck = InputBox("Enter a database user.", "Database user ..")
                .User = eCheck
                WriteIniValue p_Path, "Database", "User", eCheck
            Else
                End
            End If
        Else
            .User = ReadIniValue(p_Path, "Database", "User")
        End If

        If LenB(ReadIniValue(p_Path, "Database", "Password")) = 0 Then
            WriteIniValue p_Path, "Database", "Password", vbNullString
            If MsgBox("Database password not found, please check configuration file. Do you want to enter a value?", vbQuestion + vbYesNo) = vbYes Then
                eCheck = InputBox("Enter a database password.", "Database password ..")
                .Password = eCheck
                WriteIniValue p_Path, "Database", "Password", eCheck
            Else
                End
            End If
        Else
            .Password = Decode(ReadIniValue(p_Path, "Database", "Password"))
        End If

        If LenB(ReadIniValue(p_Path, "Database", "Host")) = 0 Then
            WriteIniValue p_Path, "Database", "Host", vbNullString
            If MsgBox("Database host not found, please check configuration file. Do you want to enter a value?", vbQuestion + vbYesNo) = vbYes Then
                eCheck = InputBox("Enter a database host.", "Database host ..")
                .Host = eCheck
                WriteIniValue p_Path, "Database", "Host", eCheck
            Else
                End
            End If
        Else
            .Host = ReadIniValue(p_Path, "Database", "Host")
        End If
    End If
End With
End Sub

Public Sub LoadConfigValue()
Dim p_Path As String
    p_Path = App.Path & "\peachConfig.conf"

If LenB(ReadIniValue(p_Path, "Options", "CapsCheck")) = 0 Then
    WriteIniValue p_Path, "Options", "CapsCheck", "1"
    MsgBox "Option caps check value not found, default set.", vbInformation
    Options.CAPS_CHECK = 1
Else
    Options.CAPS_CHECK = ReadIniValue(p_Path, "Options", "CapsCheck")
End If

If LenB(ReadIniValue(p_Path, "Options", "RepeatCheck")) = 0 Then
    WriteIniValue p_Path, "Options", "RepeatCheck", "1"
    MsgBox "Option repeat check vlaue not found, default set.", vbInformation
    Options.CAPS_CHECK = 1
Else
    Options.REPEAT_CHECK = ReadIniValue(p_Path, "Options", "RepeatCheck")
End If

If LenB(ReadIniValue(p_Path, "Options", "RootPass")) = 0 Then
    WriteIniValue p_Path, "Options", "RootPass", "alpine"
    MsgBox "Root password value not found, default set.", vbInformation
    Options.ROOT_PASSWORD = "alpine"
Else
    Options.ROOT_PASSWORD = ReadIniValue(p_Path, "Options", "RootPass")
End If
End Sub
