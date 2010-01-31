Attribute VB_Name = "modMain"
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32" ()

Sub Main()
Dim pStartTime As Long

InitCommonControls
pStartTime = timeGetTime

'Load registry values
LoadRegistry

'Connect to MySQL Database
 If ConnectMySQL(Database.Database, Database.User, Database.Password, Database.Host) Then
    If Not LoadAccounts Then
        MsgBox frmConfig.txt_log.Text, vbInformation
        frmSettings.Show 1
        End
    End If
    
    If Not LoadCommands Then
        MsgBox frmConfig.txt_log.Text, vbInformation
        frmSettings.Show 1
        End
    End If
    
    If Not LoadDeclinedNames Then
        MsgBox frmConfig.txt_log.Text, vbInformation
        frmSettings.Show 1
        End
    End If
    
    If Not LoadEmotes Then
        MsgBox frmConfig.txt_log.Text, vbInformation
        frmSettings.Show 1
        End
    End If
    
    If Not LoadFriends Then
        MsgBox frmConfig.txt_log.Text, vbInformation
        frmSettings.Show 1
        End
    End If
    
    If Not LoadIgnores Then
        MsgBox frmConfig.txt_log.Text, vbInformation
        frmSettings.Show 1
        End
    End If
Else
    MsgBox frmConfig.txt_log.Text, vbInformation
    frmSettings.Show 1
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
With Database
    '== Database ==
    If LenB(ReadFromRegistry("Server\Database", "Name")) <> 0 Then
        .Database = ReadFromRegistry("Server\Database", "Name")
    Else
        WriteLog "No Database value found."
        .Database = InputBox("The configuration file does not cotain a valid database, please insert one in the textbox below.", "Database error ..", "Database Name")
    End If
    
    If LenB(ReadFromRegistry("Server\Database", "User")) <> 0 Then
        .User = ReadFromRegistry("Server\Database", "User")
    Else
        WriteLog "No User value found."
        .User = InputBox("The configuration file does not cotain a valid user name, please insert one in the textbox below.", "Database error ..", "User Name")
    End If
    
    If LenB(DeCode(ReadFromRegistry("Server\Database", "Password"))) <> 0 Then
        .Password = DeCode(ReadFromRegistry("Server\Database", "Password"))
    Else
        WriteLog "No Password value found."
        .Password = InputBox("The configuration file does not cotain a valid password, please insert one in the textbox below.", "Database error ..", "Password")
    End If
    
    If LenB(ReadFromRegistry("Server\Database", "Host")) <> 0 Then
        .Host = ReadFromRegistry("Server\Database", "Host")
    Else
        WriteLog "No Host value found."
        .Host = InputBox("The configuration file does not cotain a host adress, please insert one in the textbox below.", "Database error ..", "Host Adress")
    End If
    
    '== Position ==
    If LenB(ReadFromRegistry("Server\Configuration", "Top")) <> 0 Then
        frmMain.Top = ReadFromRegistry("Server\Configuration", "Top")
    Else
        frmMain.Top = 1200
    End If
    
    If LenB(ReadFromRegistry("Server\Configuration", "Left")) <> 0 Then
        frmMain.Left = ReadFromRegistry("Server\Configuration", "Left")
    Else
        frmMain.Left = 1200
    End If
End With
End Sub
