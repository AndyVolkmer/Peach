Attribute VB_Name = "MainModule"
Option Explicit
Private Declare Sub InitCommonControls Lib "comctl32" ()

Sub Main()
'Load Windows own style
Call InitCommonControls

'If the application is already open then close it
If App.PrevInstance Then
    End
End If

'Define variables for switch
Dim N As String
Dim M As String

'Load configuration from .ini file into switch variables
If Len(ReadIniValue(App.Path & "\Config.ini", "Language", "Validate")) = 0 Then
    N = 1 'Default -> choose language
Else
    N = ReadIniValue(App.Path & "\Config.ini", "Language", "Validate")
End If

If Len(ReadIniValue(App.Path & "\Config.ini", "Language", "Language")) = 0 Then
    M = 1 'Default language -> english
Else
    M = ReadIniValue(App.Path & "\Config.ini", "Language", "Language")
End If

If N = 0 Then
    Select Case M
    Case 0 'German
        SetLangGerman
    Case 1 'English
        SetLangEnglish
    Case 2 'Spanish
        SetLangSpanish
    Case 3 'Swedish
        SetLangSwedish
    Case 4 'Italian
        SetLangItalian
    Case 5 'Dutch
        SetLangDutch
    Case 6 'Serbian
        SetLangSerbian
    Case 7 'French
        SetLangFrench
    End Select
    frmMain.Show
Else
    'Set language default -> english
    SetLangEnglish
    frmLanguage.Show
End If
WriteIniValue App.Path & "\Config.ini", "Revision", "Number", Rev
End Sub
