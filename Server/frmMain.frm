VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   " Peach (Server)"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7545
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   5
      Top             =   4740
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7938
            MinWidth        =   7938
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   0
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   7485
      TabIndex        =   0
      Top             =   0
      Width           =   7545
      Begin VB.CommandButton Command4 
         Caption         =   "&Panel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   4
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Send File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ch&at"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Configuration"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GWL_STYLE = (-16)
Private Const WS_SYSMENU = &H80000
Private Const WS_CAPTION = &HC00000
Private Const WS_THICKFRAME = &H40000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const SC_CLOSE As Long = &HF060&
Private Const SC_MAXIMIZE = &HF030&
Private Const SC_MINIMIZE = &HF020&
Private Const MIIM_STATE As Long = &H1&
Private Const MIIM_ID As Long = &H2&
Private Const MFS_GRAYED As Long = &H3&
Private Const WM_NCACTIVATE As Long = &H86
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type


Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage2 Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Public Prefix       As String
Public intCounter   As Integer
Public NameText     As String
Public ConverText   As String
Public Message      As String
Dim RR As Integer


Private Sub Command1_Click()
    SetupForms frmConfig
End Sub

Private Sub Command2_Click()
    SetupForms frmChat
End Sub

Private Sub SetupForms(Nix As Form)
    frmChat.Hide
    frmConfig.Hide
    frmSendFile.Hide
    Nix.Show
End Sub

Private Sub Command3_Click()
    SetupForms frmSendFile
End Sub

Private Sub Command4_Click()
    SetupForms frmPanel
End Sub

Private Sub Command5_Click()

End Sub

Private Sub MDIForm_Load()
':::::::::::::::::::::::::::::::::::
DisableFormResize Me
':::::::::::::::::::::::::::::::::::
Dim L As Long
    L = GetWindowLong(Me.hwnd, GWL_STYLE)
    'L = L And Not (WS_MINIMIZEBOX)
    L = L And Not (WS_MAXIMIZEBOX)
    L = SetWindowLong(Me.hwnd, GWL_STYLE, L)
':::::::::::::::::::::::::::::::::::
StatusBar1.Panels(1).Text = "Status : Disconnected"
frmPanel.Text1 = "0"
    SetupForms frmConfig
End Sub

Private Sub Winsock1_Close(Index As Integer)
    Dim x As Integer
    Dim NameOfUser As String
    Unload Winsock1(Index)
    For x = 1 To frmPanel.ListView1.ListItems.Count + 1
        If frmPanel.ListView1.ListItems.Item(x).SubItems(2) = Index Then
            ' Pick the user
            NameOfUser = frmPanel.ListView1.ListItems.Item(x)
            frmPanel.ListView1.ListItems.Remove (x)
            Exit For
        Else
            '
        End If
    Next x
    SendMessage " " & NameOfUser & " has disconnected!"
    StatusBar1.Panels(1).Text = "Status: Connected with  " & Winsock1.Count - 1 & " Client(s)."
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
RR = frmPanel.ListView1.ListItems.Count + 1
Dim Wsk As Winsock
intCounter = Winsock1.Count
    For Each Wsk In Winsock1
        If Wsk.State = sckConnected Then
            intCounter = Wsk.Index + 1
        End If
    Next
    Load Winsock1(intCounter)
    Winsock1(intCounter).LocalPort = frmConfig.txtPort.Text
    Winsock1(intCounter).Accept requestID
    ' New user should be listed in the panel
    With frmPanel.ListView1
        .ListItems.Add RR, , "N/A"
        .ListItems.Item(RR).SubItems(1) = Winsock1(intCounter).RemoteHostIP
        .ListItems.Item(RR).SubItems(2) = intCounter
        .ListItems.Item(RR).SubItems(3) = Time
    End With
    
    StatusBar1.Panels(1).Text = "Status: Connected with  " & Winsock1.Count - 1 & " Client(s)."
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
RR = frmPanel.ListView1.ListItems.Count
Dim strMessage As String
Dim onlineCount As String
Dim c1, c2 As String
Dim v As Variant
Dim difC As Integer

' Get Message
frmMain.Winsock1(Index).GetData strMessage

' Encode Message
v = InStr(1, strMessage, "#")

c1 = Mid(strMessage, 1, v - 1)

difC = Len(strMessage) - Len(c1) - 1

c2 = Mid(strMessage, Len(c1) + 2, difC - 1)

frmMain.NameText = c1
frmMain.ConverText = c2
' If message is to long then give warn message .. >_>
If Len(frmMain.ConverText) > 200 Then
    strMessage = "[" & frmMain.NameText & "] - You cannot spam here with older version!"
End If

Select Case Trim(frmMain.ConverText)
Case "!online"
    Select Case Winsock1.UBound
    Case 0
        SendMessage " [System] : " & " No users are online."
    Case 1
        SendMessage " [System] : 1 user is online. (You)"
    Case Else
        SendMessage " [System] : " & Winsock1.Count - 1 & " users are online."
    End Select
Case "!connected"
    SendMessage " " & frmMain.NameText & " has connected."
    frmPanel.ListView1.ListItems.Item(RR).Text = frmMain.NameText ' 1, , frmMain.NameText
Case "!namerequest"
    Dim u As Variant
    For u = 1 To frmPanel.ListView1.ListItems.Count + 1
        If frmPanel.ListView1.ListItems.Item(u) = frmMain.NameText Then
            SendRequest "!decilineD", frmMain.Winsock1(Index)
            Exit For
        Else
            SendRequest "!accepteD", frmMain.Winsock1(Index)
            Exit For
        End If
        Exit For
    Next u
Case Else
    SendMessage " [" & frmMain.NameText & "] : " & frmMain.ConverText
End Select

' Read Message
frmChat.txtConver.Text = frmChat.txtConver.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "]" & strMessage
End Sub


Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
frmMain.Prefix = "[" & Format(Time, "hh:nn:ss") & "]"
If Index < 0 Then
    Winsock1(Index).Close
    Unload Winsock1(Index)
    frmChat.txtConver.Text = frmChat.txtConver.Text & vbCrLf & frmMain.Prefix & " [System]: Disconnected with a host."
Else
    Winsock1(Index).Close
    Unload Winsock1(Index)
    frmChat.txtConver.Text = frmChat.txtConver.Text & vbCrLf & frmMain.Prefix & "[System]: Disconnected due connection problem."
    StatusBar1.Panels(1).Text = "[System]: Disconnected due connection problem."
End If
Debug.Print "Number : " & Number & vbCrLf & "Description : " & Description & " Winsock State : " & Winsock1(Index).State
End Sub

Public Sub DisableFormResize(frm As Form)

    Dim style As Long
    Dim hMenu As Long
    Dim MII As MENUITEMINFO
    Dim lngMenuID As Long
    Const xSC_MAXIMIZE As Long = -11

    style = GetWindowLong(frm.hwnd, GWL_STYLE)
    
    style = style And Not WS_THICKFRAME
    style = style And Not WS_MAXIMIZEBOX
    
    style = SetWindowLong(frm.hwnd, GWL_STYLE, style)
    
    On Error Resume Next
    
    hMenu = GetSystemMenu(frm.hwnd, 0)
    
    With MII
        .cbSize = Len(MII)
        .dwTypeData = String(80, 0)
        .cch = Len(.dwTypeData)
        .fMask = MIIM_STATE
        .wID = SC_MAXIMIZE
    End With
    If GetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then Exit Sub
    
    With MII
        lngMenuID = .wID
        .wID = xSC_MAXIMIZE
        .fMask = MIIM_ID
    End With
    If SetMenuItemInfo(hMenu, lngMenuID, False, MII) = 0 Then Exit Sub
    
    With MII
        .fState = (.fState Or MFS_GRAYED)
        .fMask = MIIM_STATE
    End With
    If SetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then Exit Sub
    
    SendMessage2 hwnd, WM_NCACTIVATE, True, 0
    
    frm.Width = frm.Width - 1
    frm.Width = frm.Width + 1
End Sub
