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
Public Command      As String
Public NameText     As String
Public ConverText   As String
Public Message      As String
Public intCounter   As Integer

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


Private Sub MDIForm_Load()
DisableFormResize Me
Dim L As Long
    L = GetWindowLong(Me.hwnd, GWL_STYLE)
'   L = L And Not (WS_MINIMIZEBOX)
    L = L And Not (WS_MAXIMIZEBOX)
    L = SetWindowLong(Me.hwnd, GWL_STYLE, L)
StatusBar1.Panels(1).Text = "Status : Disconnected"
frmPanel.Text1 = "0"
SetupForms frmConfig
End Sub

Private Sub Winsock1_Close(Index As Integer)
Dim NameOfUser      As String
Dim x               As Integer
    Unload Winsock1(Index)
    For x = 1 To frmPanel.ListView1.ListItems.Count + 1
        ' Give the message to disconnect and reload user list for all clients
        If frmPanel.ListView1.ListItems.Item(x).SubItems(2) = Index Then
            
            ' Pick the user
            NameOfUser = frmPanel.ListView1.ListItems.Item(x)
            frmPanel.ListView1.ListItems.Remove (x)
            
            ' Update Users List
            UpdateUsersList
            
            Exit For
        End If
    Next x
    VisualizeMessage "!disconnect", NameOfUser, "logout"
    StatusBar1.Panels(1).Text = "Status: Connected with  " & Winsock1.Count - 1 & " Client(s)."
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim RR As Integer
RR = frmPanel.ListView1.ListItems.Count + 1
    intCounter = loadSocket
    Winsock1(intCounter).LocalPort = frmConfig.txtPort.Text
    Winsock1(intCounter).Accept requestID
    ' New user should be listed in the panel
    With frmPanel.ListView1
        .ListItems.Add RR, , "N/A"
        .ListItems.Item(RR).SubItems(1) = Winsock1(intCounter).RemoteHostIP
        .ListItems.Item(RR).SubItems(2) = intCounter
        .ListItems.Item(RR).SubItems(3) = Format(Time, "hh:nn:ss")
    End With
    StatusBar1.Panels(1).Text = "Status: Connected with  " & Winsock1.Count - 1 & " Client(s)."
End Sub

'On error means its free so we take free socket
Private Function socketFree() As Integer
On Error GoTo HandleErrorFreeSocket
    Dim theIP As Variant
    Dim p As Integer
    For p = Winsock1.LBound + 1 To Winsock1.UBound
        theIP = Winsock1(p).LocalIP
    Next
    socketFree = Winsock1.UBound + 1
Exit Function
HandleErrorFreeSocket:
socketFree = p
End Function

Private Function loadSocket() As Integer
' Load next free socket
Dim theFreeSocket As Integer: theFreeSocket = 0
theFreeSocket = socketFree

Load Winsock1(theFreeSocket)

loadSocket = theFreeSocket
End Function
'
Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim array1()        As String
Dim strMessage      As String
Dim RR              As Integer
Dim bMatch          As Boolean
RR = frmPanel.ListView1.ListItems.Count

' Get Message
frmMain.Winsock1(Index).GetData strMessage

' We decode (split) the message into an array
array1 = Split(strMessage, "#")

' Assign the variables to the array
With frmMain
    .Command = array1(0)
    .NameText = array1(1)
    .ConverText = array1(2)
End With

' Validate: If message is to long then give warn message .. >_>
If Len(frmMain.ConverText) > 200 Then
    SendRequest " You cannot spam here with older version!", Winsock1(Index)
    VisualizeMessage "!spam", frmMain.NameText, "Spamming"
    Exit Sub
End If

Select Case frmMain.Command
Dim u As Integer
Case "!connected" ' Announce connected player and send to user online list
    frmPanel.ListView1.ListItems.Item(RR).Text = frmMain.NameText
    UpdateUsersList
    
Case "!namerequest" ' Check if the name is avaible or not
    For u = 1 To frmPanel.ListView1.ListItems.Count
        If UCase(frmPanel.ListView1.ListItems.Item(u)) = UCase(frmMain.NameText) Then
            bMatch = True
            Exit For
        End If
    Next u
    
    If bMatch = True Then ' Return yes or no to client
        bMatch = False
        SendRequest "!decilineD", frmMain.Winsock1(Index)
    Else
        SendRequest "!accepteD", frmMain.Winsock1(Index)
    End If
    
Case Else
    SendMessage " [" & frmMain.NameText & "] : " & frmMain.ConverText
End Select

' We want to read the message also , different then others tho
VisualizeMessage frmMain.Command, frmMain.NameText, frmMain.ConverText
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim i As Integer
frmMain.Prefix = "[" & Format(Time, "hh:nn:ss") & "]"
If Index < 0 Then
    Winsock1(Index).Close
    Unload Winsock1(Index)
    frmChat.txtConver.Text = frmChat.txtConver.Text & vbCrLf & frmMain.Prefix & " [System]: Disconnected with a host."
    For i = 1 To frmPanel.ListView1.ListItems.Count
        frmPanel.ListView1.ListItems.Remove (i)
    Next i
Else
    Winsock1(Index).Close
    Unload Winsock1(Index)
    frmChat.txtConver.Text = frmChat.txtConver.Text & vbCrLf & frmMain.Prefix & "[System]: Disconnected due connection problem."
    StatusBar1.Panels(1).Text = "[System]: Disconnected due connection problem."
    For i = 1 To frmPanel.ListView1.ListItems.Count
        frmPanel.ListView1.ListItems.Remove (i)
    Next
End If
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
