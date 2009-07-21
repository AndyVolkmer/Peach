VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "frmConfig"
   ClientHeight    =   4140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   Begin VB.Timer connCounter 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   2400
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Disconnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      Caption         =   "Configuration"
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "4728"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtNick 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "Server"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000C&
         Caption         =   "Offline"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   11
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000C&
         Caption         =   "IP : "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000C&
         Caption         =   "Name : "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000C&
         Caption         =   "Port :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblNick 
         BackColor       =   &H8000000C&
         Caption         =   "Nickname :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000C&
      Caption         =   "Author : Notron"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000C&
      Caption         =   "Version : 1.0.1.9"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   1335
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Integer

Private Sub Command1_Click()
connCounter.Enabled = True
' Do the buttons
txtNick.Enabled = False
txtPort.Enabled = False
Command1.Enabled = False
Command2.Enabled = True
    
' Activate the listen of sendfile form
frmSendFile.cmdConnect_Click

' Connect sockets and start listening
With frmMain
    .Winsock1(0).LocalPort = txtPort.Text
    .Winsock1(0).Listen
' frmMain.StatusBar1.Panels(1).Text = "Status: Waiting for connections .. "
    .StatusBar1.Panels(1).Text = "Status: Connected with  " & .Winsock1.Count - 1 & " Client(s)."
End With

frmConfig.Hide

With frmChat
    .txtConver = vbCrLf & vbTab & vbTab & ":::::::::::::: Welcome to the Peach Server ::::::::::::::" & vbCrLf
    .Show
    .txtToSend.SetFocus
End With
    
End Sub

Private Sub Command2_Click()
connCounter.Enabled = False
Label2.Caption = "Offline"
Dim WiSk As Winsock
With frmMain
    For Each WiSk In .Winsock1
        If WiSk.State = sckConnected Then
            WiSk.Close
            Unload WiSk
        End If
    Next
    .Winsock1(0).Close
    .StatusBar1.Panels(1).Text = "Status: Disconnected"
End With
    
' Clear panel list
frmPanel.ListView1.ListItems.Clear

' Stop the listen of sendfile form
    frmSendFile.cmdConnect_Click
    
' Do the buttons
    txtNick.Enabled = True
    txtPort.Enabled = True
    Command1.Enabled = True
    Command2.Enabled = False
End Sub

Private Sub connCounter_Timer()
x = x + 1
Label2.Caption = SetTimeFormat(x)
End Sub

Public Function SetTimeFormat(ByVal TimeValue As Double)
On Error GoTo errorhandler
Dim seconds As Integer
Dim mins As Integer
Dim secs As Integer

    seconds = Fix(TimeValue)
    mins = Fix(TimeValue / 60)
    secs = TimeValue - (mins * 60)
    If secs < 10 Then secs = "0" & secs
    SetTimeFormat = "Online : " & mins & " minutes " & secs & " seconds"
    Exit Function
errorhandler:
SetTimeFormat = "0:00"
End Function

Private Sub Form_Load()
x = 0
Me.Top = 0
Me.Left = 0

With frmMain
    Label5.Caption = "IP : " & .Winsock1(0).LocalIP
    Label6.Caption = "Name : " & .Winsock1(0).LocalHostName
End With

End Sub

Private Sub txtNick_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub
Public Function TimeString(seconds As Long, Optional Verbose _
As Boolean = False) As String

'if verbose = false, returns
'something like
'02:22.08
'if true, returns
'2 hours, 22 minutes, and 8 seconds

Dim lHrs As Long
Dim lMinutes As Long
Dim lSeconds As Long

lSeconds = seconds

lHrs = Int(lSeconds / 3600)
lMinutes = (Int(lSeconds / 60)) - (lHrs * 60)
lSeconds = Int(lSeconds Mod 60)

Dim sAns As String


If lSeconds = 60 Then
    lMinutes = lMinutes + 1
    lSeconds = 0
End If

If lMinutes = 60 Then
    lMinutes = 0
    lHrs = lHrs + 1
End If

sAns = Format(CStr(lHrs), "#####0") & ":" & _
  Format(CStr(lMinutes), "00") & "." & _
  Format(CStr(lSeconds), "00")

If Verbose Then sAns = TimeStringtoEnglish(sAns)
TimeString = sAns

End Function

Private Function TimeStringtoEnglish(sTimeString As String) As String

Dim sAns As String
Dim sHour, sMin As String, sSec As String
Dim iTemp As Integer, sTemp As String
Dim iPos As Integer
iPos = InStr(sTimeString, ":") - 1

sHour = Left$(sTimeString, iPos)
If CLng(sHour) <> 0 Then
    sAns = CLng(sHour) & " hour"
    If CLng(sHour) > 1 Then sAns = sAns & "s"
    sAns = sAns & ", "
End If

sMin = Mid$(sTimeString, iPos + 2, 2)

iTemp = sMin

If sMin = "00" Then
   sAns = IIf(Len(sAns), sAns & "0 minutes, and ", "")
Else
   sTemp = IIf(iTemp = 1, " minute", " minutes")
   sTemp = IIf(Len(sAns), sTemp & ", and ", sTemp & " and ")
   sAns = sAns & Format$(iTemp, "##") & sTemp
End If

iTemp = Val(Right$(sTimeString, 2))
sSec = Format$(iTemp, "#0")
sAns = sAns & sSec & " second"
If iTemp <> 1 Then sAns = sAns & "s"

TimeStringtoEnglish = sAns

End Function
