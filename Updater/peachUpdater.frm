VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Peach Updater"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "peachUpdater.frx":0000
   LinkTopic       =   "peachUpdater"
   MaxButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox Text1 
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3201
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"peachUpdater.frx":08CA
   End
   Begin InetCtlsObjects.Inet inetftp 
      Left            =   4440
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Open Peach"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   " Changes in latest revision:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32" ()

Dim CurRev As String
Dim NewRev As String

Dim FileLength
Dim t As String

Private Sub Command1_Click()
On Error Resume Next
Shell App.Path & "\peachClient.exe", vbNormalFocus
End Sub

Private Sub Form_Initialize()
Call InitCommonControls
End Sub

Private Sub Form_Load()
'Read current revision from registry
CurRev = ReadFromRegistry("Client\Revision", "Number")

CurRev = Left(CurRev, 3)

'Download new revision text file.
StartDownload "http://riplegion.ri.funpic.de/Peach/update.conf", App.Path & "\update.conf"

'Load new revision text file into variable
Open App.Path & "\update.conf" For Input As #1
   FileLength = LOF(1)
   t = Input(FileLength, #1)
Close #1
NewRev = Trim(t)

'Download changelog
StartDownload "http://riplegion.ri.funpic.de/Peach/pCHANGELOG.txt", App.Path & "\pCHANGELOG.txt"

'Load new revision text file into variable
Open App.Path & "\pCHANGELOG.txt" For Input As #1
   FileLength = LOF(1)
   t = Input(FileLength, #1)
Close #1
Text1.Text = Trim(t)

'Compare and check if there is an newer version or not.
Select Case CurRev
Case NewRev
    Label2.Caption = "Your Peach is up to date."
    Exit Sub
    
Case Is > NewRev
    Label2.Caption = "An unkown error occured, please try to update later."
    Exit Sub
    
Case Is < NewRev
    Label2.Caption = "Installing new version."

End Select

'Close Peach if open.
Killapp "peachClient.exe"

'Sleep it here
Label2.Caption = "Updating .."

'Download the new version *.exe
StartDownload "http://riplegion.ri.funpic.de/Peach/peachClient.exe", App.Path & "\peachClient.exe"

'Delete current file
Kill App.Path & "\update.conf"

'Rewrite into registry
InsertIntoRegistry "Client\Number", "Number", NewRev

'Update label
Label2.Caption = "Your Peach has updated from [" & CurRev & "] to [" & NewRev & "]"

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Kill App.Path & "\update.conf"
End Sub
