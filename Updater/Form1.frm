VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Peach Updater"
   ClientHeight    =   3780
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
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox Text1 
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4683
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":08CA
   End
   Begin InetCtlsObjects.Inet inetftp 
      Left            =   4440
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Open Peach"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   " Changes in latest revision:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Sub InitCommonControls Lib "comctl32" ()

Dim CurRev As String
Dim NewRev As String

Dim FileLength
Dim f As String
Dim t As String

Private Sub Command1_Click()
On Error Resume Next
Shell App.Path & "\peachClient.exe"
End Sub

Private Sub Form_Initialize()
Call InitCommonControls
End Sub

Private Sub Form_Load()

On Error GoTo HandleError

'Read current revision from text file.
f = App.Path & "\revision.conf"
Open f For Input As #1
   FileLength = LOF(1)
   t = Input(FileLength, #1)
Close #1
CurRev = Trim(t)

CurRev = Left(CurRev, 7)

'Download new revision text file.
StartDownload "http://riplegion.ri.funpic.de/Peach/update.conf", App.Path & "\update.conf"

'Load new revision text file into variable
f = App.Path & "\update.conf"
Open f For Input As #1
   FileLength = LOF(1)
   t = Input(FileLength, #1)
Close #1
NewRev = Trim(t)

'Download changelog
StartDownload "http://riplegion.ri.funpic.de/Peach/pCHANGELOG.txt", App.Path & "\pCHANGELOG.txt"

'Load new revision text file into variable
f = App.Path & "\pCHANGELOG.txt"
Open f For Input As #1
   FileLength = LOF(1)
   t = Input(FileLength, #1)
Close #1
Text1.Text = Trim(t)

'Compare and check if there is an newer version or not.
If CurRev = NewRev Then
    Label2.Caption = "Your peach is up to date."
    Exit Sub
    
ElseIf CurRev > NewRev Then
    Label2.Caption = "An unkown error occured, please try to update later."
    Exit Sub
    
ElseIf CurRev < NewRev Then
    Label2.Caption = "New version is installing."

End If

'Close Peach if open.
Killapp "peachClient.exe"

'Sleep it here
Sleep 1000

'Download the new version *.exe
StartDownload "http://riplegion.ri.funpic.de/Peach/peachClient.exe", App.Path & "\peachClient.exe"

'Delete current file
Kill App.Path & "\revision.conf"

'Rename update to revision
Name App.Path & "\update.conf" As App.Path & "\revision.conf"

'Update label
Label2.Caption = "Your Peach has updated from [" & CurRev & "] to [" & NewRev & "]"

'************************************************
Exit Sub
HandleError:
Select Case Err.Number
Case 53
    MsgBox "The file 'revision.conf' is missing. Please open Peach to replace it.", vbInformation
    Unload Me
Case Else
    MsgBox "Error : " & Err.Number & vbCrLf & Err.Description
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Kill App.Path & "\update.conf"
End Sub
