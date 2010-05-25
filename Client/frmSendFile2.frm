VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSendFile2 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Peach - File Transfer"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSendFile2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   7425
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Open File Folder"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   3360
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock SckReceiveFile 
      Index           =   0
      Left            =   6480
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrStatus 
      Interval        =   10
      Left            =   6000
      Top             =   2880
   End
   Begin MSComctlLib.ListView lstConnections 
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   5609
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Conn Stat."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Remote"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "File Name"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Progress"
         Object.Width           =   1411
      EndProperty
   End
End
Attribute VB_Name = "frmSendFile2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Client
    FileName        As String
    FileSize        As Long
    BytesReceived   As Long
    FileNum         As Integer
End Type

Private Clients() As Client

Private Sub Command1_Click()
Shell "Explorer.exe " & App.Path, vbNormalFocus
End Sub

Private Sub Form_Activate()
BackColor = Setting.SCHEME_COLOR
End Sub

Private Sub Form_Load()
lstConnections.ListItems.Add , , "0"
End Sub

Private Sub SckReceiveFile_Close(Index As Integer)
On Error Resume Next

SckReceiveFile(Index).Close

Close Clients(Index).FileNum

If Clients(Index).BytesReceived < Clients(Index).FileSize Then
    Kill App.Path & "\" & Clients(Index).FileName
    lstConnections.ListItems(Index + 1).SubItems(4) = "Incomplete, File deleted"
    frmChat.WriteText "File transfer incomplete, File deleted."
Else
    lstConnections.ListItems(Index + 1).SubItems(4) = "Transfer Complete"
    frmChat.WriteText "File transfer finished."
End If

FitTextInListView lstConnections, 4, , Index + 1

Clients(Index).FileNum = 0
Clients(Index).BytesReceived = 0
Clients(Index).FileSize = 0
Clients(Index).FileName = vbNullString
End Sub

Private Sub FitTextInListView(LV As ListView, ByVal Column As Integer, Optional ByVal Text As String, Optional ByVal ItemIndex As Long = -1)
Dim TLen   As Single
Dim CapLen As Single

CapLen = TextWidth(LV.ColumnHeaders(Column + 1).Text) + 195

If ItemIndex >= 0 Then
    If ItemIndex = 0 Then
        TLen = TextWidth(LV.ListItems(ItemIndex).Text)
    Else
        TLen = TextWidth(LV.ListItems(ItemIndex).SubItems(Column))
    End If
Else
    TLen = TextWidth(Text)
End If

TLen = TLen + 195

If CapLen > TLen Then TLen = CapLen

If LV.ColumnHeaders(Column + 1).Width < TLen Then LV.ColumnHeaders(Column + 1).Width = TLen
End Sub

Private Sub SckReceiveFile_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim i   As Long
Dim LI  As ListItem

For i = 1 To SckReceiveFile.UBound
    If SckReceiveFile(i).State = sckClosed Then Exit For
Next i

If i = SckReceiveFile.UBound + 1 Then
    Load SckReceiveFile(SckReceiveFile.UBound + 1)
    ReDim Preserve Clients(SckReceiveFile.UBound)

    i = SckReceiveFile.UBound
    lstConnections.ListItems.Add , , CStr(i)
End If

SckReceiveFile(i).Accept requestID

If LenB(SckReceiveFile(i).RemoteHost) = 0 Then
    lstConnections.ListItems(i + 1).SubItems(2) = SckReceiveFile(i).RemoteHostIP
Else
    lstConnections.ListItems(i + 1).SubItems(2) = SckReceiveFile(i).RemoteHost
End If

FitTextInListView lstConnections, 2, , i + 1
End Sub

Private Sub SckReceiveFile_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Data As String
Dim Pos  As Long
Dim Pos2 As Long

SckReceiveFile(Index).GetData Data, vbString

If Clients(Index).FileSize = 0 And InStr(1, Data, ":") > 0 Then
    Pos = InStr(1, Data, ",")

    Clients(Index).FileSize = Val(Left(Data, Pos - 1))
    Pos2 = InStr(Pos, Data, ":")
    Clients(Index).FileName = Mid(Data, Pos + 1, (Pos2 - Pos) - 1)

    Clients(Index).FileNum = FreeFile
    On Error GoTo handleErrorSendFile:
    Open App.Path & "\" & Clients(Index).FileName For Binary Access Write Lock Write As Clients(Index).FileNum

    Data = Mid(Data, Pos2 + 1)

    lstConnections.ListItems(Index + 1).SubItems(3) = Clients(Index).FileName
    FitTextInListView lstConnections, 3, , Index + 1
End If

If LenB(sData) > 0 Then
    Clients(Index).BytesReceived = Clients(Index).BytesReceived + Len(Data)
    Put Clients(Index).FileNum, , Data

    lstConnections.ListItems(Index + 1).SubItems(4) = Format$(Clients(Index).BytesReceived / Clients(Index).FileSize * 100#, "#0.00") & " %"
    FitTextInListView lstConnections, 4, , Index + 1

    If Clients(Index).BytesReceived >= Clients(Index).FileSize Then
        SckReceiveFile_Close Index
    End If
End If
Exit Sub
handleErrorSendFile:
Select Case Err.Number
    Case 75
        MsgBox "The file you are trying to get already exists in this location and is ReadOnly. Rename it and try to send again." & vbCrLf & "Current action aborted due to ReadOnly file.", vbInformation
        Exit Sub

    Case Else
        MsgBox "Error : " & Err.Number & vbCrLf & "Description : " & Err.Description & vbCrLf & "Current action aborted because of an unkown error!", vbCritical
        Exit Sub

End Select
End Sub

Private Sub tmrStatus_Timer()
Dim i      As Long
Dim TmpStr As String

For i = 0 To SckReceiveFile.UBound
    TmpStr = Choose(SckReceiveFile(i).State + 1, "Closed", "Open", "Listening", "Connection pending", "Resolving host", "Host resolved", "Connecting", "Connected", "Server is disconnecting", "Error")

    If lstConnections.ListItems(i + 1).SubItems(1) <> TmpStr Then
        lstConnections.ListItems(i + 1).SubItems(1) = TmpStr
        FitTextInListView lstConnections, 1, , i + 1
    End If
Next i
End Sub
