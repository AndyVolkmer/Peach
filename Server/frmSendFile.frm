VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSendFile 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   Caption         =   "frmSendFile"
   ClientHeight    =   4065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7470
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
   MDIChild        =   -1  'True
   ScaleHeight     =   4065
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrStatus 
      Interval        =   10
      Left            =   2775
      Top             =   240
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Start Listening for Connections"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   2940
   End
   Begin MSComctlLib.ListView lstConnections 
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   5609
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
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
   Begin MSWinsockLib.Winsock SckReceiveFile 
      Index           =   0
      Left            =   2235
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSendFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tClient
    FileName As String
    FileSize As Long
    BytesReceived As Long
    
    FileNum As Integer
End Type

Private Clients() As tClient

Private Sub cmdConnect_Click()
    If cmdConnect.Caption = "Start Listening for Connections" Then
        SckReceiveFile(0).LocalPort = iPort
        SckReceiveFile(0).Listen
        cmdConnect.Caption = "Stop Listening for Connections"
    Else
        SckReceiveFile(0).Close
        cmdConnect.Caption = "Start Listening for Connections"
    End If
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    lstConnections.ListItems.Add , , "0"
End Sub

Private Sub lstConnections_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub SckReceiveFile_Close(Index As Integer)
    On Error Resume Next
    
    SckReceiveFile(Index).Close
    
    Close Clients(Index).FileNum
    
    If Clients(Index).BytesReceived < Clients(Index).FileSize Then
        Kill App.Path & "\" & Clients(Index).FileName
        
        Me.lstConnections.ListItems(Index + 1).SubItems(4) = "Incomplete, File Deleted"
    Else
        Me.lstConnections.ListItems(Index + 1).SubItems(4) = "Transfer Complete"
    End If
    FitTextInListView Me.lstConnections, 4, , Index + 1
    
    Clients(Index).FileNum = 0
    Clients(Index).BytesReceived = 0
    Clients(Index).FileSize = 0
    Clients(Index).FileName = ""
End Sub

Private Sub FitTextInListView(LV As ListView, ByVal Column As Integer, Optional ByVal Text As String, Optional ByVal ItemIndex As Long = -1)
    Dim TLen As Single, CapLen As Single
    
    CapLen = Me.TextWidth(LV.ColumnHeaders(Column + 1).Text) + 195
    
    If ItemIndex >= 0 Then
        If ItemIndex = 0 Then
            TLen = Me.TextWidth(LV.ListItems(ItemIndex).Text)
        Else
            TLen = Me.TextWidth(LV.ListItems(ItemIndex).SubItems(Column))
        End If
    Else
        TLen = Me.TextWidth(Text)
    End If
    
    TLen = TLen + 195
    
    If CapLen > TLen Then TLen = CapLen
    
    If LV.ColumnHeaders(Column + 1).Width < TLen Then LV.ColumnHeaders(Column + 1).Width = TLen
End Sub

Private Sub SckReceiveFile_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim K As Integer, LI As ListItem
    
    For K = 1 To SckReceiveFile.UBound
        If SckReceiveFile(K).State = sckClosed Then Exit For
    Next K
    
    If K = SckReceiveFile.UBound + 1 Then
        Load SckReceiveFile(SckReceiveFile.UBound + 1)
        ReDim Preserve Clients(SckReceiveFile.UBound)
        
        K = SckReceiveFile.UBound
        lstConnections.ListItems.Add , , CStr(K)
    End If
    
    SckReceiveFile(K).Accept requestID
    
    If Len(SckReceiveFile(K).RemoteHost) = 0 Then
        Me.lstConnections.ListItems(K + 1).SubItems(2) = SckReceiveFile(K).RemoteHostIP
    Else
        Me.lstConnections.ListItems(K + 1).SubItems(2) = SckReceiveFile(K).RemoteHost
    End If
    
    FitTextInListView Me.lstConnections, 2, , K + 1
End Sub

Private Sub SckReceiveFile_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim sData As String, Pos As Long, Pos2 As Long
    
    SckReceiveFile(Index).GetData sData, vbString
    
    If Clients(Index).FileSize = 0 And InStr(1, sData, ":") > 0 Then
        Pos = InStr(1, sData, ",")
        
        Clients(Index).FileSize = Val(Left(sData, Pos - 1))
        Pos2 = InStr(Pos, sData, ":")
        Clients(Index).FileName = Mid(sData, Pos + 1, (Pos2 - Pos) - 1)
        
        Clients(Index).FileNum = FreeFile
        On Error GoTo handleErrorSendFile:
        Open App.Path & "\" & Clients(Index).FileName For Binary Access Write Lock Write As Clients(Index).FileNum
        
        sData = Mid(sData, Pos2 + 1)
        
        Me.lstConnections.ListItems(Index + 1).SubItems(3) = Clients(Index).FileName
        FitTextInListView Me.lstConnections, 3, , Index + 1
    End If
    
    If Len(sData) > 0 Then
        Clients(Index).BytesReceived = Clients(Index).BytesReceived + Len(sData)
        Put Clients(Index).FileNum, , sData
        
        Me.lstConnections.ListItems(Index + 1).SubItems(4) = Format(Clients(Index).BytesReceived / Clients(Index).FileSize * 100#, "#0.00") & " %"
        FitTextInListView Me.lstConnections, 4, , Index + 1
        
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
    Dim K As Long, TmpStr As String
    
    For K = 0 To SckReceiveFile.UBound
        TmpStr = Choose(SckReceiveFile(K).State + 1, "Closed", "Open", "Listening", "Connection pending", "Resolving host", "Host resolved", "Connecting", "Connected", "Server is disconnecting", "Error")
        
        If Me.lstConnections.ListItems(K + 1).SubItems(1) <> TmpStr Then
            Me.lstConnections.ListItems(K + 1).SubItems(1) = TmpStr
            FitTextInListView Me.lstConnections, 1, , K + 1
        End If
    Next K
End Sub
