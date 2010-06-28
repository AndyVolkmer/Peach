VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOfflineMessages 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   ClientHeight    =   5220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7635
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvOfflineMessages 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "from"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "to"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "message"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "time"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmOfflineMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Top = 0: Left = 0
End Sub

Public Sub AddOfflineMessage(fromUser As String, toUser As String, Message As String, Time As String)
With lvOfflineMessages.ListItems
    .Add , , fromUser
    .Item(.Count).SubItems(INDEX_TO) = toUser
    .Item(.Count).SubItems(INDEX_MESSAGE) = Message
    .Item(.Count).SubItems(INDEX_TIME_SENT) = Time
End With

pDB.ExecuteCommand "INSERT INTO " & DATABASE_TABLE_OFFLINE_MESSAGES & " (from1, to1, message1, time_sent1) VALUES ('" & fromUser & "', '" & toUser & "', '" & Replace(Message, "'", "\'") & "', '" & Time & "') "
End Sub

Public Sub SendAllOfflineMessages(User As String, Index As Integer)
Dim i As Long

With lvOfflineMessages.ListItems
    For i = 1 To .Count
        If i > .Count Then Exit For
        If .Item(i).SubItems(INDEX_TO) = User Then
            SendSingle "!pmessage#offline_message#" & .Item(i) & "#" & .Item(i).SubItems(INDEX_MESSAGE) & "#" & .Item(i).SubItems(INDEX_TIME_SENT) & "#", Index
            .Remove i
            i = i - 1
        End If
    Next i
End With
End Sub

Public Sub RemoveAllOfflineMessagesFromAndToUser(User As String)
Dim i As Long

With lvOfflineMessages.ListItems
    'Search for the user in from row
    For i = 1 To .Count
        If i > .Count Then Exit For
        If .Item(i) = User Then
            .Remove (i)
            i = i - 1
        End If
    Next i

    'Search for the user in to row
    For i = 1 To .Count
        If i > .Count Then Exit For
        If .Item(i).SubItems(INDEX_TO) = User Then
            .Remove (i)
            i = i - 1
        End If
    Next i
End With

pDB.ExecuteCommand "DELETE FROM " & DATABASE_TABLE_OFFLINE_MESSAGES & " WHERE from1 = '" & User & "'"
pDB.ExecuteCommand "DELETE FROM " & DATABASE_TABLE_OFFLINE_MESSAGES & " WHERE to1 = '" & User & "'"
End Sub
