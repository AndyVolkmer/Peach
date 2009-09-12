VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFriendList 
   BorderStyle     =   0  'None
   Caption         =   "Friend List Overview"
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8535
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFriendList.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3975
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView ListView1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4683
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
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Friend"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmFriendList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Top = 0: Me.Left = 0
End Sub

Public Sub AddFriend(pUser As String, pFriend As String, pIndex As Integer)
Dim IsValid  As Boolean
Dim f_table  As String

'Check if friends account exist
With frmAccountPanel.ListView1.ListItems
    For i = 1 To .Count
        If LCase(.Item(i).SubItems(1)) = LCase(pFriend) Then
            IsValid = True
            Exit For
        End If
    Next i
End With

'Check if the account is already added
With ListView1.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(2) = pFriend Then
            SendSingle "!friend_already_added#" & pFriend & "#", frmMain.Winsock1(pIndex)
            Exit Sub
        End If
    Next i
End With

'If friends account exist then process
If IsValid = True Then
    'Get friends table name
    f_table = ReadIniValue(App.Path & "\Config.ini", "Database", "F_Table")
    
    Dim j As Integer
    j = 0
    With ListView1.ListItems
        For i = 1 To .Count
            If .Item(i) = j Then
                j = i + 1
            End If
        Next i
    End With
    
    'Execute into database
    With frmMain
        .xCommand.CommandText = "INSERT INTO " & f_table & " (ID, Name, Friend) VALUES('" & j & "', '" & pUser & "', '" & pFriend & "')"
        .xCommand.Execute
    End With
    
    'Save index in variable
    i = ListView1.ListItems.Count + 1

    'Add relation to listview
    With ListView1.ListItems
        .Add , , j
        .Item(i).SubItems(1) = pUser
        .Item(i).SubItems(2) = pFriend
    End With
    UpdateFriendList pUser
Else 'Send message that the accoutn doesnt exist
    SendSingle "!friend_not_exist#" & pFriend & "#", frmMain.Winsock1(pIndex)
End If
End Sub

Public Sub RemoveFriend(pUser As String, pFriend As String, pIndex As Integer)
Dim f_table As String
Dim pID     As Integer

f_table = ReadIniValue(App.Path & "\Config.ini", "Database", "F_Table")

With ListView1.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(1) = pUser Then
            If .Item(i).SubItems(2) = pFriend Then
                pID = .Item(i)
                .Remove (i)
                Exit For
            End If
        End If
    Next i
End With

With frmMain
    .xCommand.CommandText = "DELETE FROM " & f_table & " WHERE ID = " & pID
    .xCommand.Execute
End With
UpdateFriendList pUser
End Sub
