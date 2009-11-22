VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFriendList 
   BorderStyle     =   0  'None
   Caption         =   "Friend List Overview"
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
   Icon            =   "frmFriendList.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView ListView1 
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   4207
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   4260
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Friend"
         Object.Width           =   4260
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
Dim IsValid    As Boolean
Dim j          As Long

'Check if you are trying to add urself
If LCase$(pUser) = LCase$(pFriend) Then
    SendSingle "!msgbox#You can't add yourself.#", frmMain.Winsock1(pIndex)
    Exit Sub
End If

'Check if friends account exist
With frmAccountPanel.ListView1.ListItems
    For i = 1 To .Count
        If LCase(.Item(i).SubItems(1)) = LCase(pFriend) Then
            pFriend = .Item(i).SubItems(1)
            IsValid = True
            Exit For
        End If
    Next i
End With

'Check if the account is already added
With ListView1.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(1) = pUser Then
            If .Item(i).SubItems(2) = pFriend Then
                SendSingle "!msgbox#" & pFriend & " is already in you friendlist.#", frmMain.Winsock1(pIndex)
                Exit Sub
            End If
        End If
    Next i
End With

'If friends account exist then process
If IsValid Then
    j = 0
    With ListView1.ListItems
        For i = 1 To .Count
            If .Item(i) > j Then
                j = .Item(i)
            End If
        Next i
    End With
    
    j = j + 1
    
    'Execute into database
    With frmMain
        .xCommand.CommandText = "INSERT INTO " & Database.Friend_Table & " (ID, Name, Friend) VALUES('" & j & "', '" & pUser & "', '" & pFriend & "')"
        .xCommand.Execute
    End With
    
    'Add relation to listview
    With ListView1.ListItems
        .Add , , j
        i = ListView1.ListItems.Count
        .Item(i).SubItems(1) = pUser
        .Item(i).SubItems(2) = pFriend
    End With
    
    UPDATE_FRIEND pUser, pIndex
    
Else
    'Send message that the account doesn't exist
    SendSingle "!msgbox#" & "The account '" & pFriend & "' doesnt exist.#", frmMain.Winsock1(pIndex)
End If
End Sub

Public Sub RemoveFriend(pUser As String, pFriend As String, pIndex As Integer)
Dim pID As Integer

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
    .xCommand.CommandText = "DELETE FROM " & Database.Friend_Table & " WHERE ID = " & pID
    .xCommand.Execute
End With

UPDATE_FRIEND pUser, pIndex
End Sub

Public Sub DeleteFriend(pUser As String)

With ListView1.ListItems
    'Search for the user in Name row
    For i = 1 To .Count
        If i > .Count Then Exit For
        If .Item(i).SubItems(1) = pUser Then
            .Remove (i)
            i = i - 1
        End If
    Next i

    'Search for the user in Friend row
    For i = 1 To .Count
        If i > .Count Then Exit For
        If .Item(i).SubItems(2) = pUser Then
            .Remove (i)
            i = i - 1
        End If
    Next i
End With

'Delete user from database
With frmMain.xCommand
    .CommandText = "DELETE FROM " & Database.Friend_Table & " WHERE Name ='" & pUser & "'"
    .Execute
    .CommandText = "DELETE FROM " & Database.Friend_Table & " WHERE Friend ='" & pUser & "'"
    .Execute
End With
End Sub
