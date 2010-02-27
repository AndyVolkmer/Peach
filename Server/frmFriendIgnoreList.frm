VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFriendIgnoreList 
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
   Icon            =   "frmFriendIgnoreList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView ListView1 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3625
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
         Object.Width           =   4154
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   4154
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Friend"
         Object.Width           =   4154
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3625
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
         Object.Width           =   4154
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   4154
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Ignore"
         Object.Width           =   4154
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Ignores:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Friends:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "frmFriendIgnoreList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Top = 0: Me.Left = 0
End Sub

Public Sub AddFriend(pUser As String, pFriend As String, pIndex As Integer)
Dim j As Long

'Check if you are trying to add urself
If LCase$(pUser) = LCase$(pFriend) Then
    SendSingle "!msgbox#MSG_CANT_ADD_YOU#", pIndex
    Exit Sub
End If

'Check if friends account exist
With frmAccountPanel.ListView1.ListItems
    For i = 1 To .Count
        If LCase(.Item(i).SubItems(1)) = LCase(pFriend) Then
            pFriend = .Item(i).SubItems(1)
            Exit For
        Else
            If i = .Count Then
                SendSingle "!msgbox#MSG_ACCOUNT_NOT_EXIST#'" & pFriend & "#", pIndex
                Exit Sub
            End If
        End If
    Next i
End With

'Check if the account is already added
With ListView1.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(1) = pUser Then
            If .Item(i).SubItems(2) = pFriend Then
                SendSingle "!msgbox#MSG_ALREADY_IN_FRIEND_LIST#'" & pFriend & "#", pIndex
                Exit Sub
            End If
        End If
    Next i
End With

j = pDB.GetMaxID("ID", DATABASE_TABLE_FRIENDS)

If pDB.ExecuteCommand("INSERT INTO " & DATABASE_TABLE_FRIENDS & " (ID, Name, Friend) VALUES('" & j & "', '" & pUser & "', '" & pFriend & "')") Then
    'Add relation to listview
    With ListView1.ListItems
        .Add , , j
        i = .Count
        .Item(i).SubItems(1) = pUser
        .Item(i).SubItems(2) = pFriend
    End With
    
    UPDATE_FRIEND pUser, pIndex
End If
End Sub

Public Sub RemoveFriend(pUser As String, pFriend As String, pIndex As Integer)
Dim pID As Integer

With ListView1.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(1) = pUser And .Item(i).SubItems(2) = pFriend Then
            pID = .Item(i)
            .Remove (i)
            Exit For
        End If
    Next i
End With

'This seems to be very hacky
If pDB.ExecuteCommand("DELETE FROM " & DATABASE_TABLE_FRIENDS & " WHERE ID = " & pID) Then
    UPDATE_FRIEND pUser, pIndex
End If
End Sub

Public Sub RemoveAllFriendsFromUser(pUser As String)
If pDB.ExecuteCommand("DELETE FROM " & DATABASE_TABLE_FRIENDS & " WHERE Name = '" & pUser & "'") And _
   pDB.ExecuteCommand("DELETE FROM " & DATABASE_TABLE_FRIENDS & " WHERE Friend = '" & pUser & "'") Then
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
End If
End Sub

Public Sub AddIgnore(pUser As String, pIgnore As String, pIndex As Integer)
Dim j As Long

'Check if you are trying to add urself
If LCase$(pUser) = LCase$(pIgnore) Then
    SendSingle "!msgbox#MSG_CANT_ADD_YOU#", pIndex
    Exit Sub
End If

'Check if friends account exist
With frmAccountPanel.ListView1.ListItems
    For i = 1 To .Count
        If LCase(.Item(i).SubItems(1)) = LCase(pIgnore) Then
            pIgnore = .Item(i).SubItems(1)
            Exit For
        Else
            If i = .Count Then
                'Send message that the account doesn't exist
                SendSingle "!msgbox#MSG_ACCOUNT_NOT_EXIST#'" & pIgnore & "#", pIndex
                Exit Sub
            End If
        End If
    Next i
End With

'Check if the account is already added
With ListView2.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(1) = pUser Then
            If .Item(i).SubItems(2) = pIgnore Then
                SendSingle "!msgbox#MSG_ALREADY_IN_IGNORE_LIST#'" & pIgnore & "#", pIndex
                Exit Sub
            End If
        End If
    Next i
End With

j = pDB.GetMaxID("ID", DATABASE_TABLE_IGNORES)

If pDB.ExecuteCommand("INSERT INTO " & DATABASE_TABLE_IGNORES & " (ID, Name, IgnoredName) VALUES('" & j & "', '" & pUser & "', '" & pIgnore & "')") Then
    'Add relation to listview
    With ListView2.ListItems
        .Add , , j
        i = .Count
        .Item(i).SubItems(1) = pUser
        .Item(i).SubItems(2) = pIgnore
    End With
    
    UPDATE_IGNORE pUser, pIndex
End If
End Sub

Public Sub RemoveIgnore(pUser As String, pIgnore As String, pIndex As Integer)
Dim pID As Integer

With ListView2.ListItems
    For i = 1 To .Count
        If .Item(i).SubItems(1) = pUser Then
            If .Item(i).SubItems(2) = pIgnore Then
                pID = .Item(i)
                .Remove (i)
                Exit For
            End If
        End If
    Next i
End With

'Another hacky solution
If pDB.ExecuteCommand("DELETE FROM " & DATABASE_TABLE_IGNORES & " WHERE ID = " & pID) Then
    UPDATE_IGNORE pUser, pIndex
End If
End Sub

Public Sub RemoveAllIgnoresFromUser(pUser As String)
If pDB.ExecuteCommand("DELETE FROM " & DATABASE_TABLE_IGNORES & " WHERE Name = '" & pUser & "'") And _
   pDB.ExecuteCommand("DELETE FROM " & DATABASE_TABLE_IGNORES & " WHERE IgnoredName = '" & pUser & "'") Then
    With ListView2.ListItems
        'Search for the user in Name row
        For i = 1 To .Count
            If i > .Count Then Exit For
            If .Item(i).SubItems(1) = pUser Then
                .Remove (i)
                i = i - 1
            End If
        Next i
        
        'Search for the user in Ignore row
        For i = 1 To .Count
            If i > .Count Then Exit For
            If .Item(i).SubItems(2) = pUser Then
                .Remove (i)
                i = i - 1
            End If
        Next i
    End With
End If
End Sub
