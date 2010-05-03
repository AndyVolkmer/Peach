VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmChannel 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   Caption         =   "frmChannel"
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
   MDIChild        =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvChannels 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   360
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
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Password"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Owner"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "JoinAnnounce"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvUsers 
      Height          =   2055
      Left            =   120
      TabIndex        =   2
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
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Channel"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "IsOwner?"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Users:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Channels:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Top = 0: Me.Left = 0
End Sub

Public Sub JoinChannel(Channel As String, User As String)
Dim i As Long
Dim j As Long

'If channel exist join else create new one
With lvChannels.ListItems
    For i = 1 To .Count
        If LCase$(.Item(i)) = LCase$(Channel) Then
            'Join Channel
            With lvUsers.ListItems
                .Add , , User
                .Item(.Count).SubItems(CHANNEL_USER_CHANNEL) = lvChannels.ListItems.Item(i)
                .Item(.Count).SubItems(CHANNEL_USER_IS_OWNER) = "0"
            End With
            
            If .Item(i).SubItems(CHANNEL_JOIN_ANNOUNCE) = "1" Then SendMessageToChannel Channel, User, "[" & .Item(i) & "] " & User & " has joined the channel."
            Exit Sub
        End If
    Next i
End With

'Create Channel
With lvChannels.ListItems
    .Add , , Channel
    .Item(.Count).SubItems(CHANNEL_INDEX.CHANNEL_OWNER) = User
    .Item(.Count).SubItems(CHANNEL_INDEX.CHANNEL_JOIN_ANNOUNCE) = "1"
End With

'Join Channel
With lvUsers.ListItems
    .Add , , User
    .Item(.Count).SubItems(CHANNEL_USER_CHANNEL) = Channel
    .Item(.Count).SubItems(CHANNEL_USER_IS_OWNER) = "1"
End With
    
With frmPanel.ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) = User Then
            SendSingle "[" & Channel & "] You joined the channel.", .Item(i).SubItems(INDEX_WINSOCK_ID)
            Exit For
        End If
    Next i
End With
End Sub

Public Sub LeaveChannel(Channel As String, User As String)
Dim i               As Long
Dim Counter         As Long
Dim JoinAnnounce    As Boolean

With lvChannels.ListItems
    For i = 1 To .Count
        If LCase$(.Item(i)) = LCase$(Channel) Then
            If .Item(i).SubItems(CHANNEL_JOIN_ANNOUNCE) = "1" Then JoinAnnounce = True
            Exit For
        End If
    Next i
End With

With lvUsers.ListItems
    For i = 1 To .Count
        If .Item(i) = User And LCase$(.Item(i).SubItems(CHANNEL_USER_CHANNEL)) = LCase$(Channel) Then
            'If announce is enabled then tell channel that user left
            If JoinAnnounce Then SendMessageToChannel Channel, User, "[" & .Item(i).SubItems(CHANNEL_USER_CHANNEL) & "] " & User & " left channel."
            
            'Set owner to someone else when owner is leaving
            If .Item(i).SubItems(CHANNEL_USER_IS_OWNER) = "1" And .Count > 1 Then
                .Item(2).SubItems(CHANNEL_USER_IS_OWNER) = "1"
            End If
            
            'Remove user
            .Remove i
            Exit For
        End If
    Next i
    
    'Counter for channel if it's empty then we can disband later
    For i = 1 To .Count
        If LCase$(.Item(i).SubItems(CHANNEL_USER_CHANNEL)) = LCase$(Channel) Then
            Counter = Counter + 1
        End If
    Next i
End With

'If channel is empty then delete it
If Counter = 0 Then
    With lvChannels.ListItems
        For i = 1 To .Count
            If LCase$(.Item(i)) = LCase$(Channel) Then
                .Remove i
            End If
        Next i
    End With
End If
End Sub
