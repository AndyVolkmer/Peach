VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSociety 
   BorderStyle     =   0  'None
   Caption         =   "frmSociety"
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7905
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7300
      _ExtentX        =   12885
      _ExtentY        =   7011
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   16053492
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Friend List"
      TabPicture(0)   =   "frmSociety.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lvFriendList"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdAddFriend"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdRemoveFriend"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Online List"
      TabPicture(1)   =   "frmSociety.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvOnlineList"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdAddToFriend"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdAddToIgnore"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Ignore List"
      TabPicture(2)   =   "frmSociety.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvIgnoreList"
      Tab(2).Control(1)=   "cmdAddIgnore"
      Tab(2).Control(2)=   "cmdRemoveIgnore"
      Tab(2).ControlCount=   3
      Begin VB.CommandButton cmdAddToIgnore 
         Caption         =   "&Add to Ignore"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -69960
         TabIndex        =   9
         Top             =   3360
         Width           =   2055
      End
      Begin VB.CommandButton cmdAddToFriend 
         Caption         =   "&Add to Friends"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -72000
         TabIndex        =   8
         Top             =   3360
         Width           =   2055
      End
      Begin VB.CommandButton cmdRemoveIgnore 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   -69960
         TabIndex        =   6
         Top             =   3360
         Width           =   2055
      End
      Begin VB.CommandButton cmdAddIgnore 
         Caption         =   "&Add"
         Height          =   375
         Left            =   -72000
         TabIndex        =   5
         Top             =   3360
         Width           =   2055
      End
      Begin VB.CommandButton cmdRemoveFriend 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   5040
         TabIndex        =   2
         Top             =   3360
         Width           =   2055
      End
      Begin VB.CommandButton cmdAddFriend 
         Caption         =   "&Add"
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   3360
         Width           =   2055
      End
      Begin MSComctlLib.ListView lvFriendList 
         Height          =   2775
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   7050
         _ExtentX        =   12435
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   9790
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvOnlineList 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   7055
         _ExtentX        =   12435
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   12321
         EndProperty
      End
      Begin MSComctlLib.ListView lvIgnoreList 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   7055
         _ExtentX        =   12435
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   12330
         EndProperty
      End
   End
   Begin VB.Menu lvFriendMenu 
      Caption         =   "lvFriendMenu"
      Visible         =   0   'False
      Begin VB.Menu mWhisper 
         Caption         =   "&Whisper"
      End
      Begin VB.Menu mRemoveFriend 
         Caption         =   "&Remove"
      End
   End
   Begin VB.Menu lvOnlineMenu 
      Caption         =   "lvOnlineMenu"
      Visible         =   0   'False
      Begin VB.Menu mWhisperT 
         Caption         =   "&Whisper"
      End
      Begin VB.Menu mAddToFriend 
         Caption         =   "&Add to Friends"
      End
      Begin VB.Menu mIgnoreUser 
         Caption         =   "Ignore User"
      End
   End
   Begin VB.Menu lvIgnoreMenu 
      Caption         =   "lvIgnoreMenu"
      Visible         =   0   'False
      Begin VB.Menu mRemoveIgnore 
         Caption         =   "&Remove"
      End
   End
End
Attribute VB_Name = "frmSociety"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==============================='
'======= Friend List Tab ======='
'==============================='

'======= Buttons ======='
Private Sub cmdAddFriend_Click()
Dim pVal As String
pVal = InputBox(SOC_ASK_FRIEND_TEXT, SOC_ASK_FRIEND_DEFAULT, SOC_ASK_FRIEND_DEFAULT)

If LenB(Trim$(pVal)) = 0 Then Exit Sub
SendMessage "!friend" & pSplit & "-add" & pSplit & pVal & pSplit
End Sub

Private Sub cmdRemoveFriend_Click()
With lvFriendList
    If .ListItems.Count = 0 Then Exit Sub
    pRemoveFriend .SelectedItem.Text
End With
End Sub

'======= PopUp Menu ======='
Private Sub mRemoveFriend_Click()
pRemoveFriend lvFriendList.SelectedItem.Text
End Sub

'==============================='
'======= Online List Tab ======='
'==============================='

'======= Buttons ======='
Private Sub cmdAddToFriend_Click()
With lvOnlineList
    If .ListItems.Count = 0 Then Exit Sub
    pAddToFriend .SelectedItem.Text
End With
End Sub

Private Sub cmdAddToIgnore_Click()
With lvOnlineList
    If .ListItems.Count = 0 Then Exit Sub
    pAddToIgnore .SelectedItem.Text
End With
End Sub

'======= PopUp Menu ======='
Private Sub mAddToFriend_Click()
pAddToFriend lvOnlineList.SelectedItem.Text
End Sub

Private Sub mIgnoreUser_Click()
pAddToIgnore lvOnlineList.SelectedItem.Text
End Sub

'==============================='
'======= Ignore List Tab ======='
'==============================='

'======= Buttons ======='
Private Sub cmdAddIgnore_Click()
Dim pVal As String
pVal = InputBox(SOC_ASK_IGNORE_TEXT, SOC_ASK_IGNORE_TITLE, SOC_ASK_IGNORE_DEFAULT)

If LenB(Trim$(pVal)) = 0 Then Exit Sub
SendMessage "!ignore" & pSplit & "-add" & pSplit & pVal & pSplit
End Sub

Private Sub cmdRemoveIgnore_Click()
With lvIgnoreList
    If .ListItems.Count = 0 Then Exit Sub
    pRemoveIgnore .SelectedItem.Text
End With
End Sub

'======= PopUp Menu ======='
Private Sub mRemoveIgnore_Click()
pRemoveIgnore lvIgnoreList.SelectedItem.Text
End Sub

'======================='
'====== Functions ======'
'======================='

'=== Friend List Tab ==='
Private Sub pRemoveFriend(uString As String)
If MsgBox(Replace$(SOC_ASK_DEL, "%u", uString), vbQuestion + vbYesNo) = vbYes Then
    SendMessage "!friend" & pSplit & "-remove" & pSplit & uString & pSplit
End If
End Sub

Private Sub TriggerFriendEvent()
If lvFriendList.ListItems.Count = 0 Then Exit Sub

If Len(lvFriendList.SelectedItem.SubItems(1)) = 7 Then
    mWhisper.Enabled = False
Else
    mWhisper.Enabled = True
End If

PopupMenu lvFriendMenu
End Sub

'=== Online List Tab ==='
Private Sub pAddToFriend(uString As String)
SendMessage "!friend" & pSplit & "-add" & pSplit & uString & pSplit
End Sub

Private Sub pAddToIgnore(uString As String)
SendMessage "!ignore" & pSplit & "-add" & pSplit & uString & pSplit
End Sub

Private Sub TriggerOnlineEvent()
If lvOnlineList.ListItems.Count = 0 Then Exit Sub

If lvOnlineList.SelectedItem.Text = frmMain.txtAccount.Text Then
    mWhisperT.Enabled = False
    mAddToFriend.Enabled = False
    mIgnoreUser.Enabled = False
Else
    mWhisperT.Enabled = True
    mAddToFriend.Enabled = True
    mIgnoreUser.Enabled = True
End If

PopupMenu lvOnlineMenu
End Sub

'=== Ignore List Tab ==='
Private Sub TriggerIgnoreEvent()
If lvIgnoreList.ListItems.Count = 0 Then Exit Sub
PopupMenu lvIgnoreMenu
End Sub

Private Sub pRemoveIgnore(uString As String)
If MsgBox(Replace$(SOC_ASK_DEL, "%u", uString), vbQuestion + vbYesNo) = vbYes Then
    SendMessage "!ignore" & pSplit & "-remove" & pSplit & uString & pSplit
End If
End Sub

'======================'
'======= Events ======='
'======================'

'==== Form ===='
Private Sub Form_Load()
Top = 0: Left = 0

SSTab1.TabCaption(0) = SOC_FRIEND_LIST
SSTab1.TabCaption(1) = SOC_ONLINE_LIST
SSTab1.TabCaption(2) = SOC_IGNORE_LIST

cmdAddFriend.Caption = SOC_COMMAND_FRIEND
cmdRemoveFriend.Caption = SOC_COMMAND_REMOVE
cmdAddIgnore.Caption = SOC_COMMAND_IGNORE
cmdRemoveIgnore.Caption = SOC_COMMAND_REMOVE
cmdAddToFriend.Caption = SOC_COMMAND_FRIEND
cmdAddToIgnore.Caption = SOC_COMMAND_IGNORE

lvFriendList.ColumnHeaders(1).Text = CONFIG_LABEL_ACCOUNT
lvFriendList.ColumnHeaders(2).Text = SOC_FRIEND_LIST_STATUS
lvOnlineList.ColumnHeaders(1).Text = CONFIG_LABEL_ACCOUNT
lvIgnoreList.ColumnHeaders(1).Text = CONFIG_LABEL_ACCOUNT

mWhisper.Caption = SOC_COMMAND_WHISPER
mWhisperT.Caption = SOC_COMMAND_WHISPER
mAddToFriend.Caption = SOC_COMMAND_FRIEND
mIgnoreUser.Caption = SOC_COMMAND_IGNORE
mRemoveFriend.Caption = SOC_COMMAND_REMOVE
mRemoveIgnore.Caption = SOC_COMMAND_REMOVE
End Sub

'==== Friend List ===='
Private Sub lvFriendList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 2 Then Exit Sub
If lvFriendList.HitTest(X, Y) Is Nothing Then Exit Sub

TriggerFriendEvent
End Sub

Private Sub lvFriendList_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode <> 93 Then Exit Sub

TriggerFriendEvent
End Sub

Private Sub lvFriendList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With lvFriendList
    .Sorted = True
    .SortKey = ColumnHeader.SubItemIndex

    If .SortOrder = lvwDescending Then
        .SortOrder = lvwAscending
    Else
        .SortOrder = lvwDescending
    End If

    .Sorted = False

    If .ListItems.Count <> 0 Then
        Set .SelectedItem = .ListItems.Item(1)
    End If
End With
End Sub

'==== Online List ===='
Private Sub lvOnlineList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lvOnlineList.ListItems.Count = 0 Then Exit Sub
If Button <> 2 Then Exit Sub

TriggerOnlineEvent
End Sub

Private Sub lvOnlineList_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode <> 93 Then Exit Sub

TriggerOnlineEvent
End Sub

Private Sub lvOnlineList_Click()
If lvOnlineList.SelectedItem.Text = frmMain.txtAccount.Text Then
    cmdAddToFriend.Enabled = False
    cmdAddToIgnore.Enabled = False
Else
    cmdAddToFriend.Enabled = True
    cmdAddToIgnore.Enabled = True
End If
End Sub

'==== Ignore List ===='
Private Sub lvIgnoreList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 2 Then Exit Sub
If lvIgnoreList.HitTest(X, Y) Is Nothing Then Exit Sub

TriggerIgnoreEvent
End Sub

Private Sub lvIgnoreList_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode <> 93 Then Exit Sub

TriggerIgnoreEvent
End Sub

Private Sub mWhisper_Click()
SetupChildForm frmChat
With frmChat.txtToSend
    .Text = "/whisper " & lvFriendList.SelectedItem.Text & " "
    .SelStart = Len(.Text)
    .SetFocus
End With
End Sub

Private Sub mWhisperT_Click()
SetupChildForm frmChat
With frmChat.txtToSend
    .Text = "/whisper " & lvOnlineList.SelectedItem.Text & " "
    .SelStart = Len(.Text)
    .SetFocus
End With
End Sub

