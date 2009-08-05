VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPanel 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   Caption         =   "frmPanel"
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7515
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "&Mute"
      Height          =   405
      Left            =   1440
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   4800
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4560
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Search"
      Height          =   285
      Left            =   5880
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Kick"
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2295
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IP"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Winsock ID"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Login Time"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Muted"
         Object.Width           =   2295
      EndProperty
   End
End
Attribute VB_Name = "frmPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then Exit Sub
If Text1.Text = "0" Then Exit Sub
If ListView1.ListItems.Count <= 0 Then
    Text1.Text = "0"
    Exit Sub
End If

' Disconnect the socket
frmMain.Winsock1(Text1.Text).Close
Unload frmMain.Winsock1(Text1.Text)

ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
' Update clients user online list
UpdateUsersList

' Clear entry
Text1.Text = "0"

' Update Statusbar
frmMain.StatusBar1.Panels(1).Text = "Status: Connected with  " & frmMain.Winsock1.Count - 1 & " Client(s)."
End Sub

Private Sub Command2_Click()
Dim i, Pos  As Integer
For i = 1 To ListView1.ListItems.Count
    ListView1.ListItems.Item(i).Selected = False
Next
For i = 1 To ListView1.ListItems.Count
    Pos = InStr(1, ListView1.ListItems.Item(i), Text2, vbTextCompare)
    If Pos > 0 Then
            ListView1.SelectedItem = ListView1.ListItems.Item(i)
    End If
Next i
ListView1.SetFocus
End Sub

Private Sub Command3_Click()
If Command3.Caption = "&Mute" Then
    If ListView1.ListItems.Count <= 0 Then Exit Sub
    ListView1.ListItems.Item(ListView1.SelectedItem.Index).SubItems(4) = "Yes"
    Command3.Caption = "&Unmute"
    SendMessage " " & ListView1.ListItems.Item(ListView1.SelectedItem.Index) & " got muted."
Else
    If ListView1.ListItems.Count <= 0 Then Exit Sub
    ListView1.ListItems.Item(ListView1.SelectedItem.Index).SubItems(4) = "No"
    Command3.Caption = "Mute"
    SendMessage " " & ListView1.ListItems.Item(ListView1.SelectedItem.Index) & " got unmuted."
End If
End Sub

Private Sub Form_Load()
Me.Top = "0"
Me.Left = "0"
Text1.SelStart = Len(Text1.Text)
AddBadWordsToList
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text1.Text = ListView1.SelectedItem.ListSubItems(2).Text
    If ListView1.ListItems.Item(ListView1.SelectedItem.Index).SubItems(4) = "Yes" Then
        Command3.Caption = "&Unmute"
    Else
        Command3.Caption = "&Mute"
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Command2_Click
End Sub

Private Sub AddBadWordsToList()
With List1
    .AddItem "Cunt"
    .AddItem "Penis"
    .AddItem "Defaultnick"
    .AddItem "Nickname"
    .AddItem "God"
    .AddItem "Retard"
    .AddItem "Nickname"
    .AddItem "Vagina"
    .AddItem "Schwanz"
    .AddItem "Kurac"
    .AddItem "Noob"
    .AddItem "Pizda"
    .AddItem "Fjortis"
    .AddItem "Dick"
    .AddItem "Porno"
    .AddItem "Porn"
End With
End Sub
