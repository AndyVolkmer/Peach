VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.OCX"
Begin VB.Form frmPanel 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   Caption         =   "frmPanel"
   ClientHeight    =   4020
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox txtLOG 
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmPanel.frx":0000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Mute"
      Height          =   350
      Left            =   6120
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   4800
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Search"
      Height          =   285
      Left            =   6120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Kick"
      Height          =   350
      Left            =   4920
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   4471
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
         Object.Width           =   2558
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IP"
         Object.Width           =   2558
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Winsock ID"
         Object.Width           =   2558
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Login Time"
         Object.Width           =   2558
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Muted"
         Object.Width           =   2469
      EndProperty
   End
End
Attribute VB_Name = "frmPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If ListView1.ListItems.Count = 0 Then
    txtLOG.Text = txtLOG.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "][Kick] Failed list is empty."
    Exit Sub
End If

If ListView1.SelectedItem.Index <= 0 Then
    txtLOG.Text = txtLOG.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "][Kick] Failed no user selected."
    Exit Sub
End If

'Show it in the log
txtLOG.Text = txtLOG.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "] " & ListView1.ListItems.Item(ListView1.SelectedItem.Index) & " got kicked."

' Disconnect the socket
frmMain.Winsock1(ListView1.SelectedItem.Index).Close
Unload frmMain.Winsock1(ListView1.SelectedItem.Index)

ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
' Update clients user online list
UpdateUsersList

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
If ListView1.ListItems.Count <= 0 Then
    txtLOG.Text = txtLOG.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "][Mute] List is empty."
    Exit Sub
End If
If Command3.Caption = "&Mute" Then
    ListView1.ListItems.Item(ListView1.SelectedItem.Index).SubItems(4) = "Yes"
    Command3.Caption = "&Unmute"
    txtLOG.Text = txtLOG.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "] " & ListView1.ListItems.Item(ListView1.SelectedItem.Index) & " got muted."
    SendMessage " " & ListView1.ListItems.Item(ListView1.SelectedItem.Index) & " got muted."
Else
    ListView1.ListItems.Item(ListView1.SelectedItem.Index).SubItems(4) = "No"
    Command3.Caption = "&Mute"
    txtLOG.Text = txtLOG.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "] " & ListView1.ListItems.Item(ListView1.SelectedItem.Index) & " got unmuted."
    SendMessage " " & ListView1.ListItems.Item(ListView1.SelectedItem.Index) & " got unmuted."
End If
End Sub

Private Sub Form_Load()
Me.Top = "0"
Me.Left = "0"
AddBadWordsToList
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
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

Private Sub txtLOG_Change()
txtLOG.SelStart = Len(txtLOG.Text)
End Sub
