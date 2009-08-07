VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.OCX"
Begin VB.Form frmPanel 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   Caption         =   "frmPanel"
   ClientHeight    =   4080
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
   ScaleHeight     =   4080
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "&Uncheck All"
      Height          =   350
      Left            =   6000
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Unmute"
      Height          =   350
      Left            =   2520
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Mute"
      Height          =   350
      Left            =   1320
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox txtLOG 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   1296
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmPanel.frx":0000
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   5280
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Kick"
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
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
Dim i As Integer

Private Sub Command1_Click()
If ListView1.ListItems.Count = 0 Then
    txtLOG.Text = txtLOG.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "] Failed list is empty."
    Exit Sub
End If
For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems.Item(i).Checked = True Then
        frmMain.Winsock1(i).Close
        Unload frmMain.Winsock1(i)
        txtLOG.Text = txtLOG.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "] " & ListView1.ListItems.Item(i) & " got kicked."
        ListView1.ListItems.Remove (i)
    Else
        txtLOG.Text = txtLOG.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "] No user checked."
    End If
Next i

'Update clients user online list
UpdateUsersList

'Update Statusbar
frmMain.StatusBar1.Panels(1).Text = "Status: Connected with  " & frmMain.Winsock1.Count - 1 & " Client(s)."
End Sub

Private Sub Command2_Click()
For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems.Item(i).Checked = True Then
        ListView1.ListItems.Item(i).Checked = False
    End If
Next i
End Sub

Private Sub Command3_Click()
If ListView1.ListItems.Count <= 0 Then
    txtLOG.Text = txtLOG.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "] List is empty."
    Exit Sub
End If
For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(i).Checked = True Then
        If ListView1.ListItems(i).SubItems(4) = "Yes" Then
            txtLOG.Text = txtLOG.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "] User is already muted."
        Else
            ListView1.ListItems.Item(i).SubItems(4) = "Yes"
            txtLOG.Text = txtLOG.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "] " & ListView1.ListItems.Item(i) & " got muted."
            SendMessage " " & ListView1.ListItems.Item(i) & " got muted."
        End If
    Else
        txtLOG.Text = txtLOG.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "] No user checked."
    End If
    ListView1.ListItems.Item(i).Checked = False
Next i
End Sub

Private Sub Command4_Click()
If ListView1.ListItems.Count <= 0 Then
    txtLOG.Text = txtLOG.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "] List is empty."
    Exit Sub
End If
For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(i).Checked = True Then
        If ListView1.ListItems(i).SubItems(4) = "Yes" Then
            ListView1.ListItems.Item(i).SubItems(4) = "No"
            txtLOG.Text = txtLOG.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "] " & ListView1.ListItems.Item(i) & " got unmuted."
            SendMessage " " & ListView1.ListItems.Item(i) & " got unmuted."
        Else
            txtLOG.Text = txtLOG.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "] User " & "'" & ListView1.ListItems.Item(i) & "'" & " is not muted."
        End If
    Else
        txtLOG.Text = txtLOG.Text & vbCrLf & "[" & Format(Time, "hh:nn:ss") & "] No user checked."
    End If
    ListView1.ListItems.Item(i).Checked = False
Next i
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
AddBadWordsToList
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
    .AddItem "Unknown"
End With
End Sub

Private Sub txtLOG_Change()
txtLOG.SelStart = Len(txtLOG.Text)
End Sub
