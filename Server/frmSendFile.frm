VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAccountPanel 
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5160
      TabIndex        =   12
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3960
      TabIndex        =   11
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Configuration :"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   7215
      Begin VB.ComboBox cmbBanned 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtName 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00F4F4F4&
         Caption         =   "Banned:"
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F4F4F4&
         Caption         =   "Password:"
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F4F4F4&
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "&Modify"
      Height          =   350
      Left            =   2520
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Height          =   350
      Left            =   1320
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   350
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Password"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Time"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Banned"
         Object.Width           =   1587
      EndProperty
   End
End
Attribute VB_Name = "frmAccountPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer
Dim ii As Integer
Dim Switch As String

Private Sub cmdAdd_Click()
Switch = "Add"
DoButtons1
ClearTxBoxes
End Sub

Private Sub cmdCancel_Click()
DoButtons2
End Sub

Private Sub cmdDel_Click()
Dim strName         As String
Dim lngID           As Long
Dim lngNewSelIndex  As Long

If ListView1.SelectedItem Is Nothing Then
    MsgBox "No account selected to delete.", vbInformation, "Delete"
    Exit Sub
End If
    
With ListView1.SelectedItem
    strName = .SubItems(1)
    lngID = CLng(.Text)
End With
    
If MsgBox("Are you sure that you want to delete account '" & strName & "' ?", vbYesNo + vbQuestion, "Confirm Delete") = vbNo Then
    Exit Sub
End If

frmMain.xCommand.CommandText = "DELETE FROM Accounts WHERE ID = " & lngID
frmMain.xCommand.Execute

With ListView1
    If .SelectedItem.Index = .ListItems.Count Then
        lngNewSelIndex = .ListItems.Count - 1
    Else
        lngNewSelIndex = .SelectedItem.Index
    End If
    .ListItems.Remove .SelectedItem.Index
    If .ListItems.Count > 0 Then
        Set .SelectedItem = .ListItems(lngNewSelIndex)
        ListView1_ItemClick .SelectedItem
    Else
        ClearTxBoxes
    End If
End With
End Sub

Private Sub DoButtons1()
cmdAdd.Enabled = False
cmdDel.Enabled = False
cmdMod.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
txtName.Enabled = True
txtPassword.Enabled = True
cmbBanned.Enabled = True
End Sub

Private Sub DoButtons2()
cmdAdd.Enabled = True
cmdDel.Enabled = True
cmdMod.Enabled = True
cmdSave.Enabled = False
cmdCancel.Enabled = False
txtName.Enabled = False
txtPassword.Enabled = False
cmbBanned.Enabled = False
End Sub

Private Sub ClearTxBoxes()
txtName.Text = ""
txtPassword.Text = ""
End Sub

Private Sub cmdMod_Click()
'If there is no account selected give the message
If ListView1.SelectedItem Is Nothing Then
    MsgBox "No account selected to modify.", vbInformation, "Modify"
    Exit Sub
End If
'Set switch to modify
Switch = "Modify"
DoButtons1
End Sub

Private Sub cmdSave_Click()
Select Case Switch
Case "Add"
    i = 1
    
    'Check list if the ID is already given
    For ii = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(ii) = i Then
            i = i + 1
        End If
    Next ii
    
    'Add new account to database
    frmMain.xCommand.CommandText = "INSERT INTO accounts (ID, Name1, Password1, Time1, Date1, Banned1) VALUES(" & i & ", '" & txtName.Text & "', '" & txtPassword.Text & "', '" & Format(Time, "hh:nn:ss") & "', '" & Format(Date, "dd.mm.yyyy") & "', '" & cmbBanned.Text & "')"
    frmMain.xCommand.Execute
    
    'Save index in variable
    i = ListView1.ListItems.Count + 1
    
    'Add account to the listview
    With ListView1.ListItems
        .Add , , ii
        .Item(i).SubItems(1) = txtName.Text
        .Item(i).SubItems(2) = txtPassword.Text
        .Item(i).SubItems(3) = Format(Time, "hh:nn:ss")
        .Item(i).SubItems(4) = Format(Date, "dd.mm.yyyy")
        .Item(i).SubItems(5) = cmbBanned.Text
    End With
Case "Modify"
    'Name can't be modified to nothing
    If txtName.Text = "" Then
        MsgBox "The name can't be empty.", vbInformation
        txtName.SetFocus
        Exit Sub
    End If
    
    'Password can't be modified to nothing
    If txtPassword.Text = "" Then
        MsgBox "The password can't be empty.", vbInformation
        txtPassword.SetFocus
        Exit Sub
    End If
    
    'Save index in a variable
    i = ListView1.SelectedItem.Index
    
    'Update the database
    frmMain.xCommand.CommandText = "UPDATE accounts SET Name1 = '" & txtName.Text & "', Password1 = '" & txtPassword.Text & "', Banned1 = '" & cmbBanned.Text & "' WHERE ID = " & ListView1.ListItems.Item(i)
    frmMain.xCommand.Execute
    
    'Update the listview
    With ListView1.ListItems
        .Item(i).SubItems(1) = txtName.Text
        .Item(i).SubItems(2) = txtPassword.Text
        .Item(i).SubItems(5) = cmbBanned.Text
    End With
End Select
DoButtons2
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Select Case ListView1.ListItems.Count
Case 1
    MsgBox "There is " & ListView1.ListItems.Count & " account in the list.", vbInformation
Case 0
    MsgBox "There is no account in the list.", vbInformation
Case Else
    MsgBox "There are " & ListView1.ListItems.Count & " accounts in the list.", vbInformation
End Select
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
With Item
    txtName.Text = .SubItems(1)
    txtPassword.Text = .SubItems(2)
    If .SubItems(5) = "Yes" Then
        cmbBanned.ListIndex = 0
    Else
        cmbBanned.ListIndex = 1
    End If
End With
End Sub
