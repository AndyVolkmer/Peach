VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
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
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3960
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Left            =   240
      TabIndex        =   14
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Left            =   1680
      TabIndex        =   13
      Top             =   2880
      Width           =   1215
   End
   Begin VB.ComboBox cmbBanned 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2880
      Width           =   1335
   End
   Begin VB.ComboBox cmbLevel 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2880
      Width           =   975
   End
   Begin MSWinsockLib.Winsock RegSock 
      Index           =   0
      Left            =   6480
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Configuration :"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   7215
      Begin VB.Label lblLevel 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Level:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4560
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblBanned 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Banned:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblPassword 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Password:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblName 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Name:"
         Enabled         =   0   'False
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
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   882
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
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Banned"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Level"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmAccountPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Switch As String

Private Sub cmdAdd_Click()
Switch = "Add"
DoButtons True
ClearTxBoxes
End Sub

Private Sub cmdCancel_Click()
DoButtons False
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

frmMain.xCommand.CommandText = "DELETE FROM " & Database.Account_Table & " WHERE ID = " & lngID
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

Private Sub DoButtons(Wind As Boolean)
cmdAdd.Enabled = Not Wind
cmdDel.Enabled = Not Wind
cmdMod.Enabled = Not Wind
cmdSave.Enabled = Wind
cmdCancel.Enabled = Wind
txtName.Enabled = Wind
txtPassword.Enabled = Wind
cmbBanned.Enabled = Wind
cmbLevel.Enabled = Wind
lblName.Enabled = Wind
lblPassword.Enabled = Wind
lblBanned.Enabled = Wind
lblLevel.Enabled = Wind
Frame1.Enabled = Wind
End Sub

Private Sub ClearTxBoxes()
txtName.Text = vbNullString
txtPassword.Text = vbNullString
cmbBanned.ListIndex = 0
cmbLevel.ListIndex = 0
End Sub

Private Sub cmdMod_Click()
'If there is no account selected give the message
If ListView1.SelectedItem Is Nothing Then
    MsgBox "No account selected to modify.", vbInformation, "Modify"
    Exit Sub
End If
'Set switch to modify
Switch = "Modify"
DoButtons True
End Sub

Private Sub cmdSave_Click()
Select Case Switch
Case "Add"
    For i = 1 To ListView1.ListItems.Count
        If txtName.Text = ListView1.ListItems.Item(i).SubItems(1) Then
            MsgBox "Name already given.", vbInformation
            ClearTxBoxes
            txtName.SetFocus
            Exit Sub
        End If
    Next i
    RegisterAccount txtName.Text, txtPassword.Text
Case "Modify"
    'Name can't be modified to nothing
    If Len(Trim$(txtName.Text)) = 0 Then
        MsgBox "The name can't be empty.", vbInformation
        txtName.SetFocus
        Exit Sub
    End If
    
    'Password can't be modified to nothing
    If Len(Trim$(txtPassword.Text)) = 0 Then
        MsgBox "The password can't be empty.", vbInformation
        txtPassword.SetFocus
        Exit Sub
    End If
    
    ModifyAccount txtName.Text, txtPassword.Text, CBool(cmbBanned.Text), cmbLevel.Text, ListView1.SelectedItem.Text, ListView1.SelectedItem.Index
End Select

DoButtons False
End Sub
Public Sub ModifyAccount(dName As String, dPassword As String, dBanned As Boolean, dLevel As String, MOD_ID As Long, LST_ID As Long)
'Update the database
frmMain.xCommand.CommandText = "UPDATE " & Database.Account_Table & " SET Name1 = '" & dName & "', Password1 = '" & dPassword & "', Banned1 = '" & CStr(dBanned) & "', Level1 = '" & dLevel & "' WHERE ID = " & MOD_ID
frmMain.xCommand.Execute

'Update the listview
With ListView1.ListItems
    .Item(LST_ID).SubItems(1) = dName
    .Item(LST_ID).SubItems(2) = dPassword
    .Item(LST_ID).SubItems(5) = CStr(dBanned)
    .Item(LST_ID).SubItems(6) = dLevel
End With
End Sub

Private Sub RegisterAccount(dName As String, dPassword As String)
Dim j As Long

j = 0
'Check list if the ID is already given
With ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) > j Then
            j = .Item(i)
        End If
    Next i
End With

j = j + 1

'Add new account to database
frmMain.xCommand.CommandText = "INSERT INTO " & Database.Account_Table & " (ID, Name1, Password1, Time1, Date1, Banned1, Level1) VALUES(" & j & ", '" & dName & "', '" & dPassword & "', '" & Format(Time, "hh:nn:ss") & "', '" & Format(Date, "yyyy-mm-dd") & "', 'False', '0')"
frmMain.xCommand.Execute

'Save index in variable
i = ListView1.ListItems.Count + 1

'Add account to the listview
With ListView1.ListItems
    .Add , , j
    .Item(i).SubItems(1) = dName
    .Item(i).SubItems(2) = dPassword
    .Item(i).SubItems(3) = Format(Time, "hh:nn:ss")
    .Item(i).SubItems(4) = Format(Date, "dd/mm/yyyy")
    .Item(i).SubItems(5) = "False"
    .Item(i).SubItems(6) = "0"
End With
End Sub

Private Sub Form_Activate()
If ListView1.ListItems.Count <> 0 Then
    With ListView1.SelectedItem
        txtName.Text = .SubItems(1)
        txtPassword.Text = .SubItems(2)
        cmbBanned.Text = .SubItems(5)
        cmbLevel.ListIndex = .SubItems(6)
    End With
End If
End Sub

Private Sub Form_Load()
Me.Top = 0: Me.Left = 0
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
With Item
    txtName = .SubItems(1)
    txtPassword = .SubItems(2)
    If .SubItems(5) = "False" Then
        cmbBanned.ListIndex = 0
    Else
        cmbBanned.ListIndex = 1
    End If
    Select Case .SubItems(6)
    Case 0
        cmbLevel.ListIndex = 0
    Case 1
        cmbLevel.ListIndex = 1
    Case 2
        cmbLevel.ListIndex = 2
    End Select
End With
End Sub

Private Sub RegSock_Close(Index As Integer)
RegSock(Index).Close
Unload RegSock(Index)
End Sub

Private Sub RegSock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim intCounter As Integer
    intCounter = loadSocket
    RegSock(intCounter).LocalPort = rPort
    RegSock(intCounter).Accept requestID
End Sub

Private Function socketFree() As Integer
On Error GoTo HandleErrorFreeSocket
    Dim theIP As String
    For i = RegSock.LBound + 1 To RegSock.UBound
        theIP = RegSock(i).LocalIP
    Next i
    socketFree = RegSock.UBound + 1
Exit Function
HandleErrorFreeSocket:
socketFree = i
End Function

Private Function loadSocket() As Integer
Dim theFreeSocket As Integer
theFreeSocket = 0
theFreeSocket = socketFree

Load RegSock(theFreeSocket)

loadSocket = theFreeSocket
End Function

Private Sub RegSock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim array1()    As String
Dim GetMessage  As String
Dim GetCommand  As String
Dim GetName     As String
Dim GetPassword As String

On Error GoTo HandleError
RegSock(Index).GetData GetMessage

array1 = Split(GetMessage, "#")

GetCommand = array1(0)
GetName = array1(1)
GetPassword = array1(2)

If GetCommand = "!register" Then
    For i = 1 To ListView1.ListItems.Count
        If UCase$(GetName) = UCase$(ListView1.ListItems.Item(i).SubItems(1)) Then
            If RegSock(Index).State = 7 Then
                RegSock(Index).SendData "!nameexist#"
                Exit Sub
            End If
        End If
    Next i
    
    RegisterAccount GetName, GetPassword
    RegSock(Index).SendData "!done#"
End If

Exit Sub
HandleError:
    RegSock(Index).SendData "!error#"
    Exit Sub
End Sub
