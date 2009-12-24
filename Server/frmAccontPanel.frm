VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmAccountPanel 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   Caption         =   "frmSendFile"
   ClientHeight    =   5220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7635
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
   ScaleHeight     =   5220
   ScaleWidth      =   7635
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
      Begin VB.ComboBox cmbGender 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5760
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblGender 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Gender:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5760
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblLevel 
         BackColor       =   &H00F4F4F4&
         Caption         =   " Level:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4560
         TabIndex        =   10
         Top             =   240
         Width           =   975
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
      Height          =   2220
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3916
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
      NumItems        =   10
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
         SubItemIndex    =   3
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Banned"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Level"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Secret Question"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Secret Answer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Gender"
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
Switch = "ADD"
DoButtons True
ClearTxBoxes
End Sub

Private Sub cmdCancel_Click()
DoButtons False
SetData ListView1.SelectedItem
End Sub

Private Sub cmdDel_Click()
Dim strName     As String
Dim lngID       As Long
Dim NewIndex    As Long

If ListView1.ListItems.Count = 0 Then
    MsgBox "No accounts avaible to delete.", vbInformation, "Delete"
    Exit Sub
End If
    
With ListView1.SelectedItem
    strName = .SubItems(1)
    lngID = CLng(.Text)
End With
    
If MsgBox("Are you sure that you want to delete account '" & strName & "' ?", vbYesNo + vbQuestion, "Confirm Delete") = vbNo Then Exit Sub

With frmMain.xCommand
    .CommandText = "DELETE FROM " & Database.AccountTable & " WHERE ID = " & lngID
    .Execute
End With

With frmFriendIgnoreList
    .RemoveAllFriendsFromUser strName
    .RemoveAllIgnoresFromUser strName
End With

With ListView1
    If .SelectedItem.Index = .ListItems.Count Then
        NewIndex = .ListItems.Count - 1
    Else
        NewIndex = .SelectedItem.Index
    End If
    .ListItems.Remove .SelectedItem.Index
    If .ListItems.Count > 0 Then
        Set .SelectedItem = .ListItems(NewIndex)
        ListView1_ItemClick .SelectedItem
    Else
        ClearTxBoxes
    End If
End With
End Sub

Private Sub DoButtons(Args As Boolean)
cmdAdd.Enabled = Not Args
cmdDel.Enabled = Not Args
cmdMod.Enabled = Not Args
cmdSave.Enabled = Args
cmdCancel.Enabled = Args
txtName.Enabled = Args
txtPassword.Enabled = Args
cmbBanned.Enabled = Args
cmbLevel.Enabled = Args
cmbGender.Enabled = Args
lblName.Enabled = Args
lblPassword.Enabled = Args
lblBanned.Enabled = Args
lblLevel.Enabled = Args
lblGender.Enabled = Args
Frame1.Enabled = Args
End Sub

Private Sub ClearTxBoxes()
txtName.Text = vbNullString
txtPassword.Text = vbNullString
cmbBanned.ListIndex = 0
cmbLevel.ListIndex = 0
End Sub

Private Sub cmdMod_Click()
'If there is no account selected give the message
If ListView1.ListItems.Count = 0 Then
    MsgBox "No account selected to modify.", vbInformation, "Modify"
    Exit Sub
End If
'Set switch to modify
Switch = "MOD"
DoButtons True
End Sub

Private Sub cmdSave_Click()
Select Case Switch
    Case "ADD"
        For i = 1 To ListView1.ListItems.Count
            If txtName.Text = ListView1.ListItems.Item(i).SubItems(1) Then
                MsgBox "Name already given.", vbInformation
                ClearTxBoxes
                txtName.SetFocus
                Exit Sub
            End If
        Next i
        RegisterAccount txtName.Text, txtPassword.Text, vbNullString, vbNullString, cmbGender.Text
        
    Case "MOD"
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
        
        ModifyAccount txtName.Text, txtPassword.Text, CBool(cmbBanned.Text), cmbLevel.Text, ListView1.SelectedItem.Text, ListView1.SelectedItem.Index, cmbGender.Text
End Select
DoButtons False
End Sub

Public Sub ModifyAccount(pName As String, pPassword As String, pBanned As Boolean, pLevel As String, MOD_ID As Long, LST_ID As Long, pGender As String)
'Update the database
With frmMain.xCommand
    .CommandText = "UPDATE " & Database.AccountTable & " SET Name1 = '" & pName & "', Password1 = '" & pPassword & "', Banned1 = '" & CStr(pBanned) & "', Level1 = '" & pLevel & "', Gender1 = '" & pGender & "' WHERE ID = " & MOD_ID
    .Execute
End With

'Update the listview
With ListView1.ListItems
    .Item(LST_ID).SubItems(1) = pName
    .Item(LST_ID).SubItems(2) = pPassword
    .Item(LST_ID).SubItems(5) = CStr(pBanned)
    .Item(LST_ID).SubItems(6) = pLevel
    .Item(LST_ID).SubItems(9) = pGender
End With
End Sub

Private Sub RegisterAccount(pName As String, pPassword As String, pSecretQuestion As String, pSecretAnswer As String, pGender As String)
Dim j As Long

'Check list for biggest value
With ListView1.ListItems
    For i = 1 To .Count
        If .Item(i) > j Then
            j = .Item(i)
        End If
    Next i
End With

'Create new index
j = j + 1

'Add new account to database
With frmMain.xCommand
    .CommandText = "INSERT INTO " & Database.AccountTable & " (ID, Name1, Password1, Time1, Date1, Banned1, Level1, SecretQuestion1, SecretAnswer1, Gender1) VALUES(" & j & ", '" & pName & "', '" & pPassword & "', '" & Format(Time, "hh:nn:ss") & "', '" & Format(Date, "yyyy-mm-dd") & "', 'False', '0', '" & pSecretQuestion & "', '" & pSecretAnswer & "', '" & pGender & "')"
    .Execute
End With

'Save index in variable
i = ListView1.ListItems.Count + 1

'Add account to the listview
With ListView1.ListItems
    .Add , , j
    .Item(i).SubItems(1) = pName
    .Item(i).SubItems(2) = pPassword
    .Item(i).SubItems(3) = Format(Time, "hh:nn:ss")
    .Item(i).SubItems(4) = Format(Date, "dd/mm/yyyy")
    .Item(i).SubItems(5) = "False"
    .Item(i).SubItems(6) = "0"
    .Item(i).SubItems(7) = pSecretQuestion
    .Item(i).SubItems(8) = pSecretAnswer
    .Item(i).SubItems(9) = pGender
End With
End Sub

Private Sub Form_Activate()
SetData ListView1.SelectedItem
End Sub

Private Sub Form_Load()
Me.Top = 0: Me.Left = 0
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
SetData Item
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

Private Sub SetData(pItem As ListItem)
If ListView1.ListItems.Count <> 0 Then
    With pItem
        txtName.Text = .SubItems(1)
        txtPassword.Text = .SubItems(2)
        cmbBanned.Text = .SubItems(5)
        cmbLevel.ListIndex = .SubItems(6)
        cmbGender.Text = .SubItems(9)
    End With
End If
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
Dim array1()            As String
Dim GetMessage          As String
Dim GetName             As String
Dim GetPassword         As String
Dim GetSecretQuestion   As String
Dim GetSecretAnswer     As String
Dim GetGender           As String

On Error GoTo HandleError
RegSock(Index).GetData GetMessage
array1 = Split(GetMessage, "#")

GetName = array1(1)
GetPassword = array1(2)
GetSecretQuestion = array1(3)
GetSecretAnswer = array1(4)
GetGender = array1(5)

Select Case array1(0)
    'Regster an account
    Case "!register"
        With ListView1.ListItems
            'Check if the account already exists
            For i = 1 To .Count
                If UCase$(GetName) = UCase$(.Item(i).SubItems(1)) Then
                    If RegSock(Index).State = 7 Then
                        RegSock(Index).SendData "!nameexist#"
                        Exit Sub
                    End If
                End If
            Next i
        End With
    
        RegisterAccount GetName, GetPassword, GetSecretQuestion, GetSecretAnswer, GetGender
        RegSock(Index).SendData "!done#"
        
    'Check the secret question and send password
    Case "!request_password"
        With ListView1.ListItems
            For i = 1 To .Count
                'Check if the account exists
                If UCase$(GetName) = UCase$(.Item(i).SubItems(1)) Then
                    'Check if the question chosen is the same
                    If .Item(i).SubItems(7) = GetPassword Then
                        'Check if the answer is the same
                        If .Item(i).SubItems(8) = GetSecretQuestion Then
                            If RegSock(Index).State = 7 Then
                                RegSock(Index).SendData "!successfull#" & .Item(i).SubItems(2) & ".#"
                            End If
                        Else
                            If RegSock(Index).State = 7 Then
                                RegSock(Index).SendData "!error_fp#"
                            End If
                        End If
                    Else
                        If RegSock(Index).State = 7 Then
                            RegSock(Index).SendData "!error_fp#"
                        End If
                    End If
                Else
                    If i = .Count Then
                        If RegSock(Index).State = 7 Then
                            RegSock(Index).SendData "!account_not_exist#"
                        End If
                    End If
                End If
            Next i
        End With
End Select
Exit Sub

HandleError:
    RegSock(Index).SendData "!error#"
End Sub
