VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmAccountPanel 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin MSWinsockLib.Winsock RegSock 
      Index           =   0
      Left            =   6480
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvAccounts 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
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
      NumItems        =   12
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
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Email"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Last IP"
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

Public Sub ModifyAccount(Name As String, Password As String, Banned As Long, Level As String, MOD_ID As Long, LST_ID As Long, Gender As String, Email As String)
'Modify listview values
With lvAccounts.ListItems.Item(LST_ID)
    .SubItems(INDEX_NAME) = Name
    .SubItems(INDEX_PASSWORD) = Password
    .SubItems(INDEX_BANNED) = Banned
    .SubItems(INDEX_LEVEL) = Level
    .SubItems(INDEX_GENDER) = Gender
    .SubItems(INDEX_EMAIL) = Email
End With

'Modify database values
pDB.ExecuteCommand "UPDATE " & DATABASE_TABLE_ACCOUNTS & " SET Name1 = '" & Name & "', Password1 = '" & Password & "', Banned1 = '" & Banned & "', Level1 = '" & Level & "', Gender1 = '" & Gender & "', Email1 ='" & Email & "' WHERE ID = " & MOD_ID
End Sub

Private Sub RegisterAccount(Name As String, Password As String, Banned As String, Level As String, SecretQuestion As String, SecretAnswer As String, Gender As String, Email As String)
Dim i As Long
Dim j As Long

'Check list for biggest value and create new index
j = pDB.GetMaxID("ID", DATABASE_TABLE_ACCOUNTS)

'Save index in variable
i = lvAccounts.ListItems.Count + 1

'Add account to the listview
With lvAccounts.ListItems
    .Add , , j
    .Item(i).SubItems(INDEX_NAME) = Name
    .Item(i).SubItems(INDEX_PASSWORD) = Password
    .Item(i).SubItems(INDEX_TIME) = Format(Time, "hh:nn:ss")
    .Item(i).SubItems(INDEX_DATE) = Format(Date, "dd/mm/yyyy")
    .Item(i).SubItems(INDEX_BANNED) = Banned
    .Item(i).SubItems(INDEX_LEVEL) = Level
    .Item(i).SubItems(INDEX_SECRET_QUESTION) = SecretQuestion
    .Item(i).SubItems(INDEX_SECRET_ANSWER) = SecretAnswer
    .Item(i).SubItems(INDEX_GENDER) = Gender
    .Item(i).SubItems(INDEX_EMAIL) = Email
End With

'Add account to database
pDB.ExecuteCommand "INSERT INTO " & DATABASE_TABLE_ACCOUNTS & " (ID, Name1, Password1, Time1, Date1, Banned1, Level1, SecretQuestion1, SecretAnswer1, Gender1, Email1) VALUES(" & j & ", '" & Name & "', '" & Password & "', '" & Format(Time, "hh:nn:ss") & "', '" & Format(Date, "yyyy-mm-dd") & "', '" & Banned & "', '" & Level & "', '" & SecretQuestion & "', '" & SecretAnswer & "', '" & Gender & "', '" & Email & "')"
End Sub

Private Sub Form_Load()
Top = 0: Left = 0
End Sub

Private Sub RegSock_Close(Index As Integer)
RegSock(Index).Close
Unload RegSock(Index)
End Sub

Private Sub RegSock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim i As Long
    i = LoadSocket()

RegSock(i).LocalPort = rPort
RegSock(i).Accept requestID
End Sub

Private Function GetFreeSocket() As Long
Dim i As Long

On Error GoTo HandleErrorFreeSocket
With RegSock
    For i = .LBound + 1 To .UBound
        .Item (i)
    Next i

    GetFreeSocket = .UBound + 1
End With

Exit Function
HandleErrorFreeSocket:
    GetFreeSocket = i
End Function

Private Function LoadSocket() As Long
Dim i As Long
    i = GetFreeSocket()

Load RegSock(i)
LoadSocket = i
End Function

Private Sub RegSock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim i                   As Long
Dim array1()            As String
Dim pCommand            As String
Dim pMessage            As String

On Error GoTo HandleError

RegSock(Index).GetData pMessage
array1 = Split(pMessage, pSplit)

If UBound(array1) > -1 Then
    pCommand = array1(0)
End If

Select Case pCommand
    'Regster an account
    Case "!register"
        With lvAccounts.ListItems
            'Check if the account already exists
            For i = 1 To .Count
                If LCase$(.Item(i).SubItems(INDEX_NAME)) = LCase$(array1(1)) Then
                    If RegSock(Index).State = 7 Then RegSock(Index).SendData "!nameexist" & pSplit
                    Exit Sub
                End If
            Next i

            'Check if the email is already used
            For i = 1 To .Count
                If LCase$(.Item(i).SubItems(INDEX_EMAIL)) = LCase$(array1(6)) Then
                    If RegSock(Index).State = 7 Then RegSock(Index).SendData "!emailtaken" & pSplit
                    Exit Sub
                End If
            Next i
        End With

        RegisterAccount array1(1), array1(2), "0", "0", array1(3), array1(4), array1(5), array1(6)
        RegSock(Index).SendData "!done" & pSplit

    'Check the secret question and send password
    Case "!request_password"
        With lvAccounts.ListItems
            For i = 1 To .Count
                'Check if the email exists
                If UCase$(.Item(i).SubItems(INDEX_EMAIL)) = UCase$(array1(1)) Then
                    'Check if the question chosen is the same
                    If .Item(i).SubItems(INDEX_SECRET_QUESTION) = array1(2) Then
                        'Check if the answer is the same
                        If .Item(i).SubItems(INDEX_SECRET_ANSWER) = array1(3) Then
                            If RegSock(Index).State = 7 Then RegSock(Index).SendData "!successfull" & pSplit & .Item(i).SubItems(INDEX_PASSWORD) & pSplit & .Item(i).SubItems(INDEX_NAME) & pSplit
                        Else
                            If RegSock(Index).State = 7 Then RegSock(Index).SendData "!error_fp" & pSplit
                        End If
                    Else
                        If RegSock(Index).State = 7 Then RegSock(Index).SendData "!error_fp" & pSplit
                    End If
                    Exit For
                Else
                    If i = .Count Then If RegSock(Index).State = 7 Then RegSock(Index).SendData "!email_not_exist" & pSplit
                End If
            Next i
        End With

End Select
Exit Sub

HandleError:
    If RegSock(Index).State = 7 Then RegSock(Index).SendData "!error" & pSplit
End Sub
