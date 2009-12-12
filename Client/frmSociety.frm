VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSociety 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7665
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
   ScaleHeight     =   4290
   ScaleWidth      =   7665
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
      Tab(1).Control(0)=   "cmdAddToFriend"
      Tab(1).Control(1)=   "lvOnlineList"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Ignore List"
      TabPicture(2)   =   "frmSociety.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvIgnoreList"
      Tab(2).Control(1)=   "cmdAddIgnore"
      Tab(2).Control(2)=   "cmdRemoveIgnore"
      Tab(2).ControlCount=   3
      Begin VB.CommandButton cmdAddToFriend 
         Caption         =   "&Add to Friends"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -70320
         TabIndex        =   8
         Top             =   3360
         Width           =   2415
      End
      Begin VB.CommandButton cmdRemoveIgnore 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -69600
         TabIndex        =   6
         Top             =   3360
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddIgnore 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71280
         TabIndex        =   5
         Top             =   3360
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemoveFriend 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5400
         TabIndex        =   2
         Top             =   3360
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddFriend 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   3360
         Width           =   1695
      End
      Begin MSComctlLib.ListView lvFriendList 
         Height          =   2775
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   7095
         _ExtentX        =   12515
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
            Object.Width           =   9878
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
         Width           =   7095
         _ExtentX        =   12515
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
            Object.Width           =   12409
         EndProperty
      End
      Begin MSComctlLib.ListView lvIgnoreList 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   7095
         _ExtentX        =   12515
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
            Object.Width           =   12418
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSociety"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub LoadSocietyForm()
SSTab1.TabCaption(0) = SOC_FRIEND_LIST
SSTab1.TabCaption(1) = SOC_ONLINE_LIST
SSTab1.TabCaption(2) = SOC_IGNORE_LIST
cmdAddFriend.Caption = SOC_COMMAND_ADD
cmdRemoveFriend.Caption = SOC_COMMAND_REMOVE
cmdAddIgnore.Caption = SOC_COMMAND_ADD
cmdRemoveIgnore.Caption = SOC_COMMAND_REMOVE
cmdAddToFriend.Caption = SOC_COMMAND_FRIEND
End Sub

Private Sub cmdAddFriend_Click()
Dim pVal As String
pVal = InputBox(SOC_ASK_FRIEND_TEXT, SOC_ASK_FRIEND_DEFAULT, SOC_ASK_FRIEND_DEFAULT)

If Len(Trim$(pVal)) = 0 Then Exit Sub
SendMSG "!friend#-add#" & frmConfig.txtAccount & "#" & pVal & "#"
End Sub

Private Sub cmdAddIgnore_Click()
Dim pVal As String
pVal = InputBox(SOC_ASK_IGNORE_TEXT, SOC_ASK_IGNORE_TITLE, SOC_ASK_IGNORE_DEFAULT)

If Len(Trim$(pVal)) = 0 Then Exit Sub
SendMSG "!ignore#-add#" & frmConfig.txtAccount & "#" & pVal & "#"
End Sub

Private Sub cmdAddToFriend_Click()
Dim Name    As String
Dim Full    As String
Dim MPos    As Integer

If lvOnlineList.ListItems.Count = 0 Then Exit Sub

Full = lvOnlineList.SelectedItem
MPos = InStr(1, Full, "(")
Name = Mid(Full, MPos + 2, Len(Full) - MPos - 3)

SendMSG "!friend#-add#" & frmConfig.txtAccount & "#" & Name & "#"
End Sub

Private Sub cmdRemoveFriend_Click()
Dim Name    As String
Dim MPos    As Integer
Dim Temp()  As String

With lvFriendList
    If .ListItems.Count = 0 Then Exit Sub
    
    Temp = Split(.SelectedItem.Text, " ")
    If MsgBox(SOC_ASK_DEL_1 & Temp(0) & SOC_ASK_DEL_2, vbQuestion + vbYesNo, "Deleting '" & Temp(0) & "'") = vbNo Then
        Exit Sub
    End If
    
    MPos = InStr(1, .SelectedItem.Text, " ")
    If MPos = 0 Then
        Name = .SelectedItem.Text
    Else
        Name = Left$(.SelectedItem.Text, MPos - 1)
    End If
    SendMSG "!friend#-remove#" & frmConfig.txtAccount.Text & "#" & Name & "#"
End With
End Sub

Private Sub cmdRemoveIgnore_Click()
Dim Name As String
Dim MPos As Integer
Dim Temp() As String

With lvIgnoreList
    If .ListItems.Count = 0 Then Exit Sub
    
    Temp = Split(.SelectedItem.Text, " ")
    If MsgBox(SOC_ASK_DEL_1 & Temp(0) & SOC_ASK_DEL_2, vbQuestion + vbYesNo, "Deleting '" & Temp(0) & "'") = vbNo Then
        Exit Sub
    End If
    
    MPos = InStr(1, .SelectedItem.Text, " ")
    If MPos = 0 Then
        Name = .SelectedItem.Text
    Else
        Name = Left$(.SelectedItem.Text, MPos - 1)
    End If
    SendMSG "!ignore#-remove#" & frmConfig.txtAccount.Text & "#" & Name & "#"
End With
End Sub

Private Sub Form_Load()
Me.Top = 0: Me.Left = 0
LoadSocietyForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Setting.ASK_TICK = True Then
    If MsgBox(MDI_MSG_UNLOAD, vbQuestion + vbYesNo) = vbNo Then
        Cancel = 1
        Exit Sub
    End If
End If
End Sub
