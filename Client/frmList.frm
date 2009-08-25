VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmList 
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Online List"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   3750
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView ListView1 
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   4683
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   6050
      EndProperty
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Close"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Online Users : (Check to whisper)"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmMain.UpdateListPosition.Enabled = False
frmList.Hide
End Sub

Private Sub Form_Load()
LoadListForm
End Sub

Private Sub Form_Resize()
If Not frmMain.WindowState = vbMinimized Then
    frmList.ListView1.Height = Me.Height - 1400
    Command1.Top = frmList.ListView1.Height + 400
End If
End Sub

Public Sub LoadListForm()
Me.Caption = LISTcaption
Command1.Caption = LISTcommand_close
End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If Item.Checked = True Then
    With ListView1.ListItems
        For i = 1 To .Count
            .Item(i).Checked = False
        Next i
    End With
    ListView1.SelectedItem = Item
    Item.Checked = True
Else
    Item.Checked = False
End If
End Sub
