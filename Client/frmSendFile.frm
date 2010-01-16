VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSendFile 
   Appearance      =   0  'Flat
   BackColor       =   &H00F4F4F4&
   BorderStyle     =   0  'None
   Caption         =   "frmSendFile"
   ClientHeight    =   3735
   ClientLeft      =   690
   ClientTop       =   1365
   ClientWidth     =   7500
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3735
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtFileName 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   405
      Width           =   7185
   End
   Begin VB.Timer tmrSendFile 
      Enabled         =   0   'False
      Left            =   1800
      Top             =   0
   End
   Begin VB.Timer tmrCalcSpeed 
      Interval        =   1000
      Left            =   2295
      Top             =   0
   End
   Begin VB.PictureBox picProgress 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F4F4F4&
      Enabled         =   0   'False
      FillColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   7095
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2520
      Width           =   7155
   End
   Begin VB.CommandButton cmdBrowse 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Browse"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      ToolTipText     =   "Dont click if you are not going to send something, this will make the application use more memory."
      Top             =   780
      Width           =   1680
   End
   Begin VB.CommandButton cmdSendFile 
      BackColor       =   &H00F4F4F4&
      Caption         =   "&Send File"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   1680
      Width           =   1680
   End
   Begin MSWinsockLib.Winsock SckSendFile 
      Left            =   1350
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblSendStatus 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0.0%"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00F4F4F4&
      Caption         =   "Send to :"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00F4F4F4&
      Caption         =   "File Name:"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   165
      Width           =   795
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   7320
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblSendSpeed 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00F4F4F4&
      Caption         =   "0.00 Kb/Sec"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6405
      TabIndex        =   3
      Top             =   2280
      Width           =   870
   End
End
Attribute VB_Name = "frmSendFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PacketSize    As Long = 1024 * 4
Private sFileName           As String
Private iFileNum            As Integer

Private Sub cmdBrowse_Click()
On Error GoTo ErrCancel
frmMain.CDialog.ShowOpen

txtFileName = frmMain.CDialog.FileName
txtFileName.SelStart = Len(txtFileName)

Exit Sub
ErrCancel:
End Sub

Private Sub PicShowPercentage(pic As PictureBox, ByVal Percentage As Single, Optional ByVal ForeColor As Long = -2, Optional ByVal BackColor As Long = -2)
If ForeColor = -2 Then ForeColor = pic.ForeColor
If BackColor = -2 Then BackColor = pic.BackColor

pic.Line (0, 0)-(pic.ScaleWidth * Percentage, pic.ScaleHeight), ForeColor, BF
pic.Line (pic.ScaleWidth * Percentage, pic.ScaleHeight)-(pic.ScaleWidth, 0), BackColor, BF
End Sub

Private Sub cmdSendFile_Click()
If cmdSendFile.Caption = SF_COMMAND_SENDFILE Then
    tmrCalcSpeed.Enabled = True
    
    'If no user selected then exit
    If Combo1.ListIndex < 0 Then
        MsgBox SF_MSG_USER, vbInformation
        Combo1.SetFocus
        Exit Sub
    End If
    
    If Len(txtFileName) = 0 Then
        MsgBox SF_MSG_FILE, vbInformation
        txtFileName.SetFocus
        Exit Sub
    End If
    
    'Request IP
    SendMSG "!iprequest#" & Combo1.Text & "#"
Else
    tmrCalcSpeed.Enabled = False
    SckSendFile.Close
    cmdSendFile.Caption = SF_COMMAND_SENDFILE
End If
End Sub

Public Sub SendF(IP As String)
If cmdSendFile.Caption = SF_COMMAND_SENDFILE Then
    If Len(txtFileName) > 0 Then
        If FileExists(txtFileName) Then
            cmdSendFile.Caption = SF_COMMAND_CANCEL
            SendFile txtFileName, IP
        Else
            MsgBox "File doesn't exist.", vbInformation
        End If
    End If
Else
    SckSendFile_Close
    cmdSendFile.Caption = SF_COMMAND_SENDFILE
End If
End Sub

Public Sub SendFile(ByVal FileName As String, Host As String)
tmrSendFile.Enabled = False

SckSendFile.Close
DoEvents

SckSendFile.RemoteHost = Host
SckSendFile.RemotePort = bPort
SckSendFile.Connect

sFileName = FileName
End Sub

Private Sub Combo1_Click()
If Combo1.ListCount = 0 Then
    cmdSendFile.Enabled = False
    Exit Sub
End If

If Combo1.Text = frmConfig.txtNick Then
    cmdSendFile.Enabled = False
Else
    cmdSendFile.Enabled = True
End If
End Sub

Private Sub Form_Load()
Top = 0: Left = 0
LoadSendFileForm
End Sub

Public Sub LoadSendFileForm()
Label1.Caption = SF_LABEL_FILENAME
cmdBrowse.Caption = SF_COMMAND_BROWSE
cmdSendFile.Caption = SF_COMMAND_SENDFILE
Label4.Caption = SF_LABEL_SEND_TO
End Sub

Private Sub SckSendFile_Close()
SckSendFile.Close

tmrSendFile.Enabled = False
tmrSendFile.Interval = 0
tmrCalcSpeed.Enabled = False


Close iFileNum  'Close file
iFileNum = 0    'Set file number to 0, timer will exit if another timer event

cmdSendFile.Caption = SF_COMMAND_SENDFILE
PicShowPercentage Me.picProgress, 0
lblSendStatus.Caption = "0.0%"
End Sub

Private Sub SckSendFile_Connect()
Dim Buffer() As Byte, p As Long

iFileNum = FreeFile
Open sFileName For Binary Access Read Lock Write As iFileNum

ReDim Buffer(lngMIN(LOF(iFileNum), PacketSize) - 1)
Get iFileNum, , Buffer ' read data

SckSendFile.SendData CStr(LOF(iFileNum)) & ","   ' send the file size
p = InStrRev(sFileName, "\")
SckSendFile.SendData Mid(sFileName, p + 1) & ":" ' send the file name
SckSendFile.SendData Buffer                      ' send first packet
End Sub

Private Sub SckSendFile_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
tmrCalcSpeed.Enabled = False
SckSendFile_Close
End Sub

Private Sub SckSendFile_SendComplete()
'Can't call SendData here, so enable timer to do it...
tmrSendFile.Enabled = False
tmrSendFile.Interval = 1
tmrSendFile.Enabled = True
End Sub

Private Sub tmrCalcSpeed_Timer()
Static PrevSent As Long
Dim BSentPerSec As Long, DataSent As Long, DataLen As Long
Dim SecondsLeft As Single

If iFileNum > 0 Then
    DataSent = Loc(iFileNum)
    DataLen = LOF(iFileNum)
    
    BSentPerSec = DataSent - PrevSent
    PrevSent = DataSent
    
    SecondsLeft = (CSng(DataLen) - CSng(DataSent)) / CSng(BSentPerSec)
Else
    PrevSent = 0
End If

lblSendSpeed.Caption = Format$(DataSent / 1024#, "###,###,##0.00") & SF_LABEL_KBSS & _
    Format(CDbl(BSentPerSec) / 1024#, "#0.00 ") & SF_LABEL_KBS & _
    SF_LABEL_TIME & Format$(SecondsLeft \ 3600, "00") & ":" & Format$(SecondsLeft \ 60, "00") & ":" & Format$(SecondsLeft Mod 60, "00")
End Sub

Private Sub tmrSendFile_Timer()
Dim Buffer() As Byte, BuffSize As Long

tmrSendFile.Enabled = False
If iFileNum <= 0 Or SckSendFile.State <> sckConnected Then Exit Sub

If Loc(iFileNum) >= LOF(iFileNum) Then 'FILE COMPLETE
    Close iFileNum  'Close file
    iFileNum = 0    'Set file number to 0, timer will exit if another timer event
    
    Exit Sub
End If

'if the remaining size in the file is smaller then PacketSize, the read only whatever is left
BuffSize = lngMIN(LOF(iFileNum) - Loc(iFileNum), PacketSize)

ReDim Buffer(BuffSize - 1)  'resize buffer
Get iFileNum, , Buffer      'read data
SckSendFile.SendData Buffer 'send data

'Show progress
lblSendStatus.Caption = Format$(Loc(iFileNum) / CDbl(LOF(iFileNum)), "Percent")
PicShowPercentage picProgress, Loc(iFileNum) / CDbl(LOF(iFileNum))
End Sub

Private Function lngMIN(ByVal L1 As Long, ByVal L2 As Long) As Long
If L1 < L2 Then
    lngMIN = L1
Else
    lngMIN = L2
End If
End Function
