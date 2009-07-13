VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Begin VB.Form frmSendFile 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3735
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFileName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   975
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   240
      Width           =   6450
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
   Begin VB.TextBox txtRemoteCon 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      MaxLength       =   128
      TabIndex        =   7
      Text            =   "localhost"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtRemotePort 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   6
      Text            =   "8866"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picProgress 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00C00000&
      Height          =   285
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   7215
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2520
      Width           =   7275
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   600
      Width           =   1680
   End
   Begin VB.CommandButton cmdSendFile 
      Caption         =   "&Send File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
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
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   4440
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "Remote Connection:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "Remote Port:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2280
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "File Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   285
      Width           =   750
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   7440
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblFileToSend 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "Sending File:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   915
   End
   Begin VB.Label lblProgress 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "0.0% Sent"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      TabIndex        =   4
      Top             =   2280
      Width           =   780
   End
   Begin VB.Label lblSendSpeed 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "0.00 Kb/Sec"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6525
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

' Made By Michael Ciurescu (CVMichael from vbforums.com)

Private Const PacketSize As Long = 1024 * 4

Private sFileName As String
Private iFileNum As Integer

Private Sub cmdBrowse_Click()
    On Error GoTo ErrCancel
    CDialog.ShowOpen
    
    txtFileName.Text = CDialog.FileName
    txtFileName.SelStart = Len(txtFileName.Text)
    
    Exit Sub
ErrCancel:
    If Err.Number = 32755 Then
        Exit Sub
    Else
        MsgBox "Error# " & Err.Number & vbNewLine & "Description: " & Err.Description, vbExclamation
    End If
End Sub

Private Sub PicShowPercentage(pic As PictureBox, ByVal Percentage As Single, Optional ByVal ForeColor As Long = -2, Optional ByVal BackColor As Long = -2)
    If ForeColor = -2 Then ForeColor = pic.ForeColor
    If BackColor = -2 Then BackColor = pic.BackColor
    
    pic.Line (0, 0)-(pic.ScaleWidth * Percentage, pic.ScaleHeight), ForeColor, BF
    pic.Line (pic.ScaleWidth * Percentage, pic.ScaleHeight)-(pic.ScaleWidth, 0), BackColor, BF
End Sub

Private Sub cmdSendFile_Click()
    If cmdSendFile.Caption = "&Send File" Then
        If Len(txtFileName.Text) > 0 Then
            If Len(Dir(txtFileName.Text, vbNormal + vbArchive)) > 0 Then
                cmdSendFile.Caption = "&Cancel Sending..."
                
                SendFile txtFileName.Text
            End If
        End If
    Else
        SckSendFile_Close
        cmdSendFile.Caption = "&Send File"
    End If
End Sub

Private Sub SendFile(ByVal FileName As String)
    tmrSendFile.Enabled = False
    
    SckSendFile.Close
    DoEvents
    
    SckSendFile.RemoteHost = Me.txtRemoteCon.Text
    SckSendFile.RemotePort = Val(Me.txtRemotePort.Text)
    SckSendFile.Connect
    
    sFileName = FileName
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    txtRemotePort.Text = frmConfig.txtPort.Text + 1
    txtRemoteCon.Text = frmConfig.txtIP.Text
End Sub

Private Sub SckSendFile_Close()
    tmrSendFile.Enabled = False
    tmrSendFile.Interval = 0
    
    Close iFileNum ' close file
    iFileNum = 0 ' set file number to 0, timer will exit if another timer event
    
    SckSendFile.Close
    
    cmdSendFile.Caption = "&Send File"
    PicShowPercentage Me.picProgress, 0
    lblProgress.Caption = "0.00% Done"
End Sub

Private Sub SckSendFile_Connect()
    Dim Buffer() As Byte, P As Long
    
    iFileNum = FreeFile
    Open sFileName For Binary Access Read Lock Write As iFileNum
    
    ReDim Buffer(lngMIN(LOF(iFileNum), PacketSize) - 1)
    Get iFileNum, , Buffer ' read data
    
    SckSendFile.SendData CStr(LOF(iFileNum)) & ","   ' send the file size
    P = InStrRev(sFileName, "\")
    SckSendFile.SendData Mid(sFileName, P + 1) & ":" ' send the file name
    SckSendFile.SendData Buffer                      ' send first packet
End Sub

Private Sub SckSendFile_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    SckSendFile_Close
End Sub

Private Sub SckSendFile_SendComplete()
    ' can't call SendData here, so enable timer to do it...
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
    
    lblSendSpeed.Caption = Format(DataSent / 1024#, "###,###,##0.00") & " KBytes Sent, " & _
        Format(CDbl(BSentPerSec) / 1024#, "#0.00") & " Kb/Sec, " & _
        "Time left: " & Format(SecondsLeft \ 3600, "00") & ":" & Format(SecondsLeft \ 60, "00") & ":" & Format(SecondsLeft Mod 60, "00")
End Sub

Private Sub tmrSendFile_Timer()
    Dim Buffer() As Byte, BuffSize As Long
    
    tmrSendFile.Enabled = False
    If iFileNum <= 0 Or SckSendFile.State <> sckConnected Then Exit Sub
    
    If Loc(iFileNum) >= LOF(iFileNum) Then ' FILE COMPLETE
        Close iFileNum ' close file
        iFileNum = 0 ' set file number to 0, timer will exit if another timer event
        
        Exit Sub
    End If
    
    'if the remaining size in the file is smaller then PacketSize, the read only whatever is left
    BuffSize = lngMIN(LOF(iFileNum) - Loc(iFileNum), PacketSize)
    
    ReDim Buffer(BuffSize - 1) ' resize buffer
    Get iFileNum, , Buffer ' read data
    SckSendFile.SendData Buffer ' send data
    
    ' Show progress
    lblProgress.Caption = Format(Loc(iFileNum) / CDbl(LOF(iFileNum)), "Percent") & " Done"
    PicShowPercentage picProgress, Loc(iFileNum) / CDbl(LOF(iFileNum))
End Sub

Private Sub txtRemotePort_Validate(Cancel As Boolean)
    txtRemotePort.Text = Val(txtRemotePort.Text)
End Sub

Public Function lngMIN(ByVal L1 As Long, ByVal L2 As Long) As Long
    If L1 < L2 Then
        lngMIN = L1
    Else
        lngMIN = L2
    End If
End Function


