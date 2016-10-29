VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSendFile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FileName"
   ClientHeight    =   1215
   ClientLeft      =   5850
   ClientTop       =   2730
   ClientWidth     =   4575
   Icon            =   "frmSendFile.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin MSWinsockLib.Winsock socket 
      Left            =   4080
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pgProgress 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Accept"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Accept"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblPercent 
      Caption         =   "100%"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3550
   End
End
Attribute VB_Name = "frmSendFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum FileStat
    RECV_FILE = 0
    SEND_FILE = 1
End Enum

Public sFilePath As String
Public sFileName As String
Public iFileLen As Long
Public fStat As FileStat
Public frReferer As frmChat
Dim bTransferring As Boolean
Dim bInProgress As Boolean
Dim bTransferComplete As Boolean
Dim fNum As Integer
Dim iCurLen As Integer
Dim iLenSoFar As Long

Private Sub cmdAccept_Click()
If socket.State <> sckConnected Then GoTo exit_accept
socket.SendData "[ACC_FILE]"
cmdAccept.Visible = False
cmdCancel.Top = lblStatus.Top
cmdCancel.Caption = "&Abort"
bTransferring = True

pgProgress.Visible = True
pgProgress.Value = 1
lblPercent.Visible = True
lblPercent.Caption = "0%"

fNum = FreeFile
Open RecvFilePath & sFileName For Binary Access Write As #fNum

exit_accept:
End Sub

Private Sub cmdCancel_Click()
If fStat = RECV_FILE And Not bTransferComplete Then
        If socket.State <> sckConnected Then GoTo exit_cancel
        socket.SendData "[REJ_FILE]"
        bInProgress = True
        Do While bInProgress
            DoEvents
        Loop
End If
exit_cancel:
Unload Me
End Sub

Private Sub Form_Load()
pgProgress.Visible = False
lblPercent.Visible = False
lblStatus.Caption = "Connecting to remote host..."
bTransferring = False
bTransferComplete = False
iLenSoFar = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
frReferer.mnuSendFile.Enabled = True
End Sub

Private Sub socket_Close()
If fStat = RECV_FILE Then
    If bTransferring Then
        Close #fNum
        bTransferring = False
        pgProgress.Visible = False
        lblPercent.Visible = False
        lblStatus.Caption = Me.Tag & " has aborted transferring file " & sFileName & " (" & FormatKB(iFileLen) & ")"
        cmdCancel.Caption = "E&xit"
    Else
        Unload Me
    End If
End If
End Sub

Private Sub socket_Connect()
If fStat = SEND_FILE Then lblStatus.Caption = "Waiting " & Me.Tag & " for accepting file " & sFileName & "(" & FormatKB(iFileLen) & ")..."
End Sub

Private Sub socket_ConnectionRequest(ByVal requestID As Long)
socket.Close
socket.Accept requestID
End Sub

Private Sub socket_DataArrival(ByVal bytesTotal As Long)
Dim DataRecv As String
Dim DataSend As String
socket.GetData DataRecv
If bTransferring Then
    If fStat = RECV_FILE Then
        iCurLen = Len(DataRecv)
        iLenSoFar = iLenSoFar + iCurLen
        If iCurLen >= iFileLen Then
            Put #fNum, , CStr(Left(DataRecv, iCurLen - (iLenSoFar - iFileLen)))
            Close #fNum
            bTransferring = False
            bTransferComplete = True
            cmdCancel.Caption = "E&xit"
            lblStatus.Caption = "Transferring file  " & sFileName & " (" & FormatKB(iFileLen) & ") has completed"
            pgProgress.Value = pgProgress.Max
            lblPercent.Caption = pgProgress.Max & "%"
        Else
            Put #fNum, , CStr(DataRecv)
            lblStatus.Caption = "Transferring file  " & sFileName & " (" & FormatKB(iLenSoFar) & " of " & FormatKB(iFileLen) & ")..."
            pgProgress.Value = (iLenSoFar / iFileLen) * pgProgress.Max
            lblPercent.Caption = CInt((iLenSoFar / iFileLen) * pgProgress.Max) & "%"
        End If
    End If
Else
    If DataRecv = "[ACC_FILE]" Then
        fNum = FreeFile
        iCurLen = 0
        iLenSoFar = 0
        Open sFilePath & sFileName For Binary As #fNum
        DataSend = Space$(2048)
        pgProgress.Visible = True
        pgProgress.Value = 1
        lblPercent.Visible = True
        lblPercent.Caption = "0%"
        Do While Seek(fNum) < iFileLen
            Get #fNum, , DataSend
            iCurLen = Len(DataSend)
            iLenSoFar = iLenSoFar + iCurLen
            If socket.State <> sckConnected Then GoTo exit_send
            bInProgress = True
            If iLenSoFar > iFileLen Then
                socket.SendData Left(DataSend, iCurLen - (iLenSoFar - iFileLen))
                lblStatus.Caption = "Transferring file  " & sFileName & " (" & FormatKB(iFileLen) & ") has completed"
                pgProgress.Value = pgProgress.Max
                lblPercent.Caption = pgProgress.Max & "%"
            Else
                socket.SendData DataSend
                lblStatus.Caption = "Transferring file  " & sFileName & " (" & FormatKB(iLenSoFar) & " of " & FormatKB(iFileLen) & ")..."
                pgProgress.Value = (iLenSoFar / iFileLen) * pgProgress.Max
                lblPercent.Caption = CInt((iLenSoFar / iFileLen) * pgProgress.Max) & "%"
            End If
            Do While bInProgress
                DoEvents
            Loop
        Loop
exit_send:
        Close #fNum
        DataSend = Empty
        socket.Close
        bTransferring = False
        bTransferComplete = True
        cmdCancel.Caption = "E&xit"
    ElseIf DataRecv = "[REJ_FILE]" Then
        lblStatus.Caption = Me.Tag & " has declined transferring file " & sFileName & " (" & FormatKB(iFileLen) & ")"
        cmdCancel.Caption = "E&xit"
        socket.Close
    End If
End If
End Sub

Private Sub socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If Not bTransferComplete Then
    cmdCancel.Caption = "E&xit"
    If fStat = RECV_FILE Then
        If bTransferring Then
            Close #fNum
            pgProgress.Visible = False
            lblPercent.Visible = False
            lblStatus.Caption = Me.Tag & " has aborted transferring file " & sFileName & " (" & FormatKB(iFileLen) & ")"
        Else
            Unload Me
        End If
    ElseIf fStat = SEND_FILE Then
        If bTransferring Then
            pgProgress.Visible = False
            lblPercent.Visible = False
            lblStatus.Caption = Me.Tag & " has aborted transferring file " & sFileName & " (" & FormatKB(iFileLen) & ")"
        Else
            lblStatus.Caption = Me.Tag & " has declined transferring file " & sFileName & " (" & FormatKB(iFileLen) & ")"
        End If
    End If
End If
End Sub

Private Sub socket_SendComplete()
bInProgress = False
End Sub
