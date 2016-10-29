VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChat 
   Caption         =   "Conversation"
   ClientHeight    =   5445
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar sbInfo 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   5190
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrChat 
      Left            =   480
      Top             =   4800
   End
   Begin MSWinsockLib.Winsock socket 
      Left            =   0
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   4150
      Width           =   615
   End
   Begin VB.TextBox txtSay 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   0
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4150
      Width           =   3135
   End
   Begin VB.TextBox txtChat 
      Appearance      =   0  'Flat
      Height          =   4095
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   3975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Chat"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close Chat"
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim stat As sockstat
Dim FriendlyName As String
Dim FirstCall As Boolean

Private Sub cmdSend_Click()
If stat = connected And Len(txtSay) > 0 Then
  socket.SendData "[SAY] " & txtSay.Text & vbCrLf
  txtChat.SelText = strHostname & ": " & txtSay.Text & vbCrLf
  txtSay.Text = ""
  txtSay.Refresh
ElseIf txtChat <> "" Then
    txtSay = ""
End If
End Sub

Private Sub Form_Initialize()
sbInfo.Panels(0).AutoSize
txtSay.Enabled = False
cmdSend.Enabled = False
FirstCall = True
End Sub

Public Property Let Status(ByVal sstat As sockstat)
stat = sstat
End Property

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtChat <> "" Then cmdSend_Click
If KeyAscii = 27 Then mnuClose_Click
End Sub

Private Sub Form_Resize()

On Error Resume Next
txtChat.Width = Me.ScaleWidth - (txtChat.Left * 2)
txtChat.Height = Me.ScaleHeight - txtSay.Height - 100
txtSay.Top = txtChat.Top + txtChat.Height + 50
txtSay.Width = txtChat.Width - cmdSend.Width - 200
cmdSend.Top = txtSay.Top
cmdSend.Left = txtSay.Left + txtSay.Width + 100

End Sub

Private Sub mnuClose_Click()
Dim msg
If stat = ready Then
    Unload Me
Else
    msg = MsgBox("Are you really want to end this conversation?", vbQuestion + vbYesNo, "Conversation - " & Me.Tag)
    If socket.State <> sckClosed And msg = vbYes Then
        CloseSocket
        Unload Me
    End If
End If
End Sub

Private Sub socket_Close()
If stat = connected Then
    txtChat.SelText = "** Connection was closed..."
    txtSay.Enabled = False
    stat = ready
End If
End Sub

Private Sub socket_Connect()
FriendlyName = GetFriendlyName(socket.RemoteHostIP)
If FriendlyName = "" Then FriendlyName = Me.Tag
If stat = connecting Then
    FirstCall = False
  socket.SendData "[CHAT]"
End If
End Sub

Private Sub socket_ConnectionRequest(ByVal requestID As Long)
If stat = listening Then
    CloseSocket
    socket.Accept requestID
    socket.SendData "[CHAT_OK]"
    txtChat.SelText = "** Connection Established!" & vbCrLf & vbCrLf
    Me.Caption = "Conversation - " & CallName
    txtSay.Enabled = True
    stat = connected
    tmrChat.Enabled = False
End If
End Sub

Private Sub socket_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
Dim temp

socket.GetData strData
Debug.Print "Get Data : " & strData

If Left(strData, 9) = "[CHAT_OK]" Then
  txtChat.SelText = "** Connection Established!" & vbCrLf & vbCrLf
  Me.Caption = "Conversation - " & Me.Tag
  txtSay.Enabled = True
  stat = connected
ElseIf Left(strData, 6) = "[MOVE]" Then
  temp = socket.RemoteHostIP
  CloseSocket
  socket.Connect temp, Mid(strData, 8)
ElseIf Left(strData, 5) = "[SAY]" Then
  txtChat.SelText = FriendlyName & ": " & Mid(strData, 7)
  If Right(txtChat, 2) <> vbCrLf Then txtChat.SelText = vbCrLf
  If FirstCall Then
    Dim fPop As frmPop
    Set fPop = New frmPop
    With fPop
        .lblName.Caption = FriendlyName & ":"
        If Mid(strData, 7) > 20 Then
            .lblInfo.Caption = Mid(strData, 7, 20) & "..."
        Else
            .lblInfo.Caption = Mid(strData, 7)
        End If
        PopLoad = PopLoad + 1
        If Screen.Height - (.Height * PopLoad) - 100 < 0 Then PopLoad = 0
        .Tag = Me.Tag
        .Left = Screen.Width - (.Width + 10)
        .Top = Screen.Height - (.Height * PopLoad) - 100
        .Height = 30
        .Show
        .OnTop
        .tmrMove.Enabled = True
    End With
    FirstCall = False
    Me.WindowState = vbMinimized
    Me.Show
  End If
End If

End Sub

Private Sub SOCKET_ERROR(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
txtChat.SelText = "** Socket Error : " & Description & vbCrLf
End Sub

Private Sub tmrChat_Timer()
If stat = listening Then Unload Me
End Sub

Private Sub txtSay_Change()
If Len(txtChat) = 0 Then cmdSend.Enabled = False _
Else cmdSend.Enabled = True
End Sub

Private Sub CloseSocket()
If socket.State <> sckClosed Then
    socket.Close
    Do
        DoEvents
    Loop Until socket.State = sckClosed
End If
End Sub
