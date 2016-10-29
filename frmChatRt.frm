VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmChat 
   BackColor       =   &H00A86717&
   Caption         =   "Conversation"
   ClientHeight    =   5010
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   3945
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   960
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtbSay 
      Height          =   540
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   953
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmChatRt.frx":0000
   End
   Begin RichTextLib.RichTextBox rtbChat 
      Height          =   3855
      Left            =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   165
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   6800
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmChatRt.frx":008B
   End
   Begin MSComctlLib.StatusBar sbInfo 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4755
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.Timer tmrChat 
      Left            =   480
      Top             =   5280
   End
   Begin MSWinsockLib.Winsock socket 
      Left            =   0
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Appearance      =   0  'Flat
      DisabledPicture =   "frmChatRt.frx":0116
      DownPicture     =   "frmChatRt.frx":01FD
      Height          =   540
      Left            =   3240
      Picture         =   "frmChatRt.frx":02C9
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape shRect 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   45
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Shape shChat 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   45
      Shape           =   4  'Rounded Rectangle
      Top             =   45
      Width           =   3855
   End
   Begin VB.Shape shSay 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   45
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Chat"
      End
      Begin VB.Menu mnuSendFile 
         Caption         =   "Send a &file..."
      End
      Begin VB.Menu mnuSpc 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close Chat"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
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
Dim iTimeOut As Byte
Dim isTyping As Boolean
Dim bIsSending As Boolean
Dim dLastMsgRecv As String
Dim lFilePort As Long

Private Sub cmdSend_Click()
If stat = connected And Len(rtbSay.Text) > 0 Then
  If Right(rtbSay.Text, 2) <> vbCrLf Then rtbSay.Text = rtbSay.Text & vbCrLf
  bIsSending = True
  rtbSay.Locked = True
  cmdSend.Enabled = False
  socket.SendData "[SAY]" & rtbSay.Text
  Do Until Not bIsSending
    DoEvents
  Loop
  rtbSay.Locked = False
  cmdSend.Enabled = True
  rtbChat.SelIndent = 10
  rtbChat.SelColor = &H666666
  rtbChat.SelText = strHostname & " says: " & vbCrLf
  rtbChat.SelIndent = 300
  rtbChat.SelColor = vbBlack
  rtbChat.SelText = rtbSay.Text
  rtbSay.Text = ""
  iTimeOut = 0
  tmrChat.Enabled = False
End If
End Sub

Private Sub Form_Initialize()
sbInfo.Panels.Add 1, , "", sbrText, Nothing
sbInfo.Panels.Add 2, , "", sbrText, Nothing
sbInfo.Panels.Add 3, , "", sbrText, Nothing
sbInfo.Panels(2).Visible = False
sbInfo.Panels(3).Visible = False
rtbChat.SelIndent = 10
rtbChat.Font.Name = "Arial"
rtbChat.Text = ""
rtbSay.Font.Name = "Arial"
rtbSay.Text = ""
cmdSend.Enabled = False
FirstCall = True
isTyping = False
dLastMsgRecv = ""
iTimeOut = 0
tmrChat.Interval = 1000
tmrChat.Enabled = True
End Sub

Public Property Let Status(ByVal sstat As sockstat)
stat = sstat
End Property

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And rtbChat <> "" Then cmdSend_Click
If KeyAscii = 27 Then mnuClose_Click
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If (KeyCode = vbKeyReturn And Shift = 0) Or (KeyCode = vbKeyS And Shift = 4) Then cmdSend_Click
End Sub

Private Sub Form_Resize()

On Error Resume Next
rtbChat.Top = 150
shChat.Left = rtbChat.Left
shChat.Width = Me.ScaleWidth - (shChat.Left * 2) + 10
rtbChat.Width = shChat.Width - 10
rtbChat.Height = Me.ScaleHeight - rtbChat.Top - shRect.Height - sbInfo.Height - (rtbChat.Left * 2) - 120
shRect.Top = rtbChat.Top + rtbChat.Height + rtbChat.Left
shRect.Left = rtbChat.Left
shRect.Width = shChat.Width
rtbSay.Top = shRect.Top
rtbSay.Left = shRect.Left
rtbSay.Width = shChat.Width - cmdSend.Width - (rtbChat.Left * 2)
shSay.Top = shRect.Top + 10
shSay.Left = shRect.Left
shSay.Width = shRect.Width
cmdSend.Top = rtbSay.Top
cmdSend.Left = rtbSay.Left + rtbSay.Width + rtbChat.Left
sbInfo.Panels(1).Width = Me.Width
End Sub

Private Sub mnuAbout_Click()
Dim fAbout As frmAbout
Set fAbout = New frmAbout
Load fAbout
Set fAbout.frReferer = Me
fAbout.Show vbModal, Me
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

Private Sub mnuSave_Click()
cdFile.Filter = "Text File|*.txt"
cdFile.ShowSave

If cdFile.FileName <> "" Then
    Dim f
    f = FreeFile
    On Error Resume Next
    Open cdFile.FileName For Output As #f
        Print #f, rtbChat.Text
        Print #f, dLastMsgRecv
    Close #f
    If Err.Number = 0 Then MsgBox "File successfully saved.", vbInformation, "Conversation Save" _
    Else MsgBox "Conversation cannot be saved.", vbCritical, "File Error"
    On Error GoTo 0
End If
End Sub

Private Sub mnuSendFile_Click()
cdFile.Filter = "All Files|*.*"
cdFile.ShowOpen
If cdFile.FileName <> "" Then
    Dim fSend As frmSendFile
    Dim istart
    Set fSend = New frmSendFile
    Load fSend
    With fSend
        .sFileName = cdFile.FileTitle
        .sFilePath = Replace(cdFile.FileName, cdFile.FileTitle, "")
        .iFileLen = FileLen(cdFile.FileName)
        .Tag = Me.Tag
        .fStat = SEND_FILE
        .cmdAccept.Visible = False
        .cmdCancel.Top = .lblStatus.Top
        Set .frReferer = Me
        .Show
        lFilePort = 0
        socket.SendData "[SEND_FILE];" & .sFileName & ";" & .iFileLen
        istart = (Timer)
        Do
            DoEvents
        Loop Until (lFilePort) Or (Timer - istart > 3)
        If lFilePort Then .socket.Connect socket.RemoteHostIP, lFilePort _
        Else Unload fSend
    End With
    mnuFile.Enabled = False
End If
End Sub

Private Sub rtbChat_KeyPress(KeyAscii As Integer)
    rtbSay.SetFocus
    rtbSay.SelText = Chr(KeyAscii)
End Sub

Private Sub rtbChat_KeyUp(KeyCode As Integer, Shift As Integer)
If (KeyCode = vbKeyReturn And Shift = 0) Or (KeyCode = vbKeyS And Shift = 4) Then cmdSend_Click
End Sub

Private Sub rtbSay_KeyUp(KeyCode As Integer, Shift As Integer)
If (KeyCode = vbKeyReturn And Shift = 0) Or (KeyCode = vbKeyS And Shift = 4) Then cmdSend_Click
End Sub

Private Sub socket_Close()
If stat = connected Then
    rtbChat.SelIndent = 10
    rtbChat.SelBold = True
    rtbChat.SelText = "++ Connection was closed..."
    rtbChat.SelBold = False
    rtbSay.Enabled = False
    tmrChat.Enabled = False
    stat = ready
End If
End Sub

Private Sub socket_Connect()
FriendlyName = Trim(GetFriendlyName(socket.RemoteHostIP))
If FriendlyName = "" Then FriendlyName = Me.Tag
Me.Caption = "Conversation - " & FriendlyName
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
    rtbChat.SelIndent = 10
    rtbChat.SelBold = True
    rtbChat.SelText = "++ Connection Established!" & vbCrLf & vbCrLf
    FriendlyName = Trim(GetFriendlyName(socket.RemoteHostIP))
    If FriendlyName = "" Then FriendlyName = Me.Tag
    Me.Caption = "Conversation - " & FriendlyName
    rtbChat.SelBold = False
    rtbSay.Enabled = True
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
  rtbChat.SelIndent = 10
  rtbChat.SelBold = True
  rtbChat.SelText = "++ Connection Established!" & vbCrLf & vbCrLf
  rtbChat.SelBold = False
  rtbSay.Enabled = True
  stat = connected
ElseIf Left(strData, 6) = "[MOVE]" Then
  temp = socket.RemoteHostIP
  CloseSocket
  socket.Connect temp, Mid(strData, 8)
ElseIf Left(strData, 5) = "[SAY]" Then
  rtbChat.SelIndent = 10
  rtbChat.SelColor = &H666666
  rtbChat.SelText = FriendlyName & " says: " & vbCrLf
  rtbChat.SelIndent = 300
  rtbChat.SelColor = vbBlue
  rtbChat.SelText = Mid(strData, 6)
  If Right(rtbChat, 2) <> vbCrLf Then rtbChat.SelText = vbCrLf
  If FirstCall Then
    Dim fPop As frmPop
    Set fPop = New frmPop
    With fPop
        .lblInfo.Font.Bold = False
        .lblInfo.Caption = FriendlyName & " says:" & vbCrLf
        If Len(Mid(strData, 7)) > 50 Then
            .lblInfo.Caption = .lblInfo.Caption & Mid(strData, 7, 50) & "..."
        Else
            .lblInfo.Caption = .lblInfo.Caption & Mid(strData, 7)
        End If
        .Tag = "[CHAT] " & Me.Tag
        .Left = Screen.Width - (.Width + 10)
        .Top = Screen.Height - (.Height * PopLevel)
        Set .ActiveWindow = Me
        .Height = 30
        .OnTop
        .Show
        .tmrMove.Enabled = True
    End With
    FirstCall = False
    Me.WindowState = vbMinimized
    Me.Show
    FlashWindow Me.hwnd, 3
  End If
  dLastMsgRecv = "Last message received " & CStr(Date) & " at " & CStr(Time)
  sbInfo.Panels(1).Text = dLastMsgRecv
ElseIf strData = "[TYPING]" Then
    sbInfo.Panels(1).Text = FriendlyName & " is typing a message..."
ElseIf strData = "[NOTYPING]" Then
    sbInfo.Panels(1).Text = dLastMsgRecv
ElseIf Left(strData, 11) = "[FILE_PORT]" Then
    lFilePort = CLng(Mid(strData, 12))
ElseIf Left(strData, 11) = "[SEND_FILE]" Then
    Dim tInfo() As String
    Dim fSend As frmSendFile
    tInfo = Split(strData, ";")
    Set fSend = New frmSendFile
    Load fSend
    With fSend
        .Tag = Me.Tag
        Set .frReferer = Me
        .Caption = "File Transfer - " & Me.Tag
        .fStat = RECV_FILE
        .sFileName = tInfo(1)
        .iFileLen = tInfo(2)
        .lblStatus.Caption = Me.Tag & " would like to send file " & .sFileName & " (" & FormatKB(.iFileLen) & ")"
        .cmdCancel.Caption = "&Reject"
        .socket.Close
        .socket.LocalPort = 0
        .socket.Listen
        socket.SendData "[FILE_PORT]" & .socket.LocalPort
        Debug.Print "Send Data: [FILE_PORT]" & .socket.LocalPort
        .Show
    End With
End If
End Sub

Private Sub socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
rtbChat.SelText = "** Connection Error. Reason: " & Description & vbCrLf
rtbChat.SelBold = False
rtbSay.Enabled = False
tmrChat.Enabled = False
stat = ready
End Sub

Private Sub socket_SendComplete()
bIsSending = False
End Sub

Private Sub tmrChat_Timer()
iTimeOut = iTimeOut + 1
If iTimeOut = TIME_OUT Then
    If stat = listening Then
        Unload Me
    ElseIf stat = connecting Then
        rtbChat.SelIndent = 10
        rtbChat.SelBold = True
        rtbChat.SelText = "++ Connection failed!"
        rtbChat.SelBold = False
        tmrChat.Enabled = False
        CloseSocket
    Else
        iTimeOut = 0
        If socket.State = sckConnected Then socket.SendData "[NOTYPING]"
        isTyping = False
        tmrChat.Enabled = False
    End If
End If
End Sub

Private Sub rtbSay_Change()
If Len(rtbSay.Text) = 0 Then
    cmdSend.Enabled = False
    If socket.State = sckConnected And isTyping Then socket.SendData "[NOTYPING]"
    isTyping = False
Else
    If Not cmdSend.Enabled Then cmdSend.Enabled = True
    If Not isTyping Then
        isTyping = True
        tmrChat.Enabled = True
        If socket.State = sckConnected Then socket.SendData "[TYPING]"
    End If
End If
End Sub

Private Sub CloseSocket()
If socket.State <> sckClosed Then
    socket.Close
    Do
        DoEvents
    Loop Until socket.State = sckClosed
End If
End Sub
