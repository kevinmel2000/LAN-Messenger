VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "LAN Messenger"
   ClientHeight    =   7080
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   3315
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   3315
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   6825
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   388
            MinWidth        =   388
            Text            =   "icon"
            TextSave        =   "icon"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "5:58 AM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   1560
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C42
            Key             =   "comp"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B1E
            Key             =   "green"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BAA
            Key             =   "yellow"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EFE
            Key             =   "red"
         EndProperty
      EndProperty
   End
   Begin VB.Frame frMain 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin MSComctlLib.TreeView tvComp 
         Height          =   1335
         Left            =   480
         TabIndex        =   7
         Top             =   600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2355
         _Version        =   393217
         Style           =   7
         Appearance      =   0
      End
      Begin VB.Label lblGoSign 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click To Sign In"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   795
         TabIndex        =   5
         Top             =   2400
         Width           =   1905
      End
      Begin VB.Label lblStat 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "+ Status here"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label lblSignIn 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Signing In..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   2115
         Width           =   2055
      End
      Begin VB.Label lblMyInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Computer Name <IP Address>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   190
         Width           =   2775
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   480
         X2              =   3600
         Y1              =   430
         Y2              =   430
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "My Info:"
         ForeColor       =   &H80000010&
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   0
         Width           =   1335
      End
      Begin VB.Image imgMyIcon 
         Height          =   480
         Left            =   0
         Picture         =   "frmMain.frx":2252
         Top             =   40
         Width           =   480
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   5655
         Left            =   0
         Top             =   0
         Width           =   3615
      End
   End
   Begin VB.Timer tmrPoll 
      Left            =   960
      Top             =   6240
   End
   Begin MSWinsockLib.Winsock SockListen 
      Index           =   0
      Left            =   0
      Top             =   6240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock SockConnect 
      Index           =   0
      Left            =   480
      Top             =   6240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSignIn 
         Caption         =   "&Sign In"
      End
      Begin VB.Menu mnuSignOut 
         Caption         =   "Sign &Out"
      End
      Begin VB.Menu spc1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetting 
         Caption         =   "&Setting"
      End
      Begin VB.Menu spc2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowMe 
         Caption         =   "Show Me!"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sstat As sockstat
Dim JobSck As Byte
Dim tCount As Long

Private Sub Form_Load()

Me.Caption = App.title

'*** No Network, No Way
If Not IsNetworkInstalled Then
  MsgBox "No Network Installed. You cannot use this software.", vbCritical, App.title
  End
End If

'** Load available settings
On Error Resume Next
PollingTime = CLng(GetSetting(App.title, "settings", "polltime"))
If PollingTime = 0 Then PollingTime = POLL_INTERVAL
RecvFilePath = GetSetting(App.title, "settings", "file path")

'** Preparing accessories
sstat = ready
sbStatus.Panels(1).Picture = ImageList.ListImages(4).Picture
tvComp.ImageList = ImageList
tvComp.Visible = False
PopLoad = 0
PopLevel = 0

'*** adding systray icon
ShellTrayAdd "LAN Messenger - Offline", Me
sbStatus.Panels(2).Text = "Offline"

'*** get local info
strHostname = GetComputerName
strIpAddress = GetIPFromHostName(strHostname)
lblMyInfo.Caption = UCase(strHostname) & " <" & strIpAddress & ">"

'*** set the labels
mnuSignOut.Enabled = False
lblSignIn.Visible = False
lblStat(0).Visible = False
lblGoSign.Visible = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblGoSign.ForeColor = vbBlack
lblGoSign.BackColor = vbWhite

Dim result As Long
Dim msg As Long
If Me.ScaleMode = vbPixels Then
msg = X
Else
msg = X / Screen.TwipsPerPixelX
End If
Select Case msg
Case WM_LBUTTONDBLCLK    '515 restore form window
If Me.Visible = False Then
    Me.WindowState = vbNormal
    mnuShowMe.Visible = False
    Me.Show
Else
    Me.WindowState = vbMinimized
    Me.Hide
    mnuShowMe.Visible = True
End If
Case WM_RBUTTONUP
  PopupMenu mnuFile
End Select

End Sub

Private Sub Form_Resize()
Dim i As Byte

On Error Resume Next
frMain.Width = Me.Width - 120
frMain.Height = Me.ScaleHeight - sbStatus.Height - frMain.Left
Shape1.Left = frMain.Left
Shape1.Width = frMain.Width
Shape1.Height = frMain.Height
Line1.X2 = Shape1.Left + Shape1.Width
lblSignIn.Left = (Shape1.Width / 2) - (lblSignIn.Width / 2)
lblSignIn.Top = (Shape1.Height / 2) - (lblSignIn.Height / 2)
lblGoSign.Left = (Shape1.Width / 2) - (lblGoSign.Width / 2)
lblGoSign.Top = (Shape1.Height / 2) - (lblGoSign.Height / 2)
lblStat(0).Left = lblSignIn.Left
lblStat(0).Top = lblSignIn.Top + lblSignIn.Height
tvComp.Height = Shape1.Height - tvComp.Top - 50
tvComp.Width = Shape1.Width - tvComp.Left - 50
sbStatus.Panels(2).MinWidth = Me.ScaleWidth - sbStatus.Panels(1).MinWidth - sbStatus.Panels(3).MinWidth
If Me.WindowState = vbMinimized Then
    Me.Hide
    mnuShowMe.Visible = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
ShellTrayRemove
Set frmMain = Nothing
End
End Sub

Private Sub frMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = vbArrow
lblGoSign.ForeColor = vbBlack
lblGoSign.BackColor = vbWhite
End Sub

Private Sub lblGoSign_Click()
Call GoSignIn
End Sub

Private Sub lblGoSign_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblGoSign.ForeColor = vbWhite
lblGoSign.BackColor = &H8000000D
Me.MousePointer = vbUpArrow
End Sub

Private Sub mnuAbout_Click()
Dim fAbout As frmAbout
Set fAbout = New frmAbout
Load fAbout
Set fAbout.frReferer = Me
fAbout.Show vbModal, Me
End Sub

Private Sub mnuClose_Click()
If sstat = connected Then
  If MsgBox("This will close any conversation. Are you sure?", vbOKCancel, App.title) = vbOK Then
    mnuSignOut_Click
    Unload Me
  End If
Else
    Unload Me
End If
End Sub

Private Sub mnuSetting_Click()
Load frmSetting
frmSetting.Show vbModal, Me
End Sub

Private Sub mnuShowMe_Click()
Me.WindowState = vbNormal
mnuShowMe.Visible = False
Me.Show
End Sub

Private Sub mnuSignIn_Click()
Call GoSignIn
End Sub

Private Sub mnuSignOut_Click()

sbStatus.Panels(1).Picture = ImageList.ListImages(3).Picture
sbStatus.Panels(2).Text = "Disconnecting..."
sstat = disconnecting
If FriendCount > 0 Then
    Dim i
    Dim timestart
    
    '*** say good bye to all friend(s)
    JobSck = 0
    For i = 0 To FriendCount - 1
      If i > 0 Then Load SockConnect(i)
      SockConnect(i).LocalPort = 0
      SockConnect(i).Connect arrFriendList(i + 1).ipaddress, LISTENPORT
    Next
    
    '*** wait until TIME_OUT
    timestart = (Timer)
    Do
      DoEvents
    Loop Until ((Timer) - (timestart) > TIME_OUT) Or (JobSck = FriendCount)
    
    For i = 0 To FriendCount - 1
      If SockConnect(i).State = sckConnected And _
        IsOnFriendList(SockConnect(i).RemoteHostIP) Then
        '*** Tell your buddy that you've just signed out
        SockConnect(i).SendData "[SIGNOUT]"
        Debug.Print "SIGNOUT From " & SockConnect(i).RemoteHostIP
      End If
    Next
           
    '*** Free some memory
    Erase arrFriendList()
    Erase arrHostList()
    
    '*** Reset host and friend count
    FriendCount = 0
    HostsCount = 0
    ShowList
    
    '*** Close all current connections excluding conversation and file transfer
    CleanUpSocketConnect
    CleanUpSocketListen
    CloseSockListen 0
End If

'*** set to offline status
lblGoSign.Visible = True
sbStatus.Panels(2).Text = "Offline"
ShellTrayModify "LAN Messenger - Offline", Me
tmrPoll.Enabled = False
sbStatus.Panels(1).Picture = ImageList.ListImages(4).Picture
sstat = ready
End Sub

Private Sub SockConnect_Connect(Index As Integer)
If sstat = connecting Or sstat = disconnecting Then JobSck = JobSck + 1
End Sub

Private Sub SockConnect_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If sstat = connecting Or sstat = disconnecting Then JobSck = JobSck + 1
Debug.Print "socket error: " & Description & " [ " & Index & " ]"
End Sub

Private Sub SockListen_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim tSock As Byte
tSock = 0
NEXT_SOCK:
tSock = tSock + 1
If tSock > SockListen.UBound Then Load SockListen(tSock)
If SockListen(tSock).State = sckClosed Then
  SockListen(tSock).Accept requestID
Else
  GoTo NEXT_SOCK
End If

End Sub

Private Sub SockListen_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String
Dim fPop As frmPop
Dim fChat As frmChat
Dim tHost As String, tIp As String

SockListen(Index).GetData strData
tIp = SockListen(Index).RemoteHostIP
If Not IsOnHostList(tIp) Then
    AddToHostList tIp, GetHostNameFromIP(tIp)
    AddToFriendList tIp
    ShowList
ElseIf Not IsOnFriendList(tIp) Then
    AddToFriendList tIp
    ShowList
End If

tHost = GetHostFromList(tIp)

Debug.Print "Data arrival from " & tHost & ". " & "Data : " & strData

If Left(strData, 8) = "[SIGNIN]" Then
  Set fPop = New frmPop
  Load fPop
  With fPop
    .lblInfo.Font.Bold = True
    .lblInfo.Caption = GetFriendlyNameFromList(tIp) & vbCrLf & "has just signed in"
    .Tag = "[SIGNIN]" & tHost
    .Left = Screen.Width - (.Width + 10)
    .Top = Screen.Height - (.Height * PopLevel)
    Set .ActiveWindow = Nothing
    .Height = 30
    .OnTop
    .Show
    .tmrMove.Enabled = True
  End With
ElseIf strData = "[SIGNOUT]" Then
  DeleteFromFriendList SockListen(Index).RemoteHostIP
ElseIf strData = "[CHAT]" Then
  Set fChat = New frmChat
  Load fChat
  With fChat
    .Caption = "Conversation"
    .Icon = Me.Icon
    .Tag = tHost
    .socket.Close
    .socket.LocalPort = 0
    .socket.Listen
    .Status = listening
    Debug.Print "Form Chat listening on port : " & .socket.LocalPort
    SockListen(Index).SendData "[MOVE] " & .socket.LocalPort
  End With
End If

End Sub

Private Sub tmrPoll_Timer()
tCount = tCount + 1
If tCount >= PollingTime Then
    Call Polling
    tCount = 0
End If
End Sub

Private Sub tvComp_DblClick()
If tvComp.SelectedItem <> "" Then
  Dim fChat As frmChat
  '*** start a conversation with a selected friend
  Set fChat = New frmChat
  Load fChat
  
  With fChat
    .Caption = "Conversation"
    .Icon = Me.Icon
    .Tag = tvComp.SelectedItem.Key
    .cmdSend.Enabled = False
    .Status = connecting
    .socket.Connect GetIpFromList(tvComp.SelectedItem.Key), LISTENPORT
    .tmrChat.Enabled = True
    .Show
    '.txtChat.SelText = "** Connecting to " & tvComp.SelectedItem.Text & "..." & vbCrLf
  End With
End If
End Sub

Private Sub GoSignIn()
Dim i As Byte
Dim timestart
Dim tNR As NETRESOURCE

sstat = connecting
'*** set the labels
Me.MousePointer = vbArrow
lblSignIn.Visible = True
lblGoSign.Visible = False
lblStat(0).Visible = True
sbStatus.Panels(2).Text = "Signing in..."
mnuSignIn.Enabled = False
mnuSignOut.Enabled = True
sbStatus.Panels(1).Picture = ImageList.ListImages(3).Picture

'*** Getting hosts information on Local Area Network
ReDim arrHostList(1 To 254)
ReDim arrFriendList(1 To 254)
lblStat(0).Caption = "+ Getting Hosts On Network..."
DoEvents
HostsCount = 0
GetHostList tNR
If HostsCount > 0 Then DeleteFromHostList strIpAddress

'*** This is used to debug only
Debug.Print "Host List:"
For i = 1 To HostsCount
'  DoEvents
'  tvComp.Nodes.Add , , arrHostList(i).hostname, arrHostList(i).hostname & " <" & arrHostList(i).ipaddress & ">", "comp", "comp"
Debug.Print i & ". " & arrHostList(i).hostname & "<" & arrHostList(i).ipaddress & ">"
Next

'*** Preparing the sockets
CloseSockListen (0)
CloseSockConnect (0)
SockListen(0).LocalPort = LISTENPORT
SockListen(0).Listen

'*** Call every hosts on the list for connection
lblStat(0).Caption = "+ Contacting Hosts..."
FriendCount = 0

If HostsCount > 0 Then
  JobSck = 0
  For i = 0 To HostsCount - 1
   If i > 0 Then Load SockConnect(i)
   CloseSockConnect (i)
   SockConnect(i).LocalPort = 0
   SockConnect(i).Connect arrHostList(i + 1).ipaddress, LISTENPORT
   Debug.Print "Socket " & i & " : " & arrHostList(i + 1).ipaddress
  Next

'*** Wait until TIME_OUT
timestart = (Timer)
Do
  DoEvents
Loop Until ((Timer) - (timestart) > TIME_OUT) Or (JobSck = HostsCount)

tvComp.Visible = True
For i = 0 To HostsCount - 1
    DoEvents
    If SockConnect(i).State = sckConnected And _
    Not IsOnFriendList(SockConnect(i).RemoteHostIP) Then
      '*** Tell your buddy that you've just sign in
      Debug.Print "SIGNIN to " & SockConnect(i).RemoteHost
      SockConnect(i).SendData "[SIGNIN]"
      AddToFriendList (SockConnect(i).RemoteHostIP)
    End If
Next
CleanUpSocketConnect
ShowList
End If

'Time To Online, buddy!
sstat = connected
lblSignIn.Visible = False
lblStat(0).Visible = False
tmrPoll.Interval = 1
tmrPoll.Enabled = True
sbStatus.Panels(2).Text = "Online"
ShellTrayModify "LAN Messenger - Online", Me
sbStatus.Panels(1).Picture = ImageList.ListImages(2).Picture
End Sub

Private Sub CloseSockListen(Index As Integer)
If SockListen(Index).State <> sckClosed Then
  SockListen(Index).Close
  Do
    DoEvents
  Loop Until SockListen(Index).State = sckClosed
End If
End Sub

Private Sub CloseSockConnect(Index As Integer)
  If SockConnect(Index).State <> sckClosed Then
    SockConnect(Index).Close
    Do
      DoEvents
    Loop Until SockConnect(Index).State = sckClosed
  End If
End Sub

Public Sub ShowList()
Dim i As Byte

If tvComp.Nodes.Count > 0 Then
    Do
        tvComp.Nodes.Remove 1
    Loop Until tvComp.Nodes.Count = 0
End If
If FriendCount > 0 Then
  For i = 1 To FriendCount
    DoEvents
    If arrFriendList(i).fullname = "" Then
        tvComp.Nodes.Add , , arrFriendList(i).hostname, arrFriendList(i).hostname & " <" & arrFriendList(i).ipaddress & ">", "comp", "comp"
    Else
        tvComp.Nodes.Add , , arrFriendList(i).hostname, arrFriendList(i).fullname, "comp", "comp"
    End If
  Next
  DoEvents
End If
tvComp.Refresh
If tvComp.Nodes.Count > 0 Then tvComp.Visible = True Else tvComp.Visible = False
End Sub

Private Sub CleanUpSocketConnect()
Dim i

For i = 0 To SockConnect.UBound
    CloseSockConnect (i)
    If i > 0 Then Unload SockConnect(i)
Next

End Sub

Private Sub CleanUpSocketListen()
Dim i

For i = 0 To SockListen.UBound
    CloseSockListen (i)
    If i > 0 Then Unload SockListen(i)
Next

End Sub
Private Sub Polling()
Dim i, j
Dim timestart
If FriendCount > 0 Then
    For i = 0 To FriendCount - 1
      If SockConnect(i) Is Nothing Then Load SockConnect(i)
      SockConnect(i).Connect arrFriendList(i + 1).ipaddress, LISTENPORT
    Next
    
    'tunggu 5 detik
    timestart = (Timer)
    Do
      DoEvents
    Loop Until (Timer) - (timestart) > TIME_OUT
    j = 0
    For i = 0 To FriendCount - 1
      If SockConnect(i).State <> sckConnected Then
        DeleteFromFriendList SockConnect(i).RemoteHostIP
        j = j + 1
      End If
    Next
    If FriendCount > 0 Then FriendCount = FriendCount - j
    CleanUpSocketConnect
End If
End Sub

Private Sub tvComp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then tvComp_DblClick
End Sub
