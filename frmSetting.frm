VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   6855
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5055
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "File Transfer"
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   4815
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         Height          =   300
         Left            =   3840
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtPath 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label3 
         Caption         =   "Files received from other user will be put in this folder:"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   3960
      TabIndex        =   9
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   330
      Left            =   2760
      TabIndex        =   8
      Top             =   6720
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Polling"
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   4815
      Begin VB.TextBox txtMinute 
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton optOther 
         Caption         =   "Other (in minute)"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   1200
         Width           =   1575
      End
      Begin VB.OptionButton optDefault 
         Caption         =   "Default (5 minute)"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "You can determine how often LAN Messenger make polling to verify your and  your friend's online status"
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Aliases"
      Height          =   3550
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin MSComctlLib.ListView lvAliases 
         Height          =   2655
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Ip Address"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Full Name"
            Object.Width           =   5115
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "You can give a permanent name to replace computer name of your friend to make it easier for you to recognize him/her"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Menu mnuAlias 
      Caption         =   "Alias"
      Begin VB.Menu mnuAddAlias 
         Caption         =   "Add Alias"
      End
      Begin VB.Menu mnuEditAlias 
         Caption         =   "Edit Alias"
      End
      Begin VB.Menu mnuDelAlias 
         Caption         =   "Delete Alias"
      End
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()
txtPath.Text = GetFolder("Please select a folder for storing received files:")
If txtPath.Text = "" Then txtPath.Text = App.Path
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
If optOther.Value Then
    If Not IsNumeric(txtMinute) Then
        MsgBox "Please fill with numeric value!", vbExclamation + vbOKOnly, Me.Caption
        txtMinute.SelStart = 0
        txtMinute.SelLength = Len(txtMinute)
        Exit Sub
    End If
    If Val(txtMinute) > 60 Or Val(txtMinute) < 1 Then
        MsgBox "Please enter a number between 1 and 60", vbExclamation + vbOKOnly, Me.Caption
        txtMinute.SelStart = 0
        txtMinute.SelLength = Len(txtMinute)
        Exit Sub
    End If
End If
Call WriteToFile
If optDefault.Value Then
    PollingTime = POLL_INTERVAL
ElseIf optOther.Value Then
    PollingTime = CLng(txtMinute) * 60000
End If
RecvFilePath = txtPath.Text
SaveSetting App.title, "settings", "polltime", CStr(PollingTime)
SaveSetting App.title, "settings", "file path", txtPath.Text
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdOk_Click
If KeyAscii = 27 Then cmdCancel_Click

End Sub


Private Sub Form_Load()
Dim f, lAlias, tIp, tname

Me.Top = frmMain.Top + 300
Me.Left = frmMain.Left
mnuAlias.Visible = False
f = FreeFile
lvAliases.ListItems.Clear
lvAliases.Refresh
On Error GoTo errhandler
Open App.Path & "\" & FileAlias For Input As #f
Do While Not EOF(f)
    Input #f, tIp, tname
    Set lAlias = lvAliases.ListItems.Add
    lAlias.Text = tIp
    lAlias.SubItems(1) = tname
Loop
lvAliases.Refresh
errhandler:
Close #f
On Error Resume Next
PollingTime = CLng(GetSetting(App.title, "settings", "polltime"))
RecvFilePath = GetSetting(App.title, "settings", "file path")
If PollingTime = 0 Or PollingTime = POLL_INTERVAL Then
    optDefault.Value = True
    optOther.Value = False
    txtMinute.Enabled = False
    txtMinute.BackColor = vbButtonFace
Else
    optDefault.Value = False
    optOther.Value = True
    txtMinute.Enabled = True
    txtMinute = Int(PollingTime / 60000)
End If
If RecvFilePath = "" Then txtPath.Text = App.Path Else txtPath.Text = RecvFilePath
End Sub

Private Sub lvAliases_DblClick()
mnuEditAlias_Click
End Sub

Private Sub lvAliases_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If lvAliases.SelectedItem Is Nothing And lvAliases.ListItems.Count = 0 Then
        mnuEditAlias.Enabled = False
        mnuDelAlias.Enabled = False
    End If
    PopupMenu mnuAlias
End If
End Sub

Private Sub mnuAddAlias_Click()
Load frmAlias
frmAlias.cmdOk.Caption = "Add"
frmAlias.Show vbModal, Me
End Sub

Private Sub mnuDelAlias_Click()
If lvAliases.ListItems.Count > 0 Then
    lvAliases.ListItems.Remove lvAliases.SelectedItem.Index
    lvAliases.Refresh
End If
End Sub

Private Sub mnuEditAlias_Click()
Load frmAlias
With frmAlias
    .Tag = lvAliases.SelectedItem.Index
    .txtIp = lvAliases.ListItems(lvAliases.SelectedItem.Index).Text
    .txtName = lvAliases.ListItems(lvAliases.SelectedItem.Index).SubItems(1)
    .cmdOk.Caption = "Save"
    .Show vbModal, Me
End With
End Sub

Public Sub WriteToFile()
Dim f, i

f = FreeFile
Open App.Path & "\" & FileAlias For Output As #f
With lvAliases
    For i = 1 To .ListItems.Count
        Write #f, .ListItems(i).Text, .ListItems(i).SubItems(1)
    Next
End With
Close #f
End Sub

Private Sub optDefault_Click()
txtMinute.Enabled = False
txtMinute.BackColor = vbButtonFace
End Sub

Private Sub optOther_Click()
txtMinute.Enabled = True
txtMinute.BackColor = vbWindowBackground
End Sub
