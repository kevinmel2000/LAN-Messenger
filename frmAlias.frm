VERSION 5.00
Begin VB.Form frmAlias 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4545
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtIp 
      Height          =   330
      Left            =   120
      MaxLength       =   15
      TabIndex        =   0
      ToolTipText     =   "Fill with your friend's full name"
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   330
      Left            =   1800
      TabIndex        =   3
      Top             =   885
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   3120
      TabIndex        =   5
      Top             =   885
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      Height          =   330
      Left            =   1780
      MaxLength       =   20
      TabIndex        =   1
      ToolTipText     =   "Fill with your friend's full name"
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Full Name"
      Height          =   255
      Left            =   1780
      TabIndex        =   4
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "IP Address"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmAlias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
If Len(txtName) = 0 Then
    MsgBox "Please fill with a valid name!", vbExclamation + vbOKOnly, "Aliases"
    Exit Sub
End If
If isValidIp(txtIp) Then
    If cmdOk.Caption = "Add" Then
        Dim lAlias
        With frmSetting
            Set lAlias = .lvAliases.ListItems.Add
            lAlias.Text = txtIp
            lAlias.SubItems(1) = txtName
            .lvAliases.Refresh
            Set lAlias = Nothing
        End With
        Unload Me
    Else
        With frmSetting.lvAliases
            .ListItems(Val(Me.Tag)).Text = txtIp
            .ListItems(Val(Me.Tag)).SubItems(1) = txtName
            .Refresh
        End With
        Unload Me
    End If
Else
    MsgBox "Please fill with a valid IP Address!", vbExclamation + vbOKOnly, "Aliases"
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdOk_Click
If KeyAscii = 27 Then cmdCancel_Click

End Sub

Private Sub Form_Load()
Me.Top = frmSetting.Top + 1450
Me.Left = frmSetting.Left + 300
End Sub

