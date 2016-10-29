VERSION 5.00
Begin VB.Form frmPop 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1470
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   1875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   1875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMove 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   960
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name or informations"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1665
   End
End
Attribute VB_Name = "frmPop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

Const MeHeight = 1500
Const Speed = 50
'intMove = 1 Then form is moving up
'intMove = 2 Then form is movin down
Dim intMove As Integer
Dim ActiveForm As Form

Private Sub Form_Initialize()

intMove = 1

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblInfo.FontUnderline = False
End Sub

Private Sub lblInfo_Click()

    Form_Click

End Sub

Private Sub lblInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblInfo.FontUnderline = True
End Sub

Private Sub lblName_Click()

    Form_Click

End Sub

Private Sub Form_Click()

If Left(Me.Tag, 8) = "[SIGNIN]" Then
    '*** if you'd like to talk to whom has just sign in
    Dim fChat As frmChat
    Set fChat = New frmChat
    Load fChat
    fChat.Show
    With fChat
      .Tag = Mid(Me.Tag, 9)
      .cmdSend.Enabled = False
      .rtbSay.Enabled = False
      .socket.Connect GetIPFromHostName(.Tag), LISTENPORT
      .Status = connecting
    End With
ElseIf Left(Me.Tag, 6) = "[CHAT]" Then
    '*** This will restore the chat window of current caller
    If Not ActiveForm Is Nothing Then ActiveForm.WindowState = vbNormal
End If
End Sub


Private Sub tmrMove_Timer()

    tmrMove.Interval = 5

    If intMove = 1 Then
        '*** scroll up
        Me.Top = (Me.Top - Speed)
        Me.Height = Me.Height + Speed
        If Me.Height >= MeHeight Then
            tmrMove.Interval = 5000
            intMove = 2
        End If
    ElseIf intMove = 2 Then
        '*** scroll down
        Me.Top = (Me.Top + Speed)
        Me.Height = Me.Height - Speed
        If Me.Height <= Speed Then
            tmrMove.Enabled = False
            PopLoad = PopLoad - 1
            If PopLoad <= 0 Then
                PopLoad = 0
                PopLevel = 0
            End If
            Unload Me
        End If
    End If
End Sub

Public Sub OnTop()
SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / 15, _
                        Me.Top / 15, Me.Width / 15, _
                        Me.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Property Set ActiveWindow(ByRef f As Form)
    Set ActiveForm = f
    PopLevel = PopLevel + 1
    If Screen.Height - (MeHeight * PopLevel) < 0 Then PopLevel = 0
    PopLoad = PopLoad + 1

End Property
