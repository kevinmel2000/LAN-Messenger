VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   2760
   ClientLeft      =   7965
   ClientTop       =   2280
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905.001
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   99.539
      X2              =   5330.058
      Y1              =   1499.843
      Y2              =   1499.843
   End
   Begin VB.Label lblDescription 
      Caption         =   "App Description"
      ForeColor       =   &H00000000&
      Height          =   1050
      Left            =   1050
      TabIndex        =   1
      Top             =   960
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      Index           =   0
      X1              =   99.539
      X2              =   5330.058
      Y1              =   1510.196
      Y2              =   1510.196
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1080
      TabIndex        =   4
      Top             =   600
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Warning: ..."
      ForeColor       =   &H00404040&
      Height          =   465
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   4020
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public frReferer As Form

Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub Form_Resize()
Me.Left = frReferer.Left
Me.Top = frReferer.Top + 600
If Me.Left + Me.Width > Screen.Width Then Me.Left = Screen.Width - Me.Width
If Me.Left < 0 Then Me.Left = 0
Me.Caption = "About " & App.title
lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
lblTitle.Caption = App.title
lblDescription.Caption = App.Comments
lblDisclaimer.Caption = "Developed by M. Novan Adrian " & vbCrLf _
                        & "http://www.novanadrian.com"
End Sub

Private Sub lblDisclaimer_Click()
Shell "c:\program files\internet explorer\iexplore.exe http://www.novanadrian.com"
End Sub
