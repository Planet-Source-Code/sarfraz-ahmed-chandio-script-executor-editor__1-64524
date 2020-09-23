VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00D9E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About ScriptPad v1.00"
   ClientHeight    =   6570
   ClientLeft      =   3135
   ClientTop       =   990
   ClientWidth     =   5865
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4534.731
   ScaleMode       =   0  'User
   ScaleWidth      =   5507.538
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   65
      Left            =   2280
      Top             =   2400
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By :"
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
      Left            =   120
      TabIndex        =   7
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sarfraz Ahmed Chandio   (sarfrazahmed_pk@yahoo.com)"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6240
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ScriptPad v1.00"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5655
   End
   Begin VB.Image Image2 
      Height          =   1080
      Left            =   0
      Picture         =   "frmAbout.frx":0000
      Top             =   0
      Width           =   5880
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":41A5
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2895
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   5655
   End
   Begin VB.Label lblWeb 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.angelfire.com/ultra/sarfraz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   450
      Left            =   360
      MouseIcon       =   "frmAbout.frx":454B
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   5445
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Visit the following website to check for an update of ScriptPad  or to check out other programs."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4320
      Width           =   5655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   5323.484
      Y1              =   3727.176
      Y2              =   3727.176
   End
   Begin VB.Label lblReg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Registered To:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   5535
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AnimatedText As String
Dim a As Integer
Dim b As Integer

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = frmMain.Icon

AnimatedText = lblWeb.Caption
a = Len(AnimatedText)
b = 1

'Show UserName
Dim sBuffer As String
Dim lSize As Long
sBuffer = Space$(255)
lSize = Len(sBuffer)
On Error Resume Next
Call GetUserName(sBuffer, lSize)
If lSize > 0 Then
lblReg = "Registered to:  " & Left$(sBuffer, lSize)
Else
lblReg = "Registered to:  Unknown User"
End If


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.ForeColor = vbBlue
End Sub

Private Sub Label3_Click()
On Error Resume Next
Shell "start mailto:sarfrazahmed_pk@yahoo.com", vbHide
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.ForeColor = vbRed
End Sub

Private Sub lblWeb_Click()
On Error Resume Next
Shell "start http://www.angelfire.com/ultra/sarfraz", vbHide
End Sub

Private Sub Timer1_Timer()
On Error GoTo Handler
lblWeb.Visible = True
On Error Resume Next
lblWeb = Left(AnimatedText, b)
b = b + 1

Exit Sub
Handler:
Timer1.Enabled = False
lblWeb.Visible = True
End Sub

Private Sub btnExit_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

