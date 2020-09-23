VERSION 5.00
Begin VB.Form frmShow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Output "
   ClientHeight    =   2340
   ClientLeft      =   3615
   ClientTop       =   5850
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "frmShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
Me.Left = frmMain.Left + 50
Me.Top = frmMain.StatusBar1.Top - 2140
Me.Width = frmMain.Width - 405
Text1.Width = frmMain.Text1.Width - 405
End Sub

Private Sub Form_Load()
Me.Icon = frmMain.Icon

Me.Left = frmMain.Left + 50
Me.Top = frmMain.StatusBar1.Top - 2140
Me.Width = frmMain.Width - 405
Text1.Width = frmMain.Text1.Width - 405

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = vbShiftMask And KeyCode = vbKeyEscape Then frmMain.Text1.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Me.Hide
End Sub
