VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procedure Info"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1020
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3375
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = frmMain.Icon
Dim i As Integer

On Error Resume Next
frmMain.Script1.AddCode frmMain.Text1

If frmMain.Script1.Procedures.Count = 0 Then
    MsgBox "There is no procedure info!!", vbInformation
    Unload Me
    Exit Sub
End If

For i = 1 To frmMain.Script1.Procedures.Count
    If frmMain.Script1.Procedures(i).HasReturnValue Then
        Text1.SelText = vbCrLf & frmMain.Script1.Procedures(i).Name & "  (f)"
    Else
        Text1.SelText = vbCrLf & frmMain.Script1.Procedures(i).Name & "  (s)"
    End If
Next

Text1.SelText = vbCrLf & vbCrLf & String(25, "_") & vbCrLf & vbCrLf & "[Total Procedures] = " & frmMain.Script1.Procedures.Count


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub
