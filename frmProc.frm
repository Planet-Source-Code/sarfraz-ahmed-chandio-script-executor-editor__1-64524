VERSION 5.00
Begin VB.Form frmProc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procedures"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Execute"
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
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
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
      Left            =   1560
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmProc"
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

Private Sub Command2_Click()

If List1.ListCount = 0 Then
    MsgBox "There is no procedure to execute!!", vbInformation
    Unload Me
    Exit Sub
End If

'if no entry is selected in the listbox
If List1.ListIndex = -1 Then
    MsgBox "Please select a procedure first!!", vbExclamation
    Exit Sub
End If


On Error GoTo Hell
'frmMain.Script1.AddCode List1.Text
frmMain.Script1.Run List1.Text

Exit Sub
Hell:
MsgBox Err.Description, vbCritical
End Sub

Private Sub Command2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = frmMain.Icon
Dim i As Integer

frmMain.Script1.Reset

On Error GoTo CodeError
frmMain.Script1.AddCode frmMain.Text1.Text

For i = 1 To frmMain.Script1.Procedures.Count
If frmMain.Script1.Procedures(i).HasReturnValue Then

'Text2.Text = Text2.Text & vbNewLine & "Function  " &
List1.AddItem frmMain.Script1.Procedures(i).Name
Else
'Text2.Text = Text2.Text & vbNewLine & "Subroutine  " &
List1.AddItem frmMain.Script1.Procedures(i).Name
End If
Next



Dim ShowClass As New Class1

frmMain.Script1.AddObject "OutPut", ShowClass, True


Exit Sub
CodeError:
MsgBox Err.Description, vbCritical
End Sub

Private Sub List1_DblClick()
Command2.Value = True
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command2.Value = True
If KeyAscii = 27 Then Unload Me
End Sub
