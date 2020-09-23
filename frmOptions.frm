VERSION 5.00
Begin VB.Form frmOptions 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5295
   ClientLeft      =   3165
   ClientTop       =   1425
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Defaults"
      Height          =   375
      Left            =   1800
      TabIndex        =   18
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CheckBox chkMain 
      Caption         =   "Insert &Main procedure"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      Caption         =   "StartUp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   17
      Top             =   2760
      Width           =   4455
      Begin VB.CheckBox chkWindow 
         Caption         =   "Out&put window visible"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "O&k"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Format"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdFC 
         Caption         =   "Selec&t"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   2040
         Width           =   975
      End
      Begin VB.CheckBox chkBC 
         Caption         =   "Page C&olor"
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CheckBox chkFC 
         Caption         =   "T&ext Color"
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdBC 
         Caption         =   "Se&lect"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         TabIndex        =   6
         Top             =   1680
         Width           =   975
      End
      Begin VB.ComboBox cboFontSize 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cboFontName 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   2775
      End
      Begin VB.CheckBox chkUnderline 
         Caption         =   "&Underline"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CheckBox chkItalic 
         Caption         =   "&Italic"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   975
      End
      Begin VB.CheckBox chkBold 
         Caption         =   "&Bold"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Font Si&ze"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fo&nt Name"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   420
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboFontName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
KeyAscii = 0
End Sub

Private Sub cboFontSize_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
KeyAscii = 0
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub chkBC_Click()
If chkBC.Value = 1 Then
cmdBC.Enabled = True
Else
cmdBC.Enabled = False
End If
End Sub

Private Sub chkBold_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub chkFC_Click()
If chkFC.Value = 1 Then
cmdFC.Enabled = True
Else
cmdFC.Enabled = False
End If
End Sub

Private Sub chkItalic_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub chkMain_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub chkUnderline_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub chkWindow_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdApply_Click()

frmMain.Text1.FontName = cboFontName.Text
frmMain.Text1.FontSize = cboFontSize.Text

If chkBold.Value = 1 Then frmMain.Text1.FontBold = True
If chkItalic.Value = 1 Then frmMain.Text1.FontItalic = True
If chkUnderline.Value = 1 Then frmMain.Text1.FontUnderline = True

If chkFC.Value = 0 Then frmMain.Text1.ForeColor = vbBlack
If chkBC.Value = 0 Then frmMain.Text1.BackColor = vbWhite
If chkBold.Value = 0 Then frmMain.Text1.FontBold = False
If chkItalic.Value = 0 Then frmMain.Text1.FontItalic = False
If chkUnderline.Value = 0 Then frmMain.Text1.FontUnderline = False

On Error GoTo Handler
SaveSetting App.EXEName, "FontNameSize", "FontName", cboFontName.Text
SaveSetting App.EXEName, "FontNameSize", "FontSize", cboFontSize.Text
SaveSetting App.EXEName, "FontStyle", "FontBold", chkBold.Value
SaveSetting App.EXEName, "FontStyle", "FontItalic", chkItalic.Value
SaveSetting App.EXEName, "FontStyle", "FontUnderline", chkUnderline.Value
SaveSetting App.EXEName, "Color", "ForeColor", chkFC.Value
SaveSetting App.EXEName, "Color", "BackColor", chkBC.Value
SaveSetting App.EXEName, "Color", "TextForeColor", frmMain.Text1.ForeColor
SaveSetting App.EXEName, "Color", "TextBackColor", frmMain.Text1.BackColor
SaveSetting App.EXEName, "Output Window", "Visible", chkWindow.Value
SaveSetting App.EXEName, "MainInsert", "Insert", chkMain.Value

cmdOk.SetFocus

Exit Sub
Handler:
frmMain.Text1.FontName = "Verdana"
frmMain.Text1.FontSize = 10
frmMain.Text1.FontBold = False
frmMain.Text1.FontItalic = False
frmMain.Text1.FontUnderline = False
frmMain.Text1.ForeColor = vbBlack
frmMain.Text1.BackColor = vbWhite
End Sub

Private Sub cmdApply_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdBC_Click()
On Error GoTo Handler
frmMain.cd.CancelError = True
frmMain.cd.ShowColor
frmMain.Text1.BackColor = frmMain.cd.Color

Exit Sub
Handler:
Dim a
a = GetSetting(App.EXEName, "Color", "TextBackColor", "")
If a = "" Then
frmMain.Text1.BackColor = vbWhite
Else
frmMain.Text1.BackColor = GetSetting(App.EXEName, "Color", "TextBackColor", "")
End If

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCancel_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cmdFC_Click()
On Error GoTo Handler
frmMain.cd.CancelError = True
frmMain.cd.ShowColor
frmMain.Text1.ForeColor = frmMain.cd.Color

Exit Sub
Handler:
Dim b
b = GetSetting(App.EXEName, "Color", "TextForeColor", "")
If b = "" Then
frmMain.Text1.ForeColor = vbBlack
Else
frmMain.Text1.ForeColor = GetSetting(App.EXEName, "Color", "TextForeColor", "")
End If
End Sub

Private Sub cmdOK_Click()
cmdApply_Click
Unload Me
End Sub

Private Sub cmdOK_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Command1_Click()
Dim Res As Integer
    Res = MsgBox("Are you sure you want to restore the defaults?", vbYesNo + vbQuestion)

    If Res = vbNo Then
        Exit Sub
    Else
        On Error Resume Next
        DeleteSetting App.EXEName, "FontNameSize"
        DeleteSetting App.EXEName, "FontStyle"
        DeleteSetting App.EXEName, "Color"
        DeleteSetting App.EXEName, "Output Window"
        DeleteSetting App.EXEName, "MainInsert"
        
        chkWindow.Value = 1
        SaveSetting App.EXEName, "Output Window", "Visible", chkWindow.Value
        
        MsgBox "The default settings have successfully been restored and will be applied the next time you run the ScriptPad.", vbExclamation
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Me.Icon = frmMain.Icon

Dim x%
'Get FontName
cboFontName = frmMain.Text1.FontName

For x = 1 To Screen.FontCount
cboFontName.AddItem Screen.Fonts(x)
Next
cboFontName.RemoveItem (0)

For x = 0 To cboFontName.ListCount - 1
Exit For
Next

'Get FontSize
cboFontSize = frmMain.Text1.FontSize

For x = 8 To 72
cboFontSize.AddItem Str$(x)
Next

For x = 0 To cboFontSize.ListCount - 1
Exit For
Next


Dim Values(11)

Values(0) = GetSetting(App.EXEName, "MainInsert", "Insert")
If Values(0) = 0 Or Values(0) = "" Then
chkMain.Value = 0
Else
chkMain.Value = GetSetting(App.EXEName, "MainInsert", "Insert")
End If

Values(1) = GetSetting(App.EXEName, "FontNameSize", "FontName", "")
If Values(1) = "" Then
cboFontName.Text = "Verdana"
Else
cboFontName = GetSetting(App.EXEName, "FontNameSize", "FontName", "")
End If

Values(2) = GetSetting(App.EXEName, "FontNameSize", "FontSize", "")
If Values(2) = "" Then
cboFontSize = 10
Else
cboFontSize = GetSetting(App.EXEName, "FontNameSize", "FontSize", "")
End If

Values(5) = GetSetting(App.EXEName, "FontStyle", "FontBold", "")
If Values(5) = "" Then
chkBold.Value = 0
Else
If Values(5) = "True" Then chkBold.Value = 1
chkBold.Value = GetSetting(App.EXEName, "FontStyle", "FontBold", "")
End If


Values(6) = GetSetting(App.EXEName, "FontStyle", "FontItalic", "")
If Values(6) = "" Then
chkItalic.Value = 0
Else
If Values(6) = "True" Then chkItalic.Value = 1
chkItalic.Value = GetSetting(App.EXEName, "FontStyle", "FontItalic", "")
End If


Values(7) = GetSetting(App.EXEName, "FontStyle", "FontUnderline", "")
If Values(7) = "" Then
chkUnderline.Value = 0
Else
If Values(7) = "True" Then chkUnderline.Value = 1
chkUnderline.Value = GetSetting(App.EXEName, "FontStyle", "FontUnderline", "")
End If

Values(8) = GetSetting(App.EXEName, "Output Window", "Visible")
If Values(8) = "" Then
chkWindow.Value = 0
Else
chkWindow.Value = GetSetting(App.EXEName, "Output Window", "Visible")
End If

Values(10) = GetSetting(App.EXEName, "Color", "ForeColor", "")
If Values(10) = 0 Then
chkFC.Value = 0
Else
On Error Resume Next
chkFC.Value = GetSetting(App.EXEName, "Color", "ForeColor", "")
End If

On Error Resume Next
Values(11) = GetSetting(App.EXEName, "Color", "BackColor", "")
If Values(11) = 0 Then
chkBC.Value = 0
Else
chkBC.Value = GetSetting(App.EXEName, "Color", "BackColor", "")
End If

End Sub

