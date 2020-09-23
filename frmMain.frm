VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Untitled - ScriptPad v1.00"
   ClientHeight    =   6510
   ClientLeft      =   1590
   ClientTop       =   1230
   ClientWidth     =   9165
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   6510
   ScaleWidth      =   9165
   WindowState     =   2  'Maximized
   Begin MSScriptControlCtl.ScriptControl Script1 
      Left            =   240
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   6210
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9551
            Text            =   "ScriptPad v1.00"
            TextSave        =   "ScriptPad v1.00"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   882
            TextSave        =   "12/3/05"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   6600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6120
      Top             =   5520
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   5640
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   11895
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu vbvb 
         Caption         =   "-"
      End
      Begin VB.Menu mSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mSaveSelectionAs 
         Caption         =   "Save S&election As..."
      End
      Begin VB.Menu dfgdfddfdf 
         Caption         =   "-"
      End
      Begin VB.Menu mRevertOrig 
         Caption         =   "&Revert To Original"
         Shortcut        =   {F8}
      End
      Begin VB.Menu fgfgf 
         Caption         =   "-"
      End
      Begin VB.Menu mPageSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mPrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu gfgfhfgh 
         Caption         =   "-"
      End
      Begin VB.Menu mProperties 
         Caption         =   "P&roperties..."
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu mRecentFiles 
         Caption         =   "&Recent Files"
         Begin VB.Menu Sep 
            Caption         =   "---Empty---"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuHis 
            Caption         =   "His 1"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu mnuHis 
            Caption         =   "His 2"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuHis 
            Caption         =   "His 3"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuHis 
            Caption         =   "His 4"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuHis 
            Caption         =   "His 5"
            Index           =   4
            Visible         =   0   'False
         End
      End
      Begin VB.Menu hjkhj 
         Caption         =   "-"
      End
      Begin VB.Menu mSaveExit 
         Caption         =   "Sa&ve && Exit"
         Shortcut        =   ^W
      End
      Begin VB.Menu mExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mRedo 
         Caption         =   "R&edo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu fgfgfgfg 
         Caption         =   "-"
      End
      Begin VB.Menu mSeparator 
         Caption         =   "Se&parator"
         Shortcut        =   {F2}
      End
      Begin VB.Menu jkjjk 
         Caption         =   "-"
      End
      Begin VB.Menu mCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mCopyLine 
         Caption         =   "Copy &Line"
         Shortcut        =   ^L
      End
      Begin VB.Menu mPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu ghjghjghhg 
         Caption         =   "-"
      End
      Begin VB.Menu mClear 
         Caption         =   "Clea&r"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mClearAll 
         Caption         =   "Clear &All"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mSelectAll 
         Caption         =   "&Select&All"
         Shortcut        =   ^A
      End
      Begin VB.Menu ppp 
         Caption         =   "-"
      End
      Begin VB.Menu mTimeDate 
         Caption         =   "Ti&me/Date"
         Begin VB.Menu mDate 
            Caption         =   "&Date"
         End
         Begin VB.Menu mTime 
            Caption         =   "&Time"
         End
      End
      Begin VB.Menu ggg 
         Caption         =   "-"
      End
      Begin VB.Menu mTypingSound 
         Caption         =   "T&yping Sound"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu mSearch 
      Caption         =   "&Search"
      Begin VB.Menu mFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mFindNext 
         Caption         =   "Find &Next"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mOccurrences 
         Caption         =   "Find &Occurrences"
      End
      Begin VB.Menu gdfgdfgdfgdfgdf 
         Caption         =   "-"
      End
      Begin VB.Menu mQuickReplace 
         Caption         =   "&Quick Replace"
         Shortcut        =   {F6}
      End
      Begin VB.Menu yy 
         Caption         =   "-"
      End
      Begin VB.Menu mGoto 
         Caption         =   "&Goto Line..."
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mScript 
      Caption         =   "S&cript"
      Begin VB.Menu mExecute 
         Caption         =   "E&xecute Script"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mListPro 
         Caption         =   "Proce&dures..."
         Shortcut        =   {F4}
      End
      Begin VB.Menu ghfdgdfgdf 
         Caption         =   "-"
      End
      Begin VB.Menu mFunctions 
         Caption         =   "&VBScript Functions"
         Begin VB.Menu mMsgBox 
            Caption         =   "&MsgBox"
         End
         Begin VB.Menu mInputBox 
            Caption         =   "&InputBox"
         End
         Begin VB.Menu mLen 
            Caption         =   "&Len"
         End
         Begin VB.Menu hhh 
            Caption         =   "-"
         End
         Begin VB.Menu mStringFunctions 
            Caption         =   "&String Functions"
            Begin VB.Menu mAsc 
               Caption         =   "Asc"
            End
            Begin VB.Menu mChr 
               Caption         =   "Chr"
            End
            Begin VB.Menu mInStr 
               Caption         =   "InStr"
            End
            Begin VB.Menu mLCase 
               Caption         =   "LCase"
            End
            Begin VB.Menu mLeft 
               Caption         =   "Left"
            End
            Begin VB.Menu mLTrim 
               Caption         =   "LTrim"
            End
            Begin VB.Menu mMid 
               Caption         =   "Mid"
            End
            Begin VB.Menu mRight 
               Caption         =   "Right"
            End
            Begin VB.Menu mRTrim 
               Caption         =   "RTrim"
            End
            Begin VB.Menu mStr 
               Caption         =   "Str"
            End
            Begin VB.Menu mStrComp 
               Caption         =   "StrComp"
            End
            Begin VB.Menu mString 
               Caption         =   "String"
            End
            Begin VB.Menu mTrim 
               Caption         =   "Trim"
            End
            Begin VB.Menu mUCase 
               Caption         =   "UCase"
            End
            Begin VB.Menu mVal 
               Caption         =   "Val"
            End
         End
         Begin VB.Menu mConvFunctions 
            Caption         =   "&Conversion Functions"
            Begin VB.Menu mCBoll 
               Caption         =   "CBool"
            End
            Begin VB.Menu mCByte 
               Caption         =   "CByte"
            End
            Begin VB.Menu mCDbl 
               Caption         =   "CDbl"
            End
            Begin VB.Menu mCint 
               Caption         =   "CInt"
            End
            Begin VB.Menu mCLng 
               Caption         =   "CLng"
            End
            Begin VB.Menu mCSng 
               Caption         =   "CSng"
            End
            Begin VB.Menu mCStr 
               Caption         =   "CStr"
            End
            Begin VB.Menu mCVErr 
               Caption         =   "CVErr"
            End
         End
         Begin VB.Menu mMathFunctions 
            Caption         =   "Mat&h Functions"
            Begin VB.Menu mAbs 
               Caption         =   "Abs"
            End
            Begin VB.Menu mArray 
               Caption         =   "Array"
            End
            Begin VB.Menu mAtn 
               Caption         =   "Atn"
            End
            Begin VB.Menu mExp 
               Caption         =   "Exp"
            End
            Begin VB.Menu mFix 
               Caption         =   "Fix"
            End
            Begin VB.Menu mHex 
               Caption         =   "Hex"
            End
            Begin VB.Menu mInt 
               Caption         =   "Int"
            End
            Begin VB.Menu mLog 
               Caption         =   "Log"
            End
            Begin VB.Menu mOct 
               Caption         =   "Oct"
            End
            Begin VB.Menu mRnd 
               Caption         =   "Rnd"
            End
            Begin VB.Menu mSgn 
               Caption         =   "Sgn"
            End
            Begin VB.Menu mSqr 
               Caption         =   "Sqr"
            End
            Begin VB.Menu mSin 
               Caption         =   "Sin"
            End
            Begin VB.Menu mTan 
               Caption         =   "Tan"
            End
         End
         Begin VB.Menu mTDFunctions 
            Caption         =   "&Time/Date Functions"
            Begin VB.Menu mDateF 
               Caption         =   "Date"
            End
            Begin VB.Menu mDateSerial 
               Caption         =   "DateSerial"
            End
            Begin VB.Menu mDateValue 
               Caption         =   "DateValue"
            End
            Begin VB.Menu mDay 
               Caption         =   "Day"
            End
            Begin VB.Menu mHour 
               Caption         =   "Hour"
            End
            Begin VB.Menu mYear 
               Caption         =   "Year"
            End
            Begin VB.Menu mWeekDay 
               Caption         =   "WeekDay"
            End
            Begin VB.Menu mMinute 
               Caption         =   "Minute"
            End
            Begin VB.Menu mMonth 
               Caption         =   "Month"
            End
            Begin VB.Menu mNow 
               Caption         =   "Now"
            End
            Begin VB.Menu mSecond 
               Caption         =   "Second"
            End
            Begin VB.Menu mTimeF 
               Caption         =   "Time"
            End
            Begin VB.Menu mTimeSerial 
               Caption         =   "TimeSerial"
            End
            Begin VB.Menu mTimeValue 
               Caption         =   "TimeValue"
            End
         End
         Begin VB.Menu mOtherFunctions 
            Caption         =   "&Other Functions"
            Begin VB.Menu mIsArray 
               Caption         =   "IsArray"
            End
            Begin VB.Menu mIsDate 
               Caption         =   "IsDate"
            End
            Begin VB.Menu mIsEmpty 
               Caption         =   "IsEmpty"
            End
            Begin VB.Menu mIsError 
               Caption         =   "IsError"
            End
            Begin VB.Menu mIsNull 
               Caption         =   "IsNull"
            End
            Begin VB.Menu mIsNumeric 
               Caption         =   "IsNumeric"
            End
            Begin VB.Menu mIsObject 
               Caption         =   "IsObject"
            End
            Begin VB.Menu hfg 
               Caption         =   "-"
            End
            Begin VB.Menu mVarType 
               Caption         =   "VarType"
            End
         End
      End
   End
   Begin VB.Menu mView 
      Caption         =   "&View"
      Begin VB.Menu mIncFont 
         Caption         =   "&Increase Font"
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu mDecFont 
         Caption         =   "&Decrease Font"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu jkjkjkjkj 
         Caption         =   "-"
      End
      Begin VB.Menu mFullScreen 
         Caption         =   "&Full Screen"
         Shortcut        =   {F11}
      End
      Begin VB.Menu hh 
         Caption         =   "-"
      End
      Begin VB.Menu mOutWindow 
         Caption         =   "&Output Window"
      End
      Begin VB.Menu ghfhgfhfg 
         Caption         =   "-"
      End
      Begin VB.Menu mProcInfo 
         Caption         =   "Procedure I&nfo..."
      End
   End
   Begin VB.Menu mFormat 
      Caption         =   "F&ormat"
      Begin VB.Menu mFont 
         Caption         =   "&Fonts..."
         Shortcut        =   ^T
      End
      Begin VB.Menu nbnbnbd 
         Caption         =   "-"
      End
      Begin VB.Menu mPageColor 
         Caption         =   "&Page Color..."
      End
      Begin VB.Menu mTextColor 
         Caption         =   "Te&xt Color..."
      End
      Begin VB.Menu hg 
         Caption         =   "-"
      End
      Begin VB.Menu mSelectCase 
         Caption         =   "Select Ca&se"
         Begin VB.Menu mUpperCase 
            Caption         =   "&UPPER CASE"
         End
         Begin VB.Menu mLowerCase 
            Caption         =   "&lower case"
         End
         Begin VB.Menu mProperCapitalize 
            Caption         =   "&Proper Case"
         End
         Begin VB.Menu mLowerCaps 
            Caption         =   "Lower &caps"
         End
      End
   End
   Begin VB.Menu mTools 
      Caption         =   "&Tools"
      Begin VB.Menu mSpellCheck 
         Caption         =   "&Spell Checker..."
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mTexttoHtml 
         Caption         =   "Text to &HTML..."
         Shortcut        =   ^H
      End
      Begin VB.Menu mWordCount 
         Caption         =   "&Word Count..."
         Shortcut        =   ^R
      End
      Begin VB.Menu mCalculator 
         Caption         =   "&Calculator"
      End
      Begin VB.Menu ssssssssss 
         Caption         =   "-"
      End
      Begin VB.Menu mOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mmHelp 
      Caption         =   "&Help"
      Begin VB.Menu mHelp 
         Caption         =   "&Online Help"
      End
      Begin VB.Menu mContact 
         Caption         =   "&Contact"
      End
      Begin VB.Menu mAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'                  ******************
'                   ScriptPad v1.00
'                  ******************
'
'                     Developed by:
'                "SARFRAZ AHMED CHANDIO"


'                     CONTACT ME
'Email address:- sarfrazahmed_pk@yahoo.com
'Website:- http://www.angelfire.com/ultra/sarfraz

'Distribution:
'You are free to use this source code in your projects
'as long as you specify the above email and website
'names.

'Dated:- 28/7/04

Option Explicit
Dim Changed As Boolean
Dim CancelClicked As Boolean
Dim History(4) As String
Dim UseSound As String
Dim ShowClass As New Class1
Private TrapUndo As Boolean           'flag to indicate whether actions should be trapped
Private UndoStack As New Collection   'collection of undo elements
Private RedoStack As New Collection   'collection of redo elements

Private Sub Form_Activate()
Dim Val

Val = GetSetting(App.EXEName, "MainInsert", "Insert")
If Val = 1 Then
    Changed = False
End If

End Sub

Private Sub Form_GotFocus()
Text1.SetFocus
End Sub

Private Sub Form_Load()
Dim Val

Val = GetSetting(App.EXEName, "MainInsert", "Insert")
If Val = 1 Then
    Text1 = "Sub Main ()" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "End Sub"
    Changed = False
End If

Cap = "Untitled"
Script1.Timeout = -1

On Error Resume Next
Call SetHook(Text1.hWnd, True)

TrapUndo = True     'Enable Undo Trapping
Text1_Change      'Initialize First Undo

Dim Values(9)

Values(0) = GetSetting(App.EXEName, "Output Window", "Visible")
If Values(0) <> 0 And Values(0) <> "" Then
Load frmShow
frmShow.Show , Me
Else
GoTo ResumeHere
End If

ResumeHere:
Values(1) = GetSetting(App.EXEName, "FontNameSize", "FontName", "")
If Values(1) = "" Then
Text1.FontName = "Verdana"
Else
Text1.FontName = GetSetting(App.EXEName, "FontNameSize", "FontName", "")
End If

Values(2) = GetSetting(App.EXEName, "FontNameSize", "FontSize", "")
If Values(2) = "" Then
Text1.FontSize = 10
Else
Text1.FontSize = GetSetting(App.EXEName, "FontNameSize", "FontSize", "")
End If

Values(3) = GetSetting(App.EXEName, "FontStyle", "FontBold", "")
If Values(3) = "" Then
Text1.FontBold = False
Else
Text1.FontBold = GetSetting(App.EXEName, "FontStyle", "FontBold", "")
End If

Values(4) = GetSetting(App.EXEName, "FontStyle", "FontItalic", "")
If Values(4) = "" Then
Text1.FontItalic = False
Else
Text1.FontItalic = GetSetting(App.EXEName, "FontStyle", "FontItalic", "")
End If

Values(5) = GetSetting(App.EXEName, "FontStyle", "FontUnderline", "")
If Values(5) = "" Then
Text1.FontUnderline = False
Else
Text1.FontUnderline = GetSetting(App.EXEName, "FontStyle", "FontUnderline", "")
End If

Values(6) = GetSetting(App.EXEName, "FontStyle", "FontStrikethru", "")
If Values(6) = "" Then
Text1.FontStrikethru = False
Else
Text1.FontStrikethru = GetSetting(App.EXEName, "FontStyle", "FontStrikethru", "")
End If

Values(8) = GetSetting(App.EXEName, "Color", "TextForeColor", "")
If Values(8) = "" Then
Text1.ForeColor = vbBlack
Else
Text1.ForeColor = GetSetting(App.EXEName, "Color", "TextForeColor", "")
End If

Values(9) = GetSetting(App.EXEName, "Color", "TextBackColor", "")
If Values(9) = "" Then
Text1.BackColor = vbWhite
Else
Text1.BackColor = GetSetting(App.EXEName, "Color", "TextBackColor", "")
End If


On Error Resume Next
cd.FileName = GetSetting(App.EXEName, "file", "file_pattern", "*.vbs")
History(0) = GetSetting(App.EXEName, "save", "his1", "")
History(1) = GetSetting(App.EXEName, "save", "his2", "")
History(2) = GetSetting(App.EXEName, "save", "his3", "")
History(3) = GetSetting(App.EXEName, "save", "his4", "")
History(4) = GetSetting(App.EXEName, "save", "his5", "")

Dim i As Integer
For i = 0 To 4
If History(i) <> "" Then
    mnuHis(i).Caption = History(i)
    On Error Resume Next
    mnuHis(i).Visible = True
    Sep.Visible = False
Else
    mnuHis(i).Visible = False
End If
Next i



'Here we allow the user to script the ScriptPad by
'allowing them to show the result of user-supplied
'procedures in another form.For this purpose a Class
'is required which in this case is Class1 with two
'members Show and Clear.The former shows the result
'of procedures and loads it into another output form
'and shows the result while the later clears everything
'from the output form.



Script1.AddObject "OutPut", ShowClass, True

'The result shower form can be called by the name
'Output.Show  or  Output.Clear  but because the last
'arg is True,so we can ignore the Output name and simpy
'call the class's member name ie,
'Show(Result)   or    Clear


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Changed = True Then
Dim Chandio As Integer
    Chandio = MsgBox("Do you want to save the changes to: " & vbCr & Cap, vbYesNoCancel + vbQuestion, "Save File?")

    If Chandio = vbCancel Then
        Cancel = True
    ElseIf Chandio = vbNo Then

        UnloadForms
        Exit Sub
    Else
    mSave_Click

' If cancel is clicked at the FileSaveAs box then don't
' unload.
    If CancelClicked = True Then
        Cancel = True
        CancelClicked = False
    Else
        Cancel = False
    End If
End If
End If

End Sub

Private Sub Form_Resize()
 On Error Resume Next
 Text1.Height = Me.Height - 965
 Text1.Width = Me.Width - 140
 Text1.Top = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SaveSetting App.EXEName, "save", "his1", History(0)
SaveSetting App.EXEName, "save", "his2", History(1)
SaveSetting App.EXEName, "save", "his3", History(2)
SaveSetting App.EXEName, "save", "his4", History(3)
SaveSetting App.EXEName, "save", "his5", History(4)
SaveSetting App.EXEName, "file", "file_pattern", cd.FileName
cd.FileName = Cap

UnloadForms

End Sub

Private Sub mAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mAbs_Click()
Text1.SelText = "Abs(number)"
End Sub

Private Sub mArray_Click()
Text1.SelText = "Array(arglist)"
End Sub

Private Sub mAsc_Click()
Text1.SelText = "Asc(string)"
End Sub

Private Sub mAtn_Click()
Text1.SelText = "Atn(number)"
End Sub

Private Sub mCalculator_Click()
On Error Resume Next
Shell "calc.exe", vbNormalFocus
End Sub

Private Sub mCBoll_Click()
Text1.SelText = "CBool(expression)"
End Sub

Private Sub mCByte_Click()
Text1.SelText = "CByte(expression)"
End Sub

Private Sub mCDbl_Click()
Text1.SelText = "CDbl(expression)"
End Sub

Private Sub mChr_Click()
Text1.SelText = "Chr(charcode)"
End Sub

Private Sub mCint_Click()
Text1.SelText = "CInt(expression)"
End Sub

Private Sub mClear_Click()
On Error GoTo Hell
Text1.SelText = ""

Exit Sub
Hell:
End Sub

Private Sub mClearAll_Click()
On Error GoTo Hell
Text1 = ""
Changed = True

Exit Sub
Hell:
HellError
End Sub

Private Sub mClng_Click()
Text1.SelText = "CLng(expression)"
End Sub

Private Sub mContact_Click()
On Error Resume Next
Shell "start mailto:sarfrazahmed_pk@yahoo.com", vbHide
End Sub

Private Sub mCopy_Click()
On Error GoTo Hell
Clipboard.SetText Text1.SelText

Exit Sub
Hell:
End Sub

Private Sub mCopyLine_Click()
Dim LineNumber As Integer
Dim GetLineText As String

On Error Resume Next
'get line number
LineNumber = SendMessage(Text1.hWnd, _
EM_LINEFROMCHAR, -1, ByVal 0) + 1

'get current line text
 GetLineText = GetLine(Text1, LineNumber - 1)
 
If GetLineText = "" Then
 Exit Sub
Else
 Clipboard.SetText (GetLineText)
End If

End Sub

Private Sub mCSng_Click()
Text1.SelText = "CSng(expression)"
End Sub

Private Sub mCStr_Click()
Text1.SelText = "CStr(expression)"
End Sub

Private Sub mCut_Click()
On Error GoTo Hell
Clipboard.SetText Text1.SelText
Text1.SelText = ""

Exit Sub
Hell:
End Sub

Private Sub mCVErr_Click()
Text1.SelText = "CVErr(error number)"
End Sub

Private Sub mDate_Click()
Text1.SelText = Date
End Sub

Private Sub mDateF_Click()
Text1.SelText = "Date"
End Sub

Private Sub mDateSerial_Click()
Text1.SelText = "DateSerial(year, month, day)"
End Sub

Private Sub mDateValue_Click()
Text1.SelText = "DateValue(date)"
End Sub

Private Sub mDay_Click()
Text1.SelText = "Day(date)"
End Sub

Private Sub mDecFont_Click()
On Error GoTo Hell
Text1.FontSize = Text1.FontSize - 1

Exit Sub
Hell:
If Err.Number = 380 Then
MsgBox "Can't reduce the Font anymore!", vbInformation
Else
MsgBox Err.Description, vbInformation
End If
End Sub

Private Sub mEdit_Click()
If TextSelected Then
    mCut.Enabled = True
    mCopy.Enabled = True
    mClear.Enabled = True
Else
    mCut.Enabled = False
    mCopy.Enabled = False
    mClear.Enabled = False
End If


If UndoStack.Count = 1 Then
    mUndo.Enabled = False
Else
    mUndo.Enabled = True
End If

If RedoStack.Count = 0 Then
    mRedo.Enabled = False
Else
    mRedo.Enabled = True
End If

Dim Pak As String
On Error Resume Next
Pak = Clipboard.GetText
If Pak = "" Then
    mPaste.Enabled = False
Else
    mPaste.Enabled = True
End If


If Text1 = "" Then
    mClearAll.Enabled = False
    mCopyLine.Enabled = False
Else
    mClearAll.Enabled = True
    mCopyLine.Enabled = True
End If

End Sub

Private Sub mExecute_Click()
On Error GoTo CodeError

Script1.AddCode Text1.Text
Script1.Run "Main"
Text1.SetFocus
Exit Sub

CodeError:
Dim Msg
If Script1.Error.Number <> 0 Then
Msg = Script1.Error.Description & vbCrLf
Msg = Msg & "In Line: " & Script1.Error.Line _
& ", Column" & Script1.Error.Column
MsgBox Msg, vbCritical, "Error in Script"
Else
MsgBox "ERROR #" & Err.Number & vbCrLf & Err.Description, vbCritical
End If
End Sub

Private Sub mExit_Click()
Unload Me
End Sub


Private Sub mExp_Click()
Text1.SelText = "Exp(number)"
End Sub

Private Sub mFile_Click()
If Changed = True Then
    mSave.Enabled = True
    mSaveExit.Enabled = True
Else
    mSave.Enabled = False
    mSaveExit.Enabled = False
End If

If Cap <> "Untitled" And FileExists(cd.FileName) Then
    mProperties.Enabled = True
Else
    mProperties.Enabled = False
End If

If TextSelected Then
    mSaveSelectionAs.Enabled = True
Else
    mSaveSelectionAs.Enabled = False
End If

If Text1.Text <> Text1.Tag And Cap <> "Untitled" And Text1.Tag <> "" Then
    mRevertOrig.Enabled = True
Else
    mRevertOrig.Enabled = False
End If

Dim i%
For i = 0 To 4
If History(i) <> "" Then
mnuHis(i).Caption = History(i)
On Error Resume Next
mnuHis(i).Visible = True
'mnuNoRecentFiles.Visible = False
Else
On Error Resume Next
mnuHis(i).Visible = False
End If
Next i

End Sub

Private Sub mFind_Click()
Load frmFind
frmFind.Show 0, Me
frmFind.txtFind.SetFocus

If Text1.Text = frmFind.txtFind.Text Or frmFind.cmdFindAgain.Enabled = True Then
frmFind.txtReplace.Enabled = True
frmFind.Label2.Enabled = True
Else
frmFind.txtReplace.Enabled = False
frmFind.Label2.Enabled = False
End If

If frmFind.txtFind = "" Then
mFindNext.Enabled = False
End If
End Sub

Private Sub mFindNext_Click()
frmFind.cmdFindAgain.Value = True
End Sub

Private Sub mFix_Click()
Text1.SelText = "Fix(number)"
End Sub

Private Sub mFont_Click()
On Error GoTo Hell
cd.CancelError = True
cd.Flags = cdlCFBoth + cdlCFEffects
cd.FontName = Text1.FontName
cd.FontSize = Text1.FontSize
cd.FontBold = Text1.FontBold
cd.FontItalic = Text1.FontItalic
cd.Color = Text1.ForeColor
cd.FontStrikethru = Text1.FontStrikethru
cd.FontUnderline = Text1.FontUnderline

cd.ShowFont

Text1.FontName = cd.FontName
Text1.FontBold = cd.FontBold
Text1.FontItalic = cd.FontItalic
Text1.FontSize = cd.FontSize
Text1.ForeColor = cd.Color
Text1.FontStrikethru = cd.FontStrikethru
Text1.FontUnderline = cd.FontUnderline

Exit Sub
Hell:

End Sub

Private Sub mFullScreen_Click()
On Error GoTo FullScreenError
frmScreen.txtScreen = Text1
frmScreen.Show 1

Exit Sub
FullScreenError:
If Err.Number = 7 Or Err.Description = "Out of memory" Then
MsgBox "File is too large!!", vbExclamation
Else
MsgBox Err.Description, vbCritical
End If
End Sub

Private Sub mGoto_Click()
frmGoTo.Show 0, Me
End Sub

Private Sub mHelp_Click()
On Error Resume Next
Shell "start http://www.angelfire.com/ultra/sarfraz", vbHide
End Sub

Private Sub mHex_Click()
Text1.SelText = "Hex(number)"
End Sub

Private Sub mHour_Click()
Text1.SelText = "Hour(time)"
End Sub

Private Sub mIncFont_Click()
On Error GoTo Hell
Text1.FontSize = Text1.FontSize + 1

Exit Sub
Hell:
If Err.Number = 380 Then
MsgBox "Can't reduce the Font anymore!", vbInformation
Else
MsgBox Err.Description, vbInformation
End If
End Sub

Private Sub mInputBox_Click()
Text1.SelText = "InputBox(prompt[, title] [, default] [, xpos] [, ypos] [, helpfile, context])"
End Sub

Private Sub mInStr_Click()
Text1.SelText = "InStr([start, ]string1, string2[, compare])"
End Sub

Private Sub mInt_Click()
Text1.SelText = "Int(number)"
End Sub

Private Sub mIsArray_Click()
Text1.SelText = "IsArray(varname)"
End Sub

Private Sub mIsDate_Click()
Text1.SelText = "IsDate(expression)"
End Sub

Private Sub mIsEmpty_Click()
Text1.SelText = "IsEmpty(expression)"
End Sub

Private Sub mIsError_Click()
Text1.SelText = "IsError(expression)"
End Sub

Private Sub mIsNull_Click()
Text1.SelText = "IsNull(expression)"
End Sub

Private Sub mIsNumeric_Click()
Text1.SelText = "IsNumeric(expression)"
End Sub

Private Sub mIsObject_Click()
Text1.SelText = "IsObject(identifier)"
End Sub

Private Sub mLCase_Click()
Text1.SelText = "LCase(string)"
End Sub

Private Sub mLeft_Click()
Text1.SelText = "Left(string, length)"
End Sub

Private Sub mLen_Click()
Text1.SelText = "Len(string | varname)"
End Sub

Private Sub mListPro_Click()
frmProc.Show , Me
End Sub

Private Sub mLog_Click()
Text1.SelText = "Log(number)"
End Sub

Private Sub mLowerCaps_Click()
Dim StartPoint As Integer, SelectedLength As Integer

If TextNotSelected Then Exit Sub
StartPoint = Text1.SelStart
On Error GoTo Hell
SelectedLength = Text1.SelLength
On Error GoTo Hell
Text1.SelText = StrConv(Text1.SelText, vbLowerCase)
Text1.SelStart = StartPoint
Text1.SelLength = 1
On Error GoTo Hell
Text1.SelText = StrConv(Text1.SelText, vbUpperCase)
Text1.SelStart = Text1.SelStart + SelectedLength

Exit Sub
Hell:
If Err.Number = 6 Then
MsgBox Err.Description, vbCritical, "Error"
Else
HellError
End If

End Sub

Private Sub mLowerCase_Click()
If TextNotSelected Then Exit Sub
On Error GoTo Hell
Text1.SelText = StrConv(Text1.SelText, vbLowerCase)

Exit Sub
Hell:
HellError

End Sub

Private Sub mLTrim_Click()
Text1.SelText = "LTrim(string)"
End Sub

Private Sub mMid_Click()
Text1.SelText = "Mid(string, start[, length])"
End Sub

Private Sub mMinute_Click()
Text1.SelText = "Minute(time)"
End Sub

Private Sub mMonth_Click()
Text1.SelText = "Month(date)"
End Sub

Private Sub mMsgBox_Click()
Text1.SelText = "MsgBox(prompt[, buttons] [, title] [, helpfile, context])"
End Sub

Private Sub mNew_Click()
If Changed = True Then
Dim Res As Integer
    Res = MsgBox("Do you want to save the changes to:" & vbCrLf & Cap, vbYesNoCancel + vbQuestion, "Save File?")
    
    Select Case Res
        Case vbCancel: Exit Sub
        Case vbNo: Text1 = "": Changed = False: Cap = "Untitled": Me.Caption = Cap + " - ScriptPad v1.00": cd.FileName = "": Script1.Reset: Text1.Tag = "":  Script1.AddObject "OutPut", ShowClass, True


        Case vbYes: mSave_Click
    End Select

Else
    Text1 = ""
    Cap = "Untitled"
    Me.Caption = Cap + " - ScriptPad v1.00"
    cd.FileName = ""
    Script1.Reset
    Script1.AddObject "OutPut", ShowClass, True
    Changed = False
End If

End Sub

Private Sub mNow_Click()
Text1.SelText = "Now"
End Sub

Private Sub mnuHis_Click(Index As Integer)
If Changed = True Then
Dim Temp  As Integer
Temp = MsgBox("Do you want to save the changes to: " & vbCr & Cap, vbYesNoCancel + vbQuestion, "Save File?")

If Temp = vbYes Then
    mSave_Click
ElseIf Temp = vbNo Then
    Changed = False
    Script1.Reset
    Script1.AddObject "OutPut", ShowClass, True
Else
    Exit Sub
End If
End If

If Temp = 6 Then If Changed = True Then Exit Sub

cd.FileName = mnuHis(Index).Caption

Close #1
On Error GoTo NoFile
Open mnuHis(Index).Caption For Input As #1
Text1.Text = Input$(LOF(1), #1)
On Error GoTo Handler
Text1.Tag = Text1.Text
UndoStack.Remove UndoStack.Count
mUndo.Enabled = False
mRedo.Enabled = False
On Error GoTo Handler
Cap = mnuHis(Index).Caption
Me.Caption = Cap & " - ScriptPad v1.00"
Changed = False
Script1.Reset
Script1.AddObject "OutPut", ShowClass, True
Close #1

cd.FileName = Cap

Exit Sub
NoFile:
If Not FileExists(cd.FileName) Then
MsgBox "File was not found!", vbInformation, "File not found"
cd.FileName = Cap
mnuHis(Index).Caption = ""
History(Index) = ""
ElseIf Err.Description = "Out of memory" Or Err.Number = 7 Then
MsgBox "File is too large!!", vbExclamation
Else
MsgBox Err.Description, vbInformation, "Error"
'mnuHis(Index).Caption = ""
'History(Index) = ""
'mnuHis(Index).Visible = False
End If

Exit Sub
Handler:
If Err.Number = 7 Or Err.Description = "Out of Memory" Then
Close #1
MsgBox Err.Description, vbCritical
End If
End Sub

Private Sub mOccurrences_Click()
Dim Ask$
On Error GoTo Hell
Ask = InputBox("Type the string you want to know the occurrences of in the current file." & vbCr & vbCr & "Note:-The Search is Case-Sensitive")
If Ask = "" Then
Exit Sub
Else
MsgBox "The number of OCCURRENCES is:" & vbCrLf & Format(getCountOf(Text1.Text, Ask), "###,###,###,###,###"), vbInformation, "String Occurrences"
End If

Exit Sub
Hell:
HellError

End Sub

Private Sub mOct_Click()
Text1.SelText = "Oct(number)"
End Sub

Private Sub mOpen_Click()

If Changed = True Then
Dim Res As Integer
    Res = MsgBox("Do you want to save the changes to:" & vbCrLf & Cap, vbYesNoCancel + vbQuestion, "Save File?")
    
    Select Case Res
        Case vbCancel: Exit Sub
        Case vbNo
        Case vbYes: mSave_Click
        If CancelClicked = True Then CancelClicked = False: Exit Sub
    End Select

End If

If Res = 6 Then If Changed = True Then Exit Sub

On Error Resume Next
cd.CancelError = True
cd.DefaultExt = "txt"
cd.Flags = cdlOFNFileMustExist + cdlOFNPathMustExist
cd.DialogTitle = "Open an Script File"
cd.Filter = "VB Scripts (*.vbs)|*.vbs|All Documents|*.*"
On Error GoTo Hell
cd.ShowOpen

If cd.FileName <> "" Then
Open cd.FileName For Input As #1
Text1 = Input$(LOF(1), #1)
Text1.Tag = Text1.Text
UndoStack.Remove UndoStack.Count
mUndo.Enabled = False
mRedo.Enabled = False
Cap = cd.FileName
Me.Caption = Cap + " - ScriptPad v1.00"
Script1.Reset
Script1.AddObject "OutPut", ShowClass, True
Changed = False
Close #1
End If

On Error Resume Next
AddToHis


Exit Sub
Hell:
Close #1
End Sub

Private Sub mOptions_Click()
frmOptions.Show , Me
End Sub

Private Sub mOutWindow_Click()
frmShow.Show , Me
End Sub

Private Sub mPageColor_Click()
On Error GoTo Handler
frmMain.cd.CancelError = True
frmMain.cd.ShowColor
frmMain.Text1.BackColor = frmMain.cd.Color

Exit Sub
Handler:
End Sub

Private Sub mPageSetup_Click()
On Error GoTo Hell

With cd
    .CancelError = True
    .DialogTitle = "Page Setup"
    .Flags = cdlPDPrintSetup
    .ShowPrinter
End With

Exit Sub
Hell:
End Sub

Private Sub mPaste_Click()
On Error GoTo Hell
Text1.SelText = Clipboard.GetText

Exit Sub
Hell:
End Sub

Private Sub mPrint_Click()
Dim Hheight, Hwidth
On Error Resume Next

With cd
.PrinterDefault = True
'Disable printing to file and individual page printing.
.Flags = cdlPDDisablePrintToFile Or cdlPDNoPageNums

If Text1.SelLength = 0 Then
'Hide Selection button if there is no selected text.
.Flags = .Flags Or cdlPDNoSelection
Else
'Else enable the Selection button and make it the default
'choice.
.Flags = .Flags Or cdlPDSelection
End If

'We need to know whether the user decided to print.
.CancelError = True
.ShowPrinter

If Err = 0 Then
If .Flags And cdlPDSelection Then
Printer.Print Text1.SelText

Else

On Error GoTo Hell
Hheight = Printer.TextHeight(Text1.Text)
Hwidth = Printer.TextWidth(Text1.Text)
Printer.CurrentX = 10
Printer.CurrentY = 10
Printer.Print Text1.Text

End If
End If
Printer.EndDoc
End With

Exit Sub
Hell:
MsgBox Err.Description, vbCritical
End Sub

Private Sub mProcInfo_Click()
On Error Resume Next
frmInfo.Show 1
End Sub

Private Sub mProperCapitalize_Click()
If TextNotSelected Then Exit Sub
On Error GoTo Hell
Text1.SelText = StrConv(Text1.SelText, vbProperCase)

Exit Sub
Hell:
HellError

End Sub

Private Sub mProperties_Click()
frmProperties.Show , Me
End Sub

Private Sub mQuickReplace_Click()
If TextNotSelected Then Exit Sub
Dim a$
Dim strText As String
strText = Text1.Text

a = InputBox("Enter the Text you want to replace the selected text with to replace all occurrences." & vbCr & vbCr & "Note:-The text-selection is Case-Sensitive")
If a = "" Then
Exit Sub
Else
On Error GoTo Handler
Screen.MousePointer = 11
Text1.Text = ReplaceText(strText, Text1.SelText, a)
Screen.MousePointer = 0

Handler:
If Err.Number = 0 Then
Screen.MousePointer = 0
Exit Sub
Else
HellError
Screen.MousePointer = 0
End If
Exit Sub
End If
End Sub

Private Sub mRedo_Click()
Redo
mRedo.Enabled = True
End Sub

Private Sub mRevertOrig_Click()
If Cap = "Untitled" Then Exit Sub
Dim Temp As Integer
Temp = MsgBox("Are you sure you want to revert the file to original and loose the changes made?", vbQuestion + vbYesNo, "Revert File?")

If Temp = vbYes Then
    On Error GoTo Handler
    Text1.Text = Text1.Tag
    'Changed = False
Else
    Exit Sub
End If

Exit Sub
Handler:
HellError
End Sub

Private Sub mRight_Click()
Text1.SelText = "Right(string, length)"
End Sub

Private Sub mRnd_Click()
Text1.SelText = "Rnd[(number)]"
End Sub

Private Sub mRTrim_Click()
Text1.SelText = "RTrim(string)"
End Sub

Private Sub mSave_Click()
If Cap = "Untitled" Then
    Call mSaveAs_Click
    Exit Sub
Else

On Error GoTo Handler
Open cd.FileName For Output As #1
Print #1, Text1
Close #1
mSave.Enabled = False
Changed = False
End If

Exit Sub
Handler:
MsgBox "An error occured while saving the file!", vbCritical
End Sub

Private Sub mSaveAs_Click()
On Error GoTo Hell
cd.CancelError = True
cd.DefaultExt = "vbs"
cd.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt
cd.Filter = "VB Scripts (*.vbs)|*.vbs|All Documents|*.*"
cd.ShowSave

Open cd.FileName For Output As #1
Print #1, Text1
Text1.Tag = Text1.Text
Changed = False
Script1.Reset
Script1.AddObject "OutPut", ShowClass, True
Cap = cd.FileName
Me.Caption = Cap + " - ScriptPad v1.00"
Close #1

On Error Resume Next
AddToHis

Exit Sub
Hell:
If Err.Number = 32755 Then
    CancelClicked = True
End If
End Sub

Private Sub mSaveExit_Click()
mSave_Click
Unload Me
End Sub

Private Sub mSaveSelectionAs_Click()
On Error GoTo DOWN
cd.CancelError = True
cd.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
cd.DialogTitle = "Save Selection As"
cd.DefaultExt = "vbs"
cd.Filter = "VB Scripts (*.vbs)|*.vbs|All Documents|*.*"
On Error GoTo Hell
cd.ShowSave

Open cd.FileName For Output As #1
Print #1, Text1.SelText
Close #1

cd.FileName = Cap

Exit Sub
Hell:
cd.FileName = Cap
DOWN:
If Err.Number = 32755 Then
Exit Sub
Else
HellError
cd.FileName = Cap
End If
End Sub

Private Sub mScript_Click()

If Trim(Text1) = "" Then
    mExecute.Enabled = False
    mListPro.Enabled = False
Else
    mExecute.Enabled = True
    mListPro.Enabled = True
End If

End Sub

Private Sub mSearch_Click()
If Trim(Text1) = "" Then
    mFind.Enabled = False
    mFindNext.Enabled = False
    mOccurrences.Enabled = False
    mQuickReplace.Enabled = False
    mGoto.Enabled = False
Else
    mFind.Enabled = True
    mOccurrences.Enabled = True
    mQuickReplace.Enabled = True
    mGoto.Enabled = True
End If
End Sub

Private Sub mSecond_Click()
Text1.SelText = "Second(time)"
End Sub

Private Sub mSelectAll_Click()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub mSeparator_Click()
Text1.SelText = "'" & String(65, "_")
End Sub

Private Sub mSgn_Click()
Text1.SelText = "Sgn(number)"
End Sub

Private Sub mSin_Click()
Text1.SelText = "Sin(number)"
End Sub

Private Sub mSpellCheck_Click()
Dim m_SpellCheck As clsSpellCheck

If m_SpellCheck Is Nothing Then
 Set m_SpellCheck = New clsSpellCheck

 m_SpellCheck.LoadDict App.Path & "\SpellCheck.dat"

 m_SpellCheck.CheckTextBox Me.Text1

End If

 Set m_SpellCheck = Nothing

End Sub

Private Sub mSqr_Click()
Text1.SelText = "Sqr(number)"
End Sub

Private Sub mStr_Click()
Text1.SelText = "Str(number)"
End Sub

Private Sub mStrComp_Click()
Text1.SelText = "StrComp(string1, string2[, compare])"
End Sub

Private Sub mString_Click()
Text1.SelText = "String(number, character)"
End Sub

Private Sub mTan_Click()
Text1.SelText = "Tan(number)"
End Sub

Private Sub mTextColor_Click()
On Error GoTo Handler
frmMain.cd.CancelError = True
frmMain.cd.ShowColor
frmMain.Text1.ForeColor = frmMain.cd.Color

Exit Sub
Handler:

End Sub

Private Sub mTexttoHtml_Click()
Dim a
On Error Resume Next
Text4.Text = Text1.Text

a = Text4.SelStart
Text4.SelStart = Trim(0)

Dim strTitle
Dim strFC$
Dim strBC$
Dim intSize

strTitle = InputBox("Enter the TitleName for the webpage.")
strBC = InputBox("Enter the BackColor name for the webpage." & vbNewLine & vbNewLine & "Please type the correct Spelling!")
strFC = InputBox("Enter the ForeColor name for the webpage." & vbNewLine & vbNewLine & "Please type the correct Spelling!")
intSize = InputBox("Enter the FontSize of the text for webpage.")

If Trim(strBC) = "" And Trim(strFC) = "" And Trim(intSize) = "" And Trim(strTitle) = "" Then
On Error GoTo BigSize
Text4.SelText = "<PRE>" & vbNewLine & "<BODY BGCOLOR = White >" & vbNewLine & "<FONT color = Black  face = " & Text1.FontName & "  " & "Size = 2" & "</FONT>"
Else
On Error GoTo BigSize
Text4.SelText = "<PRE>" & vbNewLine & "<Title>" & strTitle & "</Title>" & "<BODY BGCOLOR=" & strBC & ">" & vbNewLine & "<FONT color = " & strFC & "  " & "face = " & Text1.FontName & "  " & "Size = " & intSize & "</FONT>"
End If

With frmMain.cd
On Error GoTo Handler
.CancelError = True
.DefaultExt = "htm"
.FileName = ""
.Flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly
.DialogTitle = "Save As"
.Filter = "HTML Document (*.htm;*.html)|*.htm;*.html|"
.ShowSave

Open cd.FileName For Output As #1
Print #1, Text4.Text

cd.FileName = Cap

Close #1
End With


cd.FileName = Cap


Exit Sub
Handler:
Close #1
cd.FileName = Cap

Exit Sub
BigSize:
If Err.Number = 7 Or Err.Description = "Out of Memory" Then
Close #1
MsgBox "Text file is too large!! ", vbExclamation, "Can't convert!"
Else
Close #1
MsgBox Err.Description, vbCritical, "Error"
End If

End Sub

Private Sub mTime_Click()
Text1.SelText = Time
End Sub

Private Sub mTimeF_Click()
Text1.SelText = "Time"
End Sub

Private Sub mTimeSerial_Click()
Text1.SelText = "TimeSerial(hour, minute, second)"
End Sub

Private Sub mTimeValue_Click()
Text1.SelText = "TimeValue(time)"
End Sub

Private Sub mTrim_Click()
Text1.SelText = "Trim(string)"
End Sub

Private Sub mTypingSound_Click()
mTypingSound.Checked = Not mTypingSound.Checked

If mTypingSound.Checked = True Then
On Error Resume Next
    UseSound = "Yes"
Else
    UseSound = ""
End If

End Sub

Private Sub mUCase_Click()
Text1.SelText = "UCase(string)"
End Sub

Private Sub mUndo_Click()
Undo
mRedo.Enabled = True
End Sub

Private Sub mUpperCase_Click()
If TextNotSelected Then Exit Sub
On Error GoTo Hell
Text1.SelText = StrConv(Text1.SelText, vbUpperCase)

Exit Sub
Hell:
HellError

End Sub

Private Sub mVal_Click()
Text1.SelText = "Val(string)"
End Sub

Private Sub mVarType_Click()
Text1.SelText = "VarType(varname)"
End Sub

Private Sub mView_Click()
If Trim(Text1) = "" Then
    mProcInfo.Enabled = False
Else
    mProcInfo.Enabled = True
End If


mOutWindow.Caption = "&Output Window" & vbTab & "Shift+Esc"

End Sub

Private Sub mWeekDay_Click()
Text1.SelText = "Weekday(date, [firstdayofweek])"
End Sub

Private Sub mWordCount_Click()
frmWordCount.Show 1
End Sub

Private Sub mYear_Click()
Text1.SelText = "Year(date)"
End Sub

Private Sub Text1_Change()

If Text1.Text <> "" Then
    Changed = True
End If

If UseSound = "Yes" Then
Dim Play As String
On Error Resume Next
Play = sndPlaySound(App.Path + "\TypingSound.wav", SND_ASYNC)
End If


  
  If Not TrapUndo Then Exit Sub 'because trapping is disabled

    Dim newElement As New UndoElement   'create new undo element
    Dim c%, l&

On Error Resume Next
    'remove all redo items because of the change
    For c% = 1 To RedoStack.Count
        RedoStack.Remove 1
    Next c%

    'set the values of the new element
    newElement.SelStart = Me.Text1.SelStart
    newElement.TextLen = Len(Me.Text1.Text)
    newElement.Text = Me.Text1.Text

    'add it to the undo stack
    UndoStack.Add Item:=newElement
    
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = vbShiftMask And KeyCode = vbKeyEscape Then frmShow.Show , Me
End Sub

Private Sub Timer1_Timer()

On Error Resume Next
StatusBar1.Panels(2).Text = "Line#" & Format(SendMessage(frmMain.Text1.hWnd, EM_LINEFROMCHAR, -1, _
ByVal 0) + 1, "###,###,###,###") & " : Col#" & Format(GetColumnIndex(Text1), "###,###,###,###")

On Error Resume Next
StatusBar1.Panels(3).Text = "Lines:" & Format(SendMessage(Text1.hWnd, EM_GETLINECOUNT, 0, 0&), "###,###,###,###")

End Sub


Public Sub AddToHis()
Dim i As Integer

If Cap = mnuHis(0).Caption Or Cap = mnuHis(1).Caption Or Cap = mnuHis(2).Caption Or Cap = mnuHis(3).Caption Or Cap = mnuHis(4).Caption Then
Exit Sub
Else
For i = 0 To 4
If History(i) = "" Then
History(i) = cd.FileName
mnuHis(i).Caption = History(i)
On Error GoTo theError
mnuHis(i).Visible = True
Sep.Visible = False
Exit Sub
End If
Next
On Error GoTo theError
i = GetSetting(App.EXEName, "options", "add", 0)
History(i) = cd.FileName
mnuHis(i).Caption = History(i)
mnuHis(i).Visible = True
Sep.Visible = False

i = i + 1
If i >= 5 Then i = 0
SaveSetting App.EXEName, "options", "add", i
End If

Exit Sub
theError:
MsgBox Err.Description, vbInformation, "Error"
End Sub

Public Function Change(ByVal lParam1 As String, ByVal lParam2 As String, startSearch As Long) As String
Dim tempParam$
Dim d&
    If Len(lParam1) > Len(lParam2) Then 'swap
        tempParam$ = lParam1
        lParam1 = lParam2
        lParam2 = tempParam$
    End If
    d& = Len(lParam2) - Len(lParam1)
    Change = Mid(lParam2, startSearch - d&, d&)
End Function


Public Sub Undo()
Dim chg$, x&
Dim DeleteFlag As Boolean 'flag as to whether or not to delete text or append text
Dim objElement As Object, objElement2 As Object

On Error Resume Next
    If UndoStack.Count > 1 And TrapUndo Then 'we can proceed
        TrapUndo = False
        DeleteFlag = UndoStack(UndoStack.Count - 1).TextLen < UndoStack(UndoStack.Count).TextLen
        If DeleteFlag Then  'delete some text
            x& = SendMessage(Me.Text1.hWnd, EM_HIDESELECTION, 1&, 1&)
            Set objElement = UndoStack(UndoStack.Count)
            Set objElement2 = UndoStack(UndoStack.Count - 1)
            Me.Text1.SelStart = objElement.SelStart - (objElement.TextLen - objElement2.TextLen)
            Me.Text1.SelLength = objElement.TextLen - objElement2.TextLen
            Me.Text1.SelText = ""
            x& = SendMessage(Me.Text1.hWnd, EM_HIDESELECTION, 0&, 0&)
        Else 'append something
            Set objElement = UndoStack(UndoStack.Count - 1)
            Set objElement2 = UndoStack(UndoStack.Count)
            chg$ = Change(objElement.Text, objElement2.Text, _
                objElement2.SelStart + 1 + Abs(Len(objElement.Text) - Len(objElement2.Text)))
            Me.Text1.SelStart = objElement2.SelStart
            Me.Text1.SelLength = 0
            Me.Text1.SelText = chg$
            Me.Text1.SelStart = objElement2.SelStart
            If Len(chg$) > 1 And chg$ <> vbCrLf Then
                Me.Text1.SelLength = Len(chg$)
            Else
                Me.Text1.SelStart = Me.Text1.SelStart + Len(chg$)
            End If
        End If
        RedoStack.Add Item:=UndoStack(UndoStack.Count)
        UndoStack.Remove UndoStack.Count
    End If
    TrapUndo = True
    Me.Text1.SetFocus
End Sub

Public Sub Redo()
Dim chg$
Dim DeleteFlag As Boolean 'flag as to whether or not to delete text or append text
Dim objElement As Object

On Error Resume Next
    If RedoStack.Count > 0 And TrapUndo Then
        TrapUndo = False
        DeleteFlag = RedoStack(RedoStack.Count).TextLen < Len(Me.Text1.Text)
        If DeleteFlag Then  'delete last item
            Set objElement = RedoStack(RedoStack.Count)
            Me.Text1.SelStart = objElement.SelStart
            Me.Text1.SelLength = Len(Me.Text1.Text) - objElement.TextLen
            Me.Text1.SelText = ""
        Else 'append something
            Set objElement = RedoStack(RedoStack.Count)
            chg$ = Change(Me.Text1.Text, objElement.Text, objElement.SelStart + 1)
            Me.Text1.SelStart = objElement.SelStart - Len(chg$)
            Me.Text1.SelLength = 0
            Me.Text1.SelText = chg$
            Me.Text1.SelStart = objElement.SelStart - Len(chg$)
            If Len(chg$) > 1 And chg$ <> vbCrLf Then
                Me.Text1.SelLength = Len(chg$)
            Else
                Me.Text1.SelStart = Me.Text1.SelStart + Len(chg$)
            End If
        End If
        UndoStack.Add Item:=objElement
        RedoStack.Remove RedoStack.Count
    End If
    TrapUndo = True
    Me.Text1.SetFocus
End Sub


Function GetColumnIndex(txt As TextBox) As Long
'column where caret is
Dim i As Long
i = SendMessage(txt.hWnd, EM_LINEINDEX, ByVal GetCurrentLine(txt) - 1, 0&)
GetColumnIndex = txt.SelStart + 1 - i
End Function

Public Function GetCurrentLine(TextBox As TextBox)
'just what it says
    Dim CurrentLine As Long
    CurrentLine = SendMessage(TextBox.hWnd, EM_LINEFROMCHAR, -1, 0&) + 1
    GetCurrentLine = CurrentLine
End Function


