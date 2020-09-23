Attribute VB_Name = "Module1"
Option Explicit
Public Position As Double
Public Cap As String
Public Const SND_ASYNC = &H1

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Declare Function SendMessageSTRING Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const EM_LINEFROMCHAR = &HC9
Public Const WM_GETTEXTLENGTH = &HE
Public Const EM_LINEINDEX = &HBB
Public Const WM_USER = &H400
Public Const EM_HIDESELECTION = WM_USER + 63

'Get UserName
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long

'This function counts number of lines.
Public Declare Function SendMessageByVal Lib "user32" _
Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINELENGTH = &HC1


'Declarations for User-defined PopUpMenu
Public Const GWL_WNDPROC = (-4)
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
(ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
(ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_CONTEXTMENU = &H7B

Public origWndProc As Long

Function GetLine(TB As TextBox, ByVal lineNum As Long) As String
Dim charOffset As Long, linelen As Long
    
 ' Retrieve the character offset of the first character of the line.
 charOffset = SendMessageByVal(TB.hWnd, EM_LINEINDEX, lineNum, 0)
 ' Now it's possible to retrieve the length of the line.
 linelen = SendMessageByVal(TB.hWnd, EM_LINELENGTH, charOffset, 0)
 ' Extract the line text.
 GetLine = Mid$(TB.Text, charOffset + 1, linelen)
    
End Function

'True if text is selected
Public Function TextSelected()
TextSelected = frmMain.Text1.SelText <> ""
End Function

'True if text is not selected
Function TextNotSelected() As Boolean
If Len(frmMain.Text1.SelText) = 0 Then
MsgBox "Text Not Selected!", vbExclamation
TextNotSelected = True
End If
End Function

'Get FileName from the Path
Public Function GetFTitle(strFilename As String)
On Error Resume Next
Dim cbBuf As String
    
cbBuf = String(250, vbNullChar) 'Fill buffer with null chars
GetFileTitle strFilename, cbBuf, Len(cbBuf) 'Get file title
GetFTitle = Left(cbBuf, InStr(1, cbBuf, vbNullChar) - 1) 'Extract file title from buffer

End Function

'Unload all forms
Function UnloadForms()
Dim Form As Form
For Each Form In Forms
Unload Form
Set Form = Nothing
Next Form
End Function

'Checks whether a file exists
Function FileExists(ByVal strFilePath As String) As Boolean
strFilePath = Trim(strFilePath)
If strFilePath = "" Then Exit Function
If Dir(strFilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function

'Show the error message
Public Function HellError()
If Err.Number = 7 Or Err.Description = "Out of Memory" Then
MsgBox "File is too large!!", vbExclamation, "File Size Large"
Else
MsgBox Err.Description, vbCritical, "Error"
End If
End Function

'Disables the TextBox default PopUpMenu and enables the
'user-defined one.
Public Sub SetHook(hWnd, bSet As Boolean)
If bSet Then
On Error Resume Next
origWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf AppWndProc)
ElseIf origWndProc Then
Dim lRet As Long
lRet = SetWindowLong(hWnd, GWL_WNDPROC, origWndProc)
End If
End Sub
Public Function AppWndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case Msg
Case WM_CONTEXTMENU
frmMain.PopupMenu frmMain.mEdit
AppWndProc = 0
Exit Function
End Select
On Error Resume Next
AppWndProc = CallWindowProc(origWndProc, hWnd, Msg, wParam, lParam)
End Function

'This function shows Path other than FileName.
Public Function StripPath(ByVal FullPath As String) As String
If InStr(FullPath, "\") = 0 Then
StripPath = FullPath
Exit Function
End If
StripPath = Left(FullPath, InStrRev(FullPath, "\"))
End Function



'Counts the number of occurrences of a given string
Function getCountOf(OriginalString As String, StringToLookFor As String) As Long
On Error GoTo Hell
Dim i As Long
getCountOf = 0  'initilaise the return value as 0

i = 1 ' set it as the first
On Error GoTo Hell
While i <> 0
i = InStr(i, OriginalString, StringToLookFor) ' set it to the next
If i <> 0 Then ' if found then
i = i + 1   ' set it to the found place +1
getCountOf = getCountOf + 1 'increment the count by one
End If
Wend
Exit Function

Hell:
MsgBox Err.Description, vbCritical
End Function


'Replaces Text
Function ReplaceText(Text As String, TextToReplace As String, NewText As String) As String
Dim mtext As String, SpacePos As Long
mtext = Text
SpacePos = InStr(mtext, TextToReplace)
Do While SpacePos
mtext = Left(mtext, SpacePos - 1) & NewText & Mid(mtext, SpacePos + Len(TextToReplace))
SpacePos = InStr(SpacePos + Len(NewText), mtext, TextToReplace)
Loop
ReplaceText = mtext
End Function


'Note:This function might not work with non-Latin alphabets.
Public Function CountSpaces(Text As String) As Long
Dim b() As Byte, i As Long
b() = Text
For i = 0 To UBound(b) Step 2
'Consider only even-numbered items.
'Save time and code using the function name as a local
'variable.
If b(i) = 32 Then CountSpaces = CountSpaces + 1
Next
End Function

