VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Note: The code in this class must be that of VB Script
'because it is run out there by Windows Scripting Host.

Public Sub Show(Message)
frmShow.Text1.SelText = vbCrLf & Message & vbCrLf & String(70, "_") & vbCrLf
On Error Resume Next
frmShow.Show , frmMain
End Sub

Public Sub Clear()
frmShow.Text1 = ""
End Sub



