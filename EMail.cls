VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


'Here is an example of a formless plug-in.
'The run entry point is used and a simple shellexecute is run

Public Function Run() As String
On Error GoTo errtrap
    Call ShellExecute(0&, "open", "mailto:joe@you.net", vbNullString, vbNullString, vbNormalFocus)
    Run = vbNullString
Exit Function
errtrap:
    Run = Err.Description
End Function
