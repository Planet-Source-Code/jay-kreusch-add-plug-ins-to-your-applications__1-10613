VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Once again, by using a generic entry function, we can do just about anything with our
'plug-in. This one opens up a user interface


Public Function Run() As String
On Error GoTo errtrap
    frmOptions.Show
    Run = vbNullString
Exit Function
errtrap:
    Run = Err.Description
End Function
