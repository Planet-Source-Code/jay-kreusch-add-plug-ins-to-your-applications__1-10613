VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run Plug-in"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   240
      List            =   "Form1.frx":000A
      TabIndex        =   0
      Text            =   "PlugIn.EMail"
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "You MUST register any ActiveX EXE or DLL before using it as a plug in!!!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This is some VERY simple code to get you started.
'Basically, the whole plug-in idea works like this.
'You late bind your objects and use generic
'entry functions like Run or Start, or GO

Private Sub cmdRun_Click()
On Error GoTo errtrap

    'use a generic object variable
    Dim objPlugIn As Object
    'Variable holds plug-in's response
    Dim strResponse As String
    'call by prog ID. THis can be passed as ANY string
    'whether stored in a database, typed, or created in code
    'The prog ID has the following format:
    '[Project Name].[Class Name]
    Set objPlugIn = CreateObject(Combo1.Text)
    'Call the generic entry proceedure
    strResponse = objPlugIn.Run
    
    'if the plug-in returns an error, let us know
    If strResponse <> vbNullString Then
        MsgBox strResponse
    End If
    
Exit Sub
'Good error trapping is a must when trying something like this
errtrap:
    Select Case Err.Number
        Case 429 'can't create object
            'The ProgID can't be found. Either it is misspelled or the component hasn't been registered!
            MsgBox "You have selected an invalid plug-in ID. Please check that the name is correct and the component is registered."
            Exit Sub
        Case 5 'Invalid proceedure call or argument
            'The 'run' function cannot be found in the class module
            MsgBox "The plug-in you have selected does not have a valid entry point. Please verify the object module with specified guidelines."
            Exit Sub
        Case Else
            'do NOT use the stop statement except for testing purposes.
            Stop
    End Select
End Sub



