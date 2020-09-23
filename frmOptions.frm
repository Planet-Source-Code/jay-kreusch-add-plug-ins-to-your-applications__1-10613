VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Form1"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   3270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Data Conversion Code"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Whatever Command you want"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    MsgBox "This button runs custom code"
End Sub

Private Sub Command2_Click()
    MsgBox "You could make a plug-in that runs data conversion scripts"
End Sub
