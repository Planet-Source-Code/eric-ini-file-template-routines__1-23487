VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "Help"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tbHelp 
      Height          =   4335
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Dim hHelp As Integer
   Dim HelpText As String, inputLine As String
   
   tbHelp = ""
   HelpText = App.Path & "\" & "readMe.txt"
   hHelp = FreeFile
   Open HelpText For Input As hHelp
   Do While Not EOF(hHelp)
      Line Input #hHelp, inputLine
      tbHelp = tbHelp & inputLine & vbNewLine
   Loop
End Sub

Private Sub Form_Resize()
   tbHelp.Width = frmHelp.Width - 100
   tbHelp.Height = frmHelp.Height - 500
End Sub
