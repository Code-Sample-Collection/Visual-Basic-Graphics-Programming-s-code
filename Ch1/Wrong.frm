VERSION 5.00
Begin VB.Form frmWrong 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wrong"
   ClientHeight    =   1110
   ClientLeft      =   2220
   ClientTop       =   2325
   ClientWidth     =   1785
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1110
   ScaleWidth      =   1785
   ShowInTaskbar   =   0   'False
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmWrong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Make the form's exterior one inch square.
Private Sub Form_Load()
Dim i As Single

    ' Size the form.
    Width = 1440
    Height = 1440

    ' Draw some "squares."
    Scale (0, 100)-(100, 0)
    For i = 10 To 90 Step 10
        Line (0, i)-(100, i)
        Line (i, 0)-(i, 100)
    Next i
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not frmRight Is Nothing Then _
        Unload frmRight
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub


