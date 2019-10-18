VERSION 5.00
Begin VB.Form frmRight 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Right"
   ClientHeight    =   1185
   ClientLeft      =   4590
   ClientTop       =   2325
   ClientWidth     =   1695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1185
   ScaleWidth      =   1695
   ShowInTaskbar   =   0   'False
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmRight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



' Make the form's interior area one inch square.
Private Sub Form_Load()
Dim extra_wid As Single
Dim extra_hgt As Single
Dim i As Single

    ' Size the form.
    extra_wid = Width - ScaleWidth
    extra_hgt = Height - ScaleHeight
    Me.Width = 1440 + extra_wid
    Me.Height = 1440 + extra_hgt

    ' Draw some squares.
    Scale (0, 100)-(100, 0)
    For i = 10 To 90 Step 10
        Line (0, i)-(100, i)
        Line (i, 0)-(i, 100)
    Next i
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not frmWrong Is Nothing Then _
        Unload frmWrong
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub


