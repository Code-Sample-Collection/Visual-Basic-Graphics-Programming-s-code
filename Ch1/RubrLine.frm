VERSION 5.00
Begin VB.Form frmRubrLine 
   AutoRedraw      =   -1  'True
   Caption         =   "RubrLine"
   ClientHeight    =   4140
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4140
   ScaleWidth      =   6690
End
Attribute VB_Name = "frmRubrLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Rubberbanding As Boolean
Private OldMode As Integer
Private FirstX As Single
Private FirstY As Single
Private LastX As Single
Private LastY As Single

' Start rubberbanding.
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Let MouseMove know we are rubberbanding.
    Rubberbanding = True
    
    ' Save DrawMode so we can restore it later.
    OldMode = DrawMode
    DrawMode = vbInvert

    ' Save the starting coordinates.
    FirstX = X
    FirstY = Y
    
    ' Draw the initial rubberband line.
    LastX = X
    LastY = Y
    Line (FirstX, FirstY)-(LastX, LastY)
End Sub
' Continue rubberbanding.
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If we are not rubberbanding, do nothing.
    If Not Rubberbanding Then Exit Sub

    ' Erase the previous rubberband line.
    Line (FirstX, FirstY)-(LastX, LastY)

    ' Draw the new rubberband line.
    LastX = X
    LastY = Y
    Line (FirstX, FirstY)-(LastX, LastY)
End Sub
' Stop rubberbanding.
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If we are not rubberbanding, do nothing.
    If Not Rubberbanding Then Exit Sub

    ' We are no longer rubberbanding.
    Rubberbanding = False

    ' Erase the previous rubberband line.
    Line (FirstX, FirstY)-(LastX, LastY)

    ' Restore the original DrawMode.
    DrawMode = OldMode

    ' Make the final line permanent.
    Line (FirstX, FirstY)-(LastX, LastY)
End Sub
