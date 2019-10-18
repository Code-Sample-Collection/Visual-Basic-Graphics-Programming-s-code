VERSION 5.00
Begin VB.Form frmRubrBox 
   AutoRedraw      =   -1  'True
   Caption         =   "RubrBox"
   ClientHeight    =   4140
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4140
   ScaleWidth      =   6690
End
Attribute VB_Name = "frmRubrBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Rubberbanding As Boolean
Private OldMode As Integer
Private OldStyle As Integer
Private FirstX As Single
Private FirstY As Single
Private LastX As Single
Private LastY As Single
' Start rubberbanding.
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Let MouseMove know we are rubberbanding.
    Rubberbanding = True

    ' Save values so we can restore them later.
    OldMode = DrawMode
    OldStyle = DrawStyle
    DrawMode = vbInvert
    DrawStyle = vbDot

    ' Save the starting coordinates.
    FirstX = X
    FirstY = Y

    ' Draw the initial rubberband box.
    LastX = X
    LastY = Y
    Line (FirstX, FirstY)-(LastX, LastY), , B
End Sub
' Continue rubberbanding.
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If we are not rubberbanding, do nothing.
    If Not Rubberbanding Then Exit Sub

    ' Erase the previous rubberband box.
    Line (FirstX, FirstY)-(LastX, LastY), , B

    ' Draw the new rubberband box.
    LastX = X
    LastY = Y
    Line (FirstX, FirstY)-(LastX, LastY), , B
End Sub
' Stop rubberbanding.
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oldfill As Integer
Dim oldcolor As Long

    ' If we are not rubberbanding, do nothing.
    If Not Rubberbanding Then Exit Sub

    ' We are no longer rubberbanding.
    Rubberbanding = False

    ' Erase the previous rubberband box.
    Line (FirstX, FirstY)-(LastX, LastY), , B

    ' Restore the original DrawMode and DrawStyle.
    DrawMode = OldMode
    DrawStyle = OldStyle

    ' Fill the final box with a random color.
    oldfill = FillStyle
    oldcolor = FillColor
    FillStyle = vbSolid
    FillColor = QBColor(Int(Rnd * 16))

    Line (FirstX, FirstY)-(LastX, LastY), , B

    FillStyle = oldfill
    FillColor = oldcolor
End Sub
