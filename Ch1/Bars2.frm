VERSION 5.00
Begin VB.Form frmBars2 
   Caption         =   "Bars2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmBars2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NUM_VALUES = 7
Private DataValue(1 To NUM_VALUES) As Integer
' Create some random data.
Private Sub Form_Load()
Dim i As Integer

    Randomize
    For i = 1 To NUM_VALUES
        DataValue(i) = Rnd * 100
    Next i
End Sub


' Draw the bar chart.
Private Sub Form_Paint()
Dim i As Integer

    ' Define the custom coordinate system.
    ScaleLeft = 0
    ScaleWidth = NUM_VALUES + 2
    ScaleTop = 110
    ScaleHeight = -120

    ' Clear the form.
    Cls

    ' Draw the bar chart.
    For i = 1 To NUM_VALUES
        ' Pick a new fill style.
        FillStyle = i Mod 8

        ' Draw a box with i <= X <= i + 1 and
        ' 0 <= Y <= Data(i).
        Line (i, 0)-(i + 1, DataValue(i)), , B
    Next i
End Sub


