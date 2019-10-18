VERSION 5.00
Begin VB.Form frmBars3 
   Caption         =   "Bars3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmBars3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NUM_VALUES = 7
Private DataValue(1 To NUM_VALUES) As Integer

' Set the form's Scale properties.
Private Sub SetTheScale(ByVal obj As Object, ByVal upper_left_x As Single, ByVal upper_left_y As Single, ByVal lower_right_x As Single, ByVal lower_right_y As Single)
    obj.ScaleLeft = upper_left_x
    obj.ScaleTop = upper_left_y
    obj.ScaleWidth = lower_right_x - upper_left_x
    obj.ScaleHeight = lower_right_y - upper_left_y
End Sub
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
Dim wid As Single
Dim hgt As Single

    ' Define the custom coordinate system.
    SetTheScale Me, 0, 110, NUM_VALUES + 2, -10

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

    ' Draw a 5 by 5 pixel box at position
    ' (NUM_VALUES / 2 + 1, 50).
    wid = ScaleX(5, vbPixels, ScaleMode)
    hgt = ScaleY(5, vbPixels, ScaleMode)
    Line (NUM_VALUES / 2 + 1, 50)-Step(wid, hgt), vbRed, BF
End Sub


