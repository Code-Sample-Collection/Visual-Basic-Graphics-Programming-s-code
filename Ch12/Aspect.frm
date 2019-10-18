VERSION 5.00
Begin VB.Form frmAspect 
   Caption         =   "Aspect"
   ClientHeight    =   4230
   ClientLeft      =   1980
   ClientTop       =   1410
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4230
   ScaleWidth      =   4710
   Begin VB.TextBox txtAspect 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "1"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Aspect Ratio"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmAspect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SelectInProgress As Boolean
Private StartX As Single
Private StartY As Single
Private LastX As Single
Private LastY As Single
Private OldMode As Integer
Private ViewAspect As Single

' Begin selecting the region.
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelectInProgress = True

    ' For demonstration purposes, get the desired
    ' aspect ratio from a TextBox.
    ViewAspect = CSng(txtAspect.Text)

    ' Save the current drawing mode.
    OldMode = DrawMode

    ' Use invert mode for the rubberband box.
    DrawMode = vbInvert

    StartX = X
    StartY = Y
    LastX = X
    LastY = Y
    Line (StartX, StartY)-(LastX, LastY), , B
End Sub
' Update the region selected.
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim wid As Single
Dim hgt As Single

    If Not SelectInProgress Then Exit Sub

    ' Erase the old box.
    Line (StartX, StartY)-(LastX, LastY), , B

    wid = X - StartX
    hgt = Y - StartY
    AdjustAspect ViewAspect, wid, hgt
    LastX = StartX + wid
    LastY = StartY + hgt

    ' Draw the new box.
    Line (StartX, StartY)-(LastX, LastY), , B
End Sub

' Adjust ww_wid and ww_hgt so ww_hgt/ww_wid
' equals view_aspect.
Private Sub AdjustAspect(ByVal view_aspect As Single, ByRef ww_wid As Single, ByRef ww_hgt As Single)
Dim ww_aspect As Single
Dim sign As Integer

    ' Don't divide by zero.
    If ww_wid = 0 Or ww_hgt = 0 Or view_aspect = 0 _
        Then Exit Sub

    ww_aspect = ww_hgt / ww_wid
    sign = Sgn(ww_aspect)
    ww_aspect = Abs(ww_aspect)

    If ww_aspect > view_aspect Then
        ' The world window is too tall and thin. Make it wider.
        ww_wid = sign * ww_hgt / view_aspect
    Else
        ' The world window is too short and squat. Make it taller.
        ww_hgt = sign * view_aspect * ww_wid
    End If
End Sub
' Finish selecting the region.
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim wid As Single
Dim hgt As Single

    If Not SelectInProgress Then Exit Sub
    SelectInProgress = False

    ' Erase the old box.
    Line (StartX, StartY)-(LastX, LastY), , B

    ' Restore the original drawing mode.
    DrawMode = OldMode

    wid = X - StartX
    hgt = Y - StartY
    AdjustAspect ViewAspect, wid, hgt
    LastX = StartX + wid
    LastY = StartY + hgt

    ' Do something with the region
    ' (StartX, StartY) - (LastX, LastY).
    Line (StartX, StartY)-(LastX, LastY), vbRed, B
End Sub
