VERSION 5.00
Begin VB.Form frmBezier 
   Caption         =   "Bezier"
   ClientHeight    =   5490
   ClientLeft      =   2175
   ClientTop       =   645
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   366
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   322
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   0
      Width           =   495
   End
   Begin VB.CheckBox chkControlPoints 
      Caption         =   "Show Control Points"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   60
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.TextBox txtDt 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "0.01"
      Top             =   45
      Width           =   615
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   4815
      Left            =   0
      ScaleHeight     =   317
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   317
      TabIndex        =   0
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "dt"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   60
      Width           =   255
   End
End
Attribute VB_Name = "frmBezier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GAP = 2

' The endpoints are points 1 and 4. The control
' points are points 2 and 3.
Private Const NumPts = 4
Private PtX(1 To NumPts) As Single
Private PtY(1 To NumPts) As Single

' The index of the point being dragged.
Private Dragging As Integer

' The Bezier curve parameters.
Private Ax As Single
Private Bx As Single
Private Cx As Single
Private Dx As Single
Private Ay As Single
Private By As Single
Private Cy As Single
Private Dy As Single
' Draw the curve on the indicated picture box.
Private Sub DrawCurve(ByVal pic As PictureBox, ByVal start_t As Single, ByVal stop_t As Single, ByVal dt As Single)
Dim t As Single

    pic.Cls
    pic.CurrentX = X(start_t)
    pic.CurrentY = Y(start_t)

    t = start_t + dt
    Do While t < stop_t
        pic.Line -(X(t), Y(t))
        t = t + dt
    Loop

    pic.Line -(X(stop_t), Y(stop_t))
End Sub
' Compute the Bezier curve parameters.
Private Sub GetBezierValues(ByVal ex1 As Single, ByVal ey1 As Single, ByVal ex2 As Single, ByVal ey2 As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByRef Ax As Single, ByRef Bx As Single, ByRef Cx As Single, ByRef Dx As Single, ByRef Ay As Single, ByRef By As Single, ByRef Cy As Single, ByRef Dy As Single)
    Ax = ex2 - ex1 - 3 * x2 + 3 * x1
    Bx = 3 * ex1 - 6 * x1 + 3 * x2
    Cx = -3 * ex1 + 3 * x1
    Dx = ex1

    Ay = ey2 - ey1 - 3 * y2 + 3 * y1
    By = 3 * ey1 - 6 * y1 + 3 * y2
    Cy = -3 * ey1 + 3 * y1
    Dy = ey1
End Sub



' The parametric function Y(t).
Private Function Y(ByVal t As Single) As Single
    Y = Ay * t ^ 3 + By * t * t + Cy * t + Dy
End Function

' The parametric function X(t).
Private Function X(ByVal t As Single) As Single
    X = Ax * t ^ 3 + Bx * t * t + Cx * t + Dx
End Function

' Prepare to draw the Bezier curve.
Private Sub DrawBezier()
Dim dt As Single
Dim i As Integer

    ' Compute the curve parameters.
    GetBezierValues _
        PtX(1), PtY(1), _
        PtX(4), PtY(4), _
        PtX(2), PtY(2), _
        PtX(3), PtY(3), _
        Ax, Bx, Cx, Dx, Ay, By, Cy, Dy
    
    ' Draw the curve.
    dt = CSng(txtDt.Text)
    DrawCurve picCanvas, 0, 1, dt

    If chkControlPoints.Value = vbChecked Then
        ' Draw the control points.
        For i = 1 To NumPts
            picCanvas.Line _
                (PtX(i) - GAP, PtY(i) - GAP)- _
                Step(2 * GAP, 2 * GAP), , BF
        Next i
        
        ' Connect the control points.
        picCanvas.DrawStyle = vbDot
        picCanvas.CurrentX = PtX(1)
        picCanvas.CurrentY = PtY(1)
        For i = 2 To NumPts
            picCanvas.Line -(PtX(i), PtY(i))
        Next i
        picCanvas.DrawStyle = vbSolid
    End If
End Sub

' Select a point and start dragging it.
Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer

    ' Find a close point.
    For i = 1 To NumPts
        If Abs(PtX(i) - X) <= GAP And _
           Abs(PtY(i) - Y) <= GAP Then Exit For
    Next i
    If i > NumPts Then Exit Sub

    Dragging = i
    picCanvas.DrawMode = vbInvert
    PtX(Dragging) = X
    PtY(Dragging) = Y
    picCanvas.Line _
        (PtX(Dragging) - GAP, PtY(Dragging) - GAP)- _
        Step(2 * GAP, 2 * GAP), , BF
End Sub


' Continue dragging a point.
Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Dragging < 1 Then Exit Sub
    
    picCanvas.Line _
        (PtX(Dragging) - GAP, PtY(Dragging) - GAP)- _
        Step(2 * GAP, 2 * GAP), , BF
    
    PtX(Dragging) = X
    PtY(Dragging) = Y
    
    picCanvas.Line _
        (PtX(Dragging) - GAP, PtY(Dragging) - GAP)- _
        Step(2 * GAP, 2 * GAP), , BF
End Sub


' Finish the drag and redraw the curve.
Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Dragging < 1 Then Exit Sub
    
    picCanvas.DrawMode = vbCopyPen

    PtX(Dragging) = X
    PtY(Dragging) = Y
    Dragging = 0
    
    DrawBezier
End Sub




Private Sub cmdGo_Click()
    DrawBezier
End Sub

Private Sub chkControlPoints_Click()
    DrawBezier
End Sub

Private Sub Form_Load()
    PtX(1) = 0.4 * picCanvas.ScaleWidth
    PtX(2) = 0.1 * picCanvas.ScaleWidth
    PtX(3) = 0.8 * picCanvas.ScaleWidth
    PtX(4) = 0.6 * picCanvas.ScaleWidth
    PtY(1) = 0.8 * picCanvas.ScaleHeight
    PtY(2) = 0.3 * picCanvas.ScaleHeight
    PtY(3) = 0.2 * picCanvas.ScaleHeight
    PtY(4) = 0.7 * picCanvas.ScaleHeight
End Sub

' Make the picCanvas as big as possible.
Private Sub Form_Resize()
    picCanvas.Move 0, picCanvas.Top, _
        ScaleWidth, ScaleHeight - picCanvas.Top
        
    DrawBezier
End Sub
