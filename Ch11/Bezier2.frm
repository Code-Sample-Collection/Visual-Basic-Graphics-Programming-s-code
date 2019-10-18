VERSION 5.00
Begin VB.Form frmBezier2 
   Caption         =   "Bezier2"
   ClientHeight    =   5490
   ClientLeft      =   2175
   ClientTop       =   645
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   366
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   322
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
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
Attribute VB_Name = "frmBezier2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GAP = 2

' The endpoints are points 1 and 4. The control
' points are points 2 and 3.
Private MaxPt As Integer
Private PtX() As Single
Private PtY() As Single

Private MakingNew As Boolean

' The index of the point being dragged.
Private Dragging As Integer
' The blending function for i, N, and t.
Private Function Blend(ByVal i As Integer, ByVal N As Integer, ByVal t As Single) As Single
    Blend = Factorial(N) / Factorial(i) / _
        Factorial(N - i) * t ^ i * (1 - t) ^ (N - i)
End Function

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

' Return the factorial of a number.
Private Function Factorial(ByVal N As Integer) As Long
Dim value As Long
Dim i As Integer

    value = 1
    For i = 2 To N
        value = value * i
    Next i
    Factorial = value
End Function
' The parametric function Y(t).
Private Function Y(ByVal t As Single) As Single
Dim i As Integer
Dim value As Single

    For i = 0 To MaxPt
        value = value + PtY(i) * Blend(i, MaxPt, t)
    Next i
    Y = value
End Function

' The parametric function X(t).
Private Function X(ByVal t As Single) As Single
Dim i As Integer
Dim value As Single

    For i = 0 To MaxPt
        value = value + PtX(i) * Blend(i, MaxPt, t)
    Next i
    X = value
End Function

' Use DrawCurve to draw the Bezier curve.
Private Sub DrawBezier()
Dim dt As Single
Dim i As Integer

    If MaxPt < 0 Then Exit Sub

    dt = CSng(txtDt.Text)
    DrawCurve picCanvas, 0, 1, dt

    If chkControlPoints.value = vbChecked Then
        ' Draw the control points.
        For i = 0 To MaxPt
            picCanvas.Line _
                (PtX(i) - GAP, PtY(i) - GAP)- _
                Step(2 * GAP, 2 * GAP), , BF
        Next i

        ' Connect the control points.
        picCanvas.DrawStyle = vbDot
        picCanvas.CurrentX = PtX(0)
        picCanvas.CurrentY = PtY(0)
        For i = 1 To MaxPt
            picCanvas.Line -(PtX(i), PtY(i))
        Next i
        picCanvas.DrawStyle = vbSolid
    End If
End Sub

' Either collect a new point or select a point and
' start dragging it.
Private Sub picCanvas_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer

    ' If we are selecting points, do so now.
    If MakingNew Then
        MaxPt = MaxPt + 1
        ReDim Preserve PtX(0 To MaxPt)
        ReDim Preserve PtY(0 To MaxPt)
        PtX(MaxPt) = X
        PtY(MaxPt) = Y
        picCanvas.Line _
            (X - GAP, Y - GAP)- _
            Step(2 * GAP, 2 * GAP), , BF
        
        If MaxPt >= 3 Then cmdGo.Enabled = True
        
        Exit Sub
    End If

    ' Otherwise start dragging a point.
    ' Find a close point.
    For i = 0 To MaxPt
        If Abs(PtX(i) - X) <= GAP And _
           Abs(PtY(i) - Y) <= GAP Then Exit For
    Next i
    If i > MaxPt Then Exit Sub

    Dragging = i
    picCanvas.DrawMode = vbInvert
    PtX(Dragging) = X
    PtY(Dragging) = Y
    picCanvas.Line _
        (PtX(Dragging) - GAP, PtY(Dragging) - GAP)- _
        Step(2 * GAP, 2 * GAP), , BF
End Sub


' Continue dragging a point.
Private Sub picCanvas_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
    If Dragging < 0 Then Exit Sub
    
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
Private Sub picCanvas_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
    If Dragging < 0 Then Exit Sub
    
    picCanvas.DrawMode = vbCopyPen

    PtX(Dragging) = X
    PtY(Dragging) = Y
    Dragging = -1
    
    DrawBezier
End Sub




Private Sub CmdGo_Click()
    MakingNew = False
    cmdNew.Enabled = True
    DrawBezier
End Sub

' Prepare to get new points.
Private Sub CmdNew_Click()
    MaxPt = -1
    cmdGo.Enabled = False
    cmdNew.Enabled = False
    MakingNew = True
    picCanvas.Cls
End Sub

Private Sub chkControlPoints_Click()
    DrawBezier
End Sub

Private Sub Form_Load()
    MakingNew = True
    MaxPt = -1
    Dragging = -1
End Sub

' Make the picCanvas as big as possible.
Private Sub Form_Resize()
    picCanvas.Move 0, picCanvas.Top, _
        ScaleWidth, ScaleHeight - picCanvas.Top

    DrawBezier
End Sub
