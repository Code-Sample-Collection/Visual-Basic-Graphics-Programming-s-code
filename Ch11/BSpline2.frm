VERSION 5.00
Begin VB.Form frmBspline2 
   Caption         =   "BSpline2"
   ClientHeight    =   5430
   ClientLeft      =   2175
   ClientTop       =   645
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   362
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   322
   Begin VB.CheckBox chkShowT 
      Caption         =   "Show t Values"
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   300
      Width           =   1755
   End
   Begin VB.TextBox txtK 
      Height          =   285
      Left            =   1140
      TabIndex        =   6
      Text            =   "3"
      Top             =   45
      Width           =   375
   End
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
      Left            =   1680
      TabIndex        =   3
      Top             =   0
      Value           =   1  'Checked
      Width           =   1755
   End
   Begin VB.TextBox txtDt 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "0.05"
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
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "K"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   7
      Top             =   60
      Width           =   255
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
Attribute VB_Name = "frmBspline2"
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

' Kvalue determines the smoothness of the curve.
Private Kvalue As Integer

' t runs between 0 and MaxPt - Kvalue + 2.
Private MaxT As Single
' Recursively compute the blending function.
Private Function Blend(ByVal i As Integer, ByVal k As Integer, ByVal t As Single) As Single
Dim numer As Single
Dim denom As Single
Dim value1 As Single
Dim value2 As Single
Dim newt As Single

    If i > 0 Then
        newt = t - i + MaxPt + 1
        Do While newt >= MaxPt + 1
            newt = newt - (MaxPt + 1)
        Loop
        Do While newt < 0
            newt = newt + (MaxPt + 1)
        Loop
        Blend = Blend(0, k, newt)
        Exit Function
    End If

    ' Base case for the recursion.
    If k = 1 Then
        If (Knot(i) <= t) And (t < Knot(i + 1)) Then
            Blend = 1
        ElseIf (t = MaxT) And (Knot(i) <= t) And (t <= Knot(i + 1)) Then
            Blend = 1
        Else
            Blend = 0
        End If
        Exit Function
    End If
    
    denom = Knot(i + k - 1) - Knot(i)
    If denom = 0 Then
        value1 = 0
    Else
        numer = (t - Knot(i)) * Blend(i, k - 1, t)
        value1 = numer / denom
    End If
    
    denom = Knot(i + k) - Knot(i + 1)
    If denom = 0 Then
        value2 = 0
    Else
        numer = (Knot(i + k) - t) * Blend(i + 1, k - 1, t)
        value2 = numer / denom
    End If

    Blend = value1 + value2
End Function

' Draw the curve on the indicated picture box.
Private Sub DrawCurve(pic As PictureBox, start_t As Single, stop_t As Single, dt As Single)
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


' Return the ith knot value.
Private Function Knot(ByVal i As Integer) As Integer
    Knot = i
End Function
' The parametric function Y(t).
Private Function Y(ByVal t As Single) As Single
Dim i As Integer
Dim value As Single

    For i = 0 To MaxPt
        value = value + PtY(i) * Blend(i, Kvalue, t)
    Next i
    Y = value
End Function

' The parametric function X(t).
Private Function X(ByVal t As Single) As Single
Dim i As Integer
Dim value As Single

    For i = 0 To MaxPt
        value = value + PtX(i) * Blend(i, Kvalue, t)
    Next i
    X = value
End Function

' Use DrawCurve to draw the Bezier curve.
Private Sub DrawBspline()
Dim dt As Single
Dim i As Integer
Dim oldstyle As Integer

    If MaxPt < 0 Then Exit Sub

    MousePointer = vbHourglass

    Kvalue = CInt(txtK.Text)
    dt = CSng(txtDt.Text)
    MaxT = MaxPt + 1
    DrawCurve picCanvas, 0, MaxT, dt

    If chkControlPoints.value = vbChecked Then
        ' Draw the control points.
        For i = 0 To MaxPt
            picCanvas.Line _
                (PtX(i) - GAP, PtY(i) - GAP)- _
                Step(2 * GAP, 2 * GAP), , BF
        Next i

        ' Connect the control points.
        oldstyle = picCanvas.DrawStyle
        picCanvas.DrawStyle = vbDot
        picCanvas.CurrentX = PtX(MaxPt)
        picCanvas.CurrentY = PtY(MaxPt)
        For i = 0 To MaxPt
            picCanvas.Line -(PtX(i), PtY(i))
        Next i
        picCanvas.DrawStyle = oldstyle
    End If

    ' Mark the t values if desired.
    If chkShowT.value = vbChecked Then
        For dt = 0 To MaxT Step 1#
            picCanvas.Line (X(dt), Y(dt) - 5)-Step(0, 10)
            picCanvas.Line (X(dt) - 5, Y(dt))-Step(10, 0)
        Next dt
    End If

    MousePointer = vbDefault
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
        
        If MaxPt >= 2 Then cmdGo.Enabled = True
        
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

    DrawBspline
End Sub




Private Sub CmdGo_Click()
    MakingNew = False
    cmdNew.Enabled = True
    DrawBspline
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
    DrawBspline
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

    DrawBspline
End Sub


Private Sub chkShowT_Click()
    DrawBspline
End Sub
