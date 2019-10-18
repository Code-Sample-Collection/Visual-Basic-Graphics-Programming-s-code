VERSION 5.00
Begin VB.Form frmHermite 
   AutoRedraw      =   -1  'True
   Caption         =   "Hermite"
   ClientHeight    =   5685
   ClientLeft      =   1650
   ClientTop       =   360
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   379
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   322
   Begin VB.CheckBox chkControlPoints 
      Caption         =   "Draw Control Points"
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   60
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtVy2 
      Height          =   285
      Left            =   4200
      TabIndex        =   9
      Text            =   "500"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtVx2 
      Height          =   285
      Left            =   3120
      TabIndex        =   7
      Text            =   "-500"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtVy1 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "-500"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtVx1 
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Text            =   "-500"
      Top             =   480
      Width           =   615
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
      Top             =   840
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Vy2"
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   10
      Top             =   510
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Vx2"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   8
      Top             =   510
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Vy1"
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   6
      Top             =   510
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Vx1"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   510
      Width           =   375
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
Attribute VB_Name = "frmHermite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GAP = 2

' The endpoints.
Private Const NumPts = 2
Private PtX(1 To NumPts) As Single
Private PtY(1 To NumPts) As Single

' The index of the point being dragged.
Private Dragging As Integer

' The hermite curve parameters.
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
' Compute the Hermite curve parameters.
Private Sub GetHermiteValues(ByVal ex1 As Single, ByVal ey1 As Single, ByVal ex2 As Single, ByVal ey2 As Single, ByVal vx1 As Single, ByVal vy1 As Single, ByVal vx2 As Single, ByVal vy2 As Single, ByRef Ax As Single, ByRef Bx As Single, ByRef Cx As Single, ByRef Dx As Single, ByRef Ay As Single, ByRef By As Single, ByRef Cy As Single, ByRef Dy As Single)
    Ax = vx2 + vx1 - 2 * ex2 + 2 * ex1
    Bx = 3 * ex2 - 2 * vx1 - 3 * ex1 - vx2
    Cx = vx1
    Dx = ex1

    Ay = vy2 + vy1 - 2 * ey2 + 2 * ey1
    By = 3 * ey2 - 2 * vy1 - 3 * ey1 - vy2
    Cy = vy1
    Dy = ey1
End Sub

' The parametric function Y(t).
Private Function Y(t As Single) As Single
    Y = Ay * t ^ 3 + By * t * t + Cy * t + Dy
End Function

' The parametric function X(t).
Private Function X(t As Single) As Single
    X = Ax * t ^ 3 + Bx * t * t + Cx * t + Dx
End Function

' Prepare to draw the Hermite curve.
Private Sub DrawHermite()
Dim vx1 As Single
Dim vy1 As Single
Dim vx2 As Single
Dim vy2 As Single
Dim dt As Single
Dim i As Integer

    ' Compute the curve parameters.
    vx1 = CSng(txtVx1.Text)
    vy1 = CSng(txtVy1.Text)
    vx2 = CSng(txtVx2.Text)
    vy2 = CSng(txtVy2.Text)
    GetHermiteValues _
        PtX(1), PtY(1), PtX(2), PtY(2), _
        vx1, vy1, vx2, vy2, _
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

        ' Draw the tangents.
        picCanvas.DrawStyle = vbDot
        picCanvas.Line (PtX(1), PtY(1))- _
            (PtX(1) + vx1 / 5, PtY(1) + vy1 / 5)
        picCanvas.Line (PtX(2), PtY(2))- _
            (PtX(2) + vx2 / 5, PtY(2) + vy2 / 5)
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
    
    DrawHermite
End Sub




Private Sub cmdGo_Click()
    DrawHermite
End Sub


Private Sub chkControlPoints_Click()
    DrawHermite
End Sub


Private Sub Form_Load()
    PtX(1) = 0.5 * picCanvas.ScaleWidth
    PtX(2) = 0.8 * picCanvas.ScaleWidth
    PtY(1) = 0.7 * picCanvas.ScaleHeight
    PtY(2) = 0.5 * picCanvas.ScaleHeight
End Sub


' Make the picCanvas as big as possible.
Private Sub Form_Resize()
    picCanvas.Move 0, picCanvas.Top, _
        ScaleWidth, ScaleHeight - picCanvas.Top

    DrawHermite
End Sub
