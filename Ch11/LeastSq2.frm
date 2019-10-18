VERSION 5.00
Begin VB.Form frmLeastSq2 
   Caption         =   "LeastSq2"
   ClientHeight    =   5310
   ClientLeft      =   2085
   ClientTop       =   615
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5310
   ScaleWidth      =   4830
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   4920
      Width           =   615
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   2535
      Left            =   120
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmLeastSq2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NumPts As Integer
Private PtX() As Single
Private PtY() As Single
' Compute the a, b, and c values for quadratic least squares.
Private Sub GetLeastSquaresValues(X() As Single, Y() As Single, ByRef a_value As Single, ByRef b_value As Single, ByRef c_value As Single)
Dim num_points As Integer
Dim A As Single
Dim B As Single
Dim C As Single
Dim D As Single
Dim E As Single
Dim F As Single
Dim G As Single
Dim x2 As Single
Dim x3 As Single
Dim x4 As Single
Dim C2BE As Single
Dim E2CN As Single
Dim BDAF As Single
Dim CFBG As Single
Dim ACB2 As Single
Dim denom As Single
Dim i As Integer

    num_points = UBound(X)

    ' Compute the sums.
    For i = 1 To num_points
        x2 = PtX(i) * PtX(i)
        x3 = x2 * PtX(i)
        x4 = x2 * x2
        A = A + x4
        B = B + x3
        C = C + x2
        D = D + PtY(i) * x2
        E = E + PtX(i)
        F = F + PtY(i) * PtX(i)
        G = G + PtY(i)
    Next i

    ' Compute the quadratic parameters.
    C2BE = C * C - B * E
    E2CN = E * E - C * num_points
    BDAF = B * D - A * F
    CFBG = C * F - B * G
    ACB2 = A * C - B * B
    denom = (B * C - A * E) * C2BE - _
            (C * E - B * num_points) * (B * B - A * C)

    a_value = _
        ((C * D - B * F) * E2CN - (E * F - C * G) * C2BE) / _
        (ACB2 * E2CN + C2BE * C2BE)

    b_value = _
        (CFBG * (B * C - A * E) - BDAF * (C * E - B * num_points)) / _
        denom

    c_value = _
        (BDAF * (C * C - B * E) + CFBG * ACB2) / _
        denom
End Sub
Private Sub Form_Resize()
Dim hgt As Single

    cmdGo.Move (ScaleWidth - cmdGo.Width) / 2, ScaleHeight - cmdGo.Height

    hgt = cmdGo.Top - 30
    If hgt < 120 Then hgt = 120
    picCanvas.Move 0, 0, ScaleWidth, hgt
End Sub


' Add this point to the list of points.
Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const GAP = 2

    ' If this is the first point, erase the screen.
    If NumPts < 1 Then picCanvas.Cls

    ' Record the new point.
    NumPts = NumPts + 1
    ReDim Preserve PtX(1 To NumPts)
    ReDim Preserve PtY(1 To NumPts)
    PtX(NumPts) = X
    PtY(NumPts) = Y

    ' Display the point.
    picCanvas.Line (X - GAP, Y - GAP)-(X + GAP, Y + GAP), , BF

    ' If NumPts >= 2, enable the Go button.
    If NumPts >= 2 Then cmdGo.Enabled = True
End Sub


' Draw the least squares fit curve.
Private Sub cmdGo_Click()
    cmdGo.Enabled = False

    DrawCurve

    ' Prepare to get a new set of points.
    NumPts = 0
End Sub
' Draw the least squares line.
Private Sub DrawCurve()
Dim A As Single
Dim B As Single
Dim C As Single
Dim x1 As Single
Dim x2 As Single
Dim i As Integer
Dim X As Single
Dim dx As Single

    ' Get the parameters for the quadratic.
    GetLeastSquaresValues PtX, PtY, A, B, C

    ' Find the minimum and maximum X values.
    x1 = PtX(1) ' This will be the minimum X value.
    x2 = x1     ' This will be the maximum X value.
    For i = 2 To NumPts
        If x1 > PtX(i) Then x1 = PtX(i)
        If x2 < PtX(i) Then x2 = PtX(i)
    Next i

    ' Draw the curve.
    picCanvas.CurrentX = x1
    picCanvas.CurrentY = A * x1 * x1 + B * x1 + C

    ' Make dx = 1 pixel.
    dx = picCanvas.ScaleX(1, vbPixels, picCanvas.ScaleMode)

    X = x1 + dx
    Do While X < x2
        picCanvas.Line -(X, A * X * X + B * X + C)
        X = X + dx
    Loop

    picCanvas.Line -(x2, A * x2 * x2 + B * x2 + C)
End Sub
