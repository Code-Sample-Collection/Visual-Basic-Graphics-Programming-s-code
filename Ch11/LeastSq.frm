VERSION 5.00
Begin VB.Form frmLeastSq 
   Caption         =   "LeastSq"
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
Attribute VB_Name = "frmLeastSq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NumPts As Integer
Private PtX() As Single
Private PtY() As Single
' Compute the m and b values for the least squares line.
Private Sub GetLeastSquaresValues(X() As Single, Y() As Single, ByRef m_value As Single, ByRef b_value As Single)
Dim num_points As Integer
Dim A As Single
Dim B As Single
Dim C As Single
Dim D As Single
Dim i As Integer

    ' Compute the sums.
    num_points = UBound(X)
    For i = 1 To num_points
        A = A + PtX(i) * PtX(i)
        B = B + PtX(i)
        C = C + PtX(i) * PtY(i)
        D = D + PtY(i)
    Next i
    m_value = (B * D - C * num_points) / (B * B - A * num_points)
    b_value = (B * C - A * D) / (B * B - A * num_points)
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
Dim m_value As Single
Dim b_value As Single
Dim x1 As Single
Dim x2 As Single
Dim y1 As Single
Dim y2 As Single
Dim i As Integer

    ' Get the m and b values for the line.
    GetLeastSquaresValues PtX, PtY, m_value, b_value

    ' Find the minimum and maximum X values.
    x1 = PtX(1) ' This will be the minimum X value.
    x2 = x1     ' This will be the maximum X value.
    For i = 2 To NumPts
        If x1 > PtX(i) Then x1 = PtX(i)
        If x2 < PtX(i) Then x2 = PtX(i)
    Next i

    ' Draw the line.
    y1 = m_value * x1 + b_value
    y2 = m_value * x2 + b_value
    picCanvas.Line (x1, y1)-(x2, y2)
End Sub
