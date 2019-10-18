VERSION 5.00
Begin VB.Form frmLeastSq3 
   Caption         =   "LeastSq3"
   ClientHeight    =   5310
   ClientLeft      =   2085
   ClientTop       =   615
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5310
   ScaleWidth      =   4830
   Begin VB.TextBox txtDegree 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Text            =   "4"
      Top             =   5010
      Width           =   495
   End
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
   Begin VB.Label Label1 
      Caption         =   "Degree"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   5040
      Width           =   615
   End
End
Attribute VB_Name = "frmLeastSq3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NumPts As Integer
Private PtX() As Double
Private PtY() As Double
' Perform Gaussian elimination on this array.
' Return True if there is a solution.
Private Function GaussianEliminate(coeff() As Double) As Boolean
Dim max_row As Integer
Dim max_col As Integer
Dim row As Integer
Dim col As Integer
Dim i As Integer
Dim j As Integer
Dim factor As Double
Dim tmp As Double

    max_row = UBound(coeff, 1)
    max_col = UBound(coeff, 2)
    For row = 0 To max_row
        ' Make sure coeff(row, row) <> 0.
        factor = coeff(row, row)
        If Abs(factor) < 0.001 Then
            ' Switch this row with one that is not
            ' zero in position. Find this row.
            For i = row + 1 To max_row
                If Abs(coeff(i, row) > 0.001) Then
                    ' Switch rows i and row.
                    For j = 0 To max_col
                        tmp = coeff(row, j)
                        coeff(row, j) = coeff(i, j)
                        coeff(i, j) = tmp
                    Next j
                    factor = coeff(row, row)
                End If
            Next i

            ' See if we found a good row.
            If Abs(factor) < 0.001 Then
                ' We found no good row.
                ' There is no solution.
                GaussianEliminate = False
                Exit Function
            End If
        End If

        ' Divide each entry in this row by
        ' coeff(row, row).
        For i = 0 To max_col
            coeff(row, i) = coeff(row, i) / factor
        Next i

        ' Subtract this row from the others.
        For i = 0 To max_row
            If i <> row Then
                ' See what factor we will multiply
                ' by before subtracting for this row.
                factor = coeff(i, row)
                For j = 0 To max_col
                    coeff(i, j) = coeff(i, j) - factor * coeff(row, j)
                Next j
            End If
        Next i
    Next row

    ' There is a solution.
    GaussianEliminate = True
End Function
' Compute the a, b, and c values for quadratic least squares.
' Return True if there is a solution.
Private Function GetLeastSquaresValues(ByVal degree As Integer, X() As Double, Y() As Double, a_values() As Double) As Boolean
Dim max_point As Integer
Dim coeff() As Double
Dim row As Integer
Dim col As Integer
Dim i As Integer
Dim total As Double

    max_point = UBound(X) - 1

    ' Find the coefficients for the equations.
    ReDim coeff(0 To degree, 0 To degree + 1)
    For row = 0 To degree
        ' Find the coefficients for the columns.
        For col = 0 To degree
            ' Find Sum(Xi^(row + col)) over all i.
            total = 0
            For i = 0 To max_point
                total = total + X(i + 1) ^ (row + col)
            Next i
            coeff(row, col) = total
        Next col

        ' Find the constant term.
        total = 0
        For i = 0 To max_point
            total = total + Y(i + 1) * X(i + 1) ^ row
        Next i
        coeff(row, degree + 1) = total
    Next row

    ' Perform the Gaussian elimination.
    If GaussianEliminate(coeff) Then
        ' There is a solution.
        ' Save the results.
        ReDim a_values(0 To degree)
        For row = 0 To degree
            a_values(row) = coeff(row, degree + 1)
        Next row
        GetLeastSquaresValues = True
    Else
        ' There is no solution.
        GetLeastSquaresValues = False
    End If
End Function
' Find the value of the polynomial with the given
' coefficients.
Private Function PolynomialValue(a_values() As Double, ByVal X As Double) As Double
Dim max_coeff As Integer
Dim total As Double
Dim i As Integer
Dim x_power As Double

    max_coeff = UBound(a_values)
    x_power = 1#
    For i = 0 To max_coeff
        total = total + x_power * a_values(i)
        x_power = x_power * X
    Next i

    PolynomialValue = total
End Function

Private Sub Form_Resize()
Dim hgt As Double

    cmdGo.Move (ScaleWidth - cmdGo.Width) / 2, ScaleHeight - cmdGo.Height
    Label1.Top = cmdGo.Top
    txtDegree.Top = cmdGo.Top

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
Dim degree As Integer

    cmdGo.Enabled = False

    degree = CInt(txtDegree.Text)

    ' There's no point making degree >= NumPts.
    If degree >= NumPts Then
        degree = NumPts - 1
        txtDegree.Text = Format$(degree)
    End If

    DrawCurve degree

    ' Prepare to get a new set of points.
    NumPts = 0
End Sub
' Draw the least squares line.
Private Sub DrawCurve(ByVal degree As Integer)
Dim a_values() As Double
Dim x1 As Double
Dim x2 As Double
Dim i As Integer
Dim X As Double
Dim dx As Double

    ' Get the parameters for the quadratic.
    If GetLeastSquaresValues(degree, PtX, PtY, a_values) Then
        ' There is a solution.
        ' Find the minimum and maximum X values.
        x1 = PtX(1) ' This will be the minimum X value.
        x2 = x1     ' This will be the maximum X value.
        For i = 2 To NumPts
            If x1 > PtX(i) Then x1 = PtX(i)
            If x2 < PtX(i) Then x2 = PtX(i)
        Next i

        ' Draw the curve.
        picCanvas.CurrentX = x1
        picCanvas.CurrentY = PolynomialValue(a_values, x1)

        ' Make dx = 1 pixel.
        dx = picCanvas.ScaleX(1, vbPixels, picCanvas.ScaleMode)

        X = x1 + dx
        Do While X < x2
            picCanvas.Line -(X, PolynomialValue(a_values, X))
            X = X + dx
        Loop

        picCanvas.Line -(x2, PolynomialValue(a_values, x2))
    Else
        ' There is no solution.
        MsgBox "There is no solution."
    End If
End Sub
