VERSION 5.00
Begin VB.Form frmQuadMap 
   Caption         =   "QuadMap"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblT 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblS 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "T"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "S"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "frmQuadMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private x1(1 To 3) As Single
Private y1(1 To 3) As Single
Private x2(1 To 3) As Single
Private y2(1 To 3) As Single
Private x3(1 To 3) As Single
Private y3(1 To 3) As Single
Private x4(1 To 3) As Single
Private y4(1 To 3) As Single
' Using s and t values, return the coordinates of a
' point in a quadrilateral.
Private Sub STToPoints(ByRef X As Single, ByRef Y As Single, ByVal s As Single, ByVal t As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single)
Dim xa As Single
Dim ya As Single
Dim xb As Single
Dim yb As Single

    xa = x1 + t * (x2 - x1)
    ya = y1 + t * (y2 - y1)
    xb = x3 + t * (x4 - x3)
    yb = y3 + t * (y4 - y3)
    X = xa + s * (xb - xa)
    Y = ya + s * (yb - ya)
End Sub

Private Sub Form_Load()
Dim i As Integer

    ScaleMode = vbPixels
    AutoRedraw = True

    x1(1) = 20
    x2(1) = 120
    x3(1) = 10
    x4(1) = 150
    y1(1) = 50
    y2(1) = 30
    y3(1) = 130
    y4(1) = 110

    x1(2) = 120
    x2(2) = 210
    x3(2) = 100
    x4(2) = 250
    y1(2) = 150
    y2(2) = 170
    y3(2) = 240
    y4(2) = 260

    x1(3) = 200
    x2(3) = 300
    x3(3) = 200
    x4(3) = 300
    y1(3) = 20
    y2(3) = 20
    y3(3) = 120
    y4(3) = 120

    For i = 1 To 3
        Line (x1(i), y1(i))-(x2(i), y2(i))
        Line -(x4(i), y4(i))
        Line -(x3(i), y3(i))
        Line -(x1(i), y1(i))
    Next i

    Picture = Image
    DrawWidth = 3
End Sub
' Find S and T for the point (X, Y) in the
' quadrilateral with points (x1, y1), (x2, y2),
' (x3, y3), and (x4, y4). Return True if the point
' lies within the quadrilateral and False otherwise.
Private Function PointsToST(ByVal X As Single, ByVal Y As Single, ByRef s As Single, ByRef t As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single) As Boolean
Dim Ax As Single
Dim Bx As Single
Dim Cx As Single
Dim Dx As Single
Dim Ex As Single
Dim Ay As Single
Dim By As Single
Dim Cy As Single
Dim Dy As Single
Dim Ey As Single
Dim a As Single
Dim b As Single
Dim c As Single
Dim det As Single
Dim denom As Single

    Ax = x2 - x1: Ay = y2 - y1
    Bx = x4 - x3: By = y4 - y3
    Cx = x3 - x1: Cy = y3 - y1
    Dx = X - x1: Dy = Y - y1
    Ex = Bx - Ax: Ey = By - Ay

    a = -Ax * Ey + Ay * Ex
    b = Ey * Dx - Dy * Ex + Ay * Cx - Ax * Cy
    c = Dx * Cy - Dy * Cx

    det = b * b - 4 * a * c
    If det >= 0 Then
        If Abs(a) < 0.001 Then
            t = -c / b
        Else
            t = (-b - Sqr(det)) / (2 * a)
        End If
        denom = (Cx + Ex * t)
        If denom > 0.01 Then
            s = (Dx - Ax * t) / denom
        Else
            s = (Dy - Ay * t) / (Cy + Ey * t)
        End If

        PointsToST = (t >= 0# And t <= 1# And _
                  s >= 0# And s <= 1#)
    Else
        PointsToST = False
    End If
End Function
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
Dim j As Integer
Dim s As Single
Dim t As Single
Dim x0 As Single
Dim y0 As Single

    Cls
    lblS.Caption = ""
    lblT.Caption = ""

    ' See which quadrilateral holds the point.
    For i = 1 To 3
        If PointsToST(X, Y, s, t, _
            x1(i), y1(i), x2(i), y2(i), _
            x3(i), y3(i), x4(i), y4(i)) _
        Then Exit For
    Next i

    If i > 3 Then
        ' The point is not in any quadrilateral.
        Beep
    Else
        PSet (X, Y)
        lblS.Caption = Format$(s, "0.00")
        lblT.Caption = Format$(t, "0.00")

        ' Use s and t to map into the
        ' other quadrilaterals.
        For j = 1 To 3
            If i <> j Then
                STToPoints x0, y0, s, t, _
                    x1(j), y1(j), x2(j), y2(j), _
                    x3(j), y3(j), x4(j), y4(j)
                PSet (x0, y0)
            End If
        Next j
    End If
End Sub
