VERSION 5.00
Begin VB.Form frmShapes 
   Caption         =   "Shapes"
   ClientHeight    =   5925
   ClientLeft      =   1710
   ClientTop       =   465
   ClientWidth     =   7845
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5925
   ScaleWidth      =   7845
   Begin VB.OptionButton optShape 
      Caption         =   "Sphere"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   13
      Top             =   4440
      Width           =   1695
   End
   Begin VB.OptionButton optShape 
      Caption         =   "Hill and Hole"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   12
      Top             =   4080
      Width           =   1695
   End
   Begin VB.OptionButton optShape 
      Caption         =   "Monkey Saddle"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   1695
   End
   Begin VB.OptionButton optShape 
      Caption         =   "Dome"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   1695
   End
   Begin VB.OptionButton optShape 
      Caption         =   "Waves"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   1695
   End
   Begin VB.OptionButton optShape 
      Caption         =   "Stelate Octahedron"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   1695
   End
   Begin VB.OptionButton optShape 
      Caption         =   "Splash"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1695
   End
   Begin VB.OptionButton optShape 
      Caption         =   "8 Cubes"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.OptionButton optShape 
      Caption         =   "Icosahedron"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.OptionButton optShape 
      Caption         =   "Octahedron"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.OptionButton optShape 
      Caption         =   "Dodecahedron"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.OptionButton optShape 
      Caption         =   "Cube"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.OptionButton optShape 
      Caption         =   "Tetrahedron"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   5895
      Left            =   1920
      ScaleHeight     =   -800
      ScaleLeft       =   -400
      ScaleMode       =   0  'User
      ScaleTop        =   400
      ScaleWidth      =   800
      TabIndex        =   6
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "frmShapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Location of viewing eye.
Private EyeR As Single
Private EyeTheta As Single
Private EyePhi As Single

' Location of focus point.
Private Const FocusX = 0#
Private Const FocusY = 0#
Private Const FocusZ = 0#

Private Projector(1 To 4, 1 To 4) As Single

Private Polylines As Collection

Private Enum ShapeValues
    shape_Tetrahedron = 0
    shape_Cube = 1
    shape_Octahedron = 2
    shape_Dodecahedron = 3
    shape_Icosahedron = 4
    shape_8Cubes = 5
    shape_Stellate8 = 6
    shape_Surface1 = 7
    shape_Surface2 = 8
    shape_Surface3 = 9
    shape_Surface4 = 10
    shape_Surface5 = 11
    shape_Sphere = 12
End Enum

Private SelectedShape As ShapeValues
' Make a cube with the indicated center and
' side length.
Private Sub MakeCube(ByVal Cx As Single, ByVal Cy As Single, ByVal Cz As Single, ByVal side_length As Single)
Dim s2 As Single
Dim poly As Polyline3d

    s2 = side_length / 2
    Set poly = New Polyline3d
    Polylines.Add poly
    poly.AddSegment (Cx + s2), (Cy + s2), (Cz + s2), (Cx + s2), (Cy - s2), (Cz + s2)
    poly.AddSegment (Cx + s2), (Cy - s2), (Cz + s2), (Cx - s2), (Cy - s2), (Cz + s2)
    poly.AddSegment (Cx - s2), (Cy - s2), (Cz + s2), (Cx - s2), (Cy + s2), (Cz + s2)
    poly.AddSegment (Cx - s2), (Cy + s2), (Cz + s2), (Cx + s2), (Cy + s2), (Cz + s2)
    poly.AddSegment (Cx + s2), (Cy + s2), (Cz - s2), (Cx + s2), (Cy - s2), (Cz - s2)
    poly.AddSegment (Cx + s2), (Cy - s2), (Cz - s2), (Cx - s2), (Cy - s2), (Cz - s2)
    poly.AddSegment (Cx - s2), (Cy - s2), (Cz - s2), (Cx - s2), (Cy + s2), (Cz - s2)
    poly.AddSegment (Cx - s2), (Cy + s2), (Cz - s2), (Cx + s2), (Cy + s2), (Cz - s2)
    poly.AddSegment (Cx + s2), (Cy + s2), (Cz + s2), (Cx + s2), (Cy + s2), (Cz - s2)
    poly.AddSegment (Cx + s2), (Cy - s2), (Cz + s2), (Cx + s2), (Cy - s2), (Cz - s2)
    poly.AddSegment (Cx - s2), (Cy - s2), (Cz + s2), (Cx - s2), (Cy - s2), (Cz - s2)
    poly.AddSegment (Cx - s2), (Cy + s2), (Cz + s2), (Cx - s2), (Cy + s2), (Cz - s2)
End Sub

' Make a stellate octahedron.
Private Sub MakeStellate8(ByVal side_scale As Single)
Dim poly As Polyline3d

    Set poly = New Polyline3d
    Polylines.Add poly

    poly.Stellate side_scale, _
        0, side_scale, 0, _
        0, 0, side_scale, _
        side_scale, 0, 0
    poly.Stellate side_scale, _
        0, side_scale, 0, _
        side_scale, 0, 0, _
        0, 0, -side_scale
    poly.Stellate side_scale, _
        0, side_scale, 0, _
        0, 0, -side_scale, _
        -side_scale, 0, 0
    poly.Stellate side_scale, _
        0, side_scale, 0, _
        -side_scale, 0, 0, _
        0, 0, side_scale

    poly.Stellate side_scale, _
        0, -side_scale, 0, _
        side_scale, 0, 0, _
        0, 0, side_scale
    poly.Stellate side_scale, _
        0, -side_scale, 0, _
        0, 0, -side_scale, _
        side_scale, 0, 0
    poly.Stellate side_scale, _
        0, -side_scale, 0, _
        -side_scale, 0, 0, _
        0, 0, -side_scale
    poly.Stellate side_scale, _
        0, -side_scale, 0, _
        0, 0, side_scale, _
        -side_scale, 0, 0
End Sub
' Make a surface.
Private Sub MakeSurface1()
Const GAP = 25

Dim poly As Polyline3d
Dim i As Single
Dim j As Single
Dim R1 As Single
Dim r2 As Single
Dim z1 As Single
Dim z2 As Single

    Set poly = New Polyline3d
    Polylines.Add poly

    For i = -300 To 300 Step GAP
        For j = -300 To 300 - GAP Step GAP
            R1 = Sqr(i * i + j * j)
            r2 = Sqr(i * i + (j + GAP) * (j + GAP))
            z1 = 200 * Cos(R1 / 25) / (1 + R1 / 20)
            z2 = 200 * Cos(r2 / 25) / (1 + r2 / 20)
            poly.AddSegment i, z1, j, i, z2, j + GAP
            poly.AddSegment j, z1, i, j + GAP, z2, i
        Next j
    Next i
End Sub
' Make a surface.
Private Sub MakeSurface2()
Const GAP = 25

Dim poly As Polyline3d
Dim i As Single
Dim j As Single
Dim z1 As Single
Dim z2 As Single

    Set poly = New Polyline3d
    Polylines.Add poly

    For i = -300 To 300 Step GAP
        For j = -300 To 300 - GAP Step GAP
            z1 = 20 * (Cos(i * PI / 150) + Cos(j * PI / 150))
            z2 = 20 * (Cos(i * PI / 150) + Cos((j + GAP) * PI / 150))
            poly.AddSegment i, z1, j, i, z2, j + GAP
            poly.AddSegment j, z1, i, j + GAP, z2, i
        Next j
    Next i
End Sub
' Make a surface.
Private Sub MakeSurface3()
Const GAP = 25

Dim poly As Polyline3d
Dim i As Single
Dim j As Single
Dim R1 As Single
Dim r2 As Single
Dim z1 As Single
Dim z2 As Single

    Set poly = New Polyline3d
    Polylines.Add poly

    For i = -300 To 300 Step GAP
        For j = -300 To 300 - GAP Step GAP
            z1 = 300 - (i * i + j * j) / 300
            z2 = 300 - (i * i + (j + GAP) * (j + GAP)) / 300
            poly.AddSegment i, z1, j, i, z2, j + GAP
            poly.AddSegment j, z1, i, j + GAP, z2, i
        Next j
    Next i
End Sub
' Make a surface.
Private Sub MakeSurface4()
Const GAP = 25

Dim poly As Polyline3d
Dim i As Single
Dim j As Single
Dim R1 As Single
Dim r2 As Single
Dim z1 As Single
Dim z2 As Single
Dim X As Single
Dim Y As Single

    Set poly = New Polyline3d
    Polylines.Add poly

    For i = -300 To 300 Step GAP
        For j = -300 To 300 - GAP Step GAP
            X = i / 40
            Y = j / 40
            z1 = X * X * X / 3 - X * Y * Y
            Y = (j + GAP) / 40
            z2 = X * X * X / 3 - X * Y * Y
            poly.AddSegment i, z1, j, i, z2, j + GAP

            X = j / 40
            Y = i / 40
            z1 = X * X * X / 3 - X * Y * Y
            X = (j + GAP) / 40
            z2 = X * X * X / 3 - X * Y * Y
            poly.AddSegment j, z1, i, j + GAP, z2, i
        Next j
    Next i
End Sub

' Make a surface.
Private Sub MakeSurface5()
Const GAP = 25

Dim poly As Polyline3d
Dim i As Single
Dim j As Single
Dim R1 As Single
Dim r2 As Single
Dim z1 As Single
Dim z2 As Single
Dim X As Single
Dim Y As Single

    Set poly = New Polyline3d
    Polylines.Add poly

    For i = -300 To 300 Step GAP
        For j = -300 To 300 - GAP Step GAP
            X = i / 40
            Y = j / 40
            z1 = 100 * (-5 * X / (X * X + Y * Y + 1))
            Y = (j + GAP) / 40
            z2 = 100 * (-5 * X / (X * X + Y * Y + 1))
            poly.AddSegment i, z1, j, i, z2, j + GAP

            X = j / 40
            Y = i / 40
            z1 = 100 * (-5 * X / (X * X + Y * Y + 1))
            X = (j + GAP) / 40
            z2 = 100 * (-5 * X / (X * X + Y * Y + 1))
            poly.AddSegment j, z1, i, j + GAP, z2, i
        Next j
    Next i
End Sub


' Make a tetrahedron.
Private Sub MakeTetrahedron(ByVal side_length As Single)
Dim S As Single
Dim A As Single
Dim B As Single
Dim C As Single
Dim D As Single
Dim poly As Polyline3d

    Set poly = New Polyline3d
    Polylines.Add poly

    S = Sqr(6) * side_length
    A = S / Sqr(3)
    B = -A / 2
    C = A * Sqr(2) - 1
    D = S / 2
    poly.AddSegment 0, C, 0, A, -1, 0
    poly.AddSegment 0, C, 0, B, -1, D
    poly.AddSegment 0, C, 0, B, -1, -D
    poly.AddSegment B, -1, -D, B, -1, D
    poly.AddSegment B, -1, D, A, -1, 0
    poly.AddSegment A, -1, 0, B, -1, -D
End Sub

' Make an octahedron.
Private Sub MakeOctahedron(ByVal side_scale As Single)
Dim poly As Polyline3d

    Set poly = New Polyline3d
    Polylines.Add poly

    poly.AddSegment 0, side_scale, 0, side_scale, 0, 0
    poly.AddSegment 0, side_scale, 0, -side_scale, 0, 0
    poly.AddSegment 0, side_scale, 0, 0, 0, side_scale
    poly.AddSegment 0, side_scale, 0, 0, 0, -side_scale

    poly.AddSegment 0, -side_scale, 0, side_scale, 0, 0
    poly.AddSegment 0, -side_scale, 0, -side_scale, 0, 0
    poly.AddSegment 0, -side_scale, 0, 0, 0, side_scale
    poly.AddSegment 0, -side_scale, 0, 0, 0, -side_scale

    poly.AddSegment 0, 0, side_scale, side_scale, 0, 0
    poly.AddSegment 0, 0, side_scale, -side_scale, 0, 0
    poly.AddSegment 0, 0, -side_scale, side_scale, 0, 0
    poly.AddSegment 0, 0, -side_scale, -side_scale, 0, 0
End Sub
' Make an icosahedron.
Private Sub MakeIcosahedron(ByVal side_scale As Single)
Dim poly As Polyline3d
Dim theta1 As Single
Dim theta2 As Single
Dim s1 As Single
Dim s2 As Single
Dim c1 As Single
Dim c2 As Single
Dim A As Single
Dim B As Single
Dim C As Single
Dim D As Single
Dim H As Single
Dim S As Single
Dim R As Single

    Set poly = New Polyline3d
    Polylines.Add poly

    theta1 = PI * 0.4
    theta2 = PI * 0.8
    s1 = Sin(theta1)
    c1 = Cos(theta1)
    s2 = Sin(theta2)
    c2 = Cos(theta2)
    R = 2 / (2 * Sqr(1 - 2 * c1) + Sqr(3 / 4 * (2 - 2 * c1) - 2 * c2 - c2 * c2 - 1)) * side_scale
    S = R * Sqr(2 - 2 * c1)
    H = 1 - Sqr(S * S - R * R)
    A = R * s1
    B = R * s2
    C = R * c1
    D = R * c2
    poly.AddSegment R, H, 0, C, H, A        ' Top
    poly.AddSegment C, H, A, D, H, B
    poly.AddSegment D, H, B, D, H, -B
    poly.AddSegment D, H, -B, C, H, -A
    poly.AddSegment C, H, -A, R, H, 0
    poly.AddSegment R, H, 0, 0, -side_scale, 0        ' Point
    poly.AddSegment C, H, A, 0, -side_scale, 0
    poly.AddSegment D, H, B, 0, -side_scale, 0
    poly.AddSegment D, H, -B, 0, -side_scale, 0
    poly.AddSegment C, H, -A, 0, -side_scale, 0

    poly.AddSegment -R, -H, 0, -C, -H, A    ' Bottom
    poly.AddSegment -C, -H, A, -D, -H, B
    poly.AddSegment -D, -H, B, -D, -H, -B
    poly.AddSegment -D, -H, -B, -C, -H, -A
    poly.AddSegment -C, -H, -A, -R, -H, 0
    poly.AddSegment -R, -H, 0, 0, side_scale, 0     ' Point
    poly.AddSegment -C, -H, A, 0, side_scale, 0
    poly.AddSegment -D, -H, B, 0, side_scale, 0
    poly.AddSegment -D, -H, -B, 0, side_scale, 0
    poly.AddSegment -C, -H, -A, 0, side_scale, 0

    poly.AddSegment R, H, 0, -D, -H, B      ' Middle
    poly.AddSegment R, H, 0, -D, -H, -B
    poly.AddSegment C, H, A, -D, -H, B
    poly.AddSegment C, H, A, -C, -H, A
    poly.AddSegment D, H, B, -C, -H, A
    poly.AddSegment D, H, B, -R, -H, 0
    poly.AddSegment D, H, -B, -R, -H, 0
    poly.AddSegment D, H, -B, -C, -H, -A
    poly.AddSegment C, H, -A, -C, -H, -A
    poly.AddSegment C, H, -A, -D, -H, -B
End Sub

' Make a dodecahedron.
Private Sub MakeDodecahedron(ByVal side_scale As Single)
Dim poly As Polyline3d
Dim theta1 As Single
Dim theta2 As Single
Dim s1 As Single
Dim s2 As Single
Dim c1 As Single
Dim c2 As Single
Dim M As Single
Dim N As Single
Dim S As Single
Dim R As Single
Dim A As Single
Dim B As Single
Dim C As Single
Dim D As Single
Dim H As Single
Dim X As Single
Dim Y As Single
Dim y2 As Single

    Set poly = New Polyline3d
    Polylines.Add poly

    ' Dodecahedron.
    theta1 = PI * 0.4
    theta2 = PI * 0.8
    s1 = Sin(theta1)
    c1 = Cos(theta1)
    s2 = Sin(theta2)
    c2 = Cos(theta2)

    M = 1 - (2 - 2 * c1 - 4 * s1 * s1) / (2 * c1 - 2)
    N = Sqr((2 - 2 * c1) - M * M) * (1 + (1 - c2) / (c1 - c2))
    R = 2 / N * side_scale
    S = R * Sqr(2 - 2 * c1)
    A = R * s1
    B = R * s2
    C = R * c1
    D = R * c2
    H = R * (c1 - s1)

    X = (R * R * (2 - 2 * c1) - 4 * A * A) / (2 * C - 2 * R)
    Y = Sqr(S * S - (R - X) * (R - X))
    y2 = Y * (1 - c2) / (c1 - c2)

    poly.AddSegment R, side_scale, 0, C, side_scale, A        ' Top
    poly.AddSegment C, side_scale, A, D, side_scale, B
    poly.AddSegment D, side_scale, B, D, side_scale, -B
    poly.AddSegment D, side_scale, -B, C, side_scale, -A
    poly.AddSegment C, side_scale, -A, R, side_scale, 0

    poly.AddSegment R, side_scale, 0, X, side_scale - Y, 0    ' Top downward edges.
    poly.AddSegment C, side_scale, A, X * c1, side_scale - Y, X * s1
    poly.AddSegment C, side_scale, -A, X * c1, side_scale - Y, -X * s1
    poly.AddSegment D, side_scale, B, X * c2, side_scale - Y, X * s2
    poly.AddSegment D, side_scale, -B, X * c2, side_scale - Y, -X * s2

    poly.AddSegment X, side_scale - Y, 0, -X * c2, side_scale - y2, -X * s2   ' Middle.
    poly.AddSegment X, side_scale - Y, 0, -X * c2, side_scale - y2, X * s2
    poly.AddSegment X * c1, side_scale - Y, X * s1, -X * c2, side_scale - y2, X * s2
    poly.AddSegment X * c1, side_scale - Y, X * s1, -X * c1, side_scale - y2, X * s1
    poly.AddSegment X * c2, side_scale - Y, X * s2, -X * c1, side_scale - y2, X * s1
    poly.AddSegment X * c2, side_scale - Y, X * s2, -X, side_scale - y2, 0
    poly.AddSegment X * c2, side_scale - Y, -X * s2, -X, side_scale - y2, 0
    poly.AddSegment X * c2, side_scale - Y, -X * s2, -X * c1, side_scale - y2, -X * s1
    poly.AddSegment X * c1, side_scale - Y, -X * s1, -X * c1, side_scale - y2, -X * s1
    poly.AddSegment X * c1, side_scale - Y, -X * s1, -X * c2, side_scale - y2, -X * s2

    poly.AddSegment -R, -side_scale, 0, -X, side_scale - y2, 0    ' Bottom upward edges.
    poly.AddSegment -C, -side_scale, A, -X * c1, side_scale - y2, X * s1 ' Bottom upward edges.
    poly.AddSegment -D, -side_scale, B, -X * c2, side_scale - y2, X * s2
    poly.AddSegment -D, -side_scale, -B, -X * c2, side_scale - y2, -X * s2
    poly.AddSegment -C, -side_scale, -A, -X * c1, side_scale - y2, -X * s1

    poly.AddSegment -R, -side_scale, 0, -C, -side_scale, A    ' Bottom
    poly.AddSegment -C, -side_scale, A, -D, -side_scale, B
    poly.AddSegment -D, -side_scale, B, -D, -side_scale, -B
    poly.AddSegment -D, -side_scale, -B, -C, -side_scale, -A
    poly.AddSegment -C, -side_scale, -A, -R, -side_scale, 0
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Const dtheta = PI / 20

    Select Case KeyCode
        Case vbKeyLeft
            EyeTheta = EyeTheta - dtheta
        
        Case vbKeyRight
            EyeTheta = EyeTheta + dtheta
        
        Case vbKeyUp
            EyePhi = EyePhi - dtheta
        
        Case vbKeyDown
            EyePhi = EyePhi + dtheta
        
        Case Else
            Exit Sub
    End Select

    DrawData picCanvas
End Sub

Private Sub Form_Load()
    ' Initialize the eye position.
    EyeR = 1500
    EyeTheta = PI * 0.17
    EyePhi = PI * 0.16
End Sub


' Create some polylines to display.
Private Sub CreateData()
Dim poly As Polyline3d

    ' Create the polyline collection.
    Set Polylines = New Collection

    Select Case SelectedShape
        Case shape_Tetrahedron
            MakeTetrahedron 150

        Case shape_Cube
            MakeCube 0, 0, 0, 400

        Case shape_Octahedron
            MakeOctahedron 300

        Case shape_Dodecahedron
            MakeDodecahedron 300

        Case shape_Icosahedron
            MakeIcosahedron 300

        Case shape_8Cubes
            MakeCube 150, 150, 150, 100
            MakeCube 150, 150, -150, 100
            MakeCube 150, -150, 150, 100
            MakeCube -150, 150, 150, 100
            MakeCube 150, -150, -150, 100
            MakeCube -150, 150, -150, 100
            MakeCube -150, -150, 150, 100
            MakeCube -150, -150, -150, 100

        Case shape_Stellate8
            MakeStellate8 250

        Case shape_Surface1
            MakeSurface1

        Case shape_Surface2
            MakeSurface2

        Case shape_Surface3
            MakeSurface3

        Case shape_Surface4
            MakeSurface4

        Case shape_Surface5
            MakeSurface5

        Case shape_Sphere
            MakeSphere 0, 0, 0, 300, 10, 10
    End Select
End Sub
' Make a sphere.
Private Sub MakeSphere(ByVal Cx As Single, ByVal Cy As Single, ByVal Cz As Single, ByVal radius As Single, ByVal num_horizontal As Integer, ByVal num_vertical As Integer)
Dim pline As Polyline3d
Dim T As Integer
Dim theta1 As Single
Dim theta2 As Single
Dim dtheta As Single
Dim P As Integer
Dim phi1 As Single
Dim phi2 As Single
Dim dphi As Single
Dim x11 As Single   ' xij: theta = i, phi = j
Dim y11 As Single
Dim z11 As Single
Dim x12 As Single
Dim y12 As Single
Dim z12 As Single
Dim x21 As Single
Dim y21 As Single
Dim z21 As Single
Dim x22 As Single
Dim y22 As Single
Dim z22 As Single
Dim R As Single

    Set pline = New Polyline3d
    Polylines.Add pline

    theta1 = 0
    dtheta = 2 * PI / num_horizontal
    For T = 1 To num_horizontal
        theta2 = theta1 + dtheta
        phi1 = -PI / 2
        dphi = PI / num_vertical
        x11 = 0
        y11 = -radius
        z11 = 0
        x21 = 0
        y21 = -radius
        z21 = 0
        For P = 1 To num_vertical
            phi2 = phi1 + dphi

            y12 = radius * Sin(phi2)
            R = radius * Cos(phi2)
            x12 = R * Cos(theta1)
            z12 = R * Sin(theta1)

            y22 = radius * Sin(phi2)
            R = radius * Cos(phi2)
            x22 = R * Cos(theta2)
            z22 = R * Sin(theta2)

            If P = 1 Then
                ' Bottom triangle.
                pline.AddSegment _
                    Cx + x11, Cy + y11, Cz + z11, _
                    Cx + x12, Cy + y12, Cz + z12
                pline.AddSegment _
                    Cx + x11, Cy + y11, Cz + z11, _
                    Cx + x22, Cy + y22, Cz + z22
            ElseIf P = num_vertical Then
                ' Top triangle.
                pline.AddSegment _
                    Cx + x11, Cy + y11, Cz + z11, _
                    Cx + x12, Cy + y12, Cz + z12
                pline.AddSegment _
                    Cx + x11, Cy + y11, Cz + z11, _
                    Cx + x21, Cy + y21, Cz + z21
            Else
                ' Middle rectangle.
                pline.AddSegment _
                    Cx + x11, Cy + y11, Cz + z11, _
                    Cx + x12, Cy + y12, Cz + z12
                pline.AddSegment _
                    Cx + x11, Cy + y11, Cz + z11, _
                    Cx + x21, Cy + y21, Cz + z21
            End If

            x11 = x12
            y11 = y12
            z11 = z12
            x21 = x22
            y21 = y22
            z21 = z22
            phi1 = phi2
        Next P
        theta1 = theta2
    Next T
End Sub

' Display the data.
Private Sub DrawData(ByVal pic As PictureBox)
Dim pline As Polyline3d

    Screen.MousePointer = vbHourglass
    pic.Cls
    DoEvents

    ' Recreate the data.
    CreateData

    ' Build the projection transformation.
    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0

    ' Transform and draw the polylines.
    pic.Cls
    For Each pline In Polylines
        pline.ApplyFull Projector
        pline.Draw pic
    Next pline

    Screen.MousePointer = vbDefault
    pic.Refresh
End Sub
Private Sub optShape_Click(Index As Integer)
    SelectedShape = Index
    DrawData picCanvas
    picCanvas.SetFocus
End Sub


