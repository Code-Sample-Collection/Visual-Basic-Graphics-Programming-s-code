VERSION 5.00
Begin VB.Form frmDepth1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Depth1"
   ClientHeight    =   4065
   ClientLeft      =   1410
   ClientTop       =   540
   ClientWidth     =   6330
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4065
   ScaleWidth      =   6330
   Begin VB.OptionButton optSolid 
      Caption         =   "Sphere"
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   9
      Top             =   3240
      Width           =   2055
   End
   Begin VB.OptionButton optSolid 
      Caption         =   "Stellate Octahedron"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   8
      Top             =   2880
      Width           =   2055
   End
   Begin VB.OptionButton optSolid 
      Caption         =   "Platonic Solids"
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
   End
   Begin VB.OptionButton optSolid 
      Caption         =   "6 Icosahedra"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   6
      Top             =   2160
      Width           =   2055
   End
   Begin VB.OptionButton optSolid 
      Caption         =   "6 Dodecahedra"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Width           =   2055
   End
   Begin VB.OptionButton optSolid 
      Caption         =   "6 Octahedra"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.OptionButton optSolid 
      Caption         =   "8 Cubes"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.OptionButton optSolid 
      Caption         =   "6 Tetrahedra"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.CheckBox chkRemoveBackfaces 
      Caption         =   "Remove Backfaces"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   3975
      Left            =   2160
      ScaleHeight     =   261
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "frmDepth1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Location of viewing eye.
Private EyeR As Single
Private EyeTheta As Single
Private EyePhi As Single

Private Const dtheta = PI / 20
Private Const dphi = PI / 20
Private Const Dr = 1

' Location of focus point.
Private Const FocusX = 0#
Private Const FocusY = 0#
Private Const FocusZ = 0#

Private Projector(1 To 4, 1 To 4) As Single

Private Solids As Collection

Private SelectedShape As Integer
' Sort the solids in depth-sort order.
Private Sub SortSolids()
Dim solid As Solid3d
Dim ordered_solids As Collection
Dim besti As Integer
Dim bestz As Single
Dim newz As Single
Dim i As Integer

    ' Compute each solid's Zmax value.
    For Each solid In Solids
        solid.SetZmax
    Next solid

    ' Sort the objects by their Zmax values.
    Set ordered_solids = New Collection
    Do While Solids.Count > 0
        ' Find the face with the smallest Zmax
        ' left in the Faces collection.
        besti = 1
        bestz = Solids(1).zmax
        For i = 2 To Solids.Count
            newz = Solids(i).zmax
            If bestz > newz Then
                besti = i
                bestz = newz
            End If
        Next i

        ' Add the best object to the sorted list.
        ordered_solids.Add Solids(besti)
        Solids.Remove besti
    Loop

    ' Replace the Solids collection with the
    ' ordered_solids collection.
    Set Solids = ordered_solids
End Sub
' Draw the data.
Private Sub DrawData(ByVal pic As PictureBox)
Dim solid As Solid3d
Dim X As Single
Dim Y As Single
Dim Z As Single
Dim S(1 To 4, 1 To 4) As Single
Dim T(1 To 4, 1 To 4) As Single
Dim ST(1 To 4, 1 To 4) As Single
Dim PST(1 To 4, 1 To 4) As Single

    ' Prevent overflow errors when drawing lines
    ' too far out of bounds.
    On Error Resume Next

    ' Uncull the solids.
    For Each solid In Solids
        solid.Culled = False
    Next solid

    ' Cull backfaces.
    If chkRemoveBackfaces.value = vbChecked Then
        m3SphericalToCartesian EyeR, EyeTheta, EyePhi, X, Y, Z

        For Each solid In Solids
            solid.Culled = False
            solid.Cull X, Y, Z
        Next solid
    End If

    ' Scale and translate so it looks OK in pixels.
    m3Scale S, 100, -100, 1
    m3Translate T, picCanvas.ScaleWidth / 2, picCanvas.ScaleHeight / 2, 0
    m3MatMultiplyFull ST, S, T
    m3MatMultiplyFull PST, Projector, ST

    ' Transform the solids and clip faces.
    For Each solid In Solids
        solid.ApplyFull PST

        ' Clip faces behind the center of projection.
        solid.ClipEye EyeR
    Next solid

    ' Sort the solids if necessary.
    If chkRemoveBackfaces.value = vbChecked Then
        SortSolids
    End If

    ' Set the appropriate fill style.
    If chkRemoveBackfaces.value = vbChecked Then
        ' Fill to cover hidden surfaces.
        pic.FillStyle = vbFSSolid
        pic.FillColor = RGB(&H80, &HFF, &HFF)
    Else
        ' Do not fill so all lines are visible.
        pic.FillStyle = vbFSTransparent
    End If

    ' Draw the solids.
    pic.Cls
    For Each solid In Solids
        solid.Draw pic, EyeR
    Next solid
    pic.Refresh
End Sub
' Redraw the picture with culling changed.
Private Sub chkRemoveBackfaces_Click()
    DrawData picCanvas
    picCanvas.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft
            EyeTheta = EyeTheta - dtheta
        
        Case vbKeyRight
            EyeTheta = EyeTheta + dtheta
        
        Case vbKeyUp
            EyePhi = EyePhi - dphi
        
        Case vbKeyDown
            EyePhi = EyePhi + dphi
                
        Case Else
            Exit Sub
    End Select

    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0
    DrawData picCanvas
End Sub
' Make a sphere.
Private Function Sphere(ByVal Cx As Single, ByVal Cy As Single, ByVal Cz As Single, ByVal radius As Single, ByVal num_horizontal As Integer, ByVal num_vertical As Integer) As Solid3d
Dim new_solid As Solid3d
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

    Set new_solid = New Solid3d
    new_solid.IsConvex = True

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
                new_solid.AddFace _
                    Cx + x11, Cy + y11, Cz + z11, _
                    Cx + x12, Cy + y12, Cz + z12, _
                    Cx + x22, Cy + y22, Cz + z22
            ElseIf P = num_vertical Then
                ' Top triangle.
                new_solid.AddFace _
                    Cx + x11, Cy + y11, Cz + z11, _
                    Cx + x12, Cy + y12, Cz + z12, _
                    Cx + x21, Cy + y21, Cz + z21
            Else
                ' Middle rectangle.
                new_solid.AddFace _
                    Cx + x11, Cy + y11, Cz + z11, _
                    Cx + x12, Cy + y12, Cz + z12, _
                    Cx + x22, Cy + y22, Cz + z22, _
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

    Set Sphere = new_solid
End Function
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("+")
            EyeR = EyeR + Dr
        
        Case Asc("-")
            EyeR = EyeR - Dr
        
        Case Else
            Exit Sub
    End Select

    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0
    DrawData picCanvas
End Sub

Private Sub Form_Load()
    ' Initialize the eye position.
    EyeR = 10
    EyeTheta = PI * 0.2
    EyePhi = PI * 0.05

    ' Initialize the projection transformation.
    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0

    ' Start with the tetrahedron.
    Show
    optSolid(0).value = True
End Sub

' Create the data.
Private Sub CreateData()
    ' Create the new Solids collection.
    Set Solids = New Collection

    ' Create the solids.
    Select Case SelectedShape
        Case 0  ' Tetrahedra.
            Solids.Add Tetrahedron(0.75, 0.5 + 0, 0, 0.4)
            Solids.Add Tetrahedron(0, 0.5 + 0.75, 0, 0.4)
            Solids.Add Tetrahedron(0, 0.5 + 0, 0.75, 0.4)
            Solids.Add Tetrahedron(-0.75, 0.5 + 0, 0, 0.4)
            Solids.Add Tetrahedron(0, 0.5 + -0.75, 0, 0.4)
            Solids.Add Tetrahedron(0, 0.5 + 0, -0.75, 0.4)

        Case 1  ' Cubes.
            Solids.Add Cube(0.5, 0.5, 0.5, 0.4)
            Solids.Add Cube(0.5, 0.5, -0.5, 0.4)
            Solids.Add Cube(0.5, -0.5, 0.5, 0.4)
            Solids.Add Cube(-0.5, 0.5, 0.5, 0.4)
            Solids.Add Cube(0.5, -0.5, -0.5, 0.4)
            Solids.Add Cube(-0.5, 0.5, -0.5, 0.4)
            Solids.Add Cube(-0.5, -0.5, 0.5, 0.4)
            Solids.Add Cube(-0.5, -0.5, -0.5, 0.4)

        Case 2  ' Octahedra.
            Solids.Add Octahedron(0.75, 0, 0, 0.4)
            Solids.Add Octahedron(0, 0.75, 0, 0.4)
            Solids.Add Octahedron(0, 0, 0.75, 0.4)
            Solids.Add Octahedron(-0.75, 0, 0, 0.4)
            Solids.Add Octahedron(0, -0.75, 0, 0.4)
            Solids.Add Octahedron(0, 0, -0.75, 0.4)

        Case 3  ' Dodecahedra.
            Solids.Add Dodecahedron(0.75, 0, 0, 0.3)
            Solids.Add Dodecahedron(0, 0.75, 0, 0.3)
            Solids.Add Dodecahedron(0, 0, 0.75, 0.3)
            Solids.Add Dodecahedron(-0.75, 0, 0, 0.3)
            Solids.Add Dodecahedron(0, -0.75, 0, 0.3)
            Solids.Add Dodecahedron(0, 0, -0.75, 0.3)

        Case 4  ' Icosahedra.
            Solids.Add Icosahedron(0.75, 0, 0, 0.4)
            Solids.Add Icosahedron(0, 0.75, 0, 0.4)
            Solids.Add Icosahedron(0, 0, 0.75, 0.4)
            Solids.Add Icosahedron(-0.75, 0, 0, 0.4)
            Solids.Add Icosahedron(0, -0.75, 0, 0.4)
            Solids.Add Icosahedron(0, 0, -0.75, 0.4)

        Case 5  ' Platonic solids.
            Solids.Add Tetrahedron(0, 0.6 + 0.75, 0, 0.4)
            Solids.Add Cube(0.75, 0, 0, 0.6)
            Solids.Add Octahedron(0, 0, 0.75, 0.5)
            Solids.Add Dodecahedron(-0.75, 0, 0, 0.4)
            Solids.Add Icosahedron(0, 0, -0.75, 0.5)

        Case 6  ' Stellate octahedron.
            MakeStellate8 0.75

        Case 7  ' Sphere.
            Solids.Add Sphere(0, 0, 0, 1, 10, 10)

    End Select
End Sub
' Make a stellate octahedron.
Private Sub MakeStellate8(ByVal side_scale As Single)
Dim new_solid As Solid3d

    Set new_solid = New Solid3d
    Solids.Add new_solid
    new_solid.IsConvex = True
    new_solid.Stellate side_scale, _
        0, side_scale, 0, _
        0, 0, side_scale, _
        side_scale, 0, 0

    Set new_solid = New Solid3d
    Solids.Add new_solid
    new_solid.IsConvex = True
    new_solid.Stellate side_scale, _
        0, side_scale, 0, _
        side_scale, 0, 0, _
        0, 0, -side_scale

    Set new_solid = New Solid3d
    new_solid.IsConvex = True
    Solids.Add new_solid
    new_solid.Stellate side_scale, _
        0, side_scale, 0, _
        0, 0, -side_scale, _
        -side_scale, 0, 0

    Set new_solid = New Solid3d
    Solids.Add new_solid
    new_solid.IsConvex = True
    new_solid.Stellate side_scale, _
        0, side_scale, 0, _
        -side_scale, 0, 0, _
        0, 0, side_scale

    Set new_solid = New Solid3d
    Solids.Add new_solid
    new_solid.IsConvex = True
    new_solid.Stellate side_scale, _
        0, -side_scale, 0, _
        side_scale, 0, 0, _
        0, 0, side_scale

    Set new_solid = New Solid3d
    Solids.Add new_solid
    new_solid.IsConvex = True
    new_solid.Stellate side_scale, _
        0, -side_scale, 0, _
        0, 0, -side_scale, _
        side_scale, 0, 0

    Set new_solid = New Solid3d
    Solids.Add new_solid
    new_solid.IsConvex = True
    new_solid.Stellate side_scale, _
        0, -side_scale, 0, _
        -side_scale, 0, 0, _
        0, 0, -side_scale

    Set new_solid = New Solid3d
    Solids.Add new_solid
    new_solid.IsConvex = True
    new_solid.Stellate side_scale, _
        0, -side_scale, 0, _
        0, 0, side_scale, _
        -side_scale, 0, 0
End Sub
' Make a dodecahedron.
Private Function Dodecahedron(ByVal Cx As Single, ByVal Cy As Single, ByVal Cz As Single, ByVal side_scale As Single) As Solid3d
Dim new_solid As Solid3d
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

    Set new_solid = New Solid3d
    new_solid.IsConvex = True

    new_solid.AddFace _
        Cx + C, Cy + side_scale, Cz + -A, _
        Cx + D, Cy + side_scale, Cz + -B, _
        Cx + D, Cy + side_scale, Cz + B, _
        Cx + C, Cy + side_scale, Cz + A, _
        Cx + R, Cy + side_scale, Cz + 0
    new_solid.AddFace _
        Cx + C, Cy + side_scale, Cz + A, _
        Cx + X * c1, Cy + side_scale - Y, Cz + X * s1, _
        Cx + -X * c2, Cy + side_scale - y2, Cz + X * s2, _
        Cx + X, Cy + side_scale - Y, Cz + 0, _
        Cx + R, Cy + side_scale, Cz + 0
    new_solid.AddFace _
        Cx + C, Cy + side_scale, Cz + A, _
        Cx + D, Cy + side_scale, Cz + B, _
        Cx + X * c2, Cy + side_scale - Y, Cz + X * s2, _
        Cx + -X * c1, Cy + side_scale - y2, Cz + X * s1, _
        Cx + X * c1, Cy + side_scale - Y, Cz + X * s1
    new_solid.AddFace _
        Cx + D, Cy + side_scale, Cz + B, _
        Cx + D, Cy + side_scale, Cz + -B, _
        Cx + X * c2, Cy + side_scale - Y, Cz + -X * s2, _
        Cx + -X, Cy + side_scale - y2, Cz + 0, _
        Cx + X * c2, Cy + side_scale - Y, Cz + X * s2
    new_solid.AddFace _
        Cx + D, Cy + side_scale, Cz + -B, _
        Cx + C, Cy + side_scale, Cz + -A, _
        Cx + X * c1, Cy + side_scale - Y, Cz + -X * s1, _
        Cx + -X * c1, Cy + side_scale - y2, Cz + -X * s1, _
        Cx + X * c2, Cy + side_scale - Y, Cz + -X * s2, -X * c1
    new_solid.AddFace _
        Cx + C, Cy + side_scale, Cz + -A, _
        Cx + R, Cy + side_scale, Cz + 0, _
        Cx + X, Cy + side_scale - Y, Cz + 0, _
        Cx + -X * c2, Cy + side_scale - y2, Cz + -X * s2, _
        Cx + X * c1, Cy + side_scale - Y, Cz + -X * s1

    ' Bottom.
    new_solid.AddFace _
        Cx + -D, Cy + -side_scale, Cz + -B, _
        Cx + -X * c2, Cy + side_scale - y2, Cz + -X * s2, _
        Cx + X, Cy + side_scale - Y, Cz + 0, _
        Cx + -X * c2, Cy + side_scale - y2, Cz + X * s2, _
        Cx + -D, Cy + -side_scale, Cz + B
    new_solid.AddFace _
        Cx + -D, Cy + -side_scale, Cz + B, _
        Cx + -X * c2, Cy + side_scale - y2, Cz + X * s2, _
        Cx + X * c1, Cy + side_scale - Y, Cz + X * s1, _
        Cx + -X * c1, Cy + side_scale - y2, Cz + X * s1, _
        Cx + -C, Cy + -side_scale, Cz + A
    new_solid.AddFace _
        Cx + -C, Cy + -side_scale, Cz + A, _
        Cx + -X * c1, Cy + side_scale - y2, Cz + X * s1, _
        Cx + X * c2, Cy + side_scale - Y, Cz + X * s2, _
        Cx + -X, Cy + side_scale - y2, Cz + 0, _
        Cx + -R, Cy + -side_scale, Cz + 0
    new_solid.AddFace _
        Cx + -R, Cy + -side_scale, Cz + 0, _
        Cx + -X, Cy + side_scale - y2, Cz + 0, _
        Cx + X * c2, Cy + side_scale - Y, Cz + -X * s2, _
        Cx + -X * c1, Cy + side_scale - y2, Cz + -X * s1, _
        Cx + -C, Cy + -side_scale, Cz + -A
    new_solid.AddFace _
        Cx + -C, Cy + -side_scale, Cz + -A, _
        Cx + -X * c1, Cy + side_scale - y2, Cz + -X * s1, _
        Cx + X * c1, Cy + side_scale - Y, Cz + -X * s1, _
        Cx + -X * c2, Cy + side_scale - y2, Cz + -X * s2, _
        Cx + -D, Cy + -side_scale, Cz + -B
    new_solid.AddFace _
        Cx + -D, Cy + -side_scale, Cz + -B, _
        Cx + -D, Cy + -side_scale, Cz + B, _
        Cx + -C, Cy + -side_scale, Cz + A, _
        Cx + -R, Cy + -side_scale, Cz + 0, _
        Cx + -C, Cy + -side_scale, Cz + -A

    Set Dodecahedron = new_solid
End Function

' Make an icosahedron.
Private Function Icosahedron(ByVal Cx As Single, ByVal Cy As Single, ByVal Cz As Single, ByVal side_scale As Single) As Solid3d
Dim new_solid As Solid3d
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

    theta1 = PI * 0.4
    theta2 = PI * 0.8
    s1 = Sin(theta1)
    c1 = Cos(theta1)
    s2 = Sin(theta2)
    c2 = Cos(theta2)
    R = 2 / (2 * Sqr(1 - 2 * c1) + Sqr(3 / 4 * (2 - 2 * c1) - 2 * c2 - c2 * c2 - 1)) * side_scale
    S = R * Sqr(2 - 2 * c1)
    H = side_scale - Sqr(S * S - R * R)
    A = R * s1
    B = R * s2
    C = R * c1
    D = R * c2

    ' Top.
    Set new_solid = New Solid3d
    new_solid.IsConvex = True

    new_solid.AddFace _
        Cx + 0, Cy + side_scale, 0 + Cz, _
        Cx + C, Cy + H, A + Cz, _
        Cx + R, Cy + H, 0 + Cz
    new_solid.AddFace _
        Cx + 0, Cy + side_scale, 0 + Cz, _
        Cx + R, Cy + H, 0 + Cz, _
        Cx + C, Cy + H, -A + Cz
    new_solid.AddFace _
        Cx + 0, Cy + side_scale, 0 + Cz, _
        Cx + C, Cy + H, -A + Cz, _
        Cx + D, Cy + H, -B + Cz
    new_solid.AddFace _
        Cx + 0, Cy + side_scale, 0 + Cz, _
        Cx + D, Cy + H, -B + Cz, _
        Cx + D, Cy + H, B + Cz
    new_solid.AddFace _
        Cx + 0, Cy + side_scale, 0 + Cz, _
        Cx + D, Cy + H, B + Cz, _
        Cx + C, Cy + H, A + Cz

    ' Upper Middle.
    new_solid.AddFace _
        Cx + R, Cy + H, 0 + Cz, _
        Cx + C, Cy + H, A + Cz, _
        Cx + -D, Cy + -H, B + Cz
    new_solid.AddFace _
        Cx + C, Cy + H, A + Cz, _
        Cx + D, Cy + H, B + Cz, _
        Cx + -C, Cy + -H, A + Cz
    new_solid.AddFace _
        Cx + D, Cy + H, B + Cz, _
        Cx + D, Cy + H, -B + Cz, _
        Cx + -R, Cy + -H, 0 + Cz
    new_solid.AddFace _
        Cx + D, Cy + H, -B + Cz, _
        Cx + C, Cy + H, -A + Cz, _
        Cx + -C, Cy + -H, -A + Cz
    new_solid.AddFace _
        Cx + C, Cy + H, -A + Cz, _
        Cx + R, Cy + H, 0 + Cz, _
        Cx + -D, Cy + -H, -B + Cz

    ' Lower Middle.
    new_solid.AddFace _
        Cx + R, Cy + H, 0 + Cz, _
        Cx + -D, Cy + -H, B + Cz, _
        Cx + -D, Cy + -H, -B + Cz
    new_solid.AddFace _
        Cx + C, Cy + H, A + Cz, _
        Cx + -C, Cy + -H, A + Cz, _
        Cx + -D, Cy + -H, B + Cz
    new_solid.AddFace _
        Cx + D, Cy + H, B + Cz, _
        Cx + -R, Cy + -H, 0 + Cz, _
        Cx + -C, Cy + -H, A + Cz
    new_solid.AddFace _
        Cx + D, Cy + H, -B + Cz, _
        Cx + -C, Cy + -H, -A + Cz, _
        Cx + -R, Cy + -H, 0 + Cz
    new_solid.AddFace _
        Cx + C, Cy + H, -A + Cz, _
        Cx + -D, Cy + -H, -B + Cz, _
        Cx + -C, Cy + -H, -A + Cz

    ' Bottom.
    new_solid.AddFace _
        Cx + 0, Cy + -side_scale, 0 + Cz, _
        Cx + -D, Cy + -H, B + Cz, _
        Cx + -C, Cy + -H, A + Cz
    new_solid.AddFace _
        Cx + 0, Cy + -side_scale, 0 + Cz, _
        Cx + -C, Cy + -H, A + Cz, _
        Cx + -R, Cy + -H, 0 + Cz
    new_solid.AddFace _
        Cx + 0, Cy + -side_scale, 0 + Cz, _
        Cx + -R, Cy + -H, 0 + Cz, _
        Cx + -C, Cy + -H, -A + Cz
    new_solid.AddFace _
        Cx + 0, Cy + -side_scale, 0 + Cz, _
        Cx + -C, Cy + -H, -A + Cz, _
        Cx + -D, Cy + -H, -B + Cz
    new_solid.AddFace _
        Cx + 0, Cy + -side_scale, 0 + Cz, _
        Cx + -D, Cy + -H, -B + Cz, _
        Cx + -D, Cy + -H, B + Cz

    Set Icosahedron = new_solid
End Function
' Make an octahedron.
Private Function Octahedron(ByVal Cx As Single, ByVal Cy As Single, ByVal Cz As Single, ByVal side_scale As Single) As Solid3d
Dim new_solid As Solid3d

    ' Top.
    Set new_solid = New Solid3d
    new_solid.IsConvex = True

    new_solid.AddFace _
        Cx + 0, Cy + side_scale, 0 + Cz, _
        Cx + side_scale, Cy + 0, 0 + Cz, _
        Cx + 0, Cy + 0, -side_scale + Cz
    new_solid.AddFace _
        Cx + 0, Cy + side_scale, 0 + Cz, _
        Cx + 0, Cy + 0, -side_scale + Cz, _
        Cx + -side_scale, Cy + 0, 0 + Cz
    new_solid.AddFace _
        Cx + 0, Cy + side_scale, 0 + Cz, _
        Cx + -side_scale, Cy + 0, 0 + Cz, _
        Cx + 0, Cy + 0, side_scale + Cz
    new_solid.AddFace _
        Cx + 0, Cy + side_scale, 0 + Cz, _
        Cx + 0, Cy + 0, side_scale + Cz, _
        Cx + side_scale, Cy + 0, 0 + Cz

    ' Bottom.
    new_solid.AddFace _
        Cx + 0, Cy + -side_scale, 0 + Cz, _
        Cx + side_scale, Cy + 0, 0 + Cz, _
        Cx + 0, Cy + 0, side_scale + Cz
    new_solid.AddFace _
        Cx + 0, Cy + -side_scale, 0 + Cz, _
        Cx + 0, Cy + 0, side_scale + Cz, _
        Cx + -side_scale, Cy + 0, 0 + Cz
    new_solid.AddFace _
        Cx + 0, Cy + -side_scale, 0 + Cz, _
        Cx + -side_scale, Cy + 0, 0 + Cz, _
        Cx + 0, Cy + 0, -side_scale + Cz
    new_solid.AddFace _
        Cx + 0, Cy + -side_scale, 0 + Cz, _
        Cx + 0, Cy + 0, -side_scale + Cz, _
        Cx + side_scale, Cy + 0, 0 + Cz

    Set Octahedron = new_solid
End Function

' Make a cube with the indicated center and
' side length.
Private Function Cube(ByVal Cx As Single, ByVal Cy As Single, ByVal Cz As Single, ByVal side_scale As Single) As Solid3d
Dim new_solid As Solid3d
Dim s2 As Single

    s2 = side_scale / 2
    Set new_solid = New Solid3d
    new_solid.IsConvex = True

    ' Top.
    new_solid.AddFace _
        Cx + s2, Cy + s2, Cz + s2, _
        Cx + s2, Cy + s2, Cz - s2, _
        Cx - s2, Cy + s2, Cz - s2, _
        Cx - s2, Cy + s2, Cz + s2
    ' Positive X side.
    new_solid.AddFace _
        Cx + s2, Cy + s2, Cz + s2, _
        Cx + s2, Cy - s2, Cz + s2, _
        Cx + s2, Cy - s2, Cz - s2, _
        Cx + s2, Cy + s2, Cz - s2
    ' Positive Z side.
    new_solid.AddFace _
        Cx + s2, Cy + s2, Cz + s2, _
        Cx - s2, Cy + s2, Cz + s2, _
        Cx - s2, Cy - s2, Cz + s2, _
        Cx + s2, Cy - s2, Cz + s2
    ' Negative X side.
    new_solid.AddFace _
        Cx - s2, Cy - s2, Cz - s2, _
        Cx - s2, Cy - s2, Cz + s2, _
        Cx - s2, Cy + s2, Cz + s2, _
        Cx - s2, Cy + s2, Cz - s2
    ' Negative Z side.
    new_solid.AddFace _
        Cx - s2, Cy - s2, Cz - s2, _
        Cx - s2, Cy + s2, Cz - s2, _
        Cx + s2, Cy + s2, Cz - s2, _
        Cx + s2, Cy - s2, Cz - s2
    ' Bottom.
    new_solid.AddFace _
        Cx - s2, Cy - s2, Cz - s2, _
        Cx + s2, Cy - s2, Cz - s2, _
        Cx + s2, Cy - s2, Cz + s2, _
        Cx - s2, Cy - s2, Cz + s2

    Set Cube = new_solid
End Function
' Make a tetrahedron.
Private Function Tetrahedron(ByVal Cx As Single, ByVal Cy As Single, ByVal Cz As Single, ByVal side_scale As Single) As Solid3d
Dim new_solid As Solid3d
Dim S As Single
Dim A As Single
Dim B As Single
Dim C As Single
Dim D As Single

    S = Sqr(6) * side_scale
    A = S / Sqr(3)
    B = -A / 2
    C = A * Sqr(2) - 1
    D = S / 2

    Set new_solid = New Solid3d
    new_solid.IsConvex = True

    new_solid.AddFace _
        Cx + 0, Cy + C, 0 + Cz, _
        Cx + A, Cy + -1, 0 + Cz, _
        Cx + B, Cy + -1, -D + Cz
    new_solid.AddFace _
        Cx + 0, Cy + C, 0 + Cz, _
        Cx + B, Cy + -1, -D + Cz, _
        Cx + B, Cy + -1, D + Cz
    new_solid.AddFace _
        Cx + 0, Cy + C, 0 + Cz, _
        Cx + B, Cy + -1, D + Cz, _
        Cx + A, Cy + -1, 0 + Cz
    new_solid.AddFace _
        Cx + A, Cy + -1, 0 + Cz, _
        Cx + B, Cy + -1, D + Cz, _
        Cx + B, Cy + -1, -D + Cz

    Set Tetrahedron = new_solid
End Function
' Make the drawing areas as large as possible.
Private Sub Form_Resize()
Dim wid As Single

    wid = ScaleWidth - picCanvas.Left
    If wid < 120 Then wid = 120
    picCanvas.Move picCanvas.Left, 0, wid, ScaleHeight
End Sub

Private Sub optSolid_Click(Index As Integer)
    SelectedShape = Index
    CreateData
    DrawData picCanvas
    picCanvas.SetFocus
End Sub

