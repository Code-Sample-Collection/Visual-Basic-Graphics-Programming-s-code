VERSION 5.00
Begin VB.Form frmBackface 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Backface"
   ClientHeight    =   4320
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
   ScaleHeight     =   4320
   ScaleWidth      =   6330
   Begin VB.OptionButton optSolid 
      Caption         =   "Sphere"
      Height          =   255
      Index           =   9
      Left            =   0
      TabIndex        =   11
      Top             =   3960
      Width           =   2055
   End
   Begin VB.OptionButton optSolid 
      Caption         =   "8 Cubes"
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   10
      Top             =   3240
      Width           =   2055
   End
   Begin VB.OptionButton optSolid 
      Caption         =   "Stellate Octahedron"
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   9
      Top             =   3600
      Width           =   2055
   End
   Begin VB.OptionButton optSolid 
      Caption         =   "Figure 15.4b"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   8
      Top             =   2880
      Width           =   2055
   End
   Begin VB.OptionButton optSolid 
      Caption         =   "Figure 15.4a"
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
   End
   Begin VB.OptionButton optSolid 
      Caption         =   "Icosahedron"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   6
      Top             =   2160
      Width           =   2055
   End
   Begin VB.OptionButton optSolid 
      Caption         =   "Dodecahedron"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Width           =   2055
   End
   Begin VB.OptionButton optSolid 
      Caption         =   "Octahredon"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.OptionButton optSolid 
      Caption         =   "Cube"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.OptionButton optSolid 
      Caption         =   "Tetrahedron"
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
      Height          =   4215
      Left            =   2160
      ScaleHeight     =   277
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "frmBackface"
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

Private TheSolid As Solid3d

Private SelectedShape As Integer
' Make a sphere.
Private Sub MakeSphere(ByVal Cx As Single, ByVal Cy As Single, ByVal Cz As Single, ByVal radius As Single, ByVal num_horizontal As Integer, ByVal num_vertical As Integer)
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
                TheSolid.AddFace _
                    Cx + x11, Cy + y11, Cz + z11, _
                    Cx + x12, Cy + y12, Cz + z12, _
                    Cx + x22, Cy + y22, Cz + z22
            ElseIf P = num_vertical Then
                ' Top triangle.
                TheSolid.AddFace _
                    Cx + x11, Cy + y11, Cz + z11, _
                    Cx + x12, Cy + y12, Cz + z12, _
                    Cx + x21, Cy + y21, Cz + z21
            Else
                ' Middle rectangle.
                TheSolid.AddFace _
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
End Sub

' Draw the data.
Private Sub DrawData(ByVal pic As PictureBox)
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

    ' Cull backfaces.
    TheSolid.Culled = False
    If chkRemoveBackfaces.value = vbChecked Then
        m3SphericalToCartesian EyeR, EyeTheta, EyePhi, X, Y, Z
        TheSolid.Cull X, Y, Z
    End If

    ' Scale and translate so it looks OK in pixels.
    m3Scale S, 100, -100, 1
    m3Translate T, picCanvas.ScaleWidth / 2, picCanvas.ScaleHeight / 2, 0
    m3MatMultiplyFull ST, S, T
    m3MatMultiplyFull PST, Projector, ST

    ' Transform the points.
    TheSolid.ApplyFull PST

    ' Clip faces behind the center of projection.
    TheSolid.ClipEye EyeR

    ' Display the data.
    pic.Cls
    TheSolid.Draw pic, EyeR
    pic.Refresh
End Sub
' Make a solid like the one shown in Figure 15.4b.
Private Sub MakeFig15_4b()
Const S = 0.75

    TheSolid.AddFace _
        S, S, S, _
        -S, S, S, _
        -S, -S, S, _
        S, -S, S
    TheSolid.AddFace _
        S, S, S, _
        S, -S, S, _
        S, -S, -S, _
        S, S, -S
    TheSolid.AddFace _
        S, S, -S, _
        S, -S, -S, _
        -S, -S, -S, _
        -S, S, -S
    TheSolid.AddFace _
        S, S, -S, _
        -S, S, -S, _
        0, S, 0, _
        -S, S, S, _
        S, S, S
    TheSolid.AddFace _
        S, -S, S, _
        -S, -S, S, _
        0, -S, 0, _
        -S, -S, -S, _
        S, -S, -S
    TheSolid.AddFace _
        -S, S, -S, _
        -S, -S, -S, _
        0, -S, 0, _
        0, S, 0
    TheSolid.AddFace _
        0, S, 0, _
        0, -S, 0, _
        -S, -S, S, _
        -S, S, S
End Sub
' Make a solid like the one shown in Figure 15.4a.
Private Sub MakeFig15_4a()
Const S = 0.75

    TheSolid.AddFace _
        S, S, 0, _
        S, S, -S, _
        -S, S, -S, _
        -S, S, S, _
        0, S, S
    TheSolid.AddFace _
        S, S, 0, _
        0, S, S, _
        S, 0, S
    TheSolid.AddFace _
        S, S, -S, _
        S, S, 0, _
        S, 0, S, _
        S, -S, S, _
        S, -S, -S
    TheSolid.AddFace _
        S, S, -S, _
        S, -S, -S, _
        -S, -S, -S, _
        -S, S, -S
    TheSolid.AddFace _
        -S, S, -S, _
        -S, -S, -S, _
        -S, -S, 0, _
        -S, 0, S, _
        -S, S, S
    TheSolid.AddFace _
        -S, S, S, _
        -S, 0, S, _
        0, -S, S, _
        S, -S, S, _
        S, 0, S, _
        0, S, S
    TheSolid.AddFace _
        -S, 0, S, _
        -S, -S, 0, _
        0, -S, S
    TheSolid.AddFace _
        S, -S, S, _
        0, -S, S, _
        -S, -S, 0, _
        -S, -S, -S, _
        S, -S, -S
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
    optSolid(5).value = True
End Sub

' Create the data.
Private Sub CreateData()
    ' Create the new solid.
    Set TheSolid = New Solid3d

    ' Create the solid.
    Select Case SelectedShape
        Case 0  ' Tetrahedron.
            MakeTetrahedron 0.75

        Case 1  ' Cube.
            MakeCube 0, 0, 0, 1

        Case 2  ' Octahedron.
            MakeOctahedron 1

        Case 3  ' Dodecahedron.
            MakeDodecahedron 1

        Case 4  ' Icosahedron.
            MakeIcosahedron 1

        Case 5  ' Figure 15.4a.
            MakeFig15_4a

        Case 6  ' Figure 15.4b.
            MakeFig15_4b

        Case 7  ' 8 Cubes.
            MakeCube 0.5, 0.5, 0.5, 0.4
            MakeCube 0.5, 0.5, -0.5, 0.4
            MakeCube 0.5, -0.5, 0.5, 0.4
            MakeCube -0.5, 0.5, 0.5, 0.4
            MakeCube 0.5, -0.5, -0.5, 0.4
            MakeCube -0.5, 0.5, -0.5, 0.4
            MakeCube -0.5, -0.5, 0.5, 0.4
            MakeCube -0.5, -0.5, -0.5, 0.4

        Case 8  ' Stellate octahedron.
            MakeStellate8 0.75

        Case 9  ' Sphere.
            MakeSphere 0, 0, 0, 1, 10, 10

    End Select
End Sub
' Make a stellate octahedron.
Private Sub MakeStellate8(ByVal side_scale As Single)
    TheSolid.Stellate side_scale, _
        0, side_scale, 0, _
        0, 0, side_scale, _
        side_scale, 0, 0
    TheSolid.Stellate side_scale, _
        0, side_scale, 0, _
        side_scale, 0, 0, _
        0, 0, -side_scale
    TheSolid.Stellate side_scale, _
        0, side_scale, 0, _
        0, 0, -side_scale, _
        -side_scale, 0, 0
    TheSolid.Stellate side_scale, _
        0, side_scale, 0, _
        -side_scale, 0, 0, _
        0, 0, side_scale

    TheSolid.Stellate side_scale, _
        0, -side_scale, 0, _
        side_scale, 0, 0, _
        0, 0, side_scale
    TheSolid.Stellate side_scale, _
        0, -side_scale, 0, _
        0, 0, -side_scale, _
        side_scale, 0, 0
    TheSolid.Stellate side_scale, _
        0, -side_scale, 0, _
        -side_scale, 0, 0, _
        0, 0, -side_scale
    TheSolid.Stellate side_scale, _
        0, -side_scale, 0, _
        0, 0, side_scale, _
        -side_scale, 0, 0
End Sub

' Make a dodecahedron.
Private Sub MakeDodecahedron(ByVal side_scale As Single)
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

    TheSolid.AddFace _
        C, side_scale, -A, _
        D, side_scale, -B, _
        D, side_scale, B, _
        C, side_scale, A, _
        R, side_scale, 0
    TheSolid.AddFace _
        C, side_scale, A, _
        X * c1, side_scale - Y, X * s1, _
        -X * c2, side_scale - y2, X * s2, _
        X, side_scale - Y, 0, _
        R, side_scale, 0
    TheSolid.AddFace _
        C, side_scale, A, _
        D, side_scale, B, _
        X * c2, side_scale - Y, X * s2, _
        -X * c1, side_scale - y2, X * s1, _
        X * c1, side_scale - Y, X * s1
    TheSolid.AddFace _
        D, side_scale, B, _
        D, side_scale, -B, _
        X * c2, side_scale - Y, -X * s2, _
        -X, side_scale - y2, 0, _
        X * c2, side_scale - Y, X * s2
    TheSolid.AddFace _
        D, side_scale, -B, _
        C, side_scale, -A, _
        X * c1, side_scale - Y, -X * s1, _
        -X * c1, side_scale - y2, -X * s1, _
        X * c2, side_scale - Y, -X * s2, -X * c1
    TheSolid.AddFace _
        C, side_scale, -A, _
        R, side_scale, 0, _
        X, side_scale - Y, 0, _
        -X * c2, side_scale - y2, -X * s2, _
        X * c1, side_scale - Y, -X * s1

    ' Bottom.
    TheSolid.AddFace _
        -D, -side_scale, -B, _
        -X * c2, side_scale - y2, -X * s2, _
        X, side_scale - Y, 0, _
        -X * c2, side_scale - y2, X * s2, _
        -D, -side_scale, B
    TheSolid.AddFace _
        -D, -side_scale, B, _
        -X * c2, side_scale - y2, X * s2, _
        X * c1, side_scale - Y, X * s1, _
        -X * c1, side_scale - y2, X * s1, _
        -C, -side_scale, A
    TheSolid.AddFace _
        -C, -side_scale, A, _
        -X * c1, side_scale - y2, X * s1, _
        X * c2, side_scale - Y, X * s2, _
        -X, side_scale - y2, 0, _
        -R, -side_scale, 0
    TheSolid.AddFace _
        -R, -side_scale, 0, _
        -X, side_scale - y2, 0, _
        X * c2, side_scale - Y, -X * s2, _
        -X * c1, side_scale - y2, -X * s1, _
        -C, -side_scale, -A
    TheSolid.AddFace _
        -C, -side_scale, -A, _
        -X * c1, side_scale - y2, -X * s1, _
        X * c1, side_scale - Y, -X * s1, _
        -X * c2, side_scale - y2, -X * s2, _
        -D, -side_scale, -B
    TheSolid.AddFace _
        -D, -side_scale, -B, _
        -D, -side_scale, B, _
        -C, -side_scale, A, _
        -R, -side_scale, 0, _
        -C, -side_scale, -A
End Sub

' Make an icosahedron.
Private Sub MakeIcosahedron(ByVal side_scale As Single)
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
    TheSolid.AddFace _
        0, side_scale, 0, _
        C, H, A, _
        R, H, 0
    TheSolid.AddFace _
        0, side_scale, 0, _
        R, H, 0, _
        C, H, -A
    TheSolid.AddFace _
        0, side_scale, 0, _
        C, H, -A, _
        D, H, -B
    TheSolid.AddFace _
        0, side_scale, 0, _
        D, H, -B, _
        D, H, B
    TheSolid.AddFace _
        0, side_scale, 0, _
        D, H, B, _
        C, H, A

    ' Upper Middle.
    TheSolid.AddFace _
        R, H, 0, _
        C, H, A, _
        -D, -H, B
    TheSolid.AddFace _
        C, H, A, _
        D, H, B, _
        -C, -H, A
    TheSolid.AddFace _
        D, H, B, _
        D, H, -B, _
        -R, -H, 0
    TheSolid.AddFace _
        D, H, -B, _
        C, H, -A, _
        -C, -H, -A
    TheSolid.AddFace _
        C, H, -A, _
        R, H, 0, _
        -D, -H, -B

    ' Lower Middle.
    TheSolid.AddFace _
        R, H, 0, _
        -D, -H, B, _
        -D, -H, -B
    TheSolid.AddFace _
        C, H, A, _
        -C, -H, A, _
        -D, -H, B
    TheSolid.AddFace _
        D, H, B, _
        -R, -H, 0, _
        -C, -H, A
    TheSolid.AddFace _
        D, H, -B, _
        -C, -H, -A, _
        -R, -H, 0
    TheSolid.AddFace _
        C, H, -A, _
        -D, -H, -B, _
        -C, -H, -A

    ' Bottom.
    TheSolid.AddFace _
        0, -side_scale, 0, _
        -D, -H, B, _
        -C, -H, A
    TheSolid.AddFace _
        0, -side_scale, 0, _
        -C, -H, A, _
        -R, -H, 0
    TheSolid.AddFace _
        0, -side_scale, 0, _
        -R, -H, 0, _
        -C, -H, -A
    TheSolid.AddFace _
        0, -side_scale, 0, _
        -C, -H, -A, _
        -D, -H, -B
    TheSolid.AddFace _
        0, -side_scale, 0, _
        -D, -H, -B, _
        -D, -H, B
End Sub
' Make an octahedron.
Private Sub MakeOctahedron(ByVal side_scale As Single)
    ' Top.
    TheSolid.AddFace _
        0, side_scale, 0, _
        side_scale, 0, 0, _
        0, 0, -side_scale
    TheSolid.AddFace _
        0, side_scale, 0, _
        0, 0, -side_scale, _
        -side_scale, 0, 0
    TheSolid.AddFace _
        0, side_scale, 0, _
        -side_scale, 0, 0, _
        0, 0, side_scale
    TheSolid.AddFace _
        0, side_scale, 0, _
        0, 0, side_scale, _
        side_scale, 0, 0

    ' Bottom.
    TheSolid.AddFace _
        0, -side_scale, 0, _
        side_scale, 0, 0, _
        0, 0, side_scale
    TheSolid.AddFace _
        0, -side_scale, 0, _
        0, 0, side_scale, _
        -side_scale, 0, 0
    TheSolid.AddFace _
        0, -side_scale, 0, _
        -side_scale, 0, 0, _
        0, 0, -side_scale
    TheSolid.AddFace _
        0, -side_scale, 0, _
        0, 0, -side_scale, _
        side_scale, 0, 0
End Sub

' Make a cube with the indicated center and
' side length.
Private Sub MakeCube(ByVal Cx As Single, ByVal Cy As Single, ByVal Cz As Single, ByVal side_length As Single)
Dim s2 As Single

    s2 = side_length / 2

    ' Top.
    TheSolid.AddFace _
        Cx + s2, Cy + s2, Cz + s2, _
        Cx + s2, Cy + s2, Cz - s2, _
        Cx - s2, Cy + s2, Cz - s2, _
        Cx - s2, Cy + s2, Cz + s2
    ' Positive X side.
    TheSolid.AddFace _
        Cx + s2, Cy + s2, Cz + s2, _
        Cx + s2, Cy - s2, Cz + s2, _
        Cx + s2, Cy - s2, Cz - s2, _
        Cx + s2, Cy + s2, Cz - s2
    ' Positive Z side.
    TheSolid.AddFace _
        Cx + s2, Cy + s2, Cz + s2, _
        Cx - s2, Cy + s2, Cz + s2, _
        Cx - s2, Cy - s2, Cz + s2, _
        Cx + s2, Cy - s2, Cz + s2
    ' Negative X side.
    TheSolid.AddFace _
        Cx - s2, Cy - s2, Cz - s2, _
        Cx - s2, Cy - s2, Cz + s2, _
        Cx - s2, Cy + s2, Cz + s2, _
        Cx - s2, Cy + s2, Cz - s2
    ' Negative Z side.
    TheSolid.AddFace _
        Cx - s2, Cy - s2, Cz - s2, _
        Cx - s2, Cy + s2, Cz - s2, _
        Cx + s2, Cy + s2, Cz - s2, _
        Cx + s2, Cy - s2, Cz - s2
    ' Bottom.
    TheSolid.AddFace _
        Cx - s2, Cy - s2, Cz - s2, _
        Cx + s2, Cy - s2, Cz - s2, _
        Cx + s2, Cy - s2, Cz + s2, _
        Cx - s2, Cy - s2, Cz + s2
End Sub
' Make a tetrahedron.
Private Sub MakeTetrahedron(ByVal side_length As Single)
Dim S As Single
Dim A As Single
Dim B As Single
Dim C As Single
Dim D As Single

    S = Sqr(6) * side_length
    A = S / Sqr(3)
    B = -A / 2
    C = A * Sqr(2) - 1
    D = S / 2

    TheSolid.AddFace _
        0, C, 0, _
        A, -1, 0, _
        B, -1, -D
    TheSolid.AddFace _
        0, C, 0, _
        B, -1, -D, _
        B, -1, D
    TheSolid.AddFace _
        0, C, 0, _
        B, -1, D, _
        A, -1, 0
    TheSolid.AddFace _
        A, -1, 0, _
        B, -1, D, _
        B, -1, -D
End Sub
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

