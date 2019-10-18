VERSION 5.00
Begin VB.Form frmDepth2 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Depth2"
   ClientHeight    =   4065
   ClientLeft      =   1410
   ClientTop       =   540
   ClientWidth     =   6555
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
   ScaleWidth      =   6555
   Begin VB.OptionButton optSolid 
      Caption         =   "Stellate Icosahedron"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   2295
   End
   Begin VB.OptionButton optSolid 
      Caption         =   "Figure 15.4b"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   2295
   End
   Begin VB.OptionButton optSolid 
      Caption         =   "Stellate Dodecahedron"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.OptionButton optSolid 
      Caption         =   "Stellate Octahedron"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.CheckBox chkRemoveBackfaces 
      Caption         =   "Hide Surfaces"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   3975
      Left            =   2400
      ScaleHeight     =   261
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "frmDepth2"
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
Private Const Dphi = PI / 20
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
            solid.Cull X, Y, Z
            solid.HideSurfaces = True
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
' Make a solid like the one shown in Figure 15.4b.
Private Function Fig15_4b() As Solid3d
Const S = 0.75

Dim new_solid As Solid3d

    Set new_solid = New Solid3d
    new_solid.IsConvex = False
    new_solid.AddFace _
        S, S, S, _
        -S, S, S, _
        -S, -S, S, _
        S, -S, S
    new_solid.AddFace _
        S, S, S, _
        S, -S, S, _
        S, -S, -S, _
        S, S, -S
    new_solid.AddFace _
        S, S, -S, _
        S, -S, -S, _
        -S, -S, -S, _
        -S, S, -S
    new_solid.AddFace _
        S, S, -S, _
        -S, S, -S, _
        0, S, 0, _
        -S, S, S, _
        S, S, S
    new_solid.AddFace _
        S, -S, S, _
        -S, -S, S, _
        0, -S, 0, _
        -S, -S, -S, _
        S, -S, -S
    new_solid.AddFace _
        -S, S, -S, _
        -S, -S, -S, _
        0, -S, 0, _
        0, S, 0
    new_solid.AddFace _
        0, S, 0, _
        0, -S, 0, _
        -S, -S, S, _
        -S, S, S

    Set Fig15_4b = new_solid
End Function
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
            EyePhi = EyePhi - Dphi
        
        Case vbKeyDown
            EyePhi = EyePhi + Dphi
                
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
    optSolid(0).value = True
End Sub

' Create the data.
Private Sub CreateData()
    ' Create the new Solids collection.
    Set Solids = New Collection

    ' Create the solids.
    Select Case SelectedShape
        Case 0  ' Stellate Octahedron.
            Solids.Add Stellate8(0, 0, 0, 0.8)

        Case 1  ' Stellate Dodecahedron.
            Solids.Add Stellate12(0, 0, 0, 0.4)

        Case 2  ' Stellate Icosahedron.
            Solids.Add Stellate20(0, 0, 0, 0.4)

        Case 3  ' Figure 15.4b.
            Solids.Add Fig15_4b

    End Select
End Sub
' Make a stellate octahedron.
Private Function Stellate8(ByVal Cx As Single, ByVal Cy As Single, ByVal Cz As Single, ByVal side_scale As Single) As Solid3d
Dim new_solid As Solid3d

    Set new_solid = New Solid3d
    new_solid.IsConvex = False

    new_solid.Stellate side_scale, _
        Cx + 0, Cy + side_scale, 0 + Cz, _
        Cx + 0, Cy + 0, side_scale + Cz, _
        Cx + side_scale, 0, 0 + Cz
    new_solid.Stellate side_scale + Cz, _
        Cx + 0, Cy + side_scale, 0 + Cz, _
        Cx + side_scale, Cy + 0, 0 + Cz, _
        Cx + 0, Cy + 0, -side_scale + Cz
    new_solid.Stellate side_scale + Cz, _
        Cx + 0, Cy + side_scale, 0 + Cz, _
        Cx + 0, Cy + 0, -side_scale + Cz, _
        Cx + -side_scale, Cy + 0, 0 + Cz
    new_solid.Stellate side_scale + Cz, _
        Cx + 0, Cy + side_scale, 0 + Cz, _
        Cx + -side_scale, Cy + 0, 0 + Cz, _
        Cx + 0, Cy + 0, side_scale + Cz
    new_solid.Stellate side_scale + Cz, _
        Cx + 0, Cy + -side_scale, 0 + Cz, _
        Cx + side_scale, Cy + 0, 0 + Cz, _
        Cx + 0, Cy + 0, side_scale + Cz
    new_solid.Stellate side_scale + Cz, _
        Cx + 0, Cy + -side_scale, 0 + Cz, _
        Cx + 0, Cy + 0, -side_scale + Cz, _
        Cx + side_scale, 0, 0 + Cz
    new_solid.Stellate side_scale + Cz, _
        Cx + 0, Cy + -side_scale, 0 + Cz, _
        Cx + -side_scale, Cy + 0, 0 + Cz, _
        Cx + 0, Cy + 0, -side_scale + Cz
    new_solid.Stellate side_scale + Cz, _
        Cx + 0, Cy + -side_scale, 0 + Cz, _
        Cx + 0, Cy + 0, side_scale + Cz, _
        Cx + -side_scale, Cy + 0, 0 + Cz

    Set Stellate8 = new_solid
End Function
' Make a stellate dodecahedron.
Private Function Stellate12(ByVal Cx As Single, ByVal Cy As Single, ByVal Cz As Single, ByVal side_scale As Single) As Solid3d
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
    new_solid.IsConvex = False
    new_solid.Stellate side_scale * 1.5, _
        Cx + C, Cy + side_scale, Cz + -A, _
        Cx + D, Cy + side_scale, Cz + -B, _
        Cx + D, Cy + side_scale, Cz + B, _
        Cx + C, Cy + side_scale, Cz + A, _
        Cx + R, Cy + side_scale, Cz + 0
    new_solid.Stellate side_scale * 1.5, _
        Cx + C, Cy + side_scale, Cz + A, _
        Cx + X * c1, Cy + side_scale - Y, Cz + X * s1, _
        Cx + -X * c2, Cy + side_scale - y2, Cz + X * s2, _
        Cx + X, Cy + side_scale - Y, Cz + 0, _
        Cx + R, Cy + side_scale, Cz + 0
    new_solid.Stellate side_scale * 1.5, _
        Cx + C, Cy + side_scale, Cz + A, _
        Cx + D, Cy + side_scale, Cz + B, _
        Cx + X * c2, Cy + side_scale - Y, Cz + X * s2, _
        Cx + -X * c1, Cy + side_scale - y2, Cz + X * s1, _
        Cx + X * c1, Cy + side_scale - Y, Cz + X * s1
    new_solid.Stellate side_scale * 1.5, _
        Cx + D, Cy + side_scale, Cz + B, _
        Cx + D, Cy + side_scale, Cz + -B, _
        Cx + X * c2, Cy + side_scale - Y, Cz + -X * s2, _
        Cx + -X, Cy + side_scale - y2, Cz + 0, _
        Cx + X * c2, Cy + side_scale - Y, Cz + X * s2
    new_solid.Stellate side_scale * 1.5, _
        Cx + D, Cy + side_scale, Cz + -B, _
        Cx + C, Cy + side_scale, Cz + -A, _
        Cx + X * c1, Cy + side_scale - Y, Cz + -X * s1, _
        Cx + -X * c1, Cy + side_scale - y2, Cz + -X * s1, _
        Cx + X * c2, Cy + side_scale - Y, Cz + -X * s2, -X * c1
    new_solid.Stellate side_scale * 1.5, _
        Cx + C, Cy + side_scale, Cz + -A, _
        Cx + R, Cy + side_scale, Cz + 0, _
        Cx + X, Cy + side_scale - Y, Cz + 0, _
        Cx + -X * c2, Cy + side_scale - y2, Cz + -X * s2, _
        Cx + X * c1, Cy + side_scale - Y, Cz + -X * s1

    ' Bottom.
    new_solid.Stellate side_scale * 1.5, _
        Cx + -D, Cy + -side_scale, Cz + -B, _
        Cx + -X * c2, Cy + side_scale - y2, Cz + -X * s2, _
        Cx + X, Cy + side_scale - Y, Cz + 0, _
        Cx + -X * c2, Cy + side_scale - y2, Cz + X * s2, _
        Cx + -D, Cy + -side_scale, Cz + B
    new_solid.Stellate side_scale * 1.5, _
        Cx + -D, Cy + -side_scale, Cz + B, _
        Cx + -X * c2, Cy + side_scale - y2, Cz + X * s2, _
        Cx + X * c1, Cy + side_scale - Y, Cz + X * s1, _
        Cx + -X * c1, Cy + side_scale - y2, Cz + X * s1, _
        Cx + -C, Cy + -side_scale, Cz + A
    new_solid.Stellate side_scale * 1.5, _
        Cx + -C, Cy + -side_scale, Cz + A, _
        Cx + -X * c1, Cy + side_scale - y2, Cz + X * s1, _
        Cx + X * c2, Cy + side_scale - Y, Cz + X * s2, _
        Cx + -X, Cy + side_scale - y2, Cz + 0, _
        Cx + -R, Cy + -side_scale, Cz + 0
    new_solid.Stellate side_scale * 1.5, _
        Cx + -R, Cy + -side_scale, Cz + 0, _
        Cx + -X, Cy + side_scale - y2, Cz + 0, _
        Cx + X * c2, Cy + side_scale - Y, Cz + -X * s2, _
        Cx + -X * c1, Cy + side_scale - y2, Cz + -X * s1, _
        Cx + -C, Cy + -side_scale, Cz + -A
    new_solid.Stellate side_scale * 1.5, _
        Cx + -C, Cy + -side_scale, Cz + -A, _
        Cx + -X * c1, Cy + side_scale - y2, Cz + -X * s1, _
        Cx + X * c1, Cy + side_scale - Y, Cz + -X * s1, _
        Cx + -X * c2, Cy + side_scale - y2, Cz + -X * s2, _
        Cx + -D, Cy + -side_scale, Cz + -B
    new_solid.Stellate side_scale * 1.5, _
        Cx + -D, Cy + -side_scale, Cz + -B, _
        Cx + -D, Cy + -side_scale, Cz + B, _
        Cx + -C, Cy + -side_scale, Cz + A, _
        Cx + -R, Cy + -side_scale, Cz + 0, _
        Cx + -C, Cy + -side_scale, Cz + -A

    Set Stellate12 = new_solid
End Function

' Make a stellate icosahedron.
Private Function Stellate20(ByVal Cx As Single, ByVal Cy As Single, ByVal Cz As Single, ByVal side_scale As Single) As Solid3d
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
    new_solid.IsConvex = False
    new_solid.Stellate side_scale * 1.5, _
        Cx + 0, Cy + side_scale, 0 + Cz, _
        Cx + C, Cy + H, A + Cz, _
        Cx + R, Cy + H, 0 + Cz
    new_solid.Stellate side_scale * 1.5, _
        Cx + 0, Cy + side_scale, 0 + Cz, _
        Cx + R, Cy + H, 0 + Cz, _
        Cx + C, Cy + H, -A + Cz
    new_solid.Stellate side_scale * 1.5, _
        Cx + 0, Cy + side_scale, 0 + Cz, _
        Cx + C, Cy + H, -A + Cz, _
        Cx + D, Cy + H, -B + Cz
    new_solid.Stellate side_scale * 1.5, _
        Cx + 0, Cy + side_scale, 0 + Cz, _
        Cx + D, Cy + H, -B + Cz, _
        Cx + D, Cy + H, B + Cz
    new_solid.Stellate side_scale * 1.5, _
        Cx + 0, Cy + side_scale, 0 + Cz, _
        Cx + D, Cy + H, B + Cz, _
        Cx + C, Cy + H, A + Cz

    ' Upper Middle.
    new_solid.Stellate side_scale * 1.5, _
        Cx + R, Cy + H, 0 + Cz, _
        Cx + C, Cy + H, A + Cz, _
        Cx + -D, Cy + -H, B + Cz
    new_solid.Stellate side_scale * 1.5, _
        Cx + C, Cy + H, A + Cz, _
        Cx + D, Cy + H, B + Cz, _
        Cx + -C, Cy + -H, A + Cz
    new_solid.Stellate side_scale * 1.5, _
        Cx + D, Cy + H, B + Cz, _
        Cx + D, Cy + H, -B + Cz, _
        Cx + -R, Cy + -H, 0 + Cz
    new_solid.Stellate side_scale * 1.5, _
        Cx + D, Cy + H, -B + Cz, _
        Cx + C, Cy + H, -A + Cz, _
        Cx + -C, Cy + -H, -A + Cz
    new_solid.Stellate side_scale * 1.5, _
        Cx + C, Cy + H, -A + Cz, _
        Cx + R, Cy + H, 0 + Cz, _
        Cx + -D, Cy + -H, -B + Cz

    ' Lower Middle.
    new_solid.Stellate side_scale * 1.5, _
        Cx + R, Cy + H, 0 + Cz, _
        Cx + -D, Cy + -H, B + Cz, _
        Cx + -D, Cy + -H, -B + Cz
    new_solid.Stellate side_scale * 1.5, _
        Cx + C, Cy + H, A + Cz, _
        Cx + -C, Cy + -H, A + Cz, _
        Cx + -D, Cy + -H, B + Cz
    new_solid.Stellate side_scale * 1.5, _
        Cx + D, Cy + H, B + Cz, _
        Cx + -R, Cy + -H, 0 + Cz, _
        Cx + -C, Cy + -H, A + Cz
    new_solid.Stellate side_scale * 1.5, _
        Cx + D, Cy + H, -B + Cz, _
        Cx + -C, Cy + -H, -A + Cz, _
        Cx + -R, Cy + -H, 0 + Cz
    new_solid.Stellate side_scale * 1.5, _
        Cx + C, Cy + H, -A + Cz, _
        Cx + -D, Cy + -H, -B + Cz, _
        Cx + -C, Cy + -H, -A + Cz

    ' Bottom.
    new_solid.Stellate side_scale * 1.5, _
        Cx + 0, Cy + -side_scale, 0 + Cz, _
        Cx + -D, Cy + -H, B + Cz, _
        Cx + -C, Cy + -H, A + Cz
    new_solid.Stellate side_scale * 1.5, _
        Cx + 0, Cy + -side_scale, 0 + Cz, _
        Cx + -C, Cy + -H, A + Cz, _
        Cx + -R, Cy + -H, 0 + Cz
    new_solid.Stellate side_scale * 1.5, _
        Cx + 0, Cy + -side_scale, 0 + Cz, _
        Cx + -R, Cy + -H, 0 + Cz, _
        Cx + -C, Cy + -H, -A + Cz
    new_solid.Stellate side_scale * 1.5, _
        Cx + 0, Cy + -side_scale, 0 + Cz, _
        Cx + -C, Cy + -H, -A + Cz, _
        Cx + -D, Cy + -H, -B + Cz
    new_solid.Stellate side_scale * 1.5, _
        Cx + 0, Cy + -side_scale, 0 + Cz, _
        Cx + -D, Cy + -H, -B + Cz, _
        Cx + -D, Cy + -H, B + Cz

    Set Stellate20 = new_solid
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

