VERSION 5.00
Begin VB.Form frmGouraud 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Gouraud"
   ClientHeight    =   5445
   ClientLeft      =   1410
   ClientTop       =   540
   ClientWidth     =   7005
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
   ScaleHeight     =   5445
   ScaleWidth      =   7005
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
      Default         =   -1  'True
      Height          =   495
      Left            =   600
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Light Sources"
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      Top             =   0
      Width           =   4095
      Begin VB.CheckBox chkLights 
         Caption         =   "Blue"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkLights 
         Caption         =   "Green"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkLights 
         Caption         =   "Red"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkLights 
         Caption         =   "White"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scenes"
      Height          =   1335
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2415
      Begin VB.OptionButton optSolid 
         Caption         =   "Fine Sphere"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton optSolid 
         Caption         =   "Coarse Sphere"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton optSolid 
         Caption         =   "Medium Sphere"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   3615
      Left            =   2520
      ScaleHeight     =   237
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   197
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
End
Attribute VB_Name = "frmGouraud"
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
Private LightSources As Collection

Private SelectedShape As Integer

Private Const THE_AMBIENT_LIGHT = 50

' Specular reflection coefficients for all solids.
Private Const SPEC_K = 0.5
Private Const SPEC_N = 50
' Set this light source's Kdist and Rmin values.
Private Sub ScaleIntensityForDepth(ByVal Light As LightSource)
Dim solid As Solid3d
Dim Rmin As Single
Dim Rmax As Single
Dim new_rmin As Single
Dim new_rmax As Single

    Rmin = 1E+30
    Rmax = -1E+30

    For Each solid In Solids
        solid.GetRminRmax new_rmin, new_rmax, _
            Light.X, Light.Y, Light.Z
        If Rmin > new_rmin Then Rmin = new_rmin
        If Rmax < new_rmax Then Rmax = new_rmax
    Next solid

    Light.Rmin = Rmin
'    light.Kdist = (Rmax - 5 * Rmin) / 4 ' Fade to 1/5.
    Light.Kdist = Rmax - 2 * Rmin ' Fade to 1/2.
End Sub
' Set the light sources' Kdist and Rmin values.
Private Sub ScaleLightSourcesForDepth()
Dim Light As LightSource

    For Each Light In LightSources
        ScaleIntensityForDepth Light
    Next Light
End Sub
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
    m3SphericalToCartesian EyeR, EyeTheta, EyePhi, X, Y, Z
    For Each solid In Solids
        solid.Culled = False
        solid.Cull X, Y, Z
    Next solid

    ' Scale and translate so it looks OK in pixels.
    m3Scale S, 100, -100, 1
    m3Translate T, pic.ScaleWidth / 2, pic.ScaleHeight / 2, 0
    m3MatMultiplyFull ST, S, T
    m3MatMultiplyFull PST, Projector, ST

    ' Transform the solids and clip faces.
    For Each solid In Solids
        solid.ApplyFull PST

        ' Clip faces behind the center of projection.
        solid.ClipEye EyeR
    Next solid

    ' Sort the solids if necessary.
    SortSolids

    ' Set the light sources' Kdist and Rmin values
    ' used to fade colors by distance. This should
    ' happen after culling.
    ScaleLightSourcesForDepth

    ' We will set one point at a time using PSet.
    pic.FillStyle = vbFSTransparent
    pic.DrawStyle = vbSolid

    ' Draw the solids.
    pic.Cls
    For Each solid In Solids
        solid.Draw pic, LightSources, THE_AMBIENT_LIGHT, X, Y, Z
    Next solid
    pic.Refresh
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
                    Cx + x22, Cy + y22, Cz + z22, _
                    Cx + x11, Cy + y11, Cz + z11, _
                    Cx + x12, Cy + y12, Cz + z12, _
                    Cx + x22, Cy + y22, Cz + z22
            ElseIf P = num_vertical Then
                ' Top triangle.
                new_solid.AddFace _
                    Cx + x11, Cy + y11, Cz + z11, _
                    Cx + x12, Cy + y12, Cz + z12, _
                    Cx + x21, Cy + y21, Cz + z21, _
                    Cx + x11, Cy + y11, Cz + z11, _
                    Cx + x12, Cy + y12, Cz + z12, _
                    Cx + x21, Cy + y21, Cz + z21
            Else
                ' Middle rectangle.
                new_solid.AddFace _
                    Cx + x11, Cy + y11, Cz + z11, _
                    Cx + x12, Cy + y12, Cz + z12, _
                    Cx + x22, Cy + y22, Cz + z22, _
                    Cx + x21, Cy + y21, Cz + z21, _
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

    new_solid.IsConvex = True
    new_solid.HideSurfaces = True
    new_solid.SetDiffuseCoefficients 1#, 1#, 1#
    new_solid.SetAmbientCoefficients 1#, 1#, 1#
    new_solid.SetSpecularCoefficients SPEC_K, SPEC_N
    Set Sphere = new_solid
End Function
' Draw the solid.
Private Sub cmdDraw_Click()
    Screen.MousePointer = vbHourglass
    DoEvents

    CreateData
    CreateLightSources
    DrawData picCanvas

    picCanvas.SetFocus

    Screen.MousePointer = vbDefault
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

    Screen.MousePointer = vbHourglass
    DoEvents

    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0
    DrawData picCanvas

    Screen.MousePointer = vbDefault
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
    CreateLightSources
    optSolid(0).value = True
End Sub

' Create the data.
Private Sub CreateData()
    ' Create the new Solids collection.
    Set Solids = New Collection

    ' Create the solids.
    Select Case SelectedShape
        Case 0  ' Coarse Sphere.
            Solids.Add Sphere(0, 0, 0, 1, 10, 10)

        Case 1  ' Medium Sphere.
            Solids.Add Sphere(0, 0, 0, 1, 30, 30)

        Case 2  ' Fine Sphere.
            Solids.Add Sphere(0, 0, 0, 1, 100, 100)

    End Select
End Sub
' Create the light sources.
Private Sub CreateLightSources()
Dim Light As LightSource

    ' Create the new LightSources collection.
    Set LightSources = New Collection

    ' Create the light sources.
    ' White.
    If chkLights(0).value = vbChecked Then
        Set Light = New LightSource
        LightSources.Add Light
        Light.Initialize -300, 500, 1000, 200, 200, 200
    End If

    ' Red.
    If chkLights(1).value = vbChecked Then
        Set Light = New LightSource
        LightSources.Add Light
        Light.Initialize -200, 200, 1000, 200, 0, 0
    End If

    ' Green.
    If chkLights(2).value = vbChecked Then
        Set Light = New LightSource
        LightSources.Add Light
        Light.Initialize 300, -500, 300, 0, 200, 0
    End If

    ' Blue.
    If chkLights(3).value = vbChecked Then
        Set Light = New LightSource
        LightSources.Add Light
        Light.Initialize 1000, 300, -300, 0, 0, 200
    End If
End Sub
' Make the drawing areas as large as possible.
Private Sub Form_Resize()
Dim wid As Single
Dim hgt As Single

    wid = ScaleWidth - picCanvas.Left
    If wid < 120 Then wid = 120
    hgt = ScaleHeight - picCanvas.Top
    If hgt < 120 Then hgt = 120
    picCanvas.Move picCanvas.Left, picCanvas.Top, _
        wid, hgt
End Sub

Private Sub optSolid_Click(Index As Integer)
    SelectedShape = Index
End Sub

