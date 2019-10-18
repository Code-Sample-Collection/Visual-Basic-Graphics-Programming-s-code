VERSION 5.00
Begin VB.Form frmPZOrder 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "PZOrder"
   ClientHeight    =   6375
   ClientLeft      =   300
   ClientTop       =   570
   ClientWidth     =   9135
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
   ScaleHeight     =   6375
   ScaleWidth      =   9135
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
      Default         =   -1  'True
      Height          =   495
      Left            =   6840
      TabIndex        =   18
      Top             =   120
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scenes"
      Height          =   4215
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   2415
      Begin VB.OptionButton optSurface 
         Caption         =   "Saddle"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   17
         Top             =   2040
         Width           =   2055
      End
      Begin VB.OptionButton optSurface 
         Caption         =   "Cone"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   2055
      End
      Begin VB.OptionButton optSurface 
         Caption         =   "Holes"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   2055
      End
      Begin VB.OptionButton optSurface 
         Caption         =   "Hemisphere"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton optSurface 
         Caption         =   "Mounds"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton optSurface 
         Caption         =   "Splash"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optSurface 
         Caption         =   "Monkey Saddle"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   11
         Top             =   2400
         Width           =   2055
      End
      Begin VB.OptionButton optSurface 
         Caption         =   "Hill and Hole"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   10
         Top             =   2760
         Width           =   2055
      End
      Begin VB.OptionButton optSurface 
         Caption         =   "Canyons"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   9
         Top             =   3120
         Width           =   2055
      End
      Begin VB.OptionButton optSurface 
         Caption         =   "Pit"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   8
         Top             =   3480
         Width           =   2055
      End
      Begin VB.OptionButton optSurface 
         Caption         =   "Volcano"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   7
         Top             =   3840
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Light Sources"
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   0
      Width           =   4095
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
      Begin VB.CheckBox chkLights 
         Caption         =   "Red"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkLights 
         Caption         =   "Green"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkLights 
         Caption         =   "Blue"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   5655
      Left            =   2520
      ScaleHeight     =   373
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   437
      TabIndex        =   0
      Top             =   720
      Width           =   6615
   End
End
Attribute VB_Name = "frmPZOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Specular reflection coefficients.
Private Const SPEC_K = 0.5
Private Const SPEC_N = 50

' The ambient light.
Private Const THE_AMBIENT_LIGHT = 50

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

Private TheGrid As PZOrderGrid3d

Private Enum SurfaceTypes
    surface_Splash = 0
    surface_Mounds = 1
    surface_Bowl = 2
    surface_Ridges = 3
    surface_RandomRidges = 4
    surface_Hemisphere = 5
    surface_Holes = 6
    surface_Cone = 7
    surface_Saddle = 8
    surface_MonkeySaddle = 9
    surface_HillAndHole = 10
    surface_Canyons = 11
    surface_Pit = 12
    surface_Volcano = 13
End Enum
Private SelectedSurface As SurfaceTypes

Private SphereRadius As Single
Private Const Amplitude1 = 0.25
Private Const Period1 = 2 * PI / 4
Private Const Amplitude2 = 1
Private Const Period2 = 2 * PI / 16
Private Const Amplitude3 = 2
Private Const xmin = -5
Private Const zmin = -5

Private LightSources As Collection
' Set this light source's Kdist and Rmin values.
Private Sub ScaleIntensityForDepth(ByVal Light As LightSource)
Dim Rmin As Single
Dim Rmax As Single

    Rmin = Sqr(5 * 5 + 5 * 5)
    Rmax = -Rmin

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
        Light.Initialize -30, 50, 100, 200, 200, 200
    End If

    ' Red.
    If chkLights(1).value = vbChecked Then
        Set Light = New LightSource
        LightSources.Add Light
        Light.Initialize 50, 30, 100, 200, 0, 0
    End If

    ' Green.
    If chkLights(2).value = vbChecked Then
        Set Light = New LightSource
        LightSources.Add Light
        Light.Initialize 30, -50, 100, 0, 200, 0
    End If

    ' Blue.
    If chkLights(3).value = vbChecked Then
        Set Light = New LightSource
        LightSources.Add Light
        Light.Initialize -50, -30, 100, 0, 0, 200
    End If
End Sub

' Return the Y coordinate for these X and
' Z coordinates.
Private Function YValue(ByVal X As Single, ByVal Z As Single)
Dim x1 As Single
Dim z1 As Single
Dim x2 As Single
Dim z2 As Single
Dim D As Single

    Select Case SelectedSurface
        Case surface_Splash
            D = Sqr(X * X + Z * Z)
            YValue = Amplitude1 * Cos(3 * D)

        Case surface_Mounds
            YValue = Amplitude1 * (Cos(Period1 * X) + Cos(Period1 * Z))

        Case surface_Bowl
            YValue = 0.2 * (X * X + Z * Z) - 5#

        Case surface_Ridges
            YValue = Amplitude2 * Cos(Period2 * X) + Amplitude3 * Cos(Period1 * Z) / (Abs(Z) / 3 + 1)

        Case surface_RandomRidges
            YValue = Amplitude2 * Cos(Period2 * X) + Amplitude3 * Cos(Period1 * Z) / (Abs(Z) / 3 + 1) + Amplitude1 * Rnd

        Case surface_Hemisphere
            D = X * X + Z * Z
            If D >= SphereRadius Then
                YValue = 0
            Else
                YValue = Sqr(SphereRadius - D)
            End If

        Case surface_Holes
            x1 = (X + xmin / 2)
            z1 = (Z + xmin / 2)
            x2 = (X - xmin / 2)
            z2 = (Z - xmin / 2)
            YValue = Amplitude3 - _
                1 / (x1 * x1 + z1 * z1 + 0.1) - _
                1 / (x2 * x2 + z1 * z1 + 0.1) - _
                1 / (x1 * x1 + z2 * z2 + 0.1) - _
                1 / (x2 * x2 + z2 * z2 + 0.1)

        Case surface_Cone
            D = 2 * (Amplitude3 - Sqr(X * X + Z * Z))
            If D < -Amplitude3 Then D = -Amplitude3
            YValue = D

        Case surface_Saddle
            YValue = (X * X - Z * Z) / 10

        Case surface_MonkeySaddle
            x1 = 1.5 * X
            z1 = 1.5 * Z
            YValue = (x1 * x1 * x1 / 3 - x1 * z1 * z1) / 50

        Case surface_HillAndHole
            YValue = -5 * X / (X * X + Z * Z + 1)

        Case surface_Canyons
            YValue = Sin(X * 1.5) * Z * Z * Z / 30

        Case surface_Pit
            YValue = -3 + (X * X + Z * Z) / 10 + Sin(2 * Sqr(X * X + Z * Z)) / 2

        Case surface_Volcano
            YValue = -Abs(X * X + Z * Z - 9) / 10
    End Select
End Function
' Return the unit normal vector at (X, Z).
Private Function GetNormal(ByVal X As Single, ByVal Z As Single, ByRef Nx As Single, ByRef Ny As Single, ByRef Nz As Single)
Dim x1 As Single
Dim z1 As Single
Dim x2 As Single
Dim z2 As Single
Dim d2 As Single
Dim D As Single
Dim Y As Single

    Select Case SelectedSurface
        Case surface_Splash
            d2 = X * X + Z * Z
            D = Sqr(d2)
            Nx = 3 * Amplitude1 * Sin(3 * D) / _
                D * 2 * X
            Ny = 1
            Nz = 3 * Amplitude1 * Sin(3 * D) / _
                D * 2 * Z

        Case surface_Mounds
            ' Y = Amplitude1 * (Cos(Period1 * X) + Cos(Period1 * Z))
            Nx = Amplitude1 * Period1 * Sin(Period1 * X)
            Ny = 1
            Nz = Amplitude1 * Period1 * Sin(Period1 * Z)

        Case surface_Hemisphere
            D = X * X + Z * Z
            If D >= SphereRadius Then
                Nx = 0
                Ny = 1
                Nz = 0
            Else
                Nx = X / Sqr(SphereRadius - D)
                Ny = 1
                Nz = Z / Sqr(SphereRadius - D)
            End If

        Case surface_Holes
            x1 = (X + xmin / 2)
            z1 = (Z + xmin / 2)
            x2 = (X - xmin / 2)
            z2 = (Z - xmin / 2)
            Nx = 2 * x1 / (x1 * x1 + z1 * z1 + 0.1) ^ 2 - _
                 2 * x2 / (x2 * x2 + z1 * z1 + 0.1) ^ 2 - _
                 2 * x1 / (x1 * x1 + z2 * z2 + 0.1) ^ 2 - _
                 2 * x2 / (x2 * x2 + z2 * z2 + 0.1) ^ 2
            Ny = 1
            Nz = 2 * z1 / (x1 * x1 + z1 * z1 + 0.1) ^ 2 - _
                 2 * z1 / (x2 * x2 + z1 * z1 + 0.1) ^ 2 - _
                 2 * z2 / (x1 * x1 + z2 * z2 + 0.1) ^ 2 - _
                 2 * z2 / (x2 * x2 + z2 * z2 + 0.1) ^ 2

        Case surface_Cone
            D = 2 * (Amplitude3 - Sqr(X * X + Z * Z))
            If D < -Amplitude3 Then D = -Amplitude3
            Nx = 2 * X / Sqr(X * X + Z * Z)
            Ny = 1
            Nz = 2 * Z / Sqr(X * X + Z * Z)

        Case surface_Saddle
            ' Y = (X * X - Z * Z) / 10
            Nx = -2 * X / 10
            Ny = 1
            Nz = 2 * X / 10

        Case surface_MonkeySaddle
            x1 = 1.5 * X
            z1 = 1.5 * Z
            ' Y = (x1 * x1 * x1 / 3 - x1 * z1 * z1) / 50
            Nx = -(1.5 * 3 * x1 * x1 / 3 - 1.5 * z1 * z1) / 50
            Ny = 1
            Nz = (2 * x1 * z1) / 50

        Case surface_HillAndHole
            ' Y = -5 * X / (X * X + Z * Z + 1)
            Nx = -(-5 * (X * X + Z * Z + 1) - (-5 * X) * (2 * X)) / ((X * X + Z * Z + 1) * (X * X + Z * Z + 1))
            Ny = 1
            Nz = 5 * X / (X * X + Z * Z + 1) / (X * X + Z * Z + 1) * 2 * Z

        Case surface_Canyons
            ' Y = Sin(X * 1.5) * Z * Z * Z / 30
            Nx = -1.5 * Cos(X * 1.5) * Z * Z * Z / 30
            Ny = 1
            Nz = 3 * Z * Z * Sin(X * 1.5) / 30

        Case surface_Pit
            ' Y = -3 + (X * X + Z * Z) / 10 + Sin(2 * Sqr(X * X + Z * Z)) / 2
            Nx = -2 * X / 10 - Cos(2 * Sqr(X * X + Z * Z)) * 2 / 2 / Sqr(X * X + Z * Z) * 2 * X / 2
            Ny = 1
            Nz = -2 * Z / 10 - Cos(2 * Sqr(X * X + Z * Z)) * 2 / 2 / Sqr(X * X + Z * Z) * 2 * Z / 2

        Case surface_Volcano
            ' Y = -Abs(X * X + Z * Z - 9) / 10
            If (X * X + Z * Z - 9) / 10 < 0 Then
                Nx = 2 * X / 10
                Ny = 1
                Nz = 2 * Z / 10
            Else
                Nx = -2 * X / 10
                Ny = 1
                Nz = -2 * Z / 10
            End If
    End Select
End Function
' Project and display the data.
Private Sub DrawData(pic As Object)
Dim X As Single
Dim Y As Single
Dim Z As Single
Dim S(1 To 4, 1 To 4) As Single
Dim T(1 To 4, 1 To 4) As Single
Dim ST(1 To 4, 1 To 4) As Single
Dim PST(1 To 4, 1 To 4) As Single

    MousePointer = vbHourglass
    DoEvents

    ' Make the data.
    CreateData
    CreateLightSources

    ' Scale and translate so it looks OK in pixels.
    m3Scale S, 35, -35, 35
    m3Translate T, 230, 175, 0
    m3MatMultiplyFull ST, S, T
    m3MatMultiplyFull PST, Projector, ST

    ' Transform the points.
    TheGrid.ApplyFull PST

    ' Set the light sources' Kdist and Rmin values
    ' used to fade colors by distance. This should
    ' happen after culling.
    ScaleLightSourcesForDepth

    ' Prevent overflow errors when drawing lines
    ' too far out of bounds.
    On Error Resume Next

    m3SphericalToCartesian EyeR, EyeTheta, EyePhi, X, Y, Z

    ' Display the data.
    pic.Cls
    TheGrid.RemoveHidden = True
    TheGrid.Draw pic, LightSources, _
        THE_AMBIENT_LIGHT, X, Y, Z
    pic.Refresh

    MousePointer = vbDefault
    picCanvas.SetFocus
End Sub

Private Sub cmdDraw_Click()
    Screen.MousePointer = vbHourglass
    DoEvents

    DrawData picCanvas

    Screen.MousePointer = vbDefault
End Sub

Private Sub optSurface_Click(Index As Integer)
    SelectedSurface = Index
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

    m3PProject Projector, m3Parallel, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0
    cmdDraw_Click
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
    cmdDraw_Click
End Sub

Private Sub Form_Load()
    ' Initialize the eye position.
    EyeR = 10
    EyeTheta = PI * 0.2
    EyePhi = PI * 0.1

    ' Initialize the projection transformation.
    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0
End Sub

' Create the surface.
Private Sub CreateData()
Const dx = 0.3
Const dz = 0.3
Const NumX = -2 * xmin / dx
Const NumZ = -2 * zmin / dz

Dim i As Integer
Dim j As Integer
Dim X As Single
Dim Y As Single
Dim Z As Single
Dim Nx As Single
Dim Ny As Single
Dim Nz As Single

    SphereRadius = (xmin + 3 * dx) * (xmin + 3 * dx)

    Set TheGrid = New PZOrderGrid3d
    TheGrid.SetBounds xmin, dx, NumX, zmin, dz, NumZ

    X = xmin
    For i = 1 To NumX
        Z = zmin
        For j = 1 To NumZ
            Y = YValue(X, Z)
            TheGrid.SetValue X, Y, Z

            GetNormal X, Z, Nx, Ny, Nz
            TheGrid.SetNormal X, Z, Nx, Ny, Nz

            Z = Z + dz
        Next j
        X = X + dx
    Next i

    TheGrid.DiffuseKr = 1#
    TheGrid.DiffuseKg = 1#
    TheGrid.DiffuseKb = 1#
    TheGrid.AmbientKr = 1#
    TheGrid.AmbientKg = 1#
    TheGrid.AmbientKb = 1#
    TheGrid.SpecularK = SPEC_K
    TheGrid.SpecularN = SPEC_N
End Sub
