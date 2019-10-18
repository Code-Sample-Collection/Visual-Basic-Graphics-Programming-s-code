VERSION 5.00
Begin VB.Form frmSurface4 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Surface4"
   ClientHeight    =   5295
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
   ScaleHeight     =   5295
   ScaleWidth      =   9135
   Begin VB.CheckBox chkShowData 
      Caption         =   "Show Data"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   0
      Width           =   1335
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Volcano"
      Height          =   255
      Index           =   13
      Left            =   0
      TabIndex        =   14
      Top             =   5040
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Pit"
      Height          =   255
      Index           =   12
      Left            =   0
      TabIndex        =   13
      Top             =   4680
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Canyons"
      Height          =   255
      Index           =   11
      Left            =   0
      TabIndex        =   12
      Top             =   4320
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Hill and Hole"
      Height          =   255
      Index           =   10
      Left            =   0
      TabIndex        =   11
      Top             =   3960
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Monkey Saddle"
      Height          =   255
      Index           =   9
      Left            =   0
      TabIndex        =   10
      Top             =   3600
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Splash"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   360
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Mounds"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Bowl"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Ridges"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Randomized Ridges"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Hemisphere"
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   4
      Top             =   2160
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Holes"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   3
      Top             =   2520
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Cone"
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   2
      Top             =   2880
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Saddle"
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   1
      Top             =   3240
      Width           =   2055
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   5295
      Left            =   2160
      ScaleHeight     =   349
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   461
      TabIndex        =   0
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frmSurface4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Location of viewing eye.
Private EyeR As Single
Private EyeTheta As Single
Private EyePhi As Single

Private Const Dtheta = PI / 20
Private Const Dphi = PI / 20
Private Const Dr = 1

' Location of focus point.
Private Const FocusX = 0#
Private Const FocusY = 0#
Private Const FocusZ = 0#

Private Projector(1 To 4, 1 To 4) As Single

Private TheGrid As WeightedGrid3d

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
Private Const Xmin = -5
Private Const Zmin = -5

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

    ' Scale and translate so it looks OK in pixels.
    m3Scale S, 35, -35, 1
    m3Translate T, 230, 175, 0
    m3MatMultiplyFull ST, S, T
    m3MatMultiplyFull PST, Projector, ST

    ' Create the grid points.
    TheGrid.InitializeGrid 0.3, 0.3

    ' Transform the points.
    TheGrid.ApplyFull PST

    ' Prevent overflow errors when drawing lines
    ' too far out of bounds.
    On Error Resume Next

    ' Display the data.
    pic.Cls
    TheGrid.Draw pic
    pic.Refresh

    MousePointer = vbDefault
    picCanvas.SetFocus
End Sub

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
            x1 = (X + Xmin / 2)
            z1 = (Z + Xmin / 2)
            x2 = (X - Xmin / 2)
            z2 = (Z - Xmin / 2)
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
Private Sub optSurface_Click(Index As Integer)
    SelectedSurface = Index
    DrawData picCanvas
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft
            EyeTheta = EyeTheta - Dtheta
        
        Case vbKeyRight
            EyeTheta = EyeTheta + Dtheta
        
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
    EyePhi = PI * 0.1

    ' Initialize the projection transformation.
    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0

    ' Project and draw the data.
    Me.Show
    DrawData picCanvas
End Sub

' Create the surface.
Private Sub CreateData()
Const Xmin = -5
Const Zmin = -5
Const Xmax = -Xmin
Const Zmax = -Zmin
Const Dx = 0.3
Const Dz = 0.3
Const NumX = -2 * Xmin / Dx
Const NumZ = -2 * Zmin / Dz
Const NUM_PTS = NumX * NumZ / 4

Dim i As Integer
Dim X As Single
Dim Y As Single
Dim Z As Single

    SphereRadius = (Xmin + 3 * Dx) * (Xmin + 3 * Dx)

    Set TheGrid = New WeightedGrid3d
    For i = 1 To NUM_PTS
        ' Pick a random point in the area.
        X = (Xmax - Xmin) * Rnd + Xmin
        Z = (Zmax - Zmin) * Rnd + Zmin
        Y = YValue(X, Z)
        TheGrid.SetValue X, Y, Z
    Next i
End Sub
