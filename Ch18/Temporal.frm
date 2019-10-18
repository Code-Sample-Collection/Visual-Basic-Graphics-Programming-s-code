VERSION 5.00
Begin VB.Form frmTemporal 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Temporal"
   ClientHeight    =   5430
   ClientLeft      =   300
   ClientTop       =   570
   ClientWidth     =   8670
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
   ScaleHeight     =   5430
   ScaleWidth      =   8670
   Begin VB.CommandButton cmdPlayImages 
      Caption         =   "Play Images"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdCreateImages 
      Caption         =   "Create Images"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   5295
      Left            =   1680
      ScaleHeight     =   349
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   461
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label lblPlayStatus 
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.Image imgFrame 
      Height          =   855
      Index           =   0
      Left            =   240
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblCreateStatus 
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmTemporal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NUM_IMAGES = 20

' Location of viewing eye.
Private EyeR As Single
Private EyeTheta As Single
Private EyePhi As Single

' Location of focus point.
Private Const FocusX = 0#
Private Const FocusY = 0#
Private Const FocusZ = 0#

Private Projector(1 To 4, 1 To 4) As Single

Private TheGrid As ZOrderGrid3d

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

Private Playing As Boolean

' Run the animation once or until Playing is False.
Private Sub PlayImagesOnce(ByVal ms_per_frame As Integer)
Dim i As Integer
Dim next_time As Long

    ' Get the current time.
    next_time = GetTickCount

    ' Start the animation.
    For i = 0 To NUM_IMAGES - 1
        lblPlayStatus.Caption = "Frame " & Format$(i)

        ' Display the next frame.
        picCanvas.Picture = imgFrame(i).Picture

        ' Wait till we should display the next frame.
        next_time = next_time + ms_per_frame
        WaitTill next_time

        If Not Playing Then Exit For
    Next i
End Sub
' Run the animation until Playing is false.
Private Sub PlayImagesLooping(ByVal ms_per_frame As Integer)
    ' Start the animation.
    Do While Playing
        PlayImagesOnce ms_per_frame
    Loop
End Sub
' Return the Y coordinate for these X, Z, and
' T coordinates.
Private Function YValue(ByVal X As Single, ByVal Z As Single, ByVal T As Single)
    YValue = 0.25 * Cos(3 * Sqr(X * X + Z * Z) - T * PI / 10)
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

    ' Scale and translate so it looks OK in pixels.
    m3Scale S, 35, -35, 1
    m3Translate T, 230, 175, 0
    m3MatMultiplyFull ST, S, T
    m3MatMultiplyFull PST, Projector, ST

    ' Transform the points.
    TheGrid.ApplyFull PST

    ' Prevent overflow errors when drawing lines
    ' too far out of bounds.
    On Error Resume Next

    ' Display the data.
    pic.Cls
    TheGrid.RemoveHidden = True
    TheGrid.Draw pic
    pic.Refresh
End Sub
' Make the images for the animation sequence.
Private Sub cmdCreateImages_Click()
Dim T As Integer

    cmdCreateImages.Enabled = False
    Screen.MousePointer = vbHourglass

    ' Make the frames.
    For T = 0 To NUM_IMAGES - 1
        lblCreateStatus.Caption = "Frame " & Format$(T)
        DoEvents

        ' Make the data.
        CreateData T

        ' Display the adta.
        DrawData picCanvas

        ' Save the image for playback later.
        If imgFrame.UBound < T Then
            Load imgFrame(T)
        End If

        imgFrame(T).Picture = picCanvas.Image
    Next T

    lblCreateStatus.Caption = ""
    cmdPlayImages.Enabled = True
    cmdPlayImages.Default = True
    picCanvas.Cls
    Screen.MousePointer = vbDefault
End Sub
' Play the animation.
Private Sub cmdPlayImages_Click()
    If Playing Then
        ' Stop running.
        Playing = False
        cmdPlayImages.Caption = "Play Images"
    Else
        ' Start running.
        Playing = True
        cmdPlayImages.Caption = "Stop"
        PlayImagesLooping 50
    End If
End Sub

Private Sub Form_Load()
    ' Initialize the eye position.
    EyeR = 10
    EyeTheta = PI * 0.2
    EyePhi = PI * 0.1

    ' Initialize the projection transformation.
    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0
End Sub

' Create the surface for this value of t.
Private Sub CreateData(ByVal T As Single)
Const Xmin = -5
Const Zmin = -5
Const Dx = 0.3
Const Dz = 0.3
Const NumX = -2 * Xmin / Dx
Const NumZ = -2 * Zmin / Dz

Dim i As Integer
Dim j As Integer
Dim X As Single
Dim Y As Single
Dim Z As Single

    Set TheGrid = New ZOrderGrid3d
    TheGrid.SetBounds Xmin, Dx, NumX, Zmin, Dz, NumZ

    X = Xmin
    For i = 1 To NumX
        Z = Zmin
        For j = 1 To NumZ
            Y = YValue(X, Z, T)

            TheGrid.SetValue X, Y, Z
            Z = Z + Dz
        Next j
        X = X + Dx
    Next i

    ' Display the data.
    DrawData picCanvas
End Sub
