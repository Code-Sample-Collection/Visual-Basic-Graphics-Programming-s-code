VERSION 5.00
Begin VB.Form frmFractal 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Fractal"
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
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   480
      Width           =   735
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Valley"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   12
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox txtDy 
      Height          =   285
      Left            =   720
      TabIndex        =   10
      Text            =   "0.25"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtLevel 
      Height          =   285
      Left            =   720
      TabIndex        =   9
      Text            =   "3"
      Top             =   360
      Width           =   495
   End
   Begin VB.CheckBox chkRemoveHidden 
      Caption         =   "Remove Hidden"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Mountain"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Hill"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Ridge"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Peaked Ridge"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Rugged Ridge"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   2
      Top             =   2520
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Random"
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   1
      Top             =   2880
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
   Begin VB.Label Label1 
      Caption         =   "Dy"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Level"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "frmFractal"
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

Private TheGrid As FractalGrid3d

Private Enum SurfaceTypes
    surface_Mountain = 0
    surface_Hill = 1
    surface_Ridge = 2
    surface_PeakedRidge = 3
    surface_RuggedRidge = 4
    surface_Random = 5
    surface_Valley = 6
End Enum
Private SelectedSurface As SurfaceTypes

Private SphereRadius As Single
Private Const Amplitude3 = 2
Private Const Xmin = -5
Private Const Zmin = -5
' Return the Y coordinate for these X and
' Z coordinates.
Private Function YValue(ByVal X As Single, ByVal Z As Single)
Dim Y As Single
Dim D As Single
Dim d2 As Single
Dim x1 As Single
Dim x2 As Single

    Select Case SelectedSurface
        Case surface_Mountain
            x1 = X + 0.5
            D = 2 * (Amplitude3 - Sqr(x1 * x1 + Z * Z))
            x2 = X - 0.5
            d2 = 2 * (Amplitude3 - Sqr(x2 * x2 + Z * Z)) - 0.5
            If D < d2 Then D = d2
            If D < -Amplitude3 Then D = -Amplitude3
            Y = D

        Case surface_Hill
            D = X * X + Z * Z
            If D >= SphereRadius Then
                Y = 0
            ElseIf Z < 0 Then
                Y = 0.75 * Sqr(SphereRadius - D)
            Else
                Y = 0.75 * Sqr(SphereRadius - D) * (3 - Z) / 3
            End If

        Case surface_Ridge
            Y = 2 * Cos(2 * PI / 10 * Z) * (5 - Abs(Z)) / 5 + 0.5 * Rnd

        Case surface_PeakedRidge
            Y = 2 * Cos(2 * PI / 10 * Z) * (5 - Abs(Z)) / 5 + 0.25 * Sin(2 * X) + 0.25 * Sin(1# * X + x1) + 0.5 * Rnd

        Case surface_RuggedRidge
            Y = 2 * Cos(2 * PI / 10 * Z) * (5 - Abs(Z)) / 5 + Rnd

        Case surface_Random
            Y = Rnd

        Case surface_Valley
            Y = -2 * Cos(2 * PI / 10 * Z) * (5 - Abs(Z)) / 5 + 0.25 * Sin(2 * X) + 0.25 * Sin(1# * X + x1) + 0.5 * Rnd
            If Y < -1 Then Y = -1
    End Select

    YValue = Y
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
    TheGrid.RemoveHidden = (chkRemoveHidden.value = vbChecked)
    TheGrid.Draw pic
    pic.Refresh

    MousePointer = vbDefault
End Sub

Private Sub cmdDraw_Click()
    DrawData picCanvas
End Sub

Private Sub optSurface_Click(Index As Integer)
    SelectedSurface = Index
    DrawData picCanvas
    picCanvas.SetFocus
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

    m3PProject Projector, m3Parallel, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0
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
    Randomize

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
Const Dx = 1
Const Dz = 1
Const NumX = -2 * Xmin / Dx
Const NumZ = -2 * Zmin / Dz

Dim i As Integer
Dim j As Integer
Dim X As Single
Dim Y As Single
Dim Z As Single
Dim level As Integer
Dim Dy As Single

    SphereRadius = (Xmin + 3 * Dx) * (Xmin + 3 * Dx)

    Set TheGrid = New FractalGrid3d
    TheGrid.SetBounds Xmin, Dx, NumX, Zmin, Dz, NumZ

    X = Xmin
    For i = 1 To NumX
        Z = Zmin
        For j = 1 To NumZ
            Y = YValue(X, Z)
            TheGrid.SetValue X, Y, Z
            Z = Z + Dz
        Next j
        X = X + Dx
    Next i

    On Error Resume Next
    level = CInt(txtLevel.Text)
    If Err.Number <> 0 Then
        txtLevel.Text = "3"
        level = 3
    End If

    Dy = CSng(txtDy.Text)
    If Err.Number <> 0 Then
        txtDy.Text = "0.25"
        Dy = 0.25
    End If

    TheGrid.GenerateSurface level, Dy

    ' If this is the valley, flatten the bottom.
    If SelectedSurface = surface_Valley Then
        TheGrid.Flatten -1, 0.25, 0.25
    End If
End Sub
