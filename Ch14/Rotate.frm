VERSION 5.00
Begin VB.Form frmRotate 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Rotate"
   ClientHeight    =   5310
   ClientLeft      =   690
   ClientTop       =   615
   ClientWidth     =   7830
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
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   522
   Begin VB.Frame Frame2 
      Caption         =   "Curve"
      Height          =   5295
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2295
      Begin VB.OptionButton optCurve 
         Caption         =   "Tornado"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   15
         Top             =   4920
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Helix"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   14
         Top             =   4560
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Tower"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   13
         Top             =   4200
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Football"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   12
         Top             =   3840
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Goblet"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   11
         Top             =   3480
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Urn"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   10
         Top             =   3120
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Sine Wave"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   2760
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Semicircle 2"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Semicircle 1"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Circle 2"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Circle 1"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "3/4 Rectangle"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Diamond"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Rectangle"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   5295
      Left            =   2400
      ScaleHeight     =   349
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   357
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmRotate"
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

Private SelectedCurve As Integer

Private TheSurface As Rotated3d
' Create the selected surface.
Private Sub CreateSurface()
Dim R As Single
Dim offset As Single
Dim dtheta As Single
Dim theta As Single
Dim Y As Single

    Set TheSurface = New Rotated3d

    Select Case SelectedCurve
        Case 0  ' Rectangle.
            TheSurface.AddCurvePoint -3, -1.5, 0
            TheSurface.AddCurvePoint -3, 1.5, 0
            TheSurface.AddCurvePoint -1, 1.5, 0
            TheSurface.AddCurvePoint -1, -1.5, 0
            TheSurface.AddCurvePoint -3, -1.5, 0

        Case 1  ' Diamond.
            TheSurface.AddCurvePoint -3, 0, 0
            TheSurface.AddCurvePoint -2, -1, 0
            TheSurface.AddCurvePoint -1, 0, 0
            TheSurface.AddCurvePoint -2, 1, 0
            TheSurface.AddCurvePoint -3, 0, 0

        Case 2  ' 3/4 Rectangle.
            TheSurface.AddCurvePoint 0, -1.5, 0
            TheSurface.AddCurvePoint -3, -1.5, 0
            TheSurface.AddCurvePoint -3, 1.5, 0
            TheSurface.AddCurvePoint 0, 1.5, 0
            TheSurface.AddCurvePoint 0, -1.5, 0

        Case 3, 4   ' Circle 1, circle 2.
            If SelectedCurve = 3 Then
                R = 2
                offset = 2
            Else
                R = 1.5
                offset = 2.5
            End If
            dtheta = PI / 8
            TheSurface.AddCurvePoint offset + R, 0, 0
            For theta = dtheta To 2 * PI - dtheta + 0.1 Step dtheta
                TheSurface.AddCurvePoint _
                    offset + R * Cos(theta), R * Sin(theta), 0
            Next theta
            TheSurface.AddCurvePoint offset + R, 0, 0

        Case 5, 6   ' Semicircle 1, semicircle 2.
            If SelectedCurve = 5 Then
                R = 4
                offset = 0
            Else
                R = 2
                offset = 2
            End If
            dtheta = PI / 8
            TheSurface.AddCurvePoint offset, -R, 0
            For theta = -PI / 2 + dtheta To PI / 2 - dtheta + 0.1 Step dtheta
                TheSurface.AddCurvePoint _
                    offset + R * Cos(theta), _
                    R * Sin(theta), _
                    0
            Next theta
            TheSurface.AddCurvePoint offset, R, 0

        Case 7  ' Sine wave.
            R = 0.7
            dtheta = PI / 10
            For theta = -PI To PI Step dtheta
                TheSurface.AddCurvePoint _
                    1 + R + R * Sin(2 * theta), _
                    theta, _
                    0
            Next theta

        Case 8  ' Urn.
            dtheta = PI / 10
            For theta = -PI To PI Step dtheta
                TheSurface.AddCurvePoint _
                    PI / 2 + (-PI + theta) / 4 * Sin(2 * theta), _
                    theta, _
                    0
            Next theta

        Case 9  ' Goblet.
            TheSurface.AddCurvePoint 3, 3.5, 0
            TheSurface.AddCurvePoint 2.5, 3, 0
            TheSurface.AddCurvePoint 3, 1.5, 0
            TheSurface.AddCurvePoint 2.5, 1, 0
            TheSurface.AddCurvePoint 1, 1, 0
            TheSurface.AddCurvePoint 0.5, 0.5, 0
            TheSurface.AddCurvePoint 0.5, -1, 0
            TheSurface.AddCurvePoint 1, -1.5, 0
            TheSurface.AddCurvePoint 2, -1.5, 0
            TheSurface.AddCurvePoint 2.5, -2, 0

        Case 10 ' Football.
            For Y = -4 To 4 Step 0.5
                TheSurface.AddCurvePoint 16 / 5 - Y * Y / 5, Y, 0
            Next Y

        Case 11 ' Tower.
            R = 1
            dtheta = PI / 8
            For theta = -PI To -PI / 2 Step dtheta
                TheSurface.AddCurvePoint _
                    R + R * Cos(theta), _
                    4 * R + R * Sin(theta), _
                    0
            Next theta
            For theta = PI / 2 To -PI / 2 Step -dtheta
                TheSurface.AddCurvePoint _
                    R + R * Cos(theta), _
                    2 * R + R * Sin(theta), _
                    0
            Next theta
            TheSurface.AddCurvePoint R, -3, 0

        Case 12 ' Helix.
            R = 2
            dtheta = PI / 4
            For theta = -PI To PI Step dtheta
                TheSurface.AddCurvePoint _
                    R * Cos(theta / 2), _
                    theta, _
                    R * Sin(theta / 2)
            Next theta

        Case 13 ' Tornado.
            R = 2
            dtheta = PI / 4
            For theta = -PI To PI Step dtheta
                R = 2 + theta / 2
                TheSurface.AddCurvePoint _
                    R * Cos(theta / 2), _
                    theta, _
                    R * Sin(theta / 2)
            Next theta

    End Select
End Sub

' Create a new curve and rotate it.
Private Sub optCurve_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    DoEvents

    SelectedCurve = Index
    CreateSurface
    TheSurface.Rotate

    DrawData picCanvas
    picCanvas.SetFocus

    Screen.MousePointer = vbDefault
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
    
    ' Scale and translate so it looks OK in pixels.
    m3Scale S, 35, -35, 1
    m3Translate T, 180, 200, 0
    m3MatMultiplyFull ST, S, T
    m3MatMultiplyFull PST, Projector, ST

    ' Transform the points.
    TheSurface.ApplyFull PST

    ' Display the data.
    pic.Cls
    TheSurface.Draw pic, EyeR
    pic.Refresh
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
    EyePhi = PI * 0.1
    
    ' Initialize the projection transformation.
    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0

    Me.Show
    optCurve_Click 0
End Sub
