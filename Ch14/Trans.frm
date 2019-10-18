VERSION 5.00
Begin VB.Form frmTrans 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Trans"
   ClientHeight    =   5790
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
   ScaleHeight     =   386
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   522
   Begin VB.CommandButton cmdTransform 
      Caption         =   "Transform"
      Default         =   -1  'True
      Height          =   495
      Left            =   600
      TabIndex        =   12
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transformations"
      Height          =   2535
      Left            =   0
      TabIndex        =   7
      Top             =   2640
      Width           =   2295
      Begin VB.OptionButton optTransformation 
         Caption         =   "Z Rotate"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   2055
      End
      Begin VB.OptionButton optTransformation 
         Caption         =   "Y Rotate"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   2055
      End
      Begin VB.OptionButton optTransformation 
         Caption         =   "X Rotate"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   2055
      End
      Begin VB.OptionButton optTransformation 
         Caption         =   "Grow, Rotate"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   2055
      End
      Begin VB.OptionButton optTransformation 
         Caption         =   "Wierd"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   2055
      End
      Begin VB.OptionButton optTransformation 
         Caption         =   "Up, Shrink/Grow"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton optTransformation 
         Caption         =   "Up, Shrink, Twist"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optTransformation 
         Caption         =   "Up, Shrink"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2055
      End
      Begin VB.OptionButton optTransformation 
         Caption         =   "Up, Twist"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Curve"
      Height          =   2535
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2295
      Begin VB.OptionButton optCurve 
         Caption         =   "Off Center Hexagon"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Hexagon"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   20
         Top             =   1920
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Semicircle"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Line Segment"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Star"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Off Center Circle"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Circle"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Off Center Square"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Square"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   5775
      Left            =   2400
      ScaleHeight     =   381
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   357
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmTrans"
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
Private Const dR = 1

' Location of focus point.
Private Const FocusX = 0#
Private Const FocusY = 0#
Private Const FocusZ = 0#

Private Projector(1 To 4, 1 To 4) As Single

Private SelectedCurve As Integer
Private SelectedTransformation As Integer

Private NumTrans As Integer
Private trans() As Transformation

Private TheSurface As Transformed3d
' Create the selected curve.
Private Sub CreateCurve()
Dim R As Single
Dim r2 As Single
Dim dtheta As Single
Dim theta As Single
Dim Y As Single
Dim i As Integer

    Select Case SelectedCurve
        Case 0  ' Line segment.
            TheSurface.AddCurvePoint -2, 0, 0
            TheSurface.AddCurvePoint 2, 0, 0

        Case 1  ' Square.
            TheSurface.AddCurvePoint -2, 0, -2
            TheSurface.AddCurvePoint -2, 0, 2
            TheSurface.AddCurvePoint 2, 0, 2
            TheSurface.AddCurvePoint 2, 0, -2
            TheSurface.AddCurvePoint -2, 0, -2

        Case 2  ' Off Center Square.
            TheSurface.AddCurvePoint 1, 0, 1
            TheSurface.AddCurvePoint 1, 0, 3
            TheSurface.AddCurvePoint 3, 0, 3
            TheSurface.AddCurvePoint 3, 0, 1
            TheSurface.AddCurvePoint 1, 0, 1

        Case 3  ' Circle.
            R = 2
            dtheta = PI / 8
            For theta = 0 To 2 * PI - dtheta + 0.01 Step dtheta
                TheSurface.AddCurvePoint R * Cos(theta), 0, R * Sin(theta)
            Next theta
            TheSurface.AddCurvePoint R, 0, 0

        Case 4  ' Off Center Circle.
            R = 1
            dtheta = PI / 8
            For theta = 0 To 2 * PI - dtheta + 0.01 Step dtheta
                TheSurface.AddCurvePoint 2 + R * Cos(theta), 0, 2 + R * Sin(theta)
            Next theta
            TheSurface.AddCurvePoint 2 + R, 0, 2

        Case 5  ' Star.
            R = 2
            r2 = 1
            dtheta = 2 * PI / 5 / 2
            theta = PI
            For i = 1 To 5
                TheSurface.AddCurvePoint _
                    R * Cos(theta), 0, R * Sin(theta)
                theta = theta + dtheta
                TheSurface.AddCurvePoint _
                    r2 * Cos(theta), 0, r2 * Sin(theta)
                theta = theta + dtheta
            Next i
            TheSurface.AddCurvePoint _
                R * Cos(PI), 0, R * Sin(PI)

        Case 6  ' Semicircle.
            R = 2
            dtheta = PI / 8
            For theta = 0 To PI - dtheta + 0.01 Step dtheta
                TheSurface.AddCurvePoint R * Cos(theta), 0, R * Sin(theta)
            Next theta
            TheSurface.AddCurvePoint -R, 0, 0

        Case 7  ' Hexagon.
            R = 3
            dtheta = 2 * PI / 6
            theta = 0
            For i = 1 To 7
                TheSurface.AddCurvePoint _
                    R * Cos(theta), 0, R * Sin(theta)
                theta = theta + dtheta
            Next i

        Case 8  ' Off Center Hexagon.
            R = 2
            dtheta = 2 * PI / 6
            theta = 0
            For i = 1 To 7
                TheSurface.AddCurvePoint _
                    R * Cos(theta), 0, R + R * Sin(theta)
                theta = theta + dtheta
            Next i

    End Select
End Sub
' Create the array of transformations.
Private Sub CreateTransformations()
Dim A(1 To 4, 1 To 4) As Single
Dim B(1 To 4, 1 To 4) As Single
Dim C(1 To 4, 1 To 4) As Single
Dim theta As Single
Dim dtheta As Single
Dim R As Single
Dim Y As Single
Dim i As Integer

    Select Case SelectedTransformation
        Case 0  ' Up, twist.
            NumTrans = 9
            ReDim trans(1 To NumTrans)
            dtheta = PI / 12
            For i = 1 To NumTrans
                Y = i / 2
                theta = i * dtheta
                m3Translate A, 0, Y, 0  ' Translate.
                m3YRotate B, theta      ' Rotate.
                m3MatMultiply trans(i).M, A, B  ' Combine.
            Next i

        Case 1  ' Up, shrink.
            NumTrans = 9
            ReDim trans(1 To NumTrans)
            For i = 1 To NumTrans
                Y = i / 2
                R = (NumTrans - i) / NumTrans
                m3Scale A, R, 1, R      ' Scale.
                m3Translate B, 0, Y, 0  ' Translate.
                m3MatMultiply trans(i).M, A, B  ' Combine.
            Next i

        Case 2  ' Up, shrink, twist.
            NumTrans = 9
            ReDim trans(1 To NumTrans)
            dtheta = PI / 12
            For i = 1 To NumTrans
                Y = i / 2
                R = (NumTrans - i) / NumTrans
                theta = i * dtheta
                m3Scale A, R, 1, R      ' Scale.
                m3Translate B, 0, Y, 0  ' Translate.
                m3MatMultiply C, A, B   ' Combine A and B.
                m3YRotate A, theta      ' Rotate.
                m3MatMultiply trans(i).M, C, A  ' Combine all.
            Next i

        Case 3  ' Up, shrink/grow.
            NumTrans = 18
            ReDim trans(1 To NumTrans)
            dtheta = PI / 12
            For i = 1 To NumTrans
                Y = i / 4
                theta = i * dtheta
                R = 1 + Sin(2 * theta) / 2
                m3Scale A, R, 1, R      ' Scale.
                m3Translate B, 0, Y, 0  ' Translate.
                m3MatMultiply trans(i).M, A, B  ' Combine.
            Next i

        Case 4  ' Waver.
            ' Make the curve move upwards with
            ' varying rotation around the Z axis.
            NumTrans = 18
            ReDim trans(1 To NumTrans)
            dtheta = PI / 12
            R = PI / 2
            For i = 1 To NumTrans
                Y = i / 4
                theta = i * dtheta
                m3ZRotate A, R * Sin(theta)  ' Rotate.
                m3Translate B, 0, Y, 0  ' Translate.
                m3MatMultiply trans(i).M, A, B  ' Combine.
            Next i

        Case 5  ' Grow and rotate.
            ' Make the curve grow and rotate
            ' around the Z axis.
            NumTrans = 18
            ReDim trans(1 To NumTrans)
            dtheta = PI / 12
            R = PI / 2
            For i = 1 To NumTrans
                Y = i / 4
                theta = i * dtheta
                m3ZRotate A, R * Sin(theta)     ' Rotate.
                m3Scale B, i / 9, i / 9, i / 9  ' Scale
                m3MatMultiply trans(i).M, A, B  ' Combine.
            Next i

        Case 6  ' Rotate around the X axis.
            ' Rotate around the X axis.
            NumTrans = 18
            ReDim trans(1 To NumTrans)
            dtheta = 2 * PI / NumTrans
            R = PI / 2
            For i = 1 To NumTrans
                Y = i / 4
                theta = i * dtheta
                m3XRotate trans(i).M, theta ' Rotate.
            Next i

        Case 7  ' Rotate around the Y axis.
            ' Rotate around the Y axis.
            NumTrans = 18
            ReDim trans(1 To NumTrans)
            dtheta = 2 * PI / NumTrans
            R = PI / 2
            For i = 1 To NumTrans
                Y = i / 4
                theta = i * dtheta
                m3YRotate trans(i).M, theta ' Rotate.
            Next i

        Case 8  ' Rotate around the Z axis.
            ' Rotate around the Z axis.
            NumTrans = 18
            ReDim trans(1 To NumTrans)
            dtheta = 2 * PI / NumTrans
            R = PI / 2
            For i = 1 To NumTrans
                Y = i / 4
                theta = i * dtheta
                m3ZRotate trans(i).M, theta ' Rotate.
            Next i

    End Select
End Sub


' Create the surface.
Private Sub cmdTransform_Click()
Dim i As Integer

    Screen.MousePointer = vbHourglass
    DoEvents

    Set TheSurface = New Transformed3d

    CreateCurve
    CreateTransformations

    For i = 1 To NumTrans
        TheSurface.SetTransformation trans(i).M
    Next i
    TheSurface.Transform

    DrawData picCanvas
    picCanvas.SetFocus
    Screen.MousePointer = vbDefault
End Sub
' Save the curve choice.
Private Sub optCurve_Click(Index As Integer)
    SelectedCurve = Index
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
    m3Translate T, 180, 250, 0
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
            EyeR = EyeR + dR
        
        Case Asc("-")
            EyeR = EyeR - dR
        
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


' Save the current transformation choice.
Private Sub optTransformation_Click(Index As Integer)
    SelectedTransformation = Index
End Sub
