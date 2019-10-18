VERSION 5.00
Begin VB.Form frmPlatonic 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Platonic"
   ClientHeight    =   4230
   ClientLeft      =   1395
   ClientTop       =   1140
   ClientWidth     =   5850
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
   ScaleHeight     =   4230
   ScaleWidth      =   5850
   Begin VB.CheckBox Choice 
      Caption         =   "Dodecahedron"
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CheckBox Choice 
      Caption         =   "Icosahedron"
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CheckBox Choice 
      Caption         =   "Cube"
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox Choice 
      Caption         =   "Octahedron"
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CheckBox Choice 
      Caption         =   "Axes"
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox Choice 
      Caption         =   "Tetrahedron"
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   4215
      Left            =   0
      ScaleHeight     =   -4
      ScaleLeft       =   -2
      ScaleMode       =   0  'User
      ScaleTop        =   2
      ScaleWidth      =   4
      TabIndex        =   6
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmPlatonic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Location of viewing eye.
Private EyeR As Single
Private EyeTheta As Single
Private EyePhi As Single

' Location of focus point.
Private Const FocusX = 0#
Private Const FocusY = 0#
Private Const FocusZ = 0#

Private Projector(1 To 4, 1 To 4) As Single

Private FirstTet As Integer
Private FirstCube As Integer
Private FirstOct As Integer
Private FirstDod As Integer
Private FirstIco As Integer
Private LastIco As Integer
' Project and draw the cube.
Private Sub DrawData(pic As Object)
Dim i As Integer

    ' Transform the points.
    TransformAllDataFull Projector

    ' Draw the points.
    pic.Cls

    If Choice(0).value = vbChecked Then DrawSomeData pic, 1, FirstTet - 1, vbBlack, False
    If Choice(1).value = vbChecked Then DrawSomeData pic, FirstTet, FirstCube - 1, vbRed, False
    If Choice(2).value = vbChecked Then DrawSomeData pic, FirstCube, FirstOct - 1, RGB(0, 128, 0), False
    If Choice(3).value = vbChecked Then DrawSomeData pic, FirstOct, FirstDod - 1, vbBlue, False
    If Choice(4).value = vbChecked Then DrawSomeData pic, FirstDod, FirstIco - 1, vbMagenta, False
    If Choice(5).value = vbChecked Then DrawSomeData pic, FirstIco, LastIco, RGB(0, 128, 128), False
    
    pic.Refresh
End Sub


Private Sub Choice_Click(Index As Integer)
    DrawData picCanvas
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Const Dtheta = PI / 20

    Select Case KeyCode
        Case vbKeyLeft
            EyeTheta = EyeTheta - Dtheta

        Case vbKeyRight
            EyeTheta = EyeTheta + Dtheta

        Case vbKeyUp
            EyePhi = EyePhi - Dtheta

        Case vbKeyDown
            EyePhi = EyePhi + Dtheta

        Case Else
            Exit Sub
    End Select

    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0
    DrawData picCanvas
End Sub

Private Sub Form_Load()
    ' Initialize the eye position.
    EyeR = 5
    EyeTheta = PI * 0.4
    EyePhi = PI * 0.1
    
    ' Initialize the projection transformation.
    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0
    
    ' Create the data.
    CreateData

    ' Project and draw the data.
    DrawData picCanvas
End Sub

' Create the data.
Private Sub CreateData()
Dim theta1 As Single
Dim theta2 As Single
Dim s1 As Single
Dim s2 As Single
Dim c1 As Single
Dim c2 As Single
Dim S As Single
Dim R As Single
Dim H As Single
Dim A As Single
Dim B As Single
Dim C As Single
Dim D As Single
Dim X As Single
Dim Y As Single
Dim y2 As Single
Dim M As Single
Dim N As Single

    ' Axes.
    MakeSegment 0, 0, 0, 0.5, 0, 0  ' X axis.
    MakeSegment 0, 0, 0, 0, 0.5, 0  ' Y axis.
    MakeSegment 0, 0, 0, 0, 0, 0.5  ' Z axis.
    
    ' Tetrahedron.
    FirstTet = NumSegments + 1
    S = Sqr(6)
    A = S / Sqr(3)
    B = -A / 2
    C = A * Sqr(2) - 1
    D = S / 2
    MakeSegment 0, C, 0, A, -1, 0
    MakeSegment 0, C, 0, B, -1, D
    MakeSegment 0, C, 0, B, -1, -D
    MakeSegment B, -1, -D, B, -1, D
    MakeSegment B, -1, D, A, -1, 0
    MakeSegment A, -1, 0, B, -1, -D
    
    ' Cube.
    FirstCube = NumSegments + 1
    MakeSegment -1, -1, -1, -1, 1, -1
    MakeSegment -1, 1, -1, 1, 1, -1
    MakeSegment 1, 1, -1, 1, -1, -1
    MakeSegment 1, -1, -1, -1, -1, -1
    
    MakeSegment -1, -1, 1, -1, 1, 1
    MakeSegment -1, 1, 1, 1, 1, 1
    MakeSegment 1, 1, 1, 1, -1, 1
    MakeSegment 1, -1, 1, -1, -1, 1
    
    MakeSegment -1, -1, -1, -1, -1, 1
    MakeSegment -1, 1, -1, -1, 1, 1
    MakeSegment 1, 1, -1, 1, 1, 1
    MakeSegment 1, -1, -1, 1, -1, 1
    
    ' Octahedron.
    FirstOct = NumSegments + 1
    MakeSegment 0, 1, 0, 1, 0, 0
    MakeSegment 0, 1, 0, -1, 0, 0
    MakeSegment 0, 1, 0, 0, 0, 1
    MakeSegment 0, 1, 0, 0, 0, -1
    
    MakeSegment 0, -1, 0, 1, 0, 0
    MakeSegment 0, -1, 0, -1, 0, 0
    MakeSegment 0, -1, 0, 0, 0, 1
    MakeSegment 0, -1, 0, 0, 0, -1
    
    MakeSegment 0, 0, 1, 1, 0, 0
    MakeSegment 0, 0, 1, -1, 0, 0
    MakeSegment 0, 0, -1, 1, 0, 0
    MakeSegment 0, 0, -1, -1, 0, 0
    
    ' Dodecahedron.
    FirstDod = NumSegments + 1
    theta1 = PI * 0.4
    theta2 = PI * 0.8
    s1 = Sin(theta1)
    c1 = Cos(theta1)
    s2 = Sin(theta2)
    c2 = Cos(theta2)
    
    M = 1 - (2 - 2 * c1 - 4 * s1 * s1) / (2 * c1 - 2)
    N = Sqr((2 - 2 * c1) - M * M) * (1 + (1 - c2) / (c1 - c2))
    R = 2 / N
    S = R * Sqr(2 - 2 * c1)
    A = R * s1
    B = R * s2
    C = R * c1
    D = R * c2
    H = R * (c1 - s1)
    
    X = (R * R * (2 - 2 * c1) - 4 * A * A) / (2 * C - 2 * R)
    Y = Sqr(S * S - (R - X) * (R - X))
    y2 = Y * (1 - c2) / (c1 - c2)
    
    MakeSegment R, 1, 0, C, 1, A        ' Top
    MakeSegment C, 1, A, D, 1, B
    MakeSegment D, 1, B, D, 1, -B
    MakeSegment D, 1, -B, C, 1, -A
    MakeSegment C, 1, -A, R, 1, 0
    
    MakeSegment R, 1, 0, X, 1 - Y, 0    ' Top downward edges.
    MakeSegment C, 1, A, X * c1, 1 - Y, X * s1
    MakeSegment C, 1, -A, X * c1, 1 - Y, -X * s1
    MakeSegment D, 1, B, X * c2, 1 - Y, X * s2
    MakeSegment D, 1, -B, X * c2, 1 - Y, -X * s2
    
    MakeSegment X, 1 - Y, 0, -X * c2, 1 - y2, -X * s2   ' Middle.
    MakeSegment X, 1 - Y, 0, -X * c2, 1 - y2, X * s2
    MakeSegment X * c1, 1 - Y, X * s1, -X * c2, 1 - y2, X * s2
    MakeSegment X * c1, 1 - Y, X * s1, -X * c1, 1 - y2, X * s1
    MakeSegment X * c2, 1 - Y, X * s2, -X * c1, 1 - y2, X * s1
    MakeSegment X * c2, 1 - Y, X * s2, -X, 1 - y2, 0
    MakeSegment X * c2, 1 - Y, -X * s2, -X, 1 - y2, 0
    MakeSegment X * c2, 1 - Y, -X * s2, -X * c1, 1 - y2, -X * s1
    MakeSegment X * c1, 1 - Y, -X * s1, -X * c1, 1 - y2, -X * s1
    MakeSegment X * c1, 1 - Y, -X * s1, -X * c2, 1 - y2, -X * s2
        
    MakeSegment -R, -1, 0, -X, 1 - y2, 0    ' Bottom upward edges.
    MakeSegment -C, -1, A, -X * c1, 1 - y2, X * s1 ' Bottom upward edges.
    MakeSegment -D, -1, B, -X * c2, 1 - y2, X * s2
    MakeSegment -D, -1, -B, -X * c2, 1 - y2, -X * s2
    MakeSegment -C, -1, -A, -X * c1, 1 - y2, -X * s1
    
    MakeSegment -R, -1, 0, -C, -1, A    ' Bottom
    MakeSegment -C, -1, A, -D, -1, B
    MakeSegment -D, -1, B, -D, -1, -B
    MakeSegment -D, -1, -B, -C, -1, -A
    MakeSegment -C, -1, -A, -R, -1, 0
    
    ' Icosahedron.
    FirstIco = NumSegments + 1
    R = 2 / (2 * Sqr(1 - 2 * c1) + Sqr(3 / 4 * (2 - 2 * c1) - 2 * c2 - c2 * c2 - 1))
    S = R * Sqr(2 - 2 * c1)
    H = 1 - Sqr(S * S - R * R)
    A = R * s1
    B = R * s2
    C = R * c1
    D = R * c2
    MakeSegment R, H, 0, C, H, A        ' Top
    MakeSegment C, H, A, D, H, B
    MakeSegment D, H, B, D, H, -B
    MakeSegment D, H, -B, C, H, -A
    MakeSegment C, H, -A, R, H, 0
    MakeSegment R, H, 0, 0, 1, 0        ' Point
    MakeSegment C, H, A, 0, 1, 0
    MakeSegment D, H, B, 0, 1, 0
    MakeSegment D, H, -B, 0, 1, 0
    MakeSegment C, H, -A, 0, 1, 0
    
    MakeSegment -R, -H, 0, -C, -H, A    ' Bottom
    MakeSegment -C, -H, A, -D, -H, B
    MakeSegment -D, -H, B, -D, -H, -B
    MakeSegment -D, -H, -B, -C, -H, -A
    MakeSegment -C, -H, -A, -R, -H, 0
    MakeSegment -R, -H, 0, 0, -1, 0     ' Point
    MakeSegment -C, -H, A, 0, -1, 0
    MakeSegment -D, -H, B, 0, -1, 0
    MakeSegment -D, -H, -B, 0, -1, 0
    MakeSegment -C, -H, -A, 0, -1, 0

    MakeSegment R, H, 0, -D, -H, B      ' Middle
    MakeSegment R, H, 0, -D, -H, -B
    MakeSegment C, H, A, -D, -H, B
    MakeSegment C, H, A, -C, -H, A
    MakeSegment D, H, B, -C, -H, A
    MakeSegment D, H, B, -R, -H, 0
    MakeSegment D, H, -B, -R, -H, 0
    MakeSegment D, H, -B, -C, -H, -A
    MakeSegment C, H, -A, -C, -H, -A
    MakeSegment C, H, -A, -D, -H, -B
    LastIco = NumSegments

    If Not SameSideLengths(FirstTet, FirstCube - 1) Then MsgBox "Error in tetrahedron."
    If Not SameSideLengths(FirstCube, FirstOct - 1) Then MsgBox "Error in cube."
    If Not SameSideLengths(FirstOct, FirstDod - 1) Then MsgBox "Error in octahedron."
    If Not SameSideLengths(FirstDod, FirstIco - 1) Then MsgBox "Error in dodecahedron."
    If Not SameSideLengths(FirstIco, LastIco - 1) Then MsgBox "Error in icosahedron."
End Sub
