VERSION 5.00
Begin VB.Form frmExtrude 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Extrude"
   ClientHeight    =   5550
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
   ScaleHeight     =   370
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   522
   Begin VB.CommandButton cmdExtrude 
      Caption         =   "Extrude"
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   5040
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Curve"
      Height          =   2415
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2295
      Begin VB.OptionButton optCurve 
         Caption         =   "Sine Wave"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Small Circle"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Circle"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Semicircle"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Rectangle"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "Square"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Path"
      Height          =   2415
      Left            =   0
      TabIndex        =   1
      Top             =   2520
      Width           =   2295
      Begin VB.OptionButton optPath 
         Caption         =   "Circle"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   2055
      End
      Begin VB.OptionButton optPath 
         Caption         =   "Semicircle"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   2055
      End
      Begin VB.OptionButton optPath 
         Caption         =   "Helix"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   2055
      End
      Begin VB.OptionButton optPath 
         Caption         =   "Wavy Line"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton optPath 
         Caption         =   "Angled Line"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton optPath 
         Caption         =   "Vertical Line"
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
      Height          =   5535
      Left            =   2400
      ScaleHeight     =   365
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   357
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmExtrude"
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
Private SelectedPath As Integer

Private TheExtrusion As Extrusion3d
' Create the selected path.
Private Sub CreatePath()
Dim Y As Single
Dim R As Single
Dim dtheta As Single
Dim theta As Single

    Select Case SelectedPath
        Case 0  ' Vertical line.
            For Y = 0 To 4 Step 0.5
                TheExtrusion.AddPathPoint 0, Y, 0
            Next Y

        Case 1  ' Angled line.
            For Y = 0 To 4 Step 0.5
                TheExtrusion.AddPathPoint Y / 2, Y, Y / 2
            Next Y
        
        Case 2  ' Wavy line.
            R = 2
            dtheta = PI / 5
            TheExtrusion.AddPathPoint 0, 0, 0
            For theta = dtheta To 2 * PI Step dtheta
                TheExtrusion.AddPathPoint _
                    R * Sin(theta), theta * 0.7, 0
            Next theta
        
        Case 3  ' Helix.
            R = 2
            dtheta = PI / 10
            TheExtrusion.AddPathPoint R, 0, 0
            For theta = dtheta To 2 * PI Step dtheta
                TheExtrusion.AddPathPoint _
                    R * Cos(theta), _
                    theta * 0.8, _
                    R * Sin(theta)
            Next theta
            
        Case 4  ' Semicircle.
            R = 2
            dtheta = PI / 10
            TheExtrusion.AddPathPoint 0, 0, 0
            For theta = dtheta To PI Step dtheta
                TheExtrusion.AddPathPoint _
                    R * Sin(theta), _
                    R * (1 - Cos(theta)), _
                    0
            Next theta
        
        Case 5  ' Circle.
            R = 2
            dtheta = PI / 10
            TheExtrusion.AddPathPoint 0, 0, 0
            For theta = dtheta To 2 * PI - dtheta + 0.01 Step dtheta
                TheExtrusion.AddPathPoint _
                    R * Sin(theta), _
                    R * (1 - Cos(theta)), _
                    0
            Next theta
            TheExtrusion.AddPathPoint 0, 0, 0
    
    End Select
End Sub
' Create the selected curve.
Private Sub CreateCurve()
Dim R As Single
Dim dtheta As Single
Dim theta As Single

    Select Case SelectedCurve
        Case 0  ' Square.
            TheExtrusion.AddCurvePoint -2, 0, -2
            TheExtrusion.AddCurvePoint -2, 0, 2
            TheExtrusion.AddCurvePoint 2, 0, 2
            TheExtrusion.AddCurvePoint 2, 0, -2
            TheExtrusion.AddCurvePoint -2, 0, -2

        Case 1  ' Rectangle.
            TheExtrusion.AddCurvePoint -0.5, 0, -2
            TheExtrusion.AddCurvePoint -0.5, 0, 2
            TheExtrusion.AddCurvePoint 0.5, 0, 2
            TheExtrusion.AddCurvePoint 0.5, 0, -2
            TheExtrusion.AddCurvePoint -0.5, 0, -2
        
        Case 2  ' Semicircle.
            R = 2
            dtheta = PI / 10
            TheExtrusion.AddCurvePoint R, 0, 0
            For theta = dtheta To PI Step dtheta
                TheExtrusion.AddCurvePoint _
                    R * Cos(theta), 0, R * Sin(theta)
            Next theta
            
        Case 3, 4   ' Circle, small circle.
            If SelectedCurve = 3 Then
                R = 2
                dtheta = PI / 10
            Else
                R = 0.5
                dtheta = PI / 4
            End If
            TheExtrusion.AddCurvePoint R, 0, 0
            For theta = dtheta To 2 * PI - dtheta + 0.1 Step dtheta
                TheExtrusion.AddCurvePoint _
                    R * Cos(theta), 0, R * Sin(theta)
            Next theta
            TheExtrusion.AddCurvePoint R, 0, 0
            
        Case 5  ' Sine wave.
            R = 2
            dtheta = PI / 10
            theta = -PI / 2
            TheExtrusion.AddCurvePoint _
                R * Sin(theta), 0, 2 * theta
            For theta = -PI / 2 + dtheta To PI / 2 Step dtheta
                TheExtrusion.AddCurvePoint _
                    R * Sin(theta), 0, 2 * theta
            Next theta
            
    End Select
End Sub




' Create the extruded data and display it.
Private Sub cmdExtrude_Click()
Dim pline As Polyline3d

    ' If we currently have an APF file loaded,
    ' restore default settings.
    If SelectedCurve > 5 Then
        optCurve(0).value = True
        optPath(0).value = True
    End If
    
    Set TheExtrusion = New Extrusion3d

    CreateCurve
    CreatePath
    TheExtrusion.Extrude

    DrawData picCanvas
    picCanvas.SetFocus
End Sub
Private Sub optCurve_Click(Index As Integer)
    SelectedCurve = Index
    picCanvas.SetFocus
End Sub

' Draw the data.
Private Sub DrawData(pic As Object)
Dim X As Single
Dim Y As Single
Dim Z As Single
Dim S(1 To 4, 1 To 4) As Single
Dim T(1 To 4, 1 To 4) As Single
Dim ST(1 To 4, 1 To 4) As Single
Dim PST(1 To 4, 1 To 4) As Single

    Screen.MousePointer = vbHourglass
    DoEvents
    
    ' Prevent overflow errors when drawing lines
    ' too far out of bounds.
    On Error Resume Next
    
    ' Scale and translate so it looks OK in pixels.
    m3Scale S, 35, -35, 1
    m3Translate T, 180, 250, 0
    m3MatMultiplyFull ST, S, T
    m3MatMultiplyFull PST, Projector, ST
    
    ' Transform the points.
    TheExtrusion.ApplyFull PST

    ' Display the data.
    pic.Cls
    TheExtrusion.Draw pic, EyeR
    pic.Refresh

    Screen.MousePointer = vbDefault
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
End Sub


Private Sub optPath_Click(Index As Integer)
    SelectedPath = Index
    picCanvas.SetFocus
End Sub
