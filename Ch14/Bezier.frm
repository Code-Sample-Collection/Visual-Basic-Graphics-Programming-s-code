VERSION 5.00
Begin VB.Form frmBezier 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Bezier"
   ClientHeight    =   5310
   ClientLeft      =   300
   ClientTop       =   555
   ClientWidth     =   9150
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
   ScaleHeight     =   5310
   ScaleWidth      =   9150
   Begin VB.CheckBox chkShowControlPoints 
      Caption         =   "Show Control Points"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Spiral"
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   10
      Top             =   3480
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Twist"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   9
      Top             =   3120
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Cowling"
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   8
      Top             =   2760
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Pipe"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   7
      Top             =   2400
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Curl"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Wave"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Hill"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   2055
   End
   Begin VB.CheckBox chkShowControlGrid 
      Caption         =   "Show Control Grid"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Tent"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   2055
   End
   Begin VB.OptionButton optSurface 
      Caption         =   "Urn"
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   1
      Top             =   3840
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
Attribute VB_Name = "frmBezier"
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

Private TheSurface As Bezier3d

Private ShowingParameters As Boolean

Private SurfaceSelected As Integer
' Display the surface.
Private Sub DrawData(pic As Object)
Dim S(1 To 4, 1 To 4) As Single
Dim T(1 To 4, 1 To 4) As Single
Dim ST(1 To 4, 1 To 4) As Single
Dim PST(1 To 4, 1 To 4) As Single

    MousePointer = vbHourglass
    Refresh

    ' Scale and translate so it looks OK in pixels.
    m3Scale S, 35, -35, 1
    m3Translate T, 230, 175, 0
    m3MatMultiplyFull ST, S, T
    m3MatMultiplyFull PST, Projector, ST

    ' Transform the points.
    TheSurface.ApplyFull PST

    ' Prevent overflow errors when drawing lines
    ' too far out of bounds.
    On Error Resume Next

    ' Display the data.
    pic.Cls
    TheSurface.Draw pic, EyeR

    picCanvas.SetFocus
    MousePointer = vbDefault
End Sub
' Set the control points for an urn.
Private Sub MakeUrn()
Dim R(1 To 5) As Single
Dim h(1 To 5) As Single
Dim i As Integer

    TheSurface.SetBounds 5, 6

    R(1) = 1
    R(2) = 1
    R(3) = 5
    R(4) = 1.5
    R(5) = 1.5

    h(1) = 4
    h(2) = 3.5
    h(3) = 2
    h(4) = -1
    h(5) = -3

    For i = 1 To 5
        TheSurface.SetControlPoint i, 1, -R(i), h(i), 0
        TheSurface.SetControlPoint i, 2, -R(i), h(i), -1.5 * R(i)
        TheSurface.SetControlPoint i, 3, 2 * R(i), h(i), -1.5 * R(i)
        TheSurface.SetControlPoint i, 4, 2 * R(i), h(i), 1.5 * R(i)
        TheSurface.SetControlPoint i, 5, -R(i), h(i), 1.5 * R(i)
        TheSurface.SetControlPoint i, 6, -R(i), h(i), 0
    Next i
End Sub
' Set the control points for a pipe.
Private Sub MakePipe()
Const S = 3

Dim i As Integer
Dim X As Single

    TheSurface.SetBounds 4, 6

    For i = 1 To 4
        X = 1.5 * (i - 2.5)
        TheSurface.SetControlPoint i, 1, X, _
            -S, 0
        TheSurface.SetControlPoint i, 2, X, _
            -S, -S
        TheSurface.SetControlPoint i, 3, X, _
            S, -S
        TheSurface.SetControlPoint i, 4, X, _
            S, S
        TheSurface.SetControlPoint i, 5, X, _
            -S, S
        TheSurface.SetControlPoint i, 6, X, _
            -S, 0
    Next i
End Sub
' Set the control points for a curl.
Private Sub MakeCurl()
Dim ang As Integer
Dim j As Integer
Dim R As Single
Dim X As Single
Dim Y As Single
Dim Z As Single

    TheSurface.SetBounds 4, 4

    For j = 1 To 4
        Z = 1.5 * (j - 2.5)
        R = 6 - Abs(2 * j - 5)
        For ang = 1 To 4
            X = R * Cos((ang - 1) * PI / 2)
            Y = R * Sin((ang - 1) * PI / 2)
            TheSurface.SetControlPoint ang, j, X, Y, Z
        Next ang
    Next j
End Sub
' Set the control points for a wave.
Private Sub MakeWave()
Dim i As Integer
Dim j As Integer

    TheSurface.SetBounds 4, 4

    ' Start flat and modify from there.
    For i = 1 To 4
        For j = 1 To 4
            TheSurface.SetControlPoint i, j, 2 * i - 5, 0, 2 * j - 5
        Next j
    Next i

    ' Make the modifications.
    TheSurface.SetControlPoint 2, 2, -1, -10, -1
    TheSurface.SetControlPoint 2, 3, -1, 10, 1
    TheSurface.SetControlPoint 3, 2, 1, -10, -1
    TheSurface.SetControlPoint 3, 3, 1, 10, 1
End Sub
' Set the control points for a tent.
Private Sub MakeTent()
    TheSurface.SetBounds 3, 3

    TheSurface.SetControlPoint 1, 1, -3, -2, -3
    TheSurface.SetControlPoint 1, 2, -3, 2, 0
    TheSurface.SetControlPoint 1, 3, -3, -2, 3
    TheSurface.SetControlPoint 2, 1, 0, 2, -3
    TheSurface.SetControlPoint 2, 2, 0, 4, 0
    TheSurface.SetControlPoint 2, 3, 0, 2, 3
    TheSurface.SetControlPoint 3, 1, 3, -2, -3
    TheSurface.SetControlPoint 3, 2, 3, 2, 0
    TheSurface.SetControlPoint 3, 3, 3, -2, 3
End Sub




' Set the control points for a spiral.
Private Sub MakeSpiral()
    TheSurface.SetBounds 5, 2
    
    TheSurface.SetControlPoint 1, 1, -4, 2, 0
    TheSurface.SetControlPoint 1, 2, -4, -2, 0
    TheSurface.SetControlPoint 2, 1, -2, 0, -4
    TheSurface.SetControlPoint 2, 2, -2, 0, 4
    TheSurface.SetControlPoint 3, 1, 0, -6, 0
    TheSurface.SetControlPoint 3, 2, 0, 6, 0
    TheSurface.SetControlPoint 4, 1, 2, 0, 4
    TheSurface.SetControlPoint 4, 2, 2, 0, -4
    TheSurface.SetControlPoint 5, 1, 4, 2, 0
    TheSurface.SetControlPoint 5, 2, 4, -2, 0
End Sub

' Set the control points for a twist.
Private Sub MakeTwist()
    TheSurface.SetBounds 2, 2
    
    TheSurface.SetControlPoint 1, 1, -2, 3, 3
    TheSurface.SetControlPoint 1, 2, -3, 3, -3
    TheSurface.SetControlPoint 2, 1, 3, 4, -2
    TheSurface.SetControlPoint 2, 2, 2, -3, 0
End Sub


' Set the control points for a cowling.
Private Sub MakeCowl()
Dim i As Integer
Dim S As Single
Dim Y As Single

    TheSurface.SetBounds 4, 6
    
    For i = 1 To 4
        Y = 3 - 2 * Abs(i - 2.5)
        
        S = 2 + i / 2
        
        TheSurface.SetControlPoint i, 1, _
            1.25 * S - 1, Y, 0
        TheSurface.SetControlPoint i, 2, _
            1.25 * S - 1, Y, S
        TheSurface.SetControlPoint i, 3, _
            -S - 1, Y, S
        TheSurface.SetControlPoint i, 4, _
            -S - 1, Y, -S
        TheSurface.SetControlPoint i, 5, _
            1.25 * S - 1, Y, -S
        TheSurface.SetControlPoint i, 6, _
            1.25 * S - 1, Y, 0
    Next i
End Sub



' Set the control points for a hill.
Private Sub MakeHill()
Dim i As Integer
Dim j As Integer
            
    TheSurface.SetBounds 4, 4
    
    ' Start flat and modify from there.
    For i = 1 To 4
        For j = 1 To 4
            TheSurface.SetControlPoint i, j, 2 * i - 5, 0, 2 * j - 5
        Next j
    Next i
    
    ' Make the modifications.
    TheSurface.SetControlPoint 2, 2, -1, 7, -1
    TheSurface.SetControlPoint 2, 3, -1, 7, 1
    TheSurface.SetControlPoint 3, 2, 1, 7, -1
    TheSurface.SetControlPoint 3, 3, 1, 7, 1
End Sub


Private Sub Form_Resize()
Dim wid As Single

    wid = ScaleWidth - picCanvas.Left
    If wid < 120 Then wid = 120
    picCanvas.Move picCanvas.Left, 0, wid, ScaleHeight
End Sub

Private Sub optSurface_Click(Index As Integer)
    SurfaceSelected = Index
    CreateData
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
End Sub


' Create the surface.
Private Sub CreateData()
Const GapU = 0.1
Const GapV = 0.1
Const Du = GapU / 5
Const Dv = GapV / 5

    MousePointer = vbHourglass
    Refresh

    Set TheSurface = New Bezier3d

    TheSurface.DrawControls = (chkShowControlPoints.value = vbChecked)
    TheSurface.DrawGrid = (chkShowControlGrid.value = vbChecked)

    ' Set the control points.
    Select Case SurfaceSelected
        Case 0  ' Hill.
            MakeHill
    
        Case 1  ' Wave.
            MakeWave
    
        Case 2  ' Tent.
            MakeTent
            
        Case 3  ' Curl.
            MakeCurl
            
        Case 4  ' Pipe.
            MakePipe
            
        Case 5  ' Cowling.
            MakeCowl
            
        Case 6  ' Twist.
            MakeTwist
        
        Case 7  ' Spiral.
            MakeSpiral
        
        Case 8  ' Urn.
            MakeUrn
        
        Case Else  ' Something safe.
            MakeHill
    
    End Select

    ' Initialize the Bezier surface.
    TheSurface.InitializeGrid GapU, GapV, Du, Dv
End Sub

Private Sub chkShowControlPoints_Click()
    TheSurface.DrawControls = (chkShowControlPoints.value = vbChecked)
    DrawData picCanvas
End Sub
Private Sub chkshowcontrolgrid_Click()
    TheSurface.DrawGrid = (chkShowControlGrid.value = vbChecked)
    DrawData picCanvas
End Sub
