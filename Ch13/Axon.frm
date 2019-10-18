VERSION 5.00
Begin VB.Form frmAxon 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Axon"
   ClientHeight    =   5415
   ClientLeft      =   1215
   ClientTop       =   720
   ClientWidth     =   6750
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
   ScaleHeight     =   5415
   ScaleWidth      =   6750
   Begin VB.PictureBox Pict 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Index           =   5
      Left            =   4560
      ScaleHeight     =   -10
      ScaleLeft       =   -5
      ScaleMode       =   0  'User
      ScaleTop        =   5
      ScaleWidth      =   10
      TabIndex        =   10
      Top             =   2760
      Width           =   2175
   End
   Begin VB.PictureBox Pict 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Index           =   4
      Left            =   2280
      ScaleHeight     =   -10
      ScaleLeft       =   -5
      ScaleMode       =   0  'User
      ScaleTop        =   5
      ScaleWidth      =   10
      TabIndex        =   4
      Top             =   2760
      Width           =   2175
   End
   Begin VB.PictureBox Pict 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Index           =   3
      Left            =   0
      ScaleHeight     =   -10
      ScaleLeft       =   -5
      ScaleMode       =   0  'User
      ScaleTop        =   5
      ScaleWidth      =   10
      TabIndex        =   3
      Top             =   2760
      Width           =   2175
   End
   Begin VB.PictureBox Pict 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Index           =   2
      Left            =   4560
      ScaleHeight     =   -10
      ScaleLeft       =   -5
      ScaleMode       =   0  'User
      ScaleTop        =   5
      ScaleWidth      =   10
      TabIndex        =   2
      Top             =   0
      Width           =   2175
   End
   Begin VB.PictureBox Pict 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Index           =   1
      Left            =   2280
      ScaleHeight     =   -10
      ScaleLeft       =   -5
      ScaleMode       =   0  'User
      ScaleTop        =   5
      ScaleWidth      =   10
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
   Begin VB.PictureBox Pict 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Index           =   0
      Left            =   0
      ScaleHeight     =   -10
      ScaleLeft       =   -5
      ScaleMode       =   0  'User
      ScaleTop        =   5
      ScaleWidth      =   10
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "The final projection"
      Height          =   255
      Index           =   5
      Left            =   4560
      TabIndex        =   11
      Top             =   5040
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Project"
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   9
      Top             =   5040
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Rotate into Y axis"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   8
      Top             =   5040
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Rotate into Y-Z plane"
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   7
      Top             =   2280
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Translate to origin"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Top             =   2280
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Original picture"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   2280
      Width           =   2175
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAxon"
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

' The transformation matrices.
Private M(0 To 4) As Transformation

' First segment not in the axes.
Private FirstSegment As Integer
' Create the matrices used when performing an
' axonometric orthographic projection with focus
' at (f1, f2, f3) and projection direction
' <n1, n2, n3>.
Private Sub CreateMatrices(ByVal f1 As Single, ByVal f2 As Single, ByVal f3 As Single, ByVal n1 As Single, ByVal n2 As Single, ByVal n3 As Single)
Dim trans(1 To 4, 1 To 4) As Single
Dim Rot1(1 To 4, 1 To 4) As Single
Dim Rot2(1 To 4, 1 To 4) As Single
Dim Proj(1 To 4, 1 To 4) As Single
Dim D As Single
Dim L As Single

    ' Translate the focus point to the origin.
    m3Translate trans, -f1, -f2, -f3

    ' Rotate around Z-axis until the projection
    ' direction is in the Y-Z plane.
    m3Identity Rot1
    D = Sqr(n1 * n1 + n2 * n2)
    Rot1(1, 1) = n2 / D
    Rot1(1, 2) = n1 / D
    Rot1(2, 1) = -Rot1(1, 2)
    Rot1(2, 2) = Rot1(1, 1)

    ' Rotate around the X-axis until the normal
    ' lies along the Y axis.
    m3Identity Rot2
    L = Sqr(n1 * n1 + n2 * n2 + n3 * n3)
    Rot2(2, 2) = D / L
    Rot2(2, 3) = -n3 / L
    Rot2(3, 2) = -Rot2(2, 3)
    Rot2(3, 3) = Rot2(2, 2)

    ' Project into the X-Z plane.
    m3Identity Proj
    Proj(2, 2) = 0

    ' Put the matrices in the M array.
    m3Identity M(0).M
    m3MatCopy M(1).M, trans
    m3MatCopy M(2).M, Rot1
    m3MatCopy M(3).M, Rot2
    m3MatCopy M(4).M, Proj
End Sub
' Let the user change the location of the eye.
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

    ' Make a new projection matrix.
    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0
    
    ' Redraw the pictures.
    DrawAllData
End Sub


' Rotate the points in the cube and draw the cube.
Private Sub DrawTheData(ByVal pic As Object, ByVal project As Boolean)
Dim i As Integer
Dim x1 As Single
Dim y1 As Single
Dim x2 As Single
Dim y2 As Single
Dim oldwidth As Integer

    ' If we should project, do so.
    If project Then TransformData Projector, 1, NumSegments
    
    ' Draw the points.
    pic.Cls
    
    oldwidth = pic.DrawWidth
    For i = 1 To NumSegments
        x1 = Segments(i).fr_tr(1)
        y1 = Segments(i).fr_tr(2)
        x2 = Segments(i).to_tr(1)
        y2 = Segments(i).to_tr(2)
        
        ' Draw the plane's normal in bold.
        If i = 4 Then pic.DrawWidth = 3
        pic.Line (x1, y1)-(x2, y2)
        If i = 4 Then pic.DrawWidth = oldwidth
    Next i
    
    pic.Refresh
End Sub


Private Sub Form_Load()
    ' Initialize the eye position.
    EyeR = 3
    EyeTheta = PI * 0.4
    EyePhi = PI * 0.2
    
    ' Create the initial viewing transformation.
    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0

    ' Create the projection matrices.
    CreateMatrices 2, 2, 2, 2, 1, 1

    ' Create, project, and draw the data.
    DrawAllData
End Sub
' Draw all the pictures.
Private Sub DrawAllData()
Dim i As Integer
Dim P(1 To 4, 1 To 4) As Single

    ' Start with fresh data.
    CreateData
    
    For i = 0 To 4
        ' Apply the next transformation.
        TransformData M(i).M, FirstSegment, NumSegments
        SetPoints FirstSegment, NumSegments
        
        ' Display the data.
        DrawTheData Pict(i), True
    Next i

    ' Create the final, transformed picture.
    m3OrthoTop P
    TransformData P, 1, NumSegments
    DrawTheData Pict(5), False
End Sub

' Create the cube data.
Private Sub CreateData()
Dim L As Single
Dim v1x As Single
Dim v1y As Single
Dim v1z As Single
Dim v2x As Single
Dim v2y As Single
Dim v2z As Single
Dim p1x As Single
Dim p1y As Single
Dim p1z As Single
Dim p2x As Single
Dim p2y As Single
Dim p2z As Single
Dim p3x As Single
Dim p3y As Single
Dim p3z As Single
Dim p4x As Single
Dim p4y As Single
Dim p4z As Single

    ' Start with no data.
    NumSegments = 0
    
    ' Create the axes.
    MakeSegment 0, 0, 0, 5, 0, 0    ' X axis.
    MakeSegment 0, 0, 0, 0, 5, 0    ' Y axis.
    MakeSegment 0, 0, 0, 0, 0, 5    ' Z axis.
    
    FirstSegment = NumSegments + 1
    
    ' Make a projection direction vector.
    MakeSegment 2, 2, 2, 4, 3, 3

    ' Create the edges of the projection plane.
    m3Cross v1x, v1y, v1z, 2, 1, 1, 0, 1, 0
    L = Sqr(v1x * v1x + v1y * v1y + v1z * v1z)
    v1x = 3 * v1x / L
    v1y = 3 * v1y / L
    v1z = 3 * v1z / L
    
    m3Cross v2x, v2y, v2z, 2, 1, 1, v1x, v1y, v1z
    L = Sqr(v2x * v2x + v2y * v2y + v2z * v2z)
    v2x = 3 * v2x / L
    v2y = 3 * v2y / L
    v2z = 3 * v2z / L
    
    p1x = 2 + v1x + v2x
    p1y = 2 + v1y + v2y
    p1z = 2 + v1z + v2z
    p2x = 2 - v1x + v2x
    p2y = 2 - v1y + v2y
    p2z = 2 - v1z + v2z
    p3x = 2 - v1x - v2x
    p3y = 2 - v1y - v2y
    p3z = 2 - v1z - v2z
    p4x = 2 + v1x - v2x
    p4y = 2 + v1y - v2y
    p4z = 2 + v1z - v2z
    
    MakeSegment p1x, p1y, p1z, p2x, p2y, p2z
    MakeSegment p2x, p2y, p2z, p3x, p3y, p3z
    MakeSegment p3x, p3y, p3z, p4x, p4y, p4z
    MakeSegment p4x, p4y, p4z, p1x, p1y, p1z

    ' Create a cube to project.
    MakeSegment 1, 1, 1, 1, 3, 1
    MakeSegment 1, 3, 1, 3, 3, 1
    MakeSegment 3, 3, 1, 3, 1, 1
    MakeSegment 3, 1, 1, 1, 1, 1
    MakeSegment 1, 1, 3, 1, 3, 3
    MakeSegment 1, 3, 3, 3, 3, 3
    MakeSegment 3, 3, 3, 3, 1, 3
    MakeSegment 3, 1, 3, 1, 1, 3
    MakeSegment 1, 1, 1, 1, 1, 3
    MakeSegment 1, 3, 1, 1, 3, 3
    MakeSegment 3, 3, 1, 3, 3, 3
    MakeSegment 3, 1, 1, 3, 1, 3

    NumSegments = NumSegments
End Sub
