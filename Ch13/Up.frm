VERSION 5.00
Begin VB.Form frmUp 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Up"
   ClientHeight    =   5505
   ClientLeft      =   330
   ClientTop       =   735
   ClientWidth     =   9060
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
   ScaleHeight     =   5505
   ScaleWidth      =   9060
   Begin VB.PictureBox PPict 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Left            =   6840
      ScaleHeight     =   -10
      ScaleLeft       =   -5
      ScaleMode       =   0  'User
      ScaleTop        =   5
      ScaleWidth      =   10
      TabIndex        =   12
      Top             =   2760
      Width           =   2175
   End
   Begin VB.PictureBox Pict 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Index           =   0
      Left            =   1200
      ScaleHeight     =   -10
      ScaleLeft       =   -5
      ScaleMode       =   0  'User
      ScaleTop        =   5
      ScaleWidth      =   10
      TabIndex        =   5
      Top             =   0
      Width           =   2175
   End
   Begin VB.PictureBox Pict 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Index           =   1
      Left            =   3480
      ScaleHeight     =   -10
      ScaleLeft       =   -5
      ScaleMode       =   0  'User
      ScaleTop        =   5
      ScaleWidth      =   10
      TabIndex        =   4
      Top             =   0
      Width           =   2175
   End
   Begin VB.PictureBox Pict 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Index           =   2
      Left            =   5760
      ScaleHeight     =   -10
      ScaleLeft       =   -5
      ScaleMode       =   0  'User
      ScaleTop        =   5
      ScaleWidth      =   10
      TabIndex        =   3
      Top             =   0
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   2760
      Width           =   2175
   End
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
      TabIndex        =   0
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Final projection"
      Height          =   255
      Index           =   6
      Left            =   6840
      TabIndex        =   13
      Top             =   5040
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Original picture"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   11
      Top             =   2280
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Translate focus to origin"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   10
      Top             =   2280
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Rotate center of projection into Y-Z plane"
      Height          =   495
      Index           =   2
      Left            =   5760
      TabIndex        =   9
      Top             =   2280
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Rotate center of projection into Z axis"
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   8
      Top             =   5040
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Rotate UP into Y-Z plane"
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   7
      Top             =   5040
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Project onto X-Y plane"
      Height          =   255
      Index           =   5
      Left            =   4560
      TabIndex        =   6
      Top             =   5040
      Width           =   2175
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FirstCube As Integer

' Viewing parameters.
Private EyeR As Single      ' Center of projection.
Private EyeTheta As Single
Private EyePhi As Single
Private Const FocusX = 0#       ' Focus point.
Private Const FocusY = 0#
Private Const FocusZ = 0#

' Projection parameters.
Private UpX As Single       ' Up vector.
Private UpY As Single
Private UpZ As Single
Private Cx As Single        ' Center of projection.
Private Cy As Single
Private Cz As Single
Private Fx As Single        ' Focus point.
Private Fy As Single
Private Fz As Single

' Matrices used for the projection.
Private M(0 To 5) As Transformation
Private Projector(1 To 4, 1 To 4) As Single

Private P(1 To 4, 1 To 4) As Single

' Create transformation matrices for perspective
' projection with:
'       focus point             (focx, focy, focz)
'       center of projection    (ex, ey, ez)
'       up vector               <ux, uy, uz>
Private Sub CreateMatrices(ByVal focx As Single, ByVal focy As Single, ByVal focz As Single, ByVal ex As Single, ByVal ey As Single, ByVal ez As Single, ByVal ux As Single, ByVal uy As Single, ByVal uz As Single)
Dim sin1 As Single
Dim cos1 As Single
Dim sin2 As Single
Dim cos2 As Single
Dim sin3 As Single
Dim cos3 As Single
Dim A As Single
Dim B As Single
Dim C As Single
Dim d1 As Single
Dim d2 As Single
Dim d3 As Single
Dim up1(1 To 4) As Single
Dim up2(1 To 4) As Single

    ' Identity transformation.
    m3Identity M(0).M
    
    ' Translate the focus to the origin.
    m3Translate M(1).M, -focx, -focy, -focz

    A = ex - focx
    B = ey - focy
    C = ez - focz
    d1 = Sqr(A * A + C * C)
    sin1 = -A / d1
    cos1 = C / d1
    d2 = Sqr(A * A + B * B + C * C)
    sin2 = B / d2
    cos2 = d1 / d2
    
    ' Rotate around the Y axis to place the
    ' center of projection in the Y-Z plane.
    m3Identity M(2).M
    M(2).M(1, 1) = cos1
    M(2).M(1, 3) = -sin1
    M(2).M(3, 1) = sin1
    M(2).M(3, 3) = cos1

    ' Rotate around the X axis to place the
    ' center of projection in the Z axis.
    m3Identity M(3).M
    M(3).M(2, 2) = cos2
    M(3).M(2, 3) = sin2
    M(3).M(3, 2) = -sin2
    M(3).M(3, 3) = cos2

    ' Apply the rotations to the UP vector.
    up1(1) = ux
    up1(2) = uy
    up1(3) = uz
    up1(4) = 1
    m3Apply up1, M(2).M, up2
    m3Apply up2, M(3).M, up1

    ' Rotate around the Z axis to put the UP
    ' vector in the Y-Z plane.
    d3 = Sqr(up1(1) * up1(1) + up1(2) * up1(2))
    sin3 = up1(1) / d3
    cos3 = up1(2) / d3
    m3Identity M(4).M
    M(4).M(1, 1) = cos3
    M(4).M(1, 2) = sin3
    M(4).M(2, 1) = -sin3
    M(4).M(2, 2) = cos3

    ' Project.
    m3PerspectiveXZ M(5).M, d2

    ' Compute the projection all in one shot.
    m3Project P, m3Perspective, ex, ey, ez, focx, focy, focz, ux, uy, uz
End Sub
' Let the user change the location of the eye.
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Const Dtheta = PI / 20
Const Dx = 0.25

Dim inc As Single

    If Shift And 1 Then
        inc = Dx
    Else
        inc = -Dx
    End If
    
    Select Case KeyCode
        Case vbKeyLeft
            EyeTheta = EyeTheta - Dtheta
            
        Case vbKeyRight
            EyeTheta = EyeTheta + Dtheta
        
        Case vbKeyUp
            EyePhi = EyePhi - Dtheta
        
        Case vbKeyDown
            EyePhi = EyePhi + Dtheta
        
        Case Asc("X")
            UpX = UpX + inc
        Case Asc("Y")
            UpY = UpY + inc
        Case Asc("Z")
            UpZ = UpZ + inc
        
        Case Else
            Exit Sub
    End Select

    ' Redraw the pictures.
    DrawTheData
End Sub

Private Sub Form_Load()
    ' Initialize the viewing parameters.
    EyeR = 3
    EyeTheta = PI * 0.35
    EyePhi = PI * 0.1
    
    ' Initialize projection parameters.
    UpX = -1
    UpY = 1.5
    UpZ = 0
    Cx = 2
    Cy = 2.5
    Cz = 3
    Fx = 1
    Fy = 1
    Fz = 1
    
    ' Create, project, and draw the data.
    DrawTheData
End Sub
' Draw all the pictures.
Private Sub DrawTheData()
Dim i As Integer

    CreateData
    CreateMatrices Fx, Fy, Fz, Cx, Cy, Cz, UpX, UpY, UpZ
    
    ' Compute the projection matrix.
    m3PProject Projector, m3Parallel, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0

    For i = 0 To 5
        TransformData M(i).M, FirstCube, NumSegments
        SetPoints FirstCube, NumSegments
        TransformData Projector, 1, NumSegments
        DrawSomeData Pict(i), 1, NumSegments - 2, ForeColor, True
        Pict(i).DrawWidth = 3
        DrawSomeData Pict(i), NumSegments - 1, NumSegments - 1, vbRed, False
        DrawSomeData Pict(i), NumSegments, NumSegments, vbGreen, False
        Pict(i).DrawWidth = DrawWidth
        Pict(i).Refresh
    Next i

    ' For the final view use the transformation
    ' given by m3PerspectiveProjectionUp
    CreateData
    TransformData P, FirstCube, NumSegments
    DrawSomeData PPict, FirstCube, NumSegments - 2, ForeColor, True
    PPict.DrawWidth = 3
    DrawSomeData PPict, NumSegments - 1, NumSegments - 1, vbRed, False
    DrawSomeData PPict, NumSegments, NumSegments, vbGreen, False
    PPict.DrawWidth = DrawWidth
    PPict.Refresh
End Sub


' Create the data.
Private Sub CreateData()
    ' Start with no data.
    NumSegments = 0
    
    ' Create the axes.
    MakeSegment 0, 0, 0, 4, 0, 0    ' X axis.
    MakeSegment 0, 0, 0, 0, 4, 0    ' Y axis.
    MakeSegment 0, 0, 0, 0, 0, 4    ' Z axis.
        
    FirstCube = NumSegments + 1
    
    ' Create the object to reflect.
    MakeSegment -1, -1, -1, -1, -1, 3
    MakeSegment -1, -1, 3, -1, 3, 3
    MakeSegment -1, 3, 3, -1, 3, -1
    MakeSegment -1, 3, -1, -1, -1, -1
    MakeSegment 3, -1, -1, 3, -1, 3
    MakeSegment 3, -1, 3, 3, 3, 3
    MakeSegment 3, 3, 3, 3, 3, -1
    MakeSegment 3, 3, -1, 3, -1, -1
    MakeSegment -1, -1, -1, 3, -1, -1
    MakeSegment -1, -1, 3, 3, -1, 3
    MakeSegment -1, 3, 3, 3, 3, 3
    MakeSegment -1, 3, -1, 3, 3, -1
    
    ' Up vector.
    MakeSegment Fx, Fy, Fz, Fx + UpX, Fy + UpY, Fz + UpZ

    ' Center of projection.
    MakeSegment Fx, Fy, Fz, Cx, Cy, Cz
End Sub
