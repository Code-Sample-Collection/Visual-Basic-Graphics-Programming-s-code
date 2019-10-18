VERSION 5.00
Begin VB.Form frmRotOrder 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "RotOrder"
   ClientHeight    =   5670
   ClientLeft      =   825
   ClientTop       =   630
   ClientWidth     =   8190
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
   ScaleHeight     =   5670
   ScaleWidth      =   8190
   Begin VB.PictureBox Proj 
      AutoRedraw      =   -1  'True
      Height          =   2655
      Index           =   0
      Left            =   0
      ScaleHeight     =   -5
      ScaleLeft       =   -2
      ScaleMode       =   0  'User
      ScaleTop        =   2
      ScaleWidth      =   5
      TabIndex        =   5
      Top             =   3000
      Width           =   2655
   End
   Begin VB.PictureBox Proj 
      AutoRedraw      =   -1  'True
      Height          =   2655
      Index           =   1
      Left            =   2760
      ScaleHeight     =   -5
      ScaleLeft       =   -2
      ScaleMode       =   0  'User
      ScaleTop        =   2
      ScaleWidth      =   5
      TabIndex        =   4
      Top             =   3000
      Width           =   2655
   End
   Begin VB.PictureBox Proj 
      AutoRedraw      =   -1  'True
      Height          =   2655
      Index           =   2
      Left            =   5520
      ScaleHeight     =   -5
      ScaleLeft       =   -2
      ScaleMode       =   0  'User
      ScaleTop        =   2
      ScaleWidth      =   5
      TabIndex        =   3
      Top             =   3000
      Width           =   2655
   End
   Begin VB.PictureBox Pict 
      AutoRedraw      =   -1  'True
      Height          =   2655
      Index           =   2
      Left            =   5520
      ScaleHeight     =   -5
      ScaleLeft       =   -2
      ScaleMode       =   0  'User
      ScaleTop        =   2
      ScaleWidth      =   5
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.PictureBox Pict 
      AutoRedraw      =   -1  'True
      Height          =   2655
      Index           =   1
      Left            =   2760
      ScaleHeight     =   -5
      ScaleLeft       =   -2
      ScaleMode       =   0  'User
      ScaleTop        =   2
      ScaleWidth      =   5
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.PictureBox Pict 
      AutoRedraw      =   -1  'True
      Height          =   2655
      Index           =   0
      Left            =   0
      ScaleHeight     =   -5
      ScaleLeft       =   -2
      ScaleMode       =   0  'User
      ScaleTop        =   2
      ScaleWidth      =   5
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Directly around a line"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   8
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Into X-Z plane first"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   7
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Into Y-Z plane first"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmRotOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Point being rotated into the Z axis.
Private Const Px = 2
Private Const Py = 2
Private Const Pz = 1

' Line for direct rotation.
Private Const Vx = 1
Private Const Vy = 1
Private Const Vz = 2

' Location of viewing eye.
Private EyeR As Single
Private EyeTheta As Single
Private EyePhi As Single

' Location of focus point.
Private Const FocusX = 0#
Private Const FocusY = 0#
Private Const FocusZ = 0#

Private Projector(1 To 4, 1 To 4) As Single

' Matrices used for the reflection.
Private M1(1 To 4, 1 To 4) As Single
Private M2(1 To 4, 1 To 4) As Single
Private M3(1 To 4, 1 To 4) As Single
Private M4(1 To 4, 1 To 4) As Single
Private M5(1 To 4, 1 To 4) As Single
Private Identity(1 To 4, 1 To 4) As Single
Private Sub CreateMatrices()
Dim sin1 As Single
Dim cos1 As Single
Dim sin2 As Single
Dim cos2 As Single
Dim d1 As Single
Dim d2 As Single
    
    m3Identity Identity
    
    ' *************
    ' * Y-Z first *
    ' *************
    d1 = Sqr(Px * Px + Pz * Pz)
    sin1 = -Px / d1
    cos1 = Pz / d1
    d2 = Sqr(Px * Px + Py * Py + Pz * Pz)
    sin2 = Py / d2
    cos2 = d1 / d2

    m3Identity M1       ' Around Y into Y-Z plane.
    M1(1, 1) = cos1
    M1(1, 3) = -sin1
    M1(3, 1) = sin1
    M1(3, 3) = cos1
    
    m3Identity M2       ' Around X into Z axis.
    M2(2, 2) = cos2
    M2(2, 3) = sin2
    M2(3, 2) = -sin2
    M2(3, 3) = cos2
        
    ' *************
    ' * X-Z first *
    ' *************
    d1 = Sqr(Py * Py + Pz * Pz)
    sin1 = Py / d1
    cos1 = Pz / d1
    d2 = Sqr(Px * Px + Py * Py + Pz * Pz)
    sin2 = -Px / d2
    cos2 = d1 / d2

    m3Identity M3       ' Around X into X-Z plane.
    M3(2, 2) = cos1
    M3(2, 3) = sin1
    M3(3, 2) = -sin1
    M3(3, 3) = cos1
    
    m3Identity M4       ' Around Y into Z axis.
    M4(1, 1) = cos2
    M4(1, 3) = -sin2
    M4(3, 1) = sin2
    M4(3, 3) = cos2

    ' ***************
    ' * Around line *
    ' ***************
    m3LineRotate M5, 0, 0, 0, Vx, Vy, Vz, PI
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

    ' Redraw the pictures.
    DrawTheData
End Sub



Private Sub Form_Load()
    ' Initialize the eye position.
    EyeR = 3
    EyeTheta = PI * 0.4
    EyePhi = PI * 0.1
    
    ' Create the rotation matrices.
    CreateMatrices

    ' Create, project, and draw the data.
    DrawTheData
End Sub
' Draw all the pictures.
Private Sub DrawTheData()
Dim i As Integer

    ' Compute the projection matrix.
    m3PProject Projector, m3Parallel, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0
    
    ' ***********************
    ' * Around Y axis first *
    ' ***********************
    CreateData
    TransformAllData Projector
    DrawSomeData Pict(0), 1, 3, vbRed, True
    DrawSomeData Pict(0), 5, NumSegments, ForeColor, False
    
    TransformData M1, 5, NumSegments
    SetPoints 5, NumSegments
    TransformData Projector, 5, NumSegments
    DrawSomeData Pict(0), 5, NumSegments, ForeColor, False
    
    TransformData M2, 5, NumSegments
    SetPoints 5, NumSegments
    TransformData Projector, 5, NumSegments
    DrawSomeData Pict(0), 5, NumSegments, ForeColor, False

    TransformAllData Identity
    DrawSomeData Proj(0), 1, 3, vbRed, True
    DrawSomeData Proj(0), 5, NumSegments, ForeColor, False
    
    ' ***********************
    ' * Around X axis first *
    ' ***********************
    CreateData
    TransformAllData Projector
    DrawSomeData Pict(1), 1, 3, vbRed, True
    DrawSomeData Pict(1), 5, NumSegments, ForeColor, False
        
    TransformData M3, 5, NumSegments
    SetPoints 5, NumSegments
    TransformData Projector, 5, NumSegments
    DrawSomeData Pict(1), 5, NumSegments, ForeColor, False
    
    TransformData M4, 5, NumSegments
    SetPoints 5, NumSegments
    TransformData Projector, 5, NumSegments
    DrawSomeData Pict(1), 5, NumSegments, ForeColor, False

    TransformAllData Identity
    DrawSomeData Proj(1), 1, 3, vbRed, True
    DrawSomeData Proj(1), 5, NumSegments, ForeColor, False

    ' ***************
    ' * Around line *
    ' ***************
    CreateData
    TransformAllData Projector
    DrawSomeData Pict(2), 1, 3, vbRed, True
    DrawSomeData Pict(2), 4, 4, vbBlue, False
    DrawSomeData Pict(2), 5, NumSegments, ForeColor, False
        
    TransformData M5, 5, NumSegments
    SetPoints 5, NumSegments
    TransformData Projector, 5, NumSegments
    DrawSomeData Pict(2), 5, NumSegments, ForeColor, False
    
    TransformAllData Identity
    DrawSomeData Proj(2), 1, 3, vbRed, True
    DrawSomeData Proj(2), 5, NumSegments, ForeColor, False

    For i = 0 To 2
        Pict(i).Refresh
        Proj(i).Refresh
    Next i
End Sub

Private Sub CreateData()
    ' Start with no data.
    NumSegments = 0

    ' Create the axes.
    MakeSegment 0, 0, 0, 5, 0, 0    ' X axis.
    MakeSegment 0, 0, 0, 0, 5, 0    ' Y axis.
    MakeSegment 0, 0, 0, 0, 0, 5    ' Z axis.

    ' Create the line.
    MakeSegment -2 * Vx, -2 * Vy, -2 * Vz, 2 * Vx, 2 * Vy, 2 * Vz

    ' Create the object to reflect.
    MakeSegment Px, Py, Pz, Px, Py - 1, Pz - 1
    MakeSegment Px, Py - 1, Pz - 1, Px, Py - 1, Pz + 1
    MakeSegment Px, Py - 1, Pz + 1, Px, Py, Pz
End Sub

