VERSION 5.00
Begin VB.Form frmSolids 
   Caption         =   "Solids"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Index           =   3
      Left            =   5160
      ScaleHeight     =   -4
      ScaleLeft       =   -2
      ScaleMode       =   0  'User
      ScaleTop        =   2
      ScaleWidth      =   4
      TabIndex        =   16
      Top             =   1080
      Width           =   1575
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Index           =   4
      Left            =   6840
      ScaleHeight     =   -4
      ScaleLeft       =   -2
      ScaleMode       =   0  'User
      ScaleTop        =   2
      ScaleWidth      =   4
      TabIndex        =   15
      Top             =   1080
      Width           =   1575
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Index           =   2
      Left            =   3480
      ScaleHeight     =   -4
      ScaleLeft       =   -2
      ScaleMode       =   0  'User
      ScaleTop        =   2
      ScaleWidth      =   4
      TabIndex        =   10
      Top             =   1080
      Width           =   1575
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Index           =   1
      Left            =   1800
      ScaleHeight     =   -4
      ScaleLeft       =   -2
      ScaleMode       =   0  'User
      ScaleTop        =   2
      ScaleWidth      =   4
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Index           =   0
      Left            =   120
      ScaleHeight     =   -4
      ScaleLeft       =   -2
      ScaleMode       =   0  'User
      ScaleTop        =   2
      ScaleWidth      =   4
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblSegments 
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   24
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblPoints 
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   23
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblVerify 
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   22
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Dodecahedron"
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   21
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Icosahedron"
      Height          =   255
      Index           =   4
      Left            =   6840
      TabIndex        =   20
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblVerify 
      Height          =   255
      Index           =   4
      Left            =   6840
      TabIndex        =   19
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblPoints 
      Height          =   255
      Index           =   4
      Left            =   6840
      TabIndex        =   18
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblSegments 
      Height          =   255
      Index           =   4
      Left            =   6840
      TabIndex        =   17
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblSegments 
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   14
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblPoints 
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   13
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblVerify 
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   12
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Octahedron"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   11
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Cube"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblVerify 
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   8
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblPoints 
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   7
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblSegments 
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblSegments 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblPoints 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblVerify 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Tetrahedron"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmSolids"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' One polyline per solid.
Private Const MAX_SOLID = 4
Private Polylines(0 To MAX_SOLID) As Polyline3d

' Location of viewing eye.
Private EyeR As Single
Private EyeTheta As Single
Private EyePhi As Single

' Location of focus point.
Private Const FocusX = 0#
Private Const FocusY = 0#
Private Const FocusZ = 0#

Private Projector(1 To 4, 1 To 4) As Single
' Project and draw the data.
Private Sub DrawData()
Dim i As Integer

    ' Display the solids.
    For i = 0 To MAX_SOLID
        ' Transform the data for this solid.
        Polylines(i).ApplyFull Projector

        ' Draw the solid.
        picCanvas(i).Cls
        Polylines(i).Draw picCanvas(i)
        picCanvas(i).Refresh
    Next i
End Sub
' Calculate and verify the points.
Private Sub Form_Load()
Dim i As Integer

    ' Make the data.
    GetTetrahedron Polylines(0)
    GetCube Polylines(1)
    GetOctahedron Polylines(2)
    GetDodecahedron Polylines(3)
    GetIcosahedron Polylines(4)

    ' Verify and display the data.
    For i = 0 To MAX_SOLID
        If Polylines(i).SolidOk() Then
            lblVerify(i).Caption = "Verification: Ok"
        Else
            lblVerify(i).Caption = "Verification: Error"
        End If
        lblPoints(i).Caption = Format$(Polylines(i).NumPoints) & " points"
        lblSegments(i).Caption = Format$(Polylines(i).NumSegs) & " segments"
    Next i

    ' Initialize the eye position.
    EyeR = 5
    EyeTheta = PI * 0.4
    EyePhi = PI * 0.1

    ' Initialize the projection transformation.
    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0

    ' Draw the data.
    DrawData
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
    DrawData
End Sub
