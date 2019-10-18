VERSION 5.00
Begin VB.Form frmOblique 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Oblique"
   ClientHeight    =   3180
   ClientLeft      =   975
   ClientTop       =   1815
   ClientWidth     =   7470
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
   ScaleHeight     =   3180
   ScaleWidth      =   7470
   Begin VB.TextBox txtAngle 
      Height          =   285
      Left            =   480
      TabIndex        =   7
      Text            =   "30"
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox picCabinet 
      AutoRedraw      =   -1  'True
      Height          =   2415
      Left            =   5040
      ScaleHeight     =   -6
      ScaleLeft       =   -3
      ScaleMode       =   0  'User
      ScaleTop        =   3
      ScaleWidth      =   6
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.PictureBox picCavalier 
      AutoRedraw      =   -1  'True
      Height          =   2415
      Left            =   2520
      ScaleHeight     =   -6
      ScaleLeft       =   -3
      ScaleMode       =   0  'User
      ScaleTop        =   3
      ScaleWidth      =   6
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.PictureBox picOrthographic 
      AutoRedraw      =   -1  'True
      Height          =   2415
      Left            =   0
      ScaleHeight     =   -6
      ScaleLeft       =   -3
      ScaleMode       =   0  'User
      ScaleTop        =   3
      ScaleWidth      =   6
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Angle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Cabinet"
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   5
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Cavalier"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   4
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Orthographic"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   2880
      Width           =   2415
   End
End
Attribute VB_Name = "frmOblique"
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
Private CavProj(1 To 4, 1 To 4) As Single
Private CabProj(1 To 4, 1 To 4) As Single
' Transform and draw the data.
Private Sub DrawData()
Dim angle As Single

    ' Get the angle.
    On Error Resume Next
    angle = CSng(txtAngle.Text)
    If Err.Number <> 0 Then
        angle = 30
        txtAngle.Text = "30"
    End If
    angle = angle * PI / 180

    ' Create the projection matrices.
    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0
    m3ObliqueXY CavProj, 1#, angle
    m3ObliqueXY CabProj, 0.5, angle

    ' Project and draw the data.
    TransformAllData Projector
    DrawAllData picOrthographic, ForeColor, True

    TransformAllData CavProj
    DrawAllData picCavalier, ForeColor, True

    TransformAllData CabProj
    DrawAllData picCabinet, ForeColor, True
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
    TransformAllData Projector
    DrawAllData picOrthographic, ForeColor, True
End Sub

Private Sub Form_Load()
    ' Initialize the eye position.
    EyeR = 3
    EyeTheta = PI * 0.4
    EyePhi = PI * 0.1
    
    ' Create the data.
    CreateData

    ' Draw the data.
    DrawData
End Sub

Private Sub CreateData()
    ' Create the axes.
    MakeSegment 0, 0, 0, 3, 0, 0    ' X axis.
    MakeSegment 0, 0, 0, 0, 3, 0    ' Y axis.
    MakeSegment 0, 0, 0, 0, 0, 3    ' Z axis.

    ' Create the object to display.
    MakeSegment 0, 0, 0, 2, 0, 0
    MakeSegment 2, 0, 0, 2, 2, 0
    MakeSegment 2, 2, 0, 0, 2, 0
    MakeSegment 0, 2, 0, 0, 0, 0

    MakeSegment 0, 0, 2, 2, 0, 2
    MakeSegment 2, 0, 2, 2, 2, 2
    MakeSegment 2, 2, 2, 0, 2, 2
    MakeSegment 0, 2, 2, 0, 0, 2

    MakeSegment 0, 0, 0, 0, 0, 2
    MakeSegment 2, 0, 0, 2, 0, 2
    MakeSegment 2, 2, 0, 2, 2, 2
    MakeSegment 0, 2, 0, 0, 2, 2
End Sub

' Redraw the data.
Private Sub txtAngle_Change()
    DrawData
End Sub
