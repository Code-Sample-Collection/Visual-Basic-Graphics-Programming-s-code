VERSION 5.00
Begin VB.Form frmOrth 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Orth"
   ClientHeight    =   5895
   ClientLeft      =   2100
   ClientTop       =   525
   ClientWidth     =   4530
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
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5895
   ScaleWidth      =   4530
   Begin VB.PictureBox picSide 
      AutoRedraw      =   -1  'True
      Height          =   1935
      Left            =   1965
      ScaleHeight     =   -5.639
      ScaleLeft       =   -4
      ScaleMode       =   0  'User
      ScaleTop        =   4
      ScaleWidth      =   5.639
      TabIndex        =   3
      Top             =   1965
      Width           =   1935
   End
   Begin VB.PictureBox picFront 
      AutoRedraw      =   -1  'True
      Height          =   1935
      Left            =   0
      ScaleHeight     =   -5.639
      ScaleLeft       =   -2
      ScaleMode       =   0  'User
      ScaleTop        =   4
      ScaleWidth      =   5.639
      TabIndex        =   2
      Top             =   3930
      Width           =   1935
   End
   Begin VB.PictureBox picTop 
      AutoRedraw      =   -1  'True
      Height          =   1935
      Left            =   0
      ScaleHeight     =   -5.639
      ScaleLeft       =   -2
      ScaleMode       =   0  'User
      ScaleTop        =   2
      ScaleWidth      =   5.639
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.PictureBox picNormal 
      AutoRedraw      =   -1  'True
      Height          =   1935
      Left            =   0
      ScaleHeight     =   -4
      ScaleLeft       =   -1.5
      ScaleMode       =   0  'User
      ScaleTop        =   2.5
      ScaleWidth      =   4
      TabIndex        =   0
      Top             =   1965
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Side"
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   6
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Front"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   5
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Top"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "frmOrth"
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

Private Sub Form_Load()
Dim M(1 To 4, 1 To 4) As Single

    ' Initialize the eye position.
    EyeR = 3
    EyeTheta = PI * 0.37
    EyePhi = PI * 0.1

    ' Create the data.
    CreateData

    ' Create the projection matrix.
    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0

    ' Draw the data.
    TransformAllData Projector
    DrawAllData picNormal, ForeColor, True

    m3OrthoTop M
    TransformAllData M
    DrawAllData picTop, ForeColor, True

    m3OrthoFront M
    TransformAllData M
    DrawAllData picFront, ForeColor, True

    m3OrthoSide M
    TransformAllData M
    DrawAllData picSide, ForeColor, True
End Sub
' Create the object to display.
Private Sub CreateData()
'   MakeSegment 0, 0, 0, 2, 0, 0    ' Hidden.
    MakeSegment 2, 0, 0, 2, 1, 0
    MakeSegment 2, 1, 0, 1, 1, 0
    MakeSegment 1, 1, 0, 1, 2, 0
    MakeSegment 1, 2, 0, 0, 2, 0

'   MakeSegment 0, 0, 0, 0, 0, 2    ' Hidden.
    MakeSegment 0, 0, 2, 0, 1, 2
    MakeSegment 0, 1, 2, 0, 1, 1
    MakeSegment 0, 1, 1, 0, 2, 1
    MakeSegment 0, 2, 1, 0, 2, 0

'   MakeSegment 0, 0, 0, 2, 0, 0    ' Hidden.
    MakeSegment 2, 0, 0, 2, 0, 1
    MakeSegment 2, 0, 1, 1, 0, 1
    MakeSegment 1, 0, 1, 1, 0, 2
    MakeSegment 1, 0, 2, 0, 0, 2

    MakeSegment 0, 1, 1, 2, 1, 1
    MakeSegment 1, 0, 1, 1, 2, 1
    MakeSegment 1, 1, 0, 1, 1, 2
    MakeSegment 2, 1, 0, 2, 1, 1
    MakeSegment 2, 1, 1, 2, 0, 1
    MakeSegment 0, 2, 1, 1, 2, 1
    MakeSegment 1, 2, 1, 1, 2, 0
    MakeSegment 0, 1, 2, 1, 1, 2
    MakeSegment 1, 1, 2, 1, 0, 2
End Sub
