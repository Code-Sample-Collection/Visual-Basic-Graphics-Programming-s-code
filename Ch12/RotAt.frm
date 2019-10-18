VERSION 5.00
Begin VB.Form frmRotAt 
   Caption         =   "RotAt"
   ClientHeight    =   4005
   ClientLeft      =   2325
   ClientTop       =   495
   ClientWidth     =   4215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4005
   ScaleWidth      =   4215
   Begin VB.PictureBox picCanvas 
      Height          =   3975
      Left            =   0
      ScaleHeight     =   -7
      ScaleLeft       =   -1
      ScaleMode       =   0  'User
      ScaleTop        =   6
      ScaleWidth      =   7
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmRotAt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PI = 3.14159265

Private NumSegments As Integer

' Hold two points for a line segment.
Private Type Segment
    fr_pt(1 To 3) As Single
    to_pt(1 To 3) As Single
    fr_tr(1 To 3) As Single
    to_tr(1 To 3) As Single
End Type

Private Segments() As Segment

Private theta As Single
' Create the data.
Private Sub CreateData()
    ' Create the axes.
    MakeSegment 0, 0, 5, 0
    MakeSegment 0, 0, 0, 5

    ' Create an object to manipulate.
    MakeSegment 1, 1, 3, 1
    MakeSegment 3, 1, 3, 3
    MakeSegment 3, 3, 1, 3
    MakeSegment 1, 3, 1, 1

    MakeSegment 1, 1, 3, 3
    MakeSegment 3, 1, 1, 3
End Sub

' Draw the transformed data.
Private Sub DrawSegments(pic As Object)
Dim T(1 To 3, 1 To 3) As Single
Dim i As Integer
Dim x1 As Single
Dim y1 As Single
Dim x2 As Single
Dim y2 As Single

    pic.Cls

    ' Transform the picture.
    m2RotateAround T, theta, 2, 2
    TransformPicture T

    For i = 1 To NumSegments
        x1 = Segments(i).fr_tr(1)
        y1 = Segments(i).fr_tr(2)
        x2 = Segments(i).to_tr(1)
        y2 = Segments(i).to_tr(2)
        pic.Line (x1, y1)-(x2, y2)
    Next i
End Sub
' Change the angle of rotation and redraw.
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Const Dtheta = PI / 16

    Select Case KeyCode
        Case vbKeyLeft   ' Rotate left.
            theta = theta + Dtheta
        Case vbKeyRight  ' Rotate right.
            theta = theta - Dtheta
    End Select
    picCanvas.Refresh
End Sub

' Load the data.
Private Sub Form_Load()
    theta = 0

    CreateData
End Sub
' Make a new line segment.
Private Sub MakeSegment(x1 As Single, y1 As Single, x2 As Single, y2 As Single)
    NumSegments = NumSegments + 1
    ReDim Preserve Segments(1 To NumSegments)
    Segments(NumSegments).fr_pt(1) = x1
    Segments(NumSegments).fr_pt(2) = y1
    Segments(NumSegments).fr_pt(3) = 1
    Segments(NumSegments).to_pt(1) = x2
    Segments(NumSegments).to_pt(2) = y2
    Segments(NumSegments).to_pt(3) = 1
End Sub

' Transform all segments except the axes.
Private Sub TransformPicture(M() As Single)
Dim i As Integer

    For i = 1 To 2
        m2PointCopy Segments(i).fr_tr, Segments(i).fr_pt
        m2PointCopy Segments(i).to_tr, Segments(i).to_pt
    Next i
    For i = 3 To NumSegments
        m2Apply Segments(i).fr_pt, M, Segments(i).fr_tr
        m2Apply Segments(i).to_pt, M, Segments(i).to_tr
    Next i
End Sub


Private Sub picCanvas_Paint()
    DrawSegments picCanvas
End Sub
