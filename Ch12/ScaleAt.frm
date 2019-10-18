VERSION 5.00
Begin VB.Form frmScaleAt 
   Caption         =   "ScaleAt"
   ClientHeight    =   4425
   ClientLeft      =   2325
   ClientTop       =   495
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4425
   ScaleWidth      =   4215
   Begin VB.OptionButton optSize 
      Caption         =   "Normal"
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Top             =   4080
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton optSize 
      Caption         =   "Little"
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Top             =   4080
      Width           =   855
   End
   Begin VB.OptionButton optSize 
      Caption         =   "Big"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   4080
      Width           =   855
   End
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
Attribute VB_Name = "frmScaleAt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NumSegments As Integer
Private Segments() As Segment

' Hold two points for a line segment.
Private Type Segment
    fr_pt(1 To 3) As Single
    to_pt(1 To 3) As Single
    fr_tr(1 To 3) As Single
    to_tr(1 To 3) As Single
End Type

Private CurrentT(1 To 3, 1 To 3) As Single
' Make the data we will transform and draw.
Private Sub CreateData()
    ' Create the axes.
    MakeSegment 0, 0, 5, 0
    MakeSegment 0, 0, 0, 5

    ' Create an object to manipulate.
    MakeSegment 1, 1, 2, 1
    MakeSegment 2, 1, 2, 2
    MakeSegment 2, 2, 1, 2
    MakeSegment 1, 2, 1, 1

    MakeSegment 1, 1, 2, 2
    MakeSegment 2, 1, 1, 2
End Sub

' Draw all the transformed segments.
Private Sub DrawSegments(ByVal pic As PictureBox)
Dim i As Integer
Dim x1 As Single
Dim y1 As Single
Dim x2 As Single
Dim y2 As Single

    pic.Cls

    TransformPicture CurrentT

    For i = 1 To NumSegments
        x1 = Segments(i).fr_tr(1)
        y1 = Segments(i).fr_tr(2)
        x2 = Segments(i).to_tr(1)
        y2 = Segments(i).to_tr(2)
        pic.Line (x1, y1)-(x2, y2)
    Next i
End Sub
' Load the data.
Private Sub Form_Load()
    m2Identity CurrentT

    CreateData
End Sub
' Make a new segment.
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


' Redraw the picture.
Private Sub picCanvas_Paint()
    DrawSegments picCanvas
End Sub


' Change the scale.
Private Sub optSize_Click(Index As Integer)
    Select Case Index
        Case 0  ' Big.
            m2ScaleAt CurrentT, 2, 2, 1.5, 1.5
        Case 1  ' Normal.
            m2Identity CurrentT
        Case 2  ' Little.
            m2ScaleAt CurrentT, 0.5, 0.5, 1.5, 1.5
    End Select
    picCanvas.Refresh
End Sub
