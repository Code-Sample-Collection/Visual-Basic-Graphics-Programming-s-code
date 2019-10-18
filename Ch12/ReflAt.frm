VERSION 5.00
Begin VB.Form frmReflAt 
   Caption         =   "ReflAt"
   ClientHeight    =   4425
   ClientLeft      =   2325
   ClientTop       =   495
   ClientWidth     =   4215
   DrawMode        =   14  'Copy Pen
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4425
   ScaleWidth      =   4215
   Begin VB.CheckBox chkReflect 
      Caption         =   "Reflect"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   4080
      Width           =   855
   End
   Begin VB.PictureBox Pict 
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
Attribute VB_Name = "frmReflAt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NumSegments As Integer

' Hold two points for a line segment.
Private Type Segment
    fr_pt(1 To 3) As Single
    to_pt(1 To 3) As Single
    fr_tr(1 To 3) As Single
    to_tr(1 To 3) As Single
End Type

Private Segments() As Segment

Private Const LineX1 = 0.75
Private Const LineY1 = -0.5
Private Const LineX2 = 3.75
Private Const LineY2 = 5.5
' Create the data.
Private Sub CreateData()
    ' Create the axes.
    MakeSegment 0, 0, 5, 0
    MakeSegment 0, 0, 0, 5

    ' Create the line of reflection.
    MakeSegment LineX1, LineY1, LineX2, LineY2

    ' Create an object to manipulate.
    MakeSegment 1, 3, 2, 3
    MakeSegment 2, 3, 2, 4
    MakeSegment 2, 4, 1, 4
    MakeSegment 1, 4, 1, 3

    MakeSegment 1, 3, 2, 4
    MakeSegment 2, 3, 1, 4
End Sub
' Draw the transformed data.
Private Sub DrawSegments(pic As Object)
Dim dx As Single
Dim dy As Single
Dim M(1 To 3, 1 To 3) As Single
Dim i As Integer
Dim x1 As Single
Dim y1 As Single
Dim x2 As Single
Dim y2 As Single

    pic.Cls

    ' Transform the data.
    If chkReflect.Value = vbChecked Then
        dx = LineX2 - LineX1
        dy = LineY2 - LineY1
        m2ReflectAcross M, LineX1, LineY1, dx, dy
    Else
        m2Identity M
    End If
    TransformPicture M

    For i = 1 To NumSegments
        x1 = Segments(i).fr_tr(1)
        y1 = Segments(i).fr_tr(2)
        x2 = Segments(i).to_tr(1)
        y2 = Segments(i).to_tr(2)
        pic.Line (x1, y1)-(x2, y2)
    Next i
End Sub
' Redraw the picture.
Private Sub chkReflect_Click()
    Pict.Refresh
End Sub

Private Sub Form_Load()
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


Private Sub Pict_Paint()
    DrawSegments Pict
End Sub
