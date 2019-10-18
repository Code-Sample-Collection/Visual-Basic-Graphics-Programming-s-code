VERSION 5.00
Begin VB.Form frmPersp 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Persp"
   ClientHeight    =   5310
   ClientLeft      =   1410
   ClientTop       =   570
   ClientWidth     =   6870
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
   ScaleWidth      =   6870
   Begin VB.Frame Frame1 
      Caption         =   "Pre-Rotations"
      Height          =   2415
      Index           =   2
      Left            =   5400
      TabIndex        =   9
      Top             =   0
      Width           =   1455
      Begin VB.TextBox txtZW 
         Height          =   285
         Left            =   600
         MaxLength       =   6
         TabIndex        =   15
         Text            =   "0.0"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtYW 
         Height          =   285
         Left            =   600
         MaxLength       =   6
         TabIndex        =   14
         Text            =   "0.0"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtXW 
         Height          =   285
         Left            =   600
         MaxLength       =   6
         TabIndex        =   13
         Text            =   "0.0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtYZ 
         Height          =   285
         Left            =   600
         MaxLength       =   6
         TabIndex        =   12
         Text            =   "0.0"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtXZ 
         Height          =   285
         Left            =   600
         MaxLength       =   6
         TabIndex        =   11
         Text            =   "0.0"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtXY 
         Height          =   285
         Left            =   600
         MaxLength       =   6
         TabIndex        =   10
         Text            =   "0.0"
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "ZW"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "YW"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "XW"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "YZ"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   18
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "XZ"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "XY"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Post-Rotations"
      Height          =   1335
      Index           =   1
      Left            =   5400
      TabIndex        =   2
      Top             =   3120
      Width           =   1455
      Begin VB.TextBox txtXW2 
         Height          =   285
         Left            =   600
         MaxLength       =   6
         TabIndex        =   5
         Text            =   "0.2"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtYW2 
         Height          =   285
         Left            =   600
         MaxLength       =   6
         TabIndex        =   4
         Text            =   "0.1"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtZW2 
         Height          =   285
         Left            =   600
         MaxLength       =   6
         TabIndex        =   3
         Text            =   "0.0"
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "X"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Y"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Z"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   255
      End
   End
   Begin VB.TextBox txtD 
      Height          =   285
      Left            =   6000
      TabIndex        =   1
      Text            =   "3"
      Top             =   2640
      Width           =   735
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   5295
      Left            =   0
      ScaleHeight     =   349
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   349
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "D"
      Height          =   255
      Index           =   15
      Left            =   5640
      TabIndex        =   22
      Top             =   2640
      Width           =   255
   End
End
Attribute VB_Name = "frmPersp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Location of focus point.
Private Const FocusX = 0#
Private Const FocusY = 0#
Private Const FocusZ = 0#

' The points.
Private NumPoints As Integer
Private Points() As Point4D

' The segments.
Private NumSegments As Integer
Private Segments() As Segment4D
' Add a segment to the lists.
Private Sub AddSegment( _
    ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, ByVal w1 As Single, _
    ByVal x2 As Single, ByVal y2 As Single, ByVal z2 As Single, ByVal w2 As Single _
)
Dim pt1 As Integer
Dim pt2 As Integer

    ' Find the points.
    pt1 = PointNumber(x1, y1, z1, w1)
    pt2 = PointNumber(x2, y2, z2, w2)

    ' Create the segment entry.
    NumSegments = NumSegments + 1
    ReDim Preserve Segments(1 To NumSegments)
    With Segments(NumSegments)
        .pt1 = pt1
        .pt2 = pt2
    End With
End Sub

' Apply this matrix to the points.
Private Sub Apply(M() As Single)
Dim pt As Integer

    For pt = 1 To NumPoints
        m4Apply Points(pt).coord, M, Points(pt).trans
    Next pt
End Sub
' Apply this matrix to the points.
Private Sub ApplyFull(M() As Single)
Dim pt As Integer

    For pt = 1 To NumPoints
        m4ApplyFull Points(pt).coord, M, Points(pt).trans
    Next pt
End Sub

' Draw the segments.
Private Sub Draw(ByVal pic As PictureBox)
Dim seg As Integer

    For seg = 1 To NumSegments
        pic.Line ( _
            Points(Segments(seg).pt1).trans(1), _
            Points(Segments(seg).pt1).trans(2))-( _
            Points(Segments(seg).pt2).trans(1), _
            Points(Segments(seg).pt2).trans(2))
    Next seg
End Sub


' Find this point's index. If it is not here,
' create it.
Private Function PointNumber(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, ByVal W As Single)
Dim i As Integer

    ' Find the point.
    For i = 1 To NumPoints
        With Points(i)
            If .coord(1) = X And _
               .coord(2) = Y And _
               .coord(3) = Z And _
               .coord(4) = W _
            Then
                PointNumber = i
                Exit Function
            End If
        End With
    Next i

    ' We did not find the point. Create it.
    NumPoints = NumPoints + 1
    ReDim Preserve Points(1 To NumPoints)
    With Points(NumPoints)
        .coord(1) = X
        .coord(2) = Y
        .coord(3) = Z
        .coord(4) = W
        .coord(5) = 1#
    End With
    PointNumber = NumPoints
End Function
' Draw the hypercube.
Private Sub DrawData(ByVal pic As PictureBox)
Dim xw_rot As Single
Dim yw_rot As Single
Dim zw_rot As Single
Dim xy_rot As Single
Dim xz_rot As Single
Dim yz_rot As Single
Dim xw2_rot As Single
Dim yw2_rot As Single
Dim zw2_rot As Single
Dim XW(1 To 5, 1 To 5) As Single
Dim YW(1 To 5, 1 To 5) As Single
Dim ZW(1 To 5, 1 To 5) As Single
Dim XY(1 To 5, 1 To 5) As Single
Dim XZ(1 To 5, 1 To 5) As Single
Dim YZ(1 To 5, 1 To 5) As Single
Dim XW2(1 To 5, 1 To 5) As Single
Dim YW2(1 To 5, 1 To 5) As Single
Dim ZW2(1 To 5, 1 To 5) As Single
Dim S(1 To 5, 1 To 5) As Single
Dim T(1 To 5, 1 To 5) As Single
Dim P(1 To 5, 1 To 5) As Single
Dim M12(1 To 5, 1 To 5) As Single
Dim M34(1 To 5, 1 To 5) As Single
Dim M1_4(1 To 5, 1 To 5) As Single
Dim M56(1 To 5, 1 To 5) As Single
Dim M78(1 To 5, 1 To 5) As Single
Dim M5_8(1 To 5, 1 To 5) As Single
Dim M1_8(1 To 5, 1 To 5) As Single
Dim M910(1 To 5, 1 To 5) As Single
Dim M1112(1 To 5, 1 To 5) As Single
Dim M9_12(1 To 5, 1 To 5) As Single
Dim M_All(1 To 5, 1 To 5) As Single
Dim D As Single

    If Not IsNumeric(txtXW.Text) Then Exit Sub
    If Not IsNumeric(txtYW.Text) Then Exit Sub
    If Not IsNumeric(txtZW.Text) Then Exit Sub
    If Not IsNumeric(txtXY.Text) Then Exit Sub
    If Not IsNumeric(txtXZ.Text) Then Exit Sub
    If Not IsNumeric(txtYZ.Text) Then Exit Sub
    If Not IsNumeric(txtXW2.Text) Then Exit Sub
    If Not IsNumeric(txtYW2.Text) Then Exit Sub
    If Not IsNumeric(txtZW2.Text) Then Exit Sub
    If Not IsNumeric(txtD.Text) Then Exit Sub
    xw_rot = CSng(txtXW.Text)
    yw_rot = CSng(txtYW.Text)
    zw_rot = CSng(txtZW.Text)
    xy_rot = CSng(txtXY.Text)
    xz_rot = CSng(txtXZ.Text)
    yz_rot = CSng(txtYZ.Text)
    xw2_rot = CSng(txtXW2.Text)
    yw2_rot = CSng(txtYW2.Text)
    zw2_rot = CSng(txtZW2.Text)
    D = CSng(txtD.Text)

    Screen.MousePointer = vbHourglass
    DoEvents
    
    ' Prevent overflow errors when drawing lines
    ' too far out of bounds.
    On Error Resume Next
    
    ' Calculate the rotation matrices.
    m4XWRotate XW, xw_rot
    m4YWRotate YW, yw_rot
    m4ZWRotate ZW, zw_rot
    m4XYRotate XY, xy_rot
    m4XZRotate XZ, xz_rot
    m4YZRotate YZ, yz_rot
    m4XWRotate XW2, xw2_rot
    m4YWRotate YW2, yw2_rot
    m4ZWRotate ZW2, zw2_rot

    ' Calculate the projection matrix.
    m4PerspectiveW P, D

    ' Scale and translate so it looks OK in pixels.
    m4Scale S, 75, -75, 1, 1
    m4Translate T, pic.ScaleWidth / 2, pic.ScaleHeight / 2, 0, 0

    m4MatMultiply M12, XW, YW
    m4MatMultiply M34, ZW, XY
    m4MatMultiply M56, XZ, YZ
    m4MatMultiplyFull M78, P, XW2
    m4MatMultiply M1_4, M12, M34
    m4MatMultiplyFull M5_8, M56, M78
    m4MatMultiplyFull M1_8, M1_4, M5_8
    m4MatMultiply M910, YW2, ZW2
    m4MatMultiply M1112, S, T
    m4MatMultiply M9_12, M910, M1112
    m4MatMultiplyFull M_All, M1_8, M9_12

    ' Transform the points.
    ApplyFull M_All

    ' Display the data.
    pic.Cls
    Draw pic
    pic.Refresh

    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Load()
    ' Create the data.
    CreateData

    ' Project and draw the data.
    Show
    DrawData picCanvas
End Sub

' Create the hypercube.
Private Sub CreateData()
Dim X As Integer
Dim Y As Integer
Dim Z As Integer
Dim W As Integer

    MousePointer = vbHourglass
    Refresh

    For X = -1 To 1 Step 2
        For Y = -1 To 1 Step 2
            For Z = -1 To 1 Step 2
                For W = -1 To 1 Step 2
                    If X = -1 Then _
                        AddSegment _
                            X, Y, Z, W, _
                            1, Y, Z, W
                    If Y = -1 Then _
                        AddSegment _
                            X, Y, Z, W, _
                            X, 1, Z, W
                    If Z = -1 Then _
                        AddSegment _
                            X, Y, Z, W, _
                            X, Y, 1, W
                    If W = -1 Then _
                        AddSegment _
                            X, Y, Z, W, _
                            X, Y, Z, 1
                Next W
            Next Z
        Next Y
    Next X

    MousePointer = vbDefault
End Sub
Private Sub txtXY_Change()
    DrawData picCanvas
End Sub

Private Sub txtXZ_Change()
    DrawData picCanvas
End Sub


Private Sub txtYW_Change()
    DrawData picCanvas
End Sub


Private Sub txtYZ_Change()
    DrawData picCanvas
End Sub

Private Sub txtZW_Change()
    DrawData picCanvas
End Sub
Private Sub txtXW_Change()
    DrawData picCanvas
End Sub
