VERSION 5.00
Begin VB.Form frmAnimated 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Animated"
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
      Height          =   1335
      Index           =   0
      Left            =   5400
      TabIndex        =   10
      Top             =   0
      Width           =   1455
      Begin VB.CheckBox chkXY 
         Caption         =   "XY Plane"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkXZ 
         Caption         =   "XZ Plane"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkYZ 
         Caption         =   "YZ Plane"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Post-Rotations"
      Height          =   1335
      Index           =   1
      Left            =   5400
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
      Begin VB.TextBox txtXW2 
         Height          =   285
         Left            =   600
         MaxLength       =   6
         TabIndex        =   6
         Text            =   "0.2"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtYW2 
         Height          =   285
         Left            =   600
         MaxLength       =   6
         TabIndex        =   5
         Text            =   "0.1"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtZW2 
         Height          =   285
         Left            =   600
         MaxLength       =   6
         TabIndex        =   4
         Text            =   "0.0"
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "X"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Y"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Z"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   255
      End
   End
   Begin VB.TextBox txtD 
      Height          =   285
      Left            =   6000
      TabIndex        =   2
      Text            =   "3"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   3600
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
      Index           =   12
      Left            =   5640
      TabIndex        =   14
      Top             =   1560
      Width           =   255
   End
End
Attribute VB_Name = "frmAnimated"
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

Private Running As Boolean
' Animate the hypercube.
Private Sub Animate(ByVal pic As PictureBox)
Const Dtheta = PI / 40

Dim xy_rot As Single
Dim xz_rot As Single
Dim yz_rot As Single
Dim xw2_rot As Single
Dim yw2_rot As Single
Dim zw2_rot As Single
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
Dim M1_6(1 To 5, 1 To 5) As Single
Dim M_All(1 To 5, 1 To 5) As Single
Dim D As Single
Dim AnimateXY As Boolean
Dim AnimateXZ As Boolean
Dim AnimateYZ As Boolean
Dim next_time As Long

    If Not IsNumeric(txtXW2.Text) Then Exit Sub
    If Not IsNumeric(txtYW2.Text) Then Exit Sub
    If Not IsNumeric(txtZW2.Text) Then Exit Sub
    If Not IsNumeric(txtD.Text) Then Exit Sub
    xw2_rot = CSng(txtXW2.Text)
    yw2_rot = CSng(txtYW2.Text)
    zw2_rot = CSng(txtZW2.Text)
    D = CSng(txtD.Text)

    Screen.MousePointer = vbHourglass
    DoEvents

    ' Prevent overflow errors when drawing lines
    ' too far out of bounds.
    On Error Resume Next
    
    ' Calculate the matrices that don't change.
    m4XWRotate XW2, xw2_rot
    m4YWRotate YW2, yw2_rot
    m4ZWRotate ZW2, zw2_rot

    ' Calculate the projection matrix.
    m4PerspectiveW P, D

    ' Scale and translate so it looks OK in pixels.
    m4Scale S, 75, -75, 1, 1
    m4Translate T, pic.ScaleWidth / 2, pic.ScaleHeight / 2, 0, 0

    m4MatMultiplyFull M12, P, XW2
    m4MatMultiply M34, YW2, ZW2
    m4MatMultiplyFull M1_4, M12, M34
    m4MatMultiply M56, S, T
    m4MatMultiplyFull M1_6, M1_4, M56

    ' See which rotations we are animating.
    AnimateXY = (chkXY.value = vbChecked)
    AnimateXZ = (chkXZ.value = vbChecked)
    AnimateYZ = (chkYZ.value = vbChecked)

    ' Start the animation.
    Do While Running
        next_time = GetTickCount + 50

        ' Calculate the changing transformations.
        m4XYRotate XY, xy_rot
        m4XZRotate XZ, xz_rot
        m4YZRotate YZ, yz_rot

        m4MatMultiply M12, XY, XZ
        m4MatMultiply M1_4, M12, YZ
        m4MatMultiplyFull M_All, M1_4, M1_6

        If AnimateXY Then xy_rot = xy_rot + Dtheta
        If AnimateXZ Then xz_rot = xz_rot + Dtheta
        If AnimateYZ Then yz_rot = yz_rot + Dtheta

        ' Transform the points.
        ApplyFull M_All

        ' Display the data.
        pic.Cls
        Draw pic
        DoEvents

        WaitTill next_time
    Loop

    Screen.MousePointer = vbDefault
End Sub
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
Private Sub cmdGo_Click()
    If Running Then
        ' Stop it.
        cmdGo.Caption = "Go"
        Running = False
    Else
        cmdGo.Caption = "Stop"
        Running = True
        Animate picCanvas
    End If
End Sub

Private Sub Form_Load()
    ' Create the data.
    CreateData
End Sub

' Create the hypercube.
Private Sub CreateData()
Dim X As Integer
Dim Y As Integer
Dim Z As Integer
Dim W As Integer

    Screen.MousePointer = vbHourglass
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

    Screen.MousePointer = vbDefault
End Sub
