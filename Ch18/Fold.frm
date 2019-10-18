VERSION 5.00
Begin VB.Form frmFold 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Fold"
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
      Caption         =   "Post-Rotations"
      Height          =   1335
      Index           =   1
      Left            =   5400
      TabIndex        =   3
      Top             =   480
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
      Text            =   "5"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   2040
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
      TabIndex        =   10
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmFold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Location of focus point.
Private Const FocusX = 0#
Private Const FocusY = 0#
Private Const FocusZ = 0#

' The objects.
Private TheCubes(1 To 8) As Polyline4d

Private Projector(1 To 5, 1 To 5) As Single

Private Running As Boolean

Private Const MS_PER_FRAME = 50
Private Const FRAMES_PER_FOLD = 20

' Animate folding this cube across the X = x,
' W = w plane.
Private Sub FoldXW(ByVal col As Collection, ByVal X As Single, ByVal W As Single, ByVal theta As Single)
Dim cube As Polyline4d
Dim i As Single
Dim next_time As Long
Dim T1(1 To 5, 1 To 5) As Single
Dim r(1 To 5, 1 To 5) As Single
Dim T2(1 To 5, 1 To 5) As Single
Dim T1R(1 To 5, 1 To 5) As Single
Dim All(1 To 5, 1 To 5) As Single

    next_time = GetTickCount + MS_PER_FRAME

    ' Create the transformation matrices.
    m4Translate T1, -X, 0, 0, -W
    m4Translate T2, X, 0, 0, W
    m4YZRotate r, theta / FRAMES_PER_FOLD
    m4MatMultiply T1R, T1, r
    m4MatMultiply All, T1R, T2

    For i = 1 To FRAMES_PER_FOLD
        If Not Running Then Exit Sub

        ' Rotate the cubes.
        For Each cube In col
            cube.Apply All
            cube.FixPoints
        Next cube

        ' Wait until it's time for the next image.
        WaitTill next_time
        next_time = GetTickCount + MS_PER_FRAME

        ' Display the picture.
        Draw picCanvas
    Next i
End Sub

' Animate the hypercube.
Private Sub Animate(ByVal pic As PictureBox)
Dim xw2_rot As Single
Dim yw2_rot As Single
Dim zw2_rot As Single
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
Dim D As Single
Dim col As Collection

    If Not IsNumeric(txtXW2.Text) Then Exit Sub
    If Not IsNumeric(txtYW2.Text) Then Exit Sub
    If Not IsNumeric(txtZW2.Text) Then Exit Sub
    If Not IsNumeric(txtD.Text) Then Exit Sub
    xw2_rot = CSng(txtXW2.Text)
    yw2_rot = CSng(txtYW2.Text)
    zw2_rot = CSng(txtZW2.Text)
    D = CSng(txtD.Text)

    ' Create fresh data.
    CreateData

    ' Calculate the matrices.
    m4XWRotate XW2, xw2_rot
    m4YWRotate YW2, yw2_rot
    m4ZWRotate ZW2, zw2_rot
    m4PerspectiveW P, D
    m4Scale S, 25, -25, 1, 1
    m4Translate T, pic.ScaleWidth * 0.75, pic.ScaleHeight / 2, 0, 0

    m4MatMultiplyFull M12, P, XW2
    m4MatMultiply M34, YW2, ZW2
    m4MatMultiplyFull M1_4, M12, M34
    m4MatMultiply M56, S, T
    m4MatMultiplyFull Projector, M1_4, M56

    ' Present the original image.
    If Not Running Then Exit Sub
    ApplyFull Projector
    pic.Cls
    Draw pic
    DoEvents

    ' Fold up cube 5.
    Set col = New Collection
    col.Add TheCubes(5)
    FoldYW col, 1, 0, PI / 2
    If Not Running Then Exit Sub

    ' Fold up cube 6.
    Set col = New Collection
    col.Add TheCubes(6)
    FoldZW col, -1, 0, -PI / 2
    If Not Running Then Exit Sub

    ' Fold up cube 4.
    Set col = New Collection
    col.Add TheCubes(4)
    FoldXW col, 1, 0, PI / 2
    If Not Running Then Exit Sub

    ' Fold up cube 7.
    Set col = New Collection
    col.Add TheCubes(7)
    FoldYW col, -1, 0, -PI / 2
    If Not Running Then Exit Sub

    ' Fold up cube 8.
    Set col = New Collection
    col.Add TheCubes(8)
    FoldZW col, 1, 0, PI / 2
    If Not Running Then Exit Sub

    ' Fold up cubes 2 and 1 together.
    Set col = New Collection
    col.Add TheCubes(1)
    col.Add TheCubes(2)
    FoldXW col, -1, 0, -PI / 2
    If Not Running Then Exit Sub

    ' Finish folding cube 1.
    Set col = New Collection
    col.Add TheCubes(1)
    FoldXW col, -1, 2, -PI / 2
    If Not Running Then Exit Sub
End Sub
' Animate folding this cube across the Y = y,
' W = w plane.
Private Sub FoldYW(ByVal col As Collection, ByVal y As Single, ByVal W As Single, ByVal theta As Single)
Dim cube As Polyline4d
Dim i As Single
Dim next_time As Long
Dim T1(1 To 5, 1 To 5) As Single
Dim r(1 To 5, 1 To 5) As Single
Dim T2(1 To 5, 1 To 5) As Single
Dim T1R(1 To 5, 1 To 5) As Single
Dim All(1 To 5, 1 To 5) As Single

    next_time = GetTickCount + MS_PER_FRAME
    
    ' Create the transformation matrices.
    m4Translate T1, 0, -y, 0, -W
    m4Translate T2, 0, y, 0, W
    m4XZRotate r, theta / FRAMES_PER_FOLD
    m4MatMultiply T1R, T1, r
    m4MatMultiply All, T1R, T2

    For i = 1 To FRAMES_PER_FOLD
        If Not Running Then Exit Sub

        ' Rotate the cubes.
        For Each cube In col
            cube.Apply All
            cube.FixPoints
        Next cube

        ' Wait until it's time for the next image.
        WaitTill next_time
        next_time = GetTickCount + MS_PER_FRAME

        ' Display the picture.
        Draw picCanvas
    Next i
End Sub
' Animate folding this cube across the Z = z,
' W = w plane.
Private Sub FoldZW(ByVal col As Collection, ByVal z As Single, ByVal W As Single, ByVal theta As Single)
Dim cube As Polyline4d
Dim i As Single
Dim next_time As Long
Dim T1(1 To 5, 1 To 5) As Single
Dim r(1 To 5, 1 To 5) As Single
Dim T2(1 To 5, 1 To 5) As Single
Dim T1R(1 To 5, 1 To 5) As Single
Dim All(1 To 5, 1 To 5) As Single

    next_time = GetTickCount + MS_PER_FRAME
    
    ' Create the transformation matrices.
    m4Translate T1, 0, 0, -z, -W
    m4Translate T2, 0, 0, z, W
    m4XYRotate r, theta / FRAMES_PER_FOLD
    m4MatMultiply T1R, T1, r
    m4MatMultiply All, T1R, T2

    For i = 1 To FRAMES_PER_FOLD
        If Not Running Then Exit Sub

        ' Rotate the cubes.
        For Each cube In col
            cube.Apply All
            cube.FixPoints
        Next cube

        ' Wait until it's time for the next image.
        WaitTill next_time
        next_time = GetTickCount + MS_PER_FRAME

        ' Display the picture.
        Draw picCanvas
    Next i
End Sub
' Apply this matrix to the points.
Private Sub Apply(M() As Single)
Dim i As Integer

    For i = 1 To 8
        TheCubes(i).Apply M
    Next i
End Sub
' Apply this matrix to the points.
Private Sub ApplyFull(M() As Single)
Dim i As Integer

    For i = 1 To 8
        TheCubes(i).ApplyFull M
    Next i
End Sub

' Draw the data in its current form.
Private Sub Draw(ByVal pic As PictureBox)
Dim i As Integer

    ' Apply the projection transformation.
    ApplyFull Projector

    ' Draw the cubes.
    pic.Cls
    For i = 1 To 8
        TheCubes(i).Draw pic
    Next i
End Sub
' Create the projection matrix.
Private Sub CreateProjector()
Dim xw2_rot As Single
Dim yw2_rot As Single
Dim zw2_rot As Single
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
Dim D As Single

    If Not IsNumeric(txtXW2.Text) Then Exit Sub
    If Not IsNumeric(txtYW2.Text) Then Exit Sub
    If Not IsNumeric(txtZW2.Text) Then Exit Sub
    If Not IsNumeric(txtD.Text) Then Exit Sub
    xw2_rot = CSng(txtXW2.Text)
    yw2_rot = CSng(txtYW2.Text)
    zw2_rot = CSng(txtZW2.Text)
    D = CSng(txtD.Text)

    ' Calculate the matrices.
    m4XWRotate XW2, xw2_rot
    m4YWRotate YW2, yw2_rot
    m4ZWRotate ZW2, zw2_rot
    m4PerspectiveW P, D
    m4Scale S, 25, -25, 1, 1
    m4Translate T, picCanvas.ScaleWidth * 0.75, picCanvas.ScaleHeight / 2, 0, 0

    m4MatMultiplyFull M12, P, XW2
    m4MatMultiply M34, YW2, ZW2
    m4MatMultiplyFull M1_4, M12, M34
    m4MatMultiply M56, S, T
    m4MatMultiplyFull Projector, M1_4, M56
End Sub
Private Sub CmdGo_Click()
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
' Create a cube with the indicated minimum
' coordinates. W = 0 for all points.
Private Sub CreateCube(ByRef cube As Polyline4d, ByVal xmin As Single, ByVal ymin As Single, ByVal zmin As Single)
Dim X As Single
Dim y As Single
Dim z As Single

    Set cube = New Polyline4d

    For X = xmin To xmin + 2 Step 2
        For y = ymin To ymin + 2 Step 2
            For z = zmin To zmin + 2 Step 2
                If X = xmin Then _
                    cube.AddSegment _
                        X, y, z, 0, _
                        X + 2, y, z, 0
                If y = ymin Then _
                    cube.AddSegment _
                        X, y, z, 0, _
                        X, y + 2, z, 0
                If z = zmin Then _
                    cube.AddSegment _
                        X, y, z, 0, _
                        X, y, z + 2, 0
            Next z
        Next y
    Next X
End Sub
' Create the folded out hypercube.
Private Sub CreateData()
    Screen.MousePointer = vbHourglass
    Refresh

    CreateCube TheCubes(1), -5, -1, -1
    CreateCube TheCubes(2), -3, -1, -1
    CreateCube TheCubes(3), -1, -1, -1
    CreateCube TheCubes(4), 1, -1, -1
    CreateCube TheCubes(5), -1, 1, -1
    CreateCube TheCubes(6), -1, -1, -3
    CreateCube TheCubes(7), -1, -3, -1
    CreateCube TheCubes(8), -1, -1, 1

    Screen.MousePointer = vbDefault
End Sub
