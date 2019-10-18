VERSION 5.00
Begin VB.Form frmClip 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Clip"
   ClientHeight    =   5670
   ClientLeft      =   1770
   ClientTop       =   615
   ClientWidth     =   5310
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
   ScaleWidth      =   5310
   Begin VB.TextBox txtDistance 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Text            =   "10.0"
      Top             =   0
      Width           =   735
   End
   Begin VB.CheckBox ClipCheck 
      Caption         =   "Clip"
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   5295
      Left            =   0
      ScaleHeight     =   -14
      ScaleLeft       =   -7
      ScaleMode       =   0  'User
      ScaleTop        =   7
      ScaleWidth      =   14
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "Distance to origin:"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmClip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Location of viewing eye.
Private EyeR As Single
Private EyeTheta As Single
Private EyePhi As Single

Private Const Dtheta = PI / 20
Private Const Dphi = PI / 20
Private Const Dr = 1

' Location of focus point.
Const FocusX = 0#
Const FocusY = 0#
Const FocusZ = 0#

Private Projector(1 To 4, 1 To 4) As Single
' Rotate the points in the cube and draw the cube.
Private Sub DrawData(ByVal pic As Object)
Dim i As Integer
Dim x1 As Single
Dim y1 As Single
Dim z1 As Single
Dim x2 As Single
Dim y2 As Single
Dim z2 As Single
Dim do_clip As Boolean
Dim draw_seg As Boolean

    ' Get the distance.
    On Error Resume Next
    EyeR = CSng(txtDistance.Text)
    If Err.Number <> 0 Then
        EyeR = 10#
        txtDistance.Text = Format$(EyeR)
    End If
    On Error GoTo 0

    ' Get the transformation matrix.
    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0

    ' Transform the points.
    TransformAllDataFull Projector
    
    ' Draw the points, using On Error to avoid
    ' overflows when drawing lines far out
    ' of bounds.
    On Error Resume Next
    pic.Cls
    do_clip = (ClipCheck.value = vbChecked)
    draw_seg = True
    For i = 1 To NumSegments
        If do_clip Then
            z1 = Segments(i).fr_tr(3)
            z2 = Segments(i).to_tr(3)
            ' Don't draw if either point is farther
            ' from the focus point than the center of
            ' projection (which is distance EyeR away).
            draw_seg = (z1 < EyeR And z2 < EyeR)
        End If
        If draw_seg Then
            x1 = Segments(i).fr_tr(1)
            y1 = Segments(i).fr_tr(2)
            x2 = Segments(i).to_tr(1)
            y2 = Segments(i).to_tr(2)
            pic.Line (x1, y1)-(x2, y2)
        End If
    Next i
    
    pic.Refresh
End Sub



Private Sub ClipCheck_Click()
    DrawData picCanvas
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft
            EyeTheta = EyeTheta - Dtheta
        
        Case vbKeyRight
            EyeTheta = EyeTheta + Dtheta
        
        Case vbKeyUp
            EyePhi = EyePhi - Dphi
        
        Case vbKeyDown
            EyePhi = EyePhi + Dphi
                
        Case Else
            Exit Sub
    End Select

    DrawData picCanvas
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    DrawData picCanvas
End Sub

Private Sub Form_Load()
    ' Initialize the eye position.
    EyeR = 10
    EyeTheta = PI * 0.2
    EyePhi = PI * 0.05

    ' Initialize the projection transformation.
    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0

    ' Create the data.
    CreateData

    ' Project and draw the data.
    DrawData picCanvas
End Sub

Private Sub CreateData()
Const WID = 1

Dim X As Single
Dim Y As Single
Dim Z As Single

    ' Create the cubes.
    For X = -2 To 2 Step 4
        For Y = -2 To 2 Step 4
            For Z = -2 To 2 Step 4
                MakeSegment X - WID, Y - WID, Z - WID, X - WID, Y - WID, Z + WID
                MakeSegment X - WID, Y - WID, Z + WID, X - WID, Y + WID, Z + WID
                MakeSegment X - WID, Y + WID, Z + WID, X - WID, Y + WID, Z - WID
                MakeSegment X - WID, Y + WID, Z - WID, X - WID, Y - WID, Z - WID
                MakeSegment X + WID, Y - WID, Z - WID, X + WID, Y - WID, Z + WID
                MakeSegment X + WID, Y - WID, Z + WID, X + WID, Y + WID, Z + WID
                MakeSegment X + WID, Y + WID, Z + WID, X + WID, Y + WID, Z - WID
                MakeSegment X + WID, Y + WID, Z - WID, X + WID, Y - WID, Z - WID
                MakeSegment X - WID, Y - WID, Z - WID, X + WID, Y - WID, Z - WID
                MakeSegment X - WID, Y - WID, Z + WID, X + WID, Y - WID, Z + WID
                MakeSegment X - WID, Y + WID, Z + WID, X + WID, Y + WID, Z + WID
                MakeSegment X - WID, Y + WID, Z - WID, X + WID, Y + WID, Z - WID
            Next Z
        Next Y
    Next X
End Sub

Private Sub txtDistance_Change()
    DrawData picCanvas
End Sub


