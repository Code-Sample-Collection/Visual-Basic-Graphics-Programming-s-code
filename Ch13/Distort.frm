VERSION 5.00
Begin VB.Form frmDistort 
   Caption         =   "Distort"
   ClientHeight    =   6270
   ClientLeft      =   1710
   ClientTop       =   465
   ClientWidth     =   5910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6270
   ScaleWidth      =   5910
   Begin VB.OptionButton optDistortion 
      Caption         =   "Twist"
      Height          =   255
      Index           =   3
      Left            =   3240
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
   Begin VB.OptionButton optDistortion 
      Caption         =   "Sines"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin VB.OptionButton optDistortion 
      Caption         =   "Wave"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.OptionButton optDistortion 
      Caption         =   "None"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   5895
      Left            =   0
      ScaleHeight     =   -800
      ScaleLeft       =   -400
      ScaleMode       =   0  'User
      ScaleTop        =   400
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   360
      Width           =   5895
   End
End
Attribute VB_Name = "frmDistort"
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

Private Polylines As Collection
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

    DrawData picCanvas
End Sub

Private Sub Form_Load()
    Me.Show
    MousePointer = vbHourglass
    DoEvents
    
    ' Initialize the eye position.
    EyeR = 1500
    EyeTheta = PI * 0.17
    EyePhi = PI * 0.16
    
    ' Start with no distortion.
    optDistortion(0).value = True

    MousePointer = vbDefault
End Sub


' Create some polylines to display.
Private Sub CreateData()
Const GAP = 25
Dim i As Single
Dim j As Single
Dim poly As Polyline3d

    ' Create the polyline collection.
    Set Polylines = New Collection

    ' Create the top (perpendicular to Y axis).
    Set poly = New Polyline3d
    Polylines.Add poly
    For i = -200 To 200 Step GAP
        For j = -200 To 200 - GAP Step GAP
            poly.AddSegment i, 200, j, i, 200, j + GAP
            poly.AddSegment j, 200, i, j + GAP, 200, i
        Next j
    Next i

    ' Create the front (perpendicular to Z axis).
    Set poly = New Polyline3d
    Polylines.Add poly
    For i = -200 To 200 Step GAP
        For j = -200 To 200 - GAP Step GAP
            poly.AddSegment i, j, 200, i, j + GAP, 200
            poly.AddSegment j, i, 200, j + GAP, i, 200
        Next j
    Next i

    ' Create the side (perpendicular to X axis).
    Set poly = New Polyline3d
    Polylines.Add poly
    For i = -200 To 200 Step GAP
        For j = -200 To 200 - GAP Step GAP
            poly.AddSegment 200, i, j, 200, i, j + GAP
            poly.AddSegment 200, j, i, 200, j + GAP, i
        Next j
    Next i
End Sub
' Display the data.
Private Sub DrawData(ByVal pic As PictureBox)
Dim pline As Polyline3d
Dim trans As Distortion
Dim trans_sines As DistortSines
Dim trans_twist As DistortTwist
Dim trans_circle As DistortCircle

    Screen.MousePointer = vbHourglass
    pic.Cls
    DoEvents

    ' Recreate the data.
    CreateData

    ' Load the correct transformation.
    If optDistortion(1).value Then
        ' Circle.
        Set trans = New DistortCircle
        Set trans_circle = trans
        trans_circle.amplitude = 12.5
        trans_circle.period = 150
    ElseIf optDistortion(2).value Then
        ' Sines.
        Set trans = New DistortSines
        Set trans_sines = trans
        trans_sines.amplitude = 30
        trans_sines.period = 150
    ElseIf optDistortion(3).value Then
        ' Twist.
        Set trans = New DistortTwist
        Set trans_twist = trans
        trans_twist.Cx = 0
        trans_twist.Cz = 0
        trans_twist.offset = -100
        trans_twist.period = 500 * PI
    End If

    ' Apply the transformation.
    If Not (trans Is Nothing) Then
        For Each pline In Polylines
            pline.Transform trans
        Next pline
    End If

    ' Build the projection transformation.
    m3PProject Projector, m3Perspective, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0

    ' Transform and draw the polylines.
    pic.Cls
    For Each pline In Polylines
        pline.ApplyFull Projector
        pline.Draw pic
    Next pline

    Screen.MousePointer = vbDefault
End Sub

' Display the data with the new transformation.
Private Sub optDistortion_Click(Index As Integer)
    DrawData picCanvas
End Sub
