VERSION 5.00
Begin VB.Form frmFlakeAn1 
   Caption         =   "FlakeAn1"
   ClientHeight    =   4335
   ClientLeft      =   2280
   ClientTop       =   900
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4335
   ScaleWidth      =   5070
   Begin VB.TextBox txtTheta 
      Height          =   285
      Left            =   600
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "60"
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txtDepth 
      Height          =   285
      Left            =   600
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "3"
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   4335
      Left            =   1080
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   4
      Top             =   0
      Width           =   3975
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Theta"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Depth"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmFlakeAn1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PI = 3.14159

' Coordinates of the points in the initiator.
Private Const NUM_INITIATOR_POINTS = 3
Private InitiatorX(0 To NUM_INITIATOR_POINTS) As Single
Private InitiatorY(0 To NUM_INITIATOR_POINTS) As Single

' Angles and distances for the generator.
Private Const NUM_GENERATOR_ANGLES = 4
Private ScaleFactor As Single
Private GeneratorDTheta(1 To NUM_GENERATOR_ANGLES) As Single
' Draw the complete snowflake.
Private Sub DrawFlake(ByVal depth As Integer, ByVal length As Single, ByVal offset As Single)
Dim i As Integer
Dim x1 As Single
Dim y1 As Single
Dim x2 As Single
Dim y2 As Single
Dim dx As Single
Dim dy As Single
Dim theta As Single

    picCanvas.Cls

    ' Draw the snowflake.
    For i = 1 To NUM_INITIATOR_POINTS
        x1 = InitiatorX(i - 1)
        y1 = InitiatorY(i - 1)
        x2 = InitiatorX(i)
        y2 = InitiatorY(i)
        dx = x2 - x1
        dy = y2 - y1
        theta = ATan2(dy, dx)
        DrawFlakeEdge depth, x1, y1, _
            theta, length, offset
    Next i
End Sub
' Draw the animation frames.
Private Sub MakeMovie(ByVal depth As Integer, ByVal length As Single)
Const FRAMES_PER_DEPTH = 20
Const MS_PER_FRAME = 50

Dim i As Integer
Dim max_depth As Integer
Dim offset As Single
Dim doffset As Single
Dim next_time As Long

    ' Draw the animation frames.
    max_depth = depth
    next_time = GetTickCount()
    For depth = 1 To max_depth
        doffset = length * ScaleFactor ^ depth * _
            Sin(GeneratorDTheta(2)) / FRAMES_PER_DEPTH
        offset = doffset

        For i = 1 To FRAMES_PER_DEPTH
            WaitTill next_time
            DrawFlake depth, length, offset
            DoEvents
            offset = offset + doffset
            next_time = next_time + MS_PER_FRAME
        Next i
    Next depth
End Sub
Private Sub CmdGo_Click()
Dim depth As Integer
Dim length As Single
Dim theta As Single
Dim unit As Single
Dim vunit As Single
Dim hunit As Single

    picCanvas.Cls
    MousePointer = vbHourglass
    DoEvents

    ' Get the parameters.
    If Not IsNumeric(txtDepth.Text) Then txtDepth.Text = "5"
    depth = CInt(txtDepth.Text)

    ' Initialize the generator.
    If Not IsNumeric(txtTheta.Text) Then txtTheta.Text = "60"
    theta = CInt(txtTheta.Text) / 180 * PI

    ' See how big we can make the curve.
    vunit = 0.9 * picCanvas.ScaleHeight / (Sqr(3) * 4 / 3)
    hunit = 0.9 * picCanvas.ScaleWidth / 2
    If vunit < hunit Then
        unit = vunit
    Else
        unit = hunit
    End If
    length = 2 * unit

    ' Initialize the generator and initiator.
    InitializeGenerator theta, length

    ' Draw the animation frames.
    MakeMovie depth, length

    MousePointer = vbDefault
    Beep
End Sub
' Initialize the generator for the indicated angle.
Private Sub InitializeGenerator(ByVal theta As Single, ByVal length As Single)
Dim xmid As Single
Dim ymid As Single

    ' Initialize the initiator's coordinates.
    xmid = picCanvas.ScaleWidth / 2
    ymid = picCanvas.ScaleHeight / 2
    InitiatorX(1) = xmid + length / 2
    InitiatorY(1) = ymid - length / 2 * Sqr(3) / 3
    InitiatorX(2) = xmid - length / 2
    InitiatorY(2) = InitiatorY(1)
    InitiatorX(3) = xmid
    InitiatorY(3) = ymid + length / 2 * Sqr(3) * 2 / 3
    InitiatorX(0) = InitiatorX(3)
    InitiatorY(0) = InitiatorY(3)

    ScaleFactor = 1 / (2 * (1 + Cos(theta)))
    GeneratorDTheta(1) = 0
    GeneratorDTheta(2) = theta
    GeneratorDTheta(3) = -2 * theta
    GeneratorDTheta(4) = theta
End Sub

' Recursively draw a snowflake edge starting at
' (x1, y1) in direction theta and distance dist.
' Leave the coordinates of the endpoint in
' (x1, y1).
Private Sub DrawFlakeEdge(ByVal depth As Integer, ByRef x1 As Single, ByRef y1 As Single, ByVal theta As Single, ByVal dist As Single, ByVal offset As Single)
Dim status As Integer
Dim i As Integer
Dim x2 As Single
Dim y2 As Single
Dim new_theta As Single
Dim dtheta As Single
Dim hyp As Single
Dim adj As Single

    If depth <= 1 Then
        ' Draw the final depth.
        dist = dist * ScaleFactor
        adj = dist * Cos(GeneratorDTheta(2))
        hyp = Sqr(adj * adj + offset * offset)
        x2 = x1 + dist * Cos(theta)
        y2 = y1 + dist * Sin(theta)
        picCanvas.Line (x1, y1)-(x2, y2)
        x1 = x2
        y1 = y2
        
        dtheta = ATan2(offset, adj)
        new_theta = theta + dtheta
        x2 = x1 + hyp * Cos(new_theta)
        y2 = y1 + hyp * Sin(new_theta)
        picCanvas.Line (x1, y1)-(x2, y2)
        x1 = x2
        y1 = y2
        
        new_theta = theta - dtheta
        x2 = x1 + hyp * Cos(new_theta)
        y2 = y1 + hyp * Sin(new_theta)
        picCanvas.Line (x1, y1)-(x2, y2)
        x1 = x2
        y1 = y2
        
        x2 = x1 + dist * Cos(theta)
        y2 = y1 + dist * Sin(theta)
        picCanvas.Line (x1, y1)-(x2, y2)
        x1 = x2
        y1 = y2
        
        Exit Sub
    End If
    
    ' Recursively draw the edge.
    dist = dist * ScaleFactor
    For i = 1 To NUM_GENERATOR_ANGLES
        theta = theta + GeneratorDTheta(i)
        DrawFlakeEdge depth - 1, x1, y1, theta, dist, offset
    Next i
End Sub
Private Sub Form_Resize()
Dim wid As Single

    ' Make the picCanvas as big as possible.
    wid = ScaleWidth - picCanvas.Left
    If wid < 120 Then wid = 120
    picCanvas.Move picCanvas.Left, 0, wid, ScaleHeight
End Sub

