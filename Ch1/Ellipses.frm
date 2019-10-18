VERSION 5.00
Begin VB.Form frmEllipses 
   Caption         =   "Ellipses"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picLineSegments 
      Height          =   2295
      Left            =   2400
      ScaleHeight     =   2235
      ScaleWidth      =   2235
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
   Begin VB.PictureBox picScaledCircle 
      Height          =   2295
      Left            =   4800
      ScaleHeight     =   2235
      ScaleWidth      =   2235
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.PictureBox picCircle 
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Line Segments"
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   5
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ScaledCircle"
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   4
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Circle"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmEllipses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Draw an ellipse using line segments.
Private Sub EllipseWithSegments(ByVal obj As Object, ByVal xmin As Single, ByVal ymin As Single, ByVal xmax As Single, ByVal ymax As Single)
Const PI = 3.14159265
Dim theta As Single
Dim cx As Single
Dim cy As Single
Dim radius_x As Single
Dim radius_y As Single
Dim X As Single
Dim Y As Single

    ' Find the center.
    cx = (xmin + xmax) / 2
    cy = (ymin + ymax) / 2

    ' Find the X and Y half-widths.
    radius_x = (xmax - xmin) / 2
    radius_y = (ymax - ymin) / 2

    ' Draw the ellipse.
    obj.CurrentX = cx + radius_x
    obj.CurrentY = cy
    For theta = 0 To 2 * PI Step PI / 10
        X = cx + radius_x * Cos(theta)
        Y = cy + radius_y * Sin(theta)
        obj.Line -(X, Y)
    Next theta
    obj.Line -(cx + radius_x, cy)
End Sub

' Draw a circle stretched to obey the object's
' scale mode.
Private Sub ScaledCircle(ByVal obj As Object, ByVal xmin As Single, ByVal ymin As Single, ByVal xmax As Single, ByVal ymax As Single)
Dim cx As Single
Dim cy As Single
Dim wid As Single
Dim hgt As Single
Dim aspect As Single
Dim radius As Single

    ' Find the center.
    cx = (xmin + xmax) / 2
    cy = (ymin + ymax) / 2

    ' Get the ellipse's size in twips.
    wid = obj.ScaleX(xmax - xmin, obj.ScaleMode, vbTwips)
    hgt = obj.ScaleY(ymax - ymin, obj.ScaleMode, vbTwips)
    aspect = hgt / wid

    ' See which dimension is larger.
    If wid > hgt Then
        ' The major axis is horizontal.
        ' Get the radius in custom coordinates.
        radius = obj.ScaleX(wid / 2, vbTwips, obj.ScaleMode)
    Else
        ' The major axis is vertical.
        ' Get the radius in custom coordinates.
        radius = aspect * obj.ScaleX(wid / 2, vbTwips, obj.ScaleMode)
    End If

    ' Draw the circle.
    obj.Circle (cx, cy), radius, , , , aspect
End Sub

' Define the custom scale modes and draw.
Private Sub Form_Load()
    ' Prepare the Circle picture.
    picCircle.AutoRedraw = True
    picCircle.Scale (0, 0)-(200, 100)

    picCircle.Line (10, 10)-(90, 50), , B
    picCircle.Circle (50, 30), 40
    picCircle.Line (100, 10)-(190, 70), , B
    picCircle.Circle (145, 40), 45
    picCircle.Line (10, 75)-(190, 90), , B
    picCircle.Circle (100, 82.5), 15

    ' Prepare the line segments picture.
    picLineSegments.AutoRedraw = True
    picLineSegments.Scale (0, 0)-(200, 100)

    picLineSegments.Line (10, 10)-(90, 50), , B
    EllipseWithSegments picLineSegments, 10, 10, 90, 50
    picLineSegments.Line (100, 10)-(190, 70), , B
    EllipseWithSegments picLineSegments, 100, 10, 190, 70
    picLineSegments.Line (10, 75)-(190, 90), , B
    EllipseWithSegments picLineSegments, 10, 75, 190, 90

    ' Prepare the ScaledCircle picture.
    picScaledCircle.AutoRedraw = True
    picScaledCircle.Scale (0, 0)-(200, 100)

    picScaledCircle.Line (10, 10)-(90, 50), , B
    ScaledCircle picScaledCircle, 10, 10, 90, 50
    picScaledCircle.Line (100, 10)-(190, 70), , B
    ScaledCircle picScaledCircle, 100, 10, 190, 70
    picScaledCircle.Line (10, 75)-(190, 90), , B
    ScaledCircle picScaledCircle, 10, 75, 190, 90
End Sub
