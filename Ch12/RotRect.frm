VERSION 5.00
Begin VB.Form frmRotRect 
   AutoRedraw      =   -1  'True
   Caption         =   "RotRect"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAngle 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Text            =   "30"
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Angle"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmRotRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Dragging As Boolean
Private StartX As Single
Private StartY As Single
Private LastX As Single
Private LastY As Single

Private Objects As Collection

Private Sub Form_Load()
    ScaleMode = vbPixels
    Set Objects = New Collection
End Sub

' Start selecting a box.
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = True
    StartX = X
    StartY = Y
    LastX = X
    LastY = Y

    DrawMode = vbInvert
    Line (StartX, StartY)-(LastX, LastY), , B
End Sub
' Continue selecting an area.
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Dragging Then Exit Sub

    Line (StartX, StartY)-(LastX, LastY), , B
    LastX = X
    LastY = Y
    Line (StartX, StartY)-(LastX, LastY), , B
End Sub


' Finish selecting an area.
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const PI = 3.14159265

Dim pgon As TwoDPolygon
Dim M(1 To 3, 1 To 3) As Single

    If Not Dragging Then Exit Sub
    Dragging = False

    Line (StartX, StartY)-(LastX, LastY), , B
    LastX = X
    LastY = Y

    ' Create the new rectangle polygon.
    Set pgon = New TwoDPolygon
    pgon.NumPoints = 4
    pgon.X(1) = StartX
    pgon.X(2) = LastX
    pgon.X(3) = LastX
    pgon.X(4) = StartX
    pgon.Y(1) = StartY
    pgon.Y(2) = StartY
    pgon.Y(3) = LastY
    pgon.Y(4) = LastY

    ' Create the rotation transformation.
    m2RotateAround M, CSng(txtAngle.Text) * PI / 180, _
        (StartX + LastX) / 2, (StartY + LastY) / 2

    ' Rotate the polygon.
    pgon.Transform M

    ' Add the polygon to the objects collection.
    Objects.Add pgon

    ' Redraw.
    DrawObjects
End Sub

Private Sub DrawObjects()
Dim pgon As TwoDPolygon

    DrawMode = vbCopyPen
    Cls
    For Each pgon In Objects
        pgon.Draw Me
    Next pgon
End Sub


