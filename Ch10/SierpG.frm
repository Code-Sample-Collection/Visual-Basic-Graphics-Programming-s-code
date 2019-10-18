VERSION 5.00
Begin VB.Form frmSierpG 
   Caption         =   "SierpG"
   ClientHeight    =   4335
   ClientLeft      =   2280
   ClientTop       =   900
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4335
   ScaleWidth      =   5310
   Begin VB.TextBox txtDepth 
      Height          =   285
      Left            =   480
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "4"
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   4335
      Left            =   960
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   285
      TabIndex        =   3
      Top             =   0
      Width           =   4335
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Depth"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmSierpG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

' Erase the center triangle from this one.
Private Sub SierpinskiErase(ByVal depth As Integer, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single)
Dim newy As Single
Dim newx1 As Single
Dim newx2 As Single
Dim newx3 As Single
Dim points(1 To 3) As POINTAPI

    ' Find the corners of the middle triangle.
    newy = (y1 + y2) / 2
    newx1 = (3 * x1 + x3) / 4
    newx2 = (x1 + x3) / 2
    newx3 = (x1 + 3 * x3) / 4

    ' Erase the middle triangle.
    points(1).X = newx1
    points(1).Y = newy
    points(2).X = newx3
    points(2).Y = newy
    points(3).X = newx2
    points(3).Y = y1
    Polygon picCanvas.hdc, points(1), 3

    ' Recursively erase other subtriangles.
    If depth > 0 Then
        SierpinskiErase depth - 1, x1, y1, newx1, newy, newx2, y1
        SierpinskiErase depth - 1, newx1, newy, newx2, y2, newx3, newy
        SierpinskiErase depth - 1, newx2, y1, newx3, newy, x3, y1
    End If
End Sub
' Draw a complete Sierpinski gasket.
Private Sub SierpinskiGasket(ByVal depth As Integer, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single)
Dim points(1 To 3) As POINTAPI

    ' Erase the picture.
    picCanvas.Line (0, 0)-(picCanvas.ScaleWidth, picCanvas.ScaleHeight), picCanvas.BackColor, BF

    ' Draw the main filled triangle.
    picCanvas.AutoRedraw = True
    picCanvas.FillStyle = vbFSSolid
    picCanvas.FillColor = vbBlack
    points(1).X = x1
    points(1).Y = y1
    points(2).X = x2
    points(2).Y = y2
    points(3).X = x3
    points(3).Y = y3
    Polygon picCanvas.hdc, points(1), 3

    ' If depth > 0, call SierpinskiErase to
    ' erase the center of this triangle.
    If depth >= 0 Then
        picCanvas.FillColor = picCanvas.BackColor
        SierpinskiErase depth, x1, y1, x2, y2, x3, y3
    End If

    ' Make the results visible.
    picCanvas.Refresh
    picCanvas.Picture = picCanvas.Image
End Sub

Private Sub CmdGo_Click()
Dim depth As Integer
Dim x1 As Single
Dim y1 As Single
Dim x2 As Single
Dim y2 As Single
Dim x3 As Single
Dim y3 As Single

    MousePointer = vbHourglass
    DoEvents

    ' Get the parameters.
    If Not IsNumeric(txtDepth.Text) Then txtDepth.Text = "5"
    depth = CInt(txtDepth.Text)

    ' See where the first corners should be.
    x1 = picCanvas.ScaleWidth * 0.05
    x2 = picCanvas.ScaleWidth * 0.5
    x3 = picCanvas.ScaleWidth * 0.95
    y1 = picCanvas.ScaleHeight * 0.95
    y2 = picCanvas.ScaleHeight * 0.05
    y3 = y1

    ' Draw the curve.
    SierpinskiGasket depth, x1, y1, x2, y2, x3, y3

    MousePointer = vbDefault
End Sub
' Draw a hilbert curve.
Private Sub Hilbert(ByVal depth As Integer, ByVal dx As Single, ByVal dy As Single)
    If depth > 1 Then Hilbert depth - 1, dy, dx
    picCanvas.Line -Step(dx, dy)
    If depth > 1 Then Hilbert depth - 1, dx, dy
    picCanvas.Line -Step(dy, dx)
    If depth > 1 Then Hilbert depth - 1, dx, dy
    picCanvas.Line -Step(-dx, -dy)
    If depth > 1 Then Hilbert depth - 1, -dy, -dx
End Sub

Private Sub Form_Resize()
Dim wid As Single

    wid = ScaleWidth - picCanvas.Left
    If wid < 120 Then wid = 120
    picCanvas.Move picCanvas.Left, 0, _
        wid, ScaleHeight
End Sub
