VERSION 5.00
Begin VB.Form frmSierpG2 
   Caption         =   "SierpG2"
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
      Text            =   "5"
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
Attribute VB_Name = "frmSierpG2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdGo_Click()
Const PI = 3.14159265

Dim depth As Integer
Dim total_length As Single
Dim start_x As Single
Dim start_y As Single
Dim wid As Single
Dim hgt As Single

    picCanvas.Cls
    MousePointer = vbHourglass
    DoEvents

    ' Get the parameters.
    If Not IsNumeric(txtDepth.Text) Then txtDepth.Text = "5"
    depth = CInt(txtDepth.Text)

    ' See how big we can make the curve.
    wid = picCanvas.ScaleWidth * 0.9
    hgt = picCanvas.ScaleHeight * 0.9 * 2 / Sqr(3)
    If wid < hgt Then
        total_length = wid
    Else
        total_length = hgt
    End If

    start_x = picCanvas.ScaleWidth * 0.05
    start_y = picCanvas.ScaleHeight * 0.95

    ' Draw the curve.
    picCanvas.CurrentX = start_x
    picCanvas.CurrentY = start_y
    SierpinskiGasketCurve depth, 0, total_length, -PI / 3

    MousePointer = vbDefault
End Sub
' Draw a gasket-like Sierpinski curve.
Private Sub SierpinskiGasketCurve(ByVal depth As Integer, ByVal theta As Single, ByVal dist As Single, ByVal turn As Single)
    If depth > 0 Then
        SierpinskiGasketCurve depth - 1, theta + turn, dist / 2, -turn
        SierpinskiGasketCurve depth - 1, theta, dist / 2, turn
        SierpinskiGasketCurve depth - 1, theta - turn, dist / 2, -turn
    Else
        picCanvas.Line -Step(dist * Cos(theta), dist * Sin(theta))
    End If
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
