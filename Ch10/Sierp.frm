VERSION 5.00
Begin VB.Form frmSierp 
   Caption         =   "Sierp"
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
      Text            =   "3"
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
Attribute VB_Name = "frmSierp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Draw a type A sierpinski sub-curve.
Private Sub SierpA(ByVal depth As Integer, ByVal dist As Single)
    If depth = 1 Then
        picCanvas.Line -Step(-dist, dist)
        picCanvas.Line -Step(-dist, 0)
        picCanvas.Line -Step(-dist, -dist)
    Else
        SierpA depth - 1, dist
        picCanvas.Line -Step(-dist, dist)
        SierpB depth - 1, dist
        picCanvas.Line -Step(-dist, 0)
        SierpD depth - 1, dist
        picCanvas.Line -Step(-dist, -dist)
        SierpA depth - 1, dist
    End If
End Sub

' Draw a type B sierpinski sub-curve.
Private Sub SierpB(ByVal depth As Integer, ByVal dist As Single)
    If depth = 1 Then
        picCanvas.Line -Step(dist, dist)
        picCanvas.Line -Step(0, dist)
        picCanvas.Line -Step(-dist, dist)
    Else
        SierpB depth - 1, dist
        picCanvas.Line -Step(dist, dist)
        SierpC depth - 1, dist
        picCanvas.Line -Step(0, dist)
        SierpA depth - 1, dist
        picCanvas.Line -Step(-dist, dist)
        SierpB depth - 1, dist
    End If
End Sub


' Draw a type C sierpinski sub-curve.
Private Sub SierpC(ByVal depth As Integer, ByVal dist As Single)
    If depth = 1 Then
        picCanvas.Line -Step(dist, -dist)
        picCanvas.Line -Step(dist, 0)
        picCanvas.Line -Step(dist, dist)
    Else
        SierpC depth - 1, dist
        picCanvas.Line -Step(dist, -dist)
        SierpD depth - 1, dist
        picCanvas.Line -Step(dist, 0)
        SierpB depth - 1, dist
        picCanvas.Line -Step(dist, dist)
        SierpC depth - 1, dist
    End If
End Sub
' Draw a type D sierpinski sub-curve.
Private Sub SierpD(ByVal depth As Integer, ByVal dist As Single)
    If depth = 1 Then
        picCanvas.Line -Step(-dist, -dist)
        picCanvas.Line -Step(0, -dist)
        picCanvas.Line -Step(dist, -dist)
    Else
        SierpD depth - 1, dist
        picCanvas.Line -Step(-dist, -dist)
        SierpA depth - 1, dist
        picCanvas.Line -Step(0, -dist)
        SierpC depth - 1, dist
        picCanvas.Line -Step(dist, -dist)
        SierpD depth - 1, dist
    End If
End Sub
' Draw the complete Sierpinski curve.
Private Sub Sierpinski(depth As Integer, dist As Single)
    SierpB depth, dist
    picCanvas.Line -Step(dist, dist)
    SierpC depth, dist
    picCanvas.Line -Step(dist, -dist)
    SierpD depth, dist
    picCanvas.Line -Step(-dist, -dist)
    SierpA depth, dist
    picCanvas.Line -Step(-dist, dist)
End Sub

Private Sub CmdGo_Click()
Dim depth As Integer
Dim total_length As Single
Dim start_x As Single
Dim start_y As Single
Dim start_length As Single

    picCanvas.Cls
    MousePointer = vbHourglass
    DoEvents

    ' Get the parameters.
    If Not IsNumeric(txtDepth.Text) Then txtDepth.Text = "5"
    depth = CInt(txtDepth.Text)

    ' See how big we can make the curve.
    If picCanvas.ScaleHeight < picCanvas.ScaleWidth Then
        total_length = 0.9 * picCanvas.ScaleHeight
    Else
        total_length = 0.9 * picCanvas.ScaleWidth
    End If

    ' Compute the side length for this depth.
    start_length = total_length / (3 * 2 ^ depth - 1)

    start_x = (picCanvas.ScaleWidth - total_length) / 2
    start_y = (picCanvas.ScaleHeight - total_length) / 2 + start_length

    ' Draw the curve.
    picCanvas.CurrentX = start_x
    picCanvas.CurrentY = start_y
    Sierpinski depth, start_length

    MousePointer = vbDefault
End Sub
Private Sub Form_Resize()
Dim wid As Single

    wid = ScaleWidth - picCanvas.Left
    If wid < 120 Then wid = 120
    picCanvas.Move picCanvas.Left, 0, _
        wid, ScaleHeight
End Sub
