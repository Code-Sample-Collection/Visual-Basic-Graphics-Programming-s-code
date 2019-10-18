VERSION 5.00
Begin VB.Form frmSierpBox 
   Caption         =   "SierpBox"
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
Attribute VB_Name = "frmSierpBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Erase the center rectangle from this one.
Private Sub SierpinskiErase(ByVal depth As Integer, ByVal x1 As Single, ByVal y1 As Single, ByVal x4 As Single, ByVal y4 As Single)
Dim x2 As Single
Dim y2 As Single
Dim x3 As Single
Dim y3 As Single

    ' Find the corners of the middle square.
    x2 = (2 * x1 + x4) * 0.3333
    x3 = (x1 + 2 * x4) * 0.3333
    y2 = (2 * y1 + y4) * 0.3333
    y3 = (y1 + 2 * y4) * 0.3333

    ' Erase the middle rectangle.
    picCanvas.Line (x2, y2)-(x3, y3), picCanvas.BackColor, BF

    ' Recursively erase other rectangles.
    If depth > 0 Then
        SierpinskiErase depth - 1, x1, y1, x2, y2
        SierpinskiErase depth - 1, x2, y1, x3, y2
        SierpinskiErase depth - 1, x3, y1, x4, y2
        SierpinskiErase depth - 1, x1, y2, x2, y3
        SierpinskiErase depth - 1, x3, y2, x4, y3
        SierpinskiErase depth - 1, x1, y3, x2, y4
        SierpinskiErase depth - 1, x2, y3, x3, y4
        SierpinskiErase depth - 1, x3, y3, x4, y4
    End If
End Sub
' Draw a complete Sierpinski carpet.
Private Sub SierpinskiCarpet(ByVal depth As Integer, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    ' Erase the picture.
    picCanvas.Line (0, 0)-(picCanvas.ScaleWidth, picCanvas.ScaleHeight), picCanvas.BackColor, BF

    ' Draw the main filled box.
    picCanvas.AutoRedraw = True
    picCanvas.Line (x1, y1)-(x2, y2), vbBlack, BF

    ' If depth > 0, call SierpinskiErase to
    ' erase the center of this box.
    If depth >= 0 Then
        SierpinskiErase depth, x1, y1, x2, y2
    End If
End Sub

Private Sub CmdGo_Click()
Dim depth As Integer

    MousePointer = vbHourglass
    DoEvents

    ' Get the parameters.
    If Not IsNumeric(txtDepth.Text) Then txtDepth.Text = "5"
    depth = CInt(txtDepth.Text)

    ' Draw the curve.
    SierpinskiCarpet depth, _
        picCanvas.ScaleWidth * 0.02, _
        picCanvas.ScaleHeight * 0.02, _
        picCanvas.ScaleWidth * 0.98, _
        picCanvas.ScaleHeight * 0.98

    MousePointer = vbDefault
End Sub
Private Sub Form_Resize()
Dim wid As Single

    wid = ScaleWidth - picCanvas.Left
    If wid < 120 Then wid = 120
    picCanvas.Move picCanvas.Left, 0, _
        wid, ScaleHeight
End Sub
