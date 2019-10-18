VERSION 5.00
Begin VB.Form frmBinTree 
   Caption         =   "BinTree"
   ClientHeight    =   4335
   ClientLeft      =   2445
   ClientTop       =   1335
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4335
   ScaleWidth      =   7470
   Begin VB.TextBox txtDepth 
      Height          =   285
      Left            =   720
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "5"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   615
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   3495
      Left            =   1440
      ScaleHeight     =   229
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   293
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Depth"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmBinTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PI = 3.14159

Private Const LENGTH_SCALE = 0.75
Private Const DTHETA = PI / 5
' Recursively draw a binary tree branch.
Private Sub DrawBranch(ByVal depth As Integer, ByVal X As Single, ByVal Y As Single, ByVal length As Single, ByVal theta As Single)
Dim x1 As Integer
Dim y1 As Integer

    ' See where this branch should end.
    x1 = X + length * Cos(theta)
    y1 = Y + length * Sin(theta)
    picCanvas.Line (X, Y)-(x1, y1)

    ' If depth > 1, draw the attached branches.
    If depth > 1 Then
        DrawBranch depth - 1, x1, y1, length * LENGTH_SCALE, theta + DTHETA
        DrawBranch depth - 1, x1, y1, length * LENGTH_SCALE, theta - DTHETA
    End If
End Sub

Private Sub cmdGo_Click()
Dim depth As Integer
Dim start_length As Single

    picCanvas.Cls
    MousePointer = vbHourglass
    DoEvents

    If Not IsNumeric(txtDepth.Text) Then txtDepth.Text = "5"
    depth = CInt(txtDepth.Text)
    start_length = (picCanvas.ScaleHeight - 10) / _
        ((1 - LENGTH_SCALE ^ (depth + 1)) / (1 - LENGTH_SCALE))

    DrawBranch depth, picCanvas.ScaleWidth \ 2, _
        picCanvas.ScaleHeight - 5, _
        start_length, -PI / 2

    MousePointer = vbDefault
End Sub
Private Sub Form_Resize()
Dim wid As Single

    wid = ScaleWidth - picCanvas.Left
    If wid < 120 Then wid = 120
    picCanvas.Move picCanvas.Left, 0, wid, ScaleHeight
End Sub

