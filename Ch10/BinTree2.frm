VERSION 5.00
Begin VB.Form frmBinTree2 
   Caption         =   "BinTree2"
   ClientHeight    =   4335
   ClientLeft      =   1095
   ClientTop       =   990
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4335
   ScaleWidth      =   7470
   Begin VB.CheckBox chkTaper 
      Caption         =   "Taper Branches"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtDTheta 
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "36"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtLengthScale 
      Height          =   285
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "0.75"
      Top             =   360
      Width           =   615
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   4335
      Left            =   2040
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   357
      TabIndex        =   6
      Top             =   0
      Width           =   5415
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtDepth 
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "5"
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "DTHETA"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "LENGTH_SCALE"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Depth"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmBinTree2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PI = 3.14159
' Recursively draw a binary tree branch.
Private Sub DrawBranch(ByVal thickness As Single, ByVal depth As Integer, ByVal X As Single, ByVal Y As Single, ByVal length As Single, ByVal length_scale As Single, ByVal theta As Single, ByVal dtheta As Single)
Dim x1 As Single
Dim y1 As Single
Dim status As Integer

    ' See where this branch should end.
    x1 = X + length * Cos(theta)
    y1 = Y + length * Sin(theta)
    If thickness > 0 Then picCanvas.DrawWidth = thickness
    picCanvas.Line (X, Y)-(x1, y1)
    
    ' If depth > 1, draw the attached branches.
    If depth > 1 Then
        DrawBranch thickness - 1, depth - 1, _
            x1, y1, length * length_scale, _
            length_scale, theta + dtheta, dtheta
        DrawBranch thickness - 1, depth - 1, _
            x1, y1, length * length_scale, _
            length_scale, theta - dtheta, dtheta
    End If
End Sub


Private Sub cmdGo_Click()
Dim taper As Integer
Dim depth As Integer
Dim dtheta As Single
Dim length As Single
Dim length_scale As Single

    picCanvas.Cls
    MousePointer = vbHourglass
    DoEvents

    If Not IsNumeric(txtDepth.Text) Then txtDepth.Text = "5"
    depth = CInt(txtDepth.Text)

    If Not IsNumeric(txtLengthScale.Text) Then txtLengthScale.Text = "0.75"
    length_scale = CSng(txtLengthScale.Text)

    If Not IsNumeric(txtDTheta.Text) Then txtDTheta.Text = "36"
    dtheta = CSng(txtDTheta.Text) * PI / 180#

    If chkTaper.Value = vbChecked Then
        taper = depth
    Else
        taper = 0
    End If

    length = (picCanvas.ScaleHeight - 10) / _
        ((1 - length_scale ^ (depth + 1)) / (1 - length_scale))
    DrawBranch taper, depth, _
        picCanvas.ScaleWidth \ 2, _
        picCanvas.ScaleHeight - 5, length, _
        length_scale, -PI / 2, dtheta

    MousePointer = vbDefault
End Sub
Private Sub Form_Resize()
Dim wid As Single

    wid = ScaleWidth - picCanvas.Left
    If wid < 120 Then wid = 120
    picCanvas.Move picCanvas.Left, 0, wid, ScaleHeight
End Sub
