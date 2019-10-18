VERSION 5.00
Begin VB.Form frmRndTree 
   Caption         =   "RndTree"
   ClientHeight    =   4050
   ClientLeft      =   1140
   ClientTop       =   1050
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4050
   ScaleWidth      =   7470
   Begin VB.TextBox txtRndDTheta 
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   14
      Text            =   "10"
      Top             =   1800
      Width           =   615
   End
   Begin VB.CheckBox chkBend 
      Caption         =   "Bend Branches"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtMaxBranches 
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   11
      Text            =   "3"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtRndScale 
      Height          =   285
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   9
      Text            =   "0.20"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox txtDTheta 
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "36"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox txtLengthScale 
      Height          =   285
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "0.75"
      Top             =   720
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
   Begin VB.CheckBox chkTaper 
      Caption         =   "Taper Branches"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
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
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Rnd DTheta"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   15
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Max Branches"
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   12
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Rnd Scale"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "DTHETA"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "LENGTH_SCALE"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Depth 
      Caption         =   "Level"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmRndTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PI = 3.14159
' Recursively draw a tree branch.
Private Sub DrawBranch(ByVal bend As Single, ByVal thickness As Single, ByVal Depth As Integer, ByVal X As Single, ByVal Y As Single, ByVal length As Single, ByVal length_scale As Single, ByVal rnd_scale As Single, ByVal theta As Single, ByVal dtheta As Single, ByVal rnd_dtheta As Single, ByVal max_branches As Integer)
Const DIST_PER_BEND = 5#
Const BEND_FACTOR = 2#
Const MAX_BEND = PI / 6

Dim x1 As Integer
Dim y1 As Integer
Dim x2 As Integer
Dim y2 As Integer
Dim status As Integer
Dim num_bends As Integer
Dim num_branches As Integer
Dim i As Integer
Dim new_length As Integer
Dim new_theta As Single
Dim new_bend As Single
Dim dt As Single
Dim t As Single

    If thickness > 0 Then picCanvas.DrawWidth = thickness

    ' Draw the branch.
    If bend > 0 Then
        ' This is a bending branch.
        num_bends = length / DIST_PER_BEND
        t = theta
        x1 = X
        y1 = Y
        For i = 1 To num_bends
            x2 = x1 + DIST_PER_BEND * Cos(t)
            y2 = y1 + DIST_PER_BEND * Sin(t)
            picCanvas.Line (x1, y1)-(x2, y2)
        
            t = t + bend * (Rnd - 0.5)
            x1 = x2
            y1 = y2
        Next i
    Else
        ' This is a straight branch.
        x1 = X + length * Cos(theta)
        y1 = Y + length * Sin(theta)
        picCanvas.Line (X, Y)-(x1, y1)
    End If

    ' If depth > 1, draw the attached branches.
    If Depth > 1 Then
        num_branches = Int((max_branches - 1) * Rnd + 2)
        dt = 2 * dtheta / (num_branches - 1)
        t = theta - dtheta
        For i = 1 To num_branches
            new_length = length * (length_scale + rnd_scale * (Rnd - 0.5))
            new_theta = t + rnd_dtheta * (Rnd - 0.5)
            t = t + dt
            If bend > 0 Then
                new_bend = bend * BEND_FACTOR
                If new_bend > MAX_BEND Then new_bend = MAX_BEND
            Else
                new_bend = bend
            End If
            DrawBranch new_bend, thickness - 1, _
                Depth - 1, x1, y1, new_length, _
                length_scale, rnd_scale, new_theta, _
                dtheta, rnd_dtheta, max_branches
        Next i
    End If
End Sub
Private Sub CmdGo_Click()
Dim thickness As Integer
Dim bend As Single
Dim Depth As Integer
Dim length As Single
Dim length_scale As Single
Dim rnd_scale As Single
Dim dtheta As Single
Dim rnd_dtheta As Single
Dim max_branches As Integer

    picCanvas.Cls
    MousePointer = vbHourglass
    DoEvents

    ' Get the tree parameters.
    If Not IsNumeric(txtDepth.Text) Then txtDepth.Text = "5"
    Depth = CInt(txtDepth.Text)

    If Not IsNumeric(txtLengthScale.Text) Then txtLengthScale.Text = "0.75"
    length_scale = CSng(txtLengthScale.Text)

    If Not IsNumeric(txtDTheta.Text) Then txtDTheta.Text = "36"
    dtheta = CSng(txtDTheta.Text) * PI / 180#

    If Not IsNumeric(txtRndScale.Text) Then txtRndScale.Text = "0.2"
    rnd_scale = CSng(txtRndScale.Text)

    If Not IsNumeric(txtRndDTheta.Text) Then txtRndDTheta.Text = "20"
    rnd_dtheta = CSng(txtRndDTheta.Text) * PI / 180#

    If Not IsNumeric(txtMaxBranches.Text) Then txtMaxBranches.Text = "3"
    max_branches = CInt(txtMaxBranches.Text)

    If chkTaper.Value = vbChecked Then
        thickness = Depth
    Else
        thickness = 0
    End If

    If chkBend.Value = vbChecked Then
        bend = PI / 90
    Else
        bend = 0
    End If

    length = (picCanvas.ScaleHeight - 10) / _
        ((1 - length_scale ^ (Depth + 1)) / (1 - length_scale))

    ' Draw the tree.
    DrawBranch bend, thickness, Depth, _
        picCanvas.ScaleWidth \ 2, _
        picCanvas.ScaleHeight - 5, _
        length, length_scale, rnd_scale, _
        -PI / 2, dtheta, rnd_dtheta, max_branches

    MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Randomize
End Sub

Private Sub Form_Resize()
Dim wid As Single

    wid = ScaleWidth - picCanvas.Left
    If wid < 120 Then wid = 120
    picCanvas.Move picCanvas.Left, 0, wid, ScaleHeight
End Sub

