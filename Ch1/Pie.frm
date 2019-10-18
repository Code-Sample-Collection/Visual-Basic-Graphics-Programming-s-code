VERSION 5.00
Begin VB.Form frmPie 
   Caption         =   "Pie"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPieSliceStartAngle 
      Height          =   285
      Left            =   3600
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtPieSliceEndAngle 
      Height          =   285
      Left            =   3600
      TabIndex        =   3
      Text            =   "0.5"
      Top             =   480
      Width           =   1455
   End
   Begin VB.PictureBox picPieSlice 
      Height          =   2415
      Left            =   2640
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   6
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Draw"
      Default         =   -1  'True
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtCircleEndAngle 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "-0.5"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtCircleStartAngle 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Text            =   "-0"
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox picCircle 
      Height          =   2415
      Left            =   120
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   5
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Start Angle"
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "End Angle"
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   11
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PieSlice Subroutine"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   10
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Circle Method"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "End Angle"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Start Angle"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmPie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const PI = 3.14159265

' Draw a pie slice.
Public Sub PieSlice(ByVal obj As Object, ByVal X As Single, ByVal Y As Single, ByVal radius As Single, ByVal start_angle As Single, ByVal end_angle As Single)
    ' Make both angles <= 2 * PI.
    Do While start_angle > 2 * PI
        start_angle = start_angle - 2 * PI
    Loop
    Do While end_angle > 2 * PI
        end_angle = end_angle - 2 * PI
    Loop

    ' Make both angles strictly positive.
    Do While start_angle <= 0
        start_angle = start_angle + 2 * PI
    Loop
    Do While end_angle <= 0
        end_angle = end_angle + 2 * PI
    Loop

    ' Draw the slice
    obj.Circle (X, Y), radius, obj.ForeColor, -start_angle, -end_angle
End Sub
Private Sub cmdDraw_Click()
Dim X As Single
Dim radius As Single
Dim start_angle As Single
Dim end_angle As Single

    ' Clear the PictureBoxes.
    picCircle.Cls
    picPieSlice.Cls

    ' Get the pie slice geometry parameters.
    X = picCircle.ScaleWidth / 2
    radius = picCircle.ScaleWidth * 0.45

    ' Draw using Circle.
    start_angle = CSng(txtCircleStartAngle.Text)
    If start_angle < -2 * PI Then
        start_angle = -2 * PI
        txtCircleStartAngle.Text = Format$(start_angle)
    ElseIf start_angle > 2 * PI Then
        start_angle = 2 * PI
        txtCircleStartAngle.Text = Format$(start_angle)
    End If

    end_angle = CSng(txtCircleEndAngle.Text)
    If end_angle < -2 * PI Then
        end_angle = -2 * PI
        txtCircleEndAngle.Text = Format$(end_angle)
    ElseIf end_angle > 2 * PI Then
        end_angle = 2 * PI
        txtCircleEndAngle.Text = Format$(end_angle)
    End If

    picCircle.Circle (X, X), radius, , _
        start_angle, end_angle

    ' Draw using PieSlice.
    start_angle = CSng(txtPieSliceStartAngle.Text)
    end_angle = CSng(txtPieSliceEndAngle.Text)
    PieSlice picPieSlice, X, X, radius, _
        start_angle, end_angle
End Sub


' Initialize the drawing PictureBox.
Private Sub Form_Load()
    picPieSlice.AutoRedraw = True
    picPieSlice.FillColor = vbWhite
    picPieSlice.FillStyle = vbFSSolid

    picCircle.AutoRedraw = True
    picCircle.FillColor = vbWhite
    picCircle.FillStyle = vbFSSolid
End Sub
