VERSION 5.00
Begin VB.Form frmBditch 
   Caption         =   "Bditch"
   ClientHeight    =   5310
   ClientLeft      =   2175
   ClientTop       =   645
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5310
   ScaleWidth      =   4830
   Begin VB.TextBox txtDt 
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Text            =   "0.1"
      Top             =   45
      Width           =   615
   End
   Begin VB.TextBox txtTmin 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Text            =   "0"
      Top             =   45
      Width           =   615
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox txtTmax 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "6.2832"
      Top             =   45
      Width           =   615
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   4815
      Left            =   0
      ScaleHeight     =   4755
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "dt"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   5
      Top             =   60
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "<= t <="
      Height          =   255
      Index           =   0
      Left            =   645
      TabIndex        =   1
      Top             =   60
      Width           =   495
   End
End
Attribute VB_Name = "frmBditch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Draw the curve on the indicated picture box.
Private Sub DrawCurve(ByVal pic As PictureBox, ByVal start_t As Single, ByVal stop_t As Single, ByVal dt As Single)
Dim cx As Single
Dim cy As Single
Dim t As Single

    cx = pic.ScaleLeft + pic.ScaleWidth / 2
    cy = pic.ScaleTop + pic.ScaleHeight / 2

    pic.Cls
    pic.CurrentX = cx + X(start_t)
    pic.CurrentY = cy + Y(start_t)

    t = start_t + dt
    Do While t < stop_t
        pic.Line -(cx + X(t), cy + Y(t))
        t = t + dt
    Loop

    pic.Line -(cx + X(stop_t), cy + Y(stop_t))
End Sub


' The parametric function Y(t).
Private Function Y(ByVal t As Single) As Single
    Y = 2000 * Sin(5 * t)
End Function

' The parametric function X(t).
Private Function X(ByVal t As Single) As Single
    X = 2000 * Sin(4 * t)
End Function

Private Sub cmdGo_Click()
Dim tmin As Single
Dim tmax As Single
Dim dt As Single

    tmin = CSng(txtTmin.Text)
    tmax = CSng(txtTmax.Text)
    dt = CSng(txtDt.Text)

    DrawCurve picCanvas, tmin, tmax, dt
End Sub

Private Sub Form_Load()
Const PI = 3.14159265

    txtTmin.Text = Format$(0, "0.00")
    txtTmax.Text = Format$(2 * PI, "0.00")
    txtDt.Text = "0.01"
End Sub

Private Sub Form_Resize()
Dim lft As Single
Dim hgt As Single

    lft = txtDt.Left + txtDt.Width
    If lft < ScaleWidth - cmdGo.Width Then lft = ScaleWidth - cmdGo.Width
    cmdGo.Left = lft

    hgt = ScaleHeight - picCanvas.Top
    If hgt < 120 Then hgt = 120
    picCanvas.Move 0, picCanvas.Top, ScaleWidth, hgt
End Sub


