VERSION 5.00
Begin VB.Form frmCircle3 
   Caption         =   "Circle3"
   ClientHeight    =   5310
   ClientLeft      =   2175
   ClientTop       =   645
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5310
   ScaleWidth      =   4830
   Begin VB.TextBox txtTheta 
      Height          =   285
      Left            =   3720
      TabIndex        =   11
      Text            =   "0.0"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtYScale 
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Text            =   "1.0"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtXScale 
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Text            =   "2.0"
      Top             =   480
      Width           =   615
   End
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
      Top             =   840
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Theta"
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   12
      Top             =   495
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Y Scale"
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   10
      Top             =   495
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "X Scale"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   615
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
Attribute VB_Name = "frmCircle3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PI = 3.14159265

Private Radius As Single
' Draw the curve on the indicated picture box,
' scaled and rotated.
Private Sub DrawCurve(ByVal pic As PictureBox, ByVal start_t As Single, ByVal stop_t As Single, ByVal dt As Single, ByVal x_scale As Single, ByVal y_scale As Single, ByVal theta As Single)
Dim cx As Single
Dim cy As Single
Dim sin_theta As Single
Dim cos_theta As Single
Dim old_x As Single
Dim old_y As Single
Dim new_x As Single
Dim new_y As Single
Dim t As Single

    cx = pic.ScaleLeft + pic.ScaleWidth / 2
    cy = pic.ScaleTop + pic.ScaleHeight / 2

    ' Get Cos(theta) and Sin(theta)
    sin_theta = Sin(theta)
    cos_theta = Cos(theta)

    pic.Cls
    old_x = x_scale * X(start_t)
    old_y = y_scale * Y(start_t)
    pic.CurrentX = cx + old_x * cos_theta - old_y * sin_theta
    pic.CurrentY = cy + old_x * sin_theta + old_y * cos_theta

    t = start_t + dt
    Do While t < stop_t
        old_x = x_scale * X(t)
        old_y = y_scale * Y(t)
        pic.Line -( _
            cx + old_x * cos_theta - old_y * sin_theta, _
            cy + old_x * sin_theta + old_y * cos_theta)
        t = t + dt
    Loop

    old_x = x_scale * X(stop_t)
    old_y = y_scale * Y(stop_t)
    pic.Line -( _
        cx + old_x * cos_theta - old_y * sin_theta, _
        cy + old_x * sin_theta + old_y * cos_theta)
End Sub
' The parametric function Y(t).
Private Function Y(ByVal t As Single) As Single
    Y = Radius * Sin(t)
End Function

' The parametric function X(t).
Private Function X(ByVal t As Single) As Single
    X = Radius * Cos(t)
End Function

Private Sub cmdGo_Click()
Dim tmin As Single
Dim tmax As Single
Dim dt As Single
Dim x_scale As Single
Dim y_scale As Single
Dim max_scale As Single
Dim theta As Single

    tmin = CSng(txtTmin.Text)
    tmax = CSng(txtTmax.Text)
    dt = CSng(txtDt.Text)
    x_scale = CSng(txtXScale.Text)
    y_scale = CSng(txtYScale.Text)
    If x_scale > y_scale Then
        max_scale = x_scale
    Else
        max_scale = y_scale
    End If

    If picCanvas.ScaleWidth / max_scale < picCanvas.ScaleHeight / max_scale Then
        Radius = picCanvas.ScaleWidth * 0.45 / max_scale
    Else
        Radius = picCanvas.ScaleHeight * 0.45 / max_scale
    End If
    theta = -CSng(txtTheta.Text) * 180 / PI

    DrawCurve picCanvas, tmin, tmax, dt, x_scale, y_scale, theta
End Sub

Private Sub Form_Load()
    txtTmin.Text = Format$(0, "0.00")
    txtTmax.Text = Format$(2 * PI, "0.00")
    txtDt.Text = "0.1"
    txtTheta.Text = "30"
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


