VERSION 5.00
Begin VB.Form frmCannon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cannon"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   615
   End
   Begin VB.PictureBox picHouseHit 
      AutoSize        =   -1  'True
      Height          =   330
      Left            =   6720
      Picture         =   "Cannon.frx":0000
      ScaleHeight     =   270
      ScaleWidth      =   285
      TabIndex        =   7
      Top             =   3360
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picHouseOk 
      AutoSize        =   -1  'True
      Height          =   330
      Left            =   6240
      Picture         =   "Cannon.frx":015A
      ScaleHeight     =   270
      ScaleWidth      =   285
      TabIndex        =   6
      Top             =   3360
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txtSpeed 
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Text            =   "100"
      Top             =   3390
      Width           =   495
   End
   Begin VB.PictureBox picCanvas 
      Height          =   3255
      Left            =   0
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   469
      TabIndex        =   4
      Top             =   0
      Width           =   7095
   End
   Begin VB.TextBox txtAngle 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Text            =   "45"
      Top             =   3390
      Width           =   495
   End
   Begin VB.CommandButton cmdFire 
      Caption         =   "Fire"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Speed"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   5
      Top             =   3390
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Angle"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   3
      Top             =   3390
      Width           =   495
   End
End
Attribute VB_Name = "frmCannon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DISTANCE_SCALE = 10
Private Const CANNON_SCALE = 10

Private TargetX As Single

Private BitmapWid As Long
Private BitmapHgt As Long
Private BitmapNumBytes As Long
Private Bytes() As Byte

' ------------------
' Bitmap Information
' ------------------
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

' Get the initial velocity components from the
' speed and angle.
Private Sub GetInitialVelocity(ByRef vx As Single, ByRef vy As Single)
Const PI = 3.14159365

Dim angle As Single
Dim speed As Single

    ' Get the angle in radians and the speed.
    On Error Resume Next
    angle = CSng(txtAngle.Text) * PI / 180
    speed = CSng(txtSpeed.Text)

    vx = Cos(angle) * speed / DISTANCE_SCALE
    vy = -Sin(angle) * speed / DISTANCE_SCALE
End Sub

' Start the animation.
Private Sub PlayImages()
Const MS_PER_FRAME = 50
Const SCALED_F = 16 / DISTANCE_SCALE

Dim X As Single
Dim Y As Single
Dim hitx As Single
Dim hity As Single
Dim dhity As Single
Dim not_hit As Boolean
Dim vx As Single
Dim vy As Single
Dim dist As Single
Dim next_time As Long
Dim test_color As Long

    ' Get the initial velocity and position.
    GetInitialVelocity vx, vy

    ' Start the point at the end of the cannon.
    dist = Sqr(vx * vx + vy * vy)
    X = vx / dist * CANNON_SCALE
    Y = BitmapHgt + vy / dist * CANNON_SCALE

    not_hit = True
    next_time = GetTickCount()
    Do
        ' Subtract the force of gravity from the
        ' Y velocity component.
        vy = vy + SCALED_F

        ' Restore the background.
        SetBitmapBits picCanvas.Image, BitmapNumBytes, Bytes(1, 1)

        ' See if we will hit the house.
        If not_hit Then
            dhity = vy / vx
            hity = Y
            For hitx = X To X + vx
                ' See if (hitx, hity) is a hit.
                test_color = picCanvas.Point(hitx, hity)
                If (test_color > 0) And _
                    (test_color <> picCanvas.BackColor) _
                Then
                    not_hit = False
                    picCanvas.PaintPicture _
                        picHouseHit.Picture, TargetX, _
                        picCanvas.ScaleHeight - picHouseOk.ScaleHeight
                    DoEvents

                    ' Save the new background.
                    SaveBackground
                    Beep
                    Exit For
                End If
                hity = hity + dhity
            Next hitx
        End If

        ' Calculate the next position.
        X = X + vx
        Y = Y + vy

        ' Draw the projectile.
        picCanvas.PSet (X, Y), vbBlue

        ' Wait until it's time for the next frame.
        next_time = next_time + MS_PER_FRAME
        WaitTill next_time
    Loop While Y < BitmapHgt + 3
End Sub

' Start the animation.
Private Sub cmdFire_Click()
    DrawBackground

    PlayImages
End Sub


' Move the target.
Private Sub cmdReset_Click()
    TargetX = picCanvas.ScaleWidth * (0.3 + Rnd * 0.6)
    DrawBackground

    cmdFire.SetFocus
End Sub

Private Sub Form_Load()
    Randomize
    Show

    picCanvas.AutoRedraw = True
    picCanvas.ScaleMode = vbPixels
    picCanvas.DrawWidth = 3
    picCanvas.FillStyle = vbSolid
    picCanvas.BackColor = &HC0C0C0

    picHouseOk.ScaleMode = vbPixels
    picHouseHit.ScaleMode = vbPixels

    cmdReset_Click
End Sub

' Save the background bitmap data.
Private Sub SaveBackground()
Dim bm As BITMAP

    GetObject picCanvas.Image, Len(bm), bm
    BitmapWid = bm.bmWidthBytes
    BitmapHgt = bm.bmHeight
    BitmapNumBytes = BitmapWid * BitmapHgt
    ReDim Bytes(1 To bm.bmWidthBytes, 1 To bm.bmHeight)
    GetBitmapBits picCanvas.Image, BitmapNumBytes, Bytes(1, 1)
End Sub
' Draw the target and the cannon pointed in the
' direction of the current angle.
Private Sub DrawBackground()
Dim vx As Single
Dim vy As Single
Dim dist As Single
Dim bm As BITMAP

    ' Clear the canvas.
    picCanvas.Line (0, 0)-(picCanvas.ScaleWidth, picCanvas.ScaleHeight), picCanvas.BackColor, BF

    ' Get the initial velocity components.
    GetInitialVelocity vx, vy

    ' Draw the target.
    picCanvas.PaintPicture _
        picHouseOk.Picture, TargetX, _
        picCanvas.ScaleHeight - picHouseOk.ScaleHeight

    ' Draw the cannon.
    dist = Sqr(vx * vx + vy * vy)
    vx = vx / dist
    vy = vy / dist
    picCanvas.Line (0, picCanvas.ScaleHeight)-Step(vx * CANNON_SCALE, vy * CANNON_SCALE), vbBlack

    ' Save the background bitmap data.
    SaveBackground
End Sub

Private Sub txtAngle_Change()
    DrawBackground
End Sub
