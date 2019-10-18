VERSION 5.00
Begin VB.Form SpriteForm 
   Caption         =   "Sprites"
   ClientHeight    =   5235
   ClientLeft      =   1320
   ClientTop       =   825
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   349
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   458
   Begin VB.TextBox txtFramesPerSecond 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "20"
      Top             =   4920
      Width           =   375
   End
   Begin VB.TextBox txtNumObjects 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "20"
      Top             =   4560
      Width           =   375
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   4620
      Width           =   855
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   4455
      Left            =   0
      ScaleHeight     =   293
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   453
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "Frames per second:"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Number of objects:"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   4560
      Width           =   1455
   End
End
Attribute VB_Name = "SpriteForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private xmin As Integer
Private ymin As Integer
Private xmax As Integer
Private ymax As Integer

Private NumSprites As Integer
Private Sprites() As Sprite

Private Playing As Boolean
Private NumPlayed As Long

Private BitmapWid As Long
Private BitmapHgt As Long
Private BitmapNumBytes As Long
Private Bytes() As Byte

' Bitmap Information
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

' Play the animation.
Private Sub PlayImages(ByVal ms_per_frame As Long)
Dim sprite_number As Integer
Dim next_time As Long

    ' Get the current time.
    next_time = GetTickCount()

    ' Start the animation.
    Do While Playing
        NumPlayed = NumPlayed + 1

        ' Restore the background.
        SetBitmapBits picCanvas.Image, BitmapNumBytes, Bytes(1, 1)

        ' Draw and move the sprites.
        For sprite_number = 1 To NumSprites
            Sprites(sprite_number).DrawSprite picCanvas
            Sprites(sprite_number).MoveSprite xmin, xmax, ymin, ymax
        Next sprite_number

        ' Wait until it's time for the next frame.
        next_time = next_time + ms_per_frame
        WaitTill next_time
    Loop
End Sub

' Draw some random rectangles on the bacground.
Private Sub DrawBackground()
Dim i As Integer
Dim Wid As Single
Dim Hgt As Single

    ' Start with a clean slate.
    picCanvas.Line (0, 0)-(picCanvas.ScaleWidth, picCanvas.ScaleHeight), picCanvas.BackColor, BF

    ' Draw some rectangles.
    For i = 1 To 10
        Hgt = 10 + Rnd * xmax / 3
        Wid = 10 + Rnd * ymax / 3
        picCanvas.Line (Int(Rnd * xmax), Int(Rnd * ymax))-Step(Hgt, Wid), QBColor(Int(Rnd * 16)), BF
    Next i

    ' Make the rectangles part of the permanent background.
    picCanvas.Picture = picCanvas.Image
End Sub

' Generate some random data.
Private Sub InitializeData()
Dim obj As Object
Dim i As Integer

    ' See how many objects there should be.
    If Not IsNumeric(txtNumObjects.Text) Then Exit Sub
    NumSprites = CInt(txtNumObjects.Text)
    If NumSprites < 1 Then Exit Sub

    ' Create the sprites.
    ReDim Sprites(1 To NumSprites)
    For i = 1 To NumSprites
        ' Pick a random sprite type.
        Select Case Int(3 * Rnd)
            Case 0
                Set Sprites(i) = NewRectangle()
            Case 1
                Set Sprites(i) = NewTriangle()
            Case 2
                Set Sprites(i) = NewBall()
        End Select
    Next i
End Sub



' Create and initialize a random BallSprite.
Private Function NewBall() As BallSprite
Dim new_sprite As BallSprite
Dim new_color As Long

    ' Make the new sprite.
    Set new_sprite = New BallSprite

    ' Pick a color other than 7 (gray).
    new_color = Int(15 * Rnd)
    If new_color >= 7 Then new_color = new_color + 1

    ' Initialize the sprite.
    new_sprite.InitializeBall _
        Int(15 * Rnd + 5), _
        Int(xmax * Rnd), _
        Int(ymax * Rnd), _
        Int(11 * Rnd - 5), Int(11 * Rnd - 5), _
        QBColor(new_color)

    Set NewBall = new_sprite
End Function


' Create and initialize a random TriangleSprite.
Private Function NewTriangle() As TriangleSprite
Const PI = 3.14159265
Const THIRD_CIRCLE = 2 * PI / 3
Const PI_OVER_8 = PI / 8
Const PI_OVER_16 = PI / 16

Dim new_sprite As TriangleSprite
Dim new_color As Long

    ' Make the new sprite.
    Set new_sprite = New TriangleSprite

    ' Pick a color other than 7 (gray).
    new_color = Int(15 * Rnd)
    If new_color >= 7 Then new_color = new_color + 1

    ' Initialize the sprite.
    new_sprite.InitializeTriangle _
        Int(xmax * Rnd), Int(ymax * Rnd), _
        Int(11 * Rnd - 5), Int(11 * Rnd - 5), _
        Int(15 * Rnd + 10), THIRD_CIRCLE * Rnd, _
        Int(15 * Rnd + 10), THIRD_CIRCLE * (1 + Rnd), _
        Int(15 * Rnd + 10), THIRD_CIRCLE * (2 + Rnd), _
        0, PI_OVER_8 * Rnd - PI_OVER_16, _
        QBColor(new_color)

    Set NewTriangle = new_sprite
End Function

' Create and initialize a random RectangleSprite.
Private Function NewRectangle() As RectangleSprite
Const PI = 3.14159265
Const PI_OVER_2 = PI / 2
Const PI_OVER_8 = PI / 8
Const PI_OVER_16 = PI / 16

Dim new_sprite As RectangleSprite
Dim new_color As Integer

    ' Make the new sprite.
    Set new_sprite = New RectangleSprite

    ' Pick a color other than 7 (gray).
    new_color = Int(15 * Rnd)
    If new_color >= 7 Then new_color = new_color + 1

    ' Initialize the sprite.
    new_sprite.InitializeRectangle _
        Int(20 * Rnd + 10), _
        Int(20 * Rnd + 10), _
        Int(xmax * Rnd), Int(ymax * Rnd), _
        Int(11 * Rnd - 5), Int(11 * Rnd - 5), _
        PI_OVER_2 * Rnd, _
        PI_OVER_8 * Rnd - PI_OVER_16, _
        QBColor(new_color)

    Set NewRectangle = new_sprite
End Function



' Start the animation.
Private Sub cmdStart_Click()
    If Playing Then
        Playing = False
        cmdStart.Caption = "Stopped"
        cmdStart.Enabled = False
    Else
        cmdStart.Caption = "Stop"
        Playing = True
        InitializeData
        PlayData
        Playing = False
        cmdStart.Caption = "Start"
        cmdStart.Enabled = True
    End If
End Sub

' Play the animation.
Private Sub PlayData()
Dim ms_per_frame As Long
Dim start_time As Single
Dim stop_time As Single
Dim bm As BITMAP

    ' Draw a random background.
    DrawBackground

    ' Save the background bitmap data.
    GetObject picCanvas.Image, Len(bm), bm
    BitmapWid = bm.bmWidthBytes
    BitmapHgt = bm.bmHeight
    BitmapNumBytes = BitmapWid * BitmapHgt
    ReDim Bytes(1 To bm.bmWidthBytes, 1 To bm.bmHeight)
    GetBitmapBits picCanvas.Image, BitmapNumBytes, Bytes(1, 1)

    ' See how fast we should go.
    If Not IsNumeric(txtFramesPerSecond.Text) Then _
        txtFramesPerSecond.Text = "10"
    ms_per_frame = 1000 \ CLng(txtFramesPerSecond.Text)

    ' Start the animation.
    NumPlayed = 0
    start_time = Timer
    PlayImages ms_per_frame

    ' Display results.
    stop_time = Timer
    MsgBox "Displayed" & Str$(NumPlayed) & _
        " frames in " & _
        Format$(stop_time - start_time, "0.00") & _
        " seconds (" & _
        Format$(NumPlayed / (stop_time - start_time), "0.00") & _
        " FPS)."
End Sub

Private Sub Form_Load()
    picCanvas.FillStyle = vbFSSolid
End Sub
' Make the ball picCanvas nice and big.
Private Sub Form_Resize()
Const GAP = 3

    txtFramesPerSecond.Top = ScaleHeight - GAP - txtFramesPerSecond.Height
    Label1(0).Top = txtFramesPerSecond.Top
    txtNumObjects.Top = txtFramesPerSecond.Top - GAP - txtNumObjects.Height
    Label1(1).Top = txtNumObjects.Top
    cmdStart.Top = (txtNumObjects.Top + txtFramesPerSecond.Top + txtFramesPerSecond.Height - cmdStart.Height) / 2
    picCanvas.Move 0, 0, ScaleWidth, txtNumObjects.Top - GAP

    xmax = picCanvas.ScaleWidth - 1
    ymax = picCanvas.ScaleHeight - 1
End Sub
