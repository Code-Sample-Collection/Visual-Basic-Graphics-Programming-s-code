VERSION 5.00
Begin VB.Form frmBounce2b 
   Caption         =   "Bounce2b"
   ClientHeight    =   5235
   ClientLeft      =   1320
   ClientTop       =   825
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   349
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   458
   Begin VB.PictureBox picHidden 
      Height          =   495
      Index           =   0
      Left            =   480
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtFramesPerSecond 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "20"
      Top             =   4920
      Width           =   375
   End
   Begin VB.TextBox txtNumBalls 
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
   Begin VB.PictureBox picCourt 
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
      Caption         =   "Number of balls:"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   4560
      Width           =   1455
   End
End
Attribute VB_Name = "frmBounce2b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private xmax As Integer
Private ymax As Integer

Private NumBalls As Integer
Private BallX() As Integer
Private BallY() As Integer
Private BallDx() As Integer
Private BallDy() As Integer
Private BallRadius() As Integer
Private BallColor() As Long

Private Playing As Boolean
Private NumPlayed As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

' Draw some random rectangles on the bacground.
Private Sub DrawBackground()
Dim i As Integer
Dim wid As Single
Dim hgt As Single

    ' Start with a clean slate.
    picCourt.Line (0, 0)-(picCourt.ScaleWidth, picCourt.ScaleHeight), picCourt.BackColor, BF

    ' Draw some rectangles.
    For i = 1 To 10
        hgt = 10 + Rnd * xmax / 3
        wid = 10 + Rnd * ymax / 3
        picCourt.Line (Int(Rnd * xmax), Int(Rnd * ymax))-Step(hgt, wid), QBColor(Int(Rnd * 16)), BF
    Next i

    ' Make the rectangles part of the permanent background.
    picCourt.Picture = picCourt.Image
End Sub


' Generate some random data.
Private Sub InitData()
Dim ball As Integer
Dim R As Integer
Dim clr As Integer

    ' See how many balls there should be.
    If Not IsNumeric(txtNumBalls.Text) Then _
        txtNumBalls.Text = "10"
    NumBalls = CInt(txtNumBalls.Text)
    ReDim BallRadius(1 To NumBalls)
    ReDim BallX(1 To NumBalls)
    ReDim BallY(1 To NumBalls)
    ReDim BallDx(1 To NumBalls)
    ReDim BallDy(1 To NumBalls)
    ReDim BallColor(1 To NumBalls)
    
    ' Set the initial ball data.
    For ball = 1 To NumBalls
        R = Int(10 * Rnd + 5)
        BallRadius(ball) = R
        BallX(ball) = Int((xmax - R + 1) * Rnd)
        BallY(ball) = Int((ymax - R + 1) * Rnd)
        BallDx(ball) = Int(21 * Rnd - 10)
        BallDy(ball) = Int(21 * Rnd - 10)
        clr = Int(15 * Rnd)
        If clr >= 7 Then clr = clr + 1
        BallColor(ball) = QBColor(clr)

        ' Create a hidden PictureBox for this ball.
        If ball > picHidden.UBound Then
            Load picHidden(ball)
        End If

        ' Make the picture big enough.
        picHidden(ball).Width = 2 * BallRadius(ball) + 4
        picHidden(ball).Height = 2 * BallRadius(ball) + 4
    Next ball

    ' Unload any hidden PictureBoxes we no longer need.
    For ball = NumBalls + 1 To picHidden.UBound
        Unload picHidden(ball)
    Next ball
End Sub
' Start the animation.
Private Sub cmdStart_Click()
    If Playing Then
        Playing = False
        cmdStart.Caption = "Stopped"
        cmdStart.Enabled = False
    Else
        cmdStart.Caption = "Stop"
        Playing = True
        InitData
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

    ' Draw a random background.
    DrawBackground

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
' Play the animation.
Private Sub PlayImages(ByVal ms_per_frame As Long)
Dim ball As Integer
Dim next_time As Long

    ' Get the current time.
    next_time = GetTickCount()

    ' Start the animation.
    Do While Playing
        NumPlayed = NumPlayed + 1

        ' Save the background where the balls
        ' will be placed.
        For ball = 1 To NumBalls
            BitBlt picHidden(ball).hDC, _
                0, 0, _
                2 * BallRadius(ball) + 4, _
                2 * BallRadius(ball) + 4, _
                picCourt.hDC, _
                BallX(ball) - BallRadius(ball) - 2, _
                BallY(ball) - BallRadius(ball) - 2, _
                vbSrcCopy

            picHidden(ball).Picture = picHidden(ball).Image
        Next ball

        ' Draw the balls.
        For ball = 1 To NumBalls
            picCourt.FillColor = BallColor(ball)
            picCourt.Circle _
                (BallX(ball), BallY(ball)), _
                BallRadius(ball), BallColor(ball)
        Next ball

        ' Wait until it's time for the next frame.
        next_time = next_time + ms_per_frame
        WaitTill next_time

        ' Restore the background information.
        For ball = 1 To NumBalls
            BitBlt picCourt.hDC, _
                BallX(ball) - BallRadius(ball) - 2, _
                BallY(ball) - BallRadius(ball) - 2, _
                2 * BallRadius(ball) + 4, _
                2 * BallRadius(ball) + 4, _
                picHidden(ball).hDC, _
                0, 0, _
                vbSrcCopy
        Next ball

        ' Move the balls for the next frame,
        ' keeping them within picCourt.
        For ball = 1 To NumBalls
            BallX(ball) = BallX(ball) + BallDx(ball)
            If BallX(ball) < BallRadius(ball) Then
                BallX(ball) = 2 * BallRadius(ball) - BallX(ball)
                BallDx(ball) = -BallDx(ball)
            ElseIf BallX(ball) > xmax - BallRadius(ball) Then
                BallX(ball) = 2 * (xmax - BallRadius(ball)) - BallX(ball)
                BallDx(ball) = -BallDx(ball)
            End If

            BallY(ball) = BallY(ball) + BallDy(ball)
            If BallY(ball) < BallRadius(ball) Then
                BallY(ball) = 2 * BallRadius(ball) - BallY(ball)
                BallDy(ball) = -BallDy(ball)
            ElseIf BallY(ball) > ymax - BallRadius(ball) Then
                BallY(ball) = 2 * (ymax - BallRadius(ball)) - BallY(ball)
                BallDy(ball) = -BallDy(ball)
            End If
        Next ball

        If Not Playing Then Exit Do
    Loop
End Sub
Private Sub Form_Load()
    Randomize

    picCourt.FillStyle = vbSolid
    picCourt.ScaleMode = vbPixels
    With picHidden(0)
        .AutoRedraw = True
        .Visible = False
        .ScaleMode = vbPixels
        .BorderStyle = vbBSNone
    End With
End Sub

' Make the ball court nice and big.
Private Sub Form_Resize()
Const GAP = 3

    txtFramesPerSecond.Top = ScaleHeight - GAP - txtFramesPerSecond.Height
    Label1(0).Top = txtFramesPerSecond.Top
    txtNumBalls.Top = txtFramesPerSecond.Top - GAP - txtNumBalls.Height
    Label1(1).Top = txtNumBalls.Top
    cmdStart.Top = (txtNumBalls.Top + txtFramesPerSecond.Top + txtFramesPerSecond.Height - cmdStart.Height) / 2
    picCourt.Move 0, 0, ScaleWidth, txtNumBalls.Top - GAP

    xmax = picCourt.ScaleWidth - 1
    ymax = picCourt.ScaleHeight - 1

    picCourt.Picture = picCourt.Image
End Sub
