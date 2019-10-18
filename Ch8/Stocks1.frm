VERSION 5.00
Begin VB.Form frmStocks1 
   Caption         =   "Stocks1"
   ClientHeight    =   3990
   ClientLeft      =   1875
   ClientTop       =   1380
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3990
   ScaleWidth      =   5190
   Begin VB.TextBox txtFramesPerSecond 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "10"
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   3480
      Width           =   855
   End
   Begin VB.PictureBox picGraph 
      Height          =   3375
      Left            =   0
      ScaleHeight     =   -100
      ScaleLeft       =   0.5
      ScaleMode       =   0  'User
      ScaleTop        =   100
      ScaleWidth      =   10.75
      TabIndex        =   0
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Frames per second:"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblDay 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Day:"
      Height          =   255
      Index           =   0
      Left            =   4440
      TabIndex        =   2
      Top             =   3600
      Width           =   375
   End
End
Attribute VB_Name = "frmStocks1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NUM_STOCKS = 10
Private Const NUM_FRAMES = 100

Private Data(1 To NUM_FRAMES, 1 To NUM_STOCKS) As Integer
Private Playing As Boolean
' Generate some random data.
Private Sub InitData()
Dim stock As Integer
Dim frame As Integer

    ' Set the initial values between 30 and 70.
    For stock = 1 To NUM_STOCKS
        Data(1, stock) = Int(41 * Rnd + 30)
    Next stock

    ' Make values for the other frames.
    ' Each value is up to +/- 5 different than the
    ' previous value for the same stock.
    For frame = 2 To NUM_FRAMES
        For stock = 1 To NUM_STOCKS
            Data(frame, stock) = _
                Data(frame - 1, stock) + _
                Int(11 * Rnd - 5)
            If Data(frame, stock) < 0 Then _
                Data(frame, stock) = 5
            If Data(frame, stock) > 100 Then _
                Data(frame, stock) = 95
        Next stock
    Next frame
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
Dim milliseconds_per_frame As Long
Dim frame As Integer
Dim stock As Integer
Dim next_time As Long
Dim old_style As Integer

    ' Set FillStyle to vbSolid.
    old_style = picGraph.FillStyle
    picGraph.FillStyle = vbSolid
    picGraph.AutoRedraw = True
    
    ' See how fast we should go.
    If Not IsNumeric(txtFramesPerSecond.Text) Then _
        txtFramesPerSecond.Text = "10"
    milliseconds_per_frame = 1000 \ CLng(txtFramesPerSecond.Text)
    
    ' Draw the background.
    '
    ' Note that we must cover all of the background
    ' so it becomes part of the image. Then Cls
    ' will restore it.
    picGraph.Line (0, 0)-(NUM_STOCKS + 1.25, 50), RGB(128, 128, 128), BF
    picGraph.Line (0, 50)-(NUM_STOCKS + 1.25, 100), picGraph.BackColor, BF
    
    ' Make the picture a permanent part of the
    ' background.
    picGraph.Picture = picGraph.Image

    ' Start the animation.
    next_time = GetTickCount()
    For frame = 1 To NUM_FRAMES
        If Not Playing Then Exit For

        ' Draw the graph.
        picGraph.Cls
        For stock = 1 To NUM_STOCKS
            If Data(frame, stock) > 50 Then
                picGraph.Line (stock, 0)-(stock + 0.75, 50), vbRed, BF
                picGraph.Line (stock, 50)-(stock + 0.75, Data(frame, stock)), vbGreen, BF
            Else
                picGraph.Line (stock, 0)-(stock + 0.75, Data(frame, stock)), vbRed, BF
            End If
        Next stock
        picGraph.Line (0, 50)-(NUM_STOCKS + 1.25, 50), vbBlack, BF
        lblDay.Caption = Format$(frame)

        ' Wait until it's time for the next frame.
        next_time = next_time + milliseconds_per_frame
        WaitTill next_time
    Next frame

    ' Restore the old FillStyle.
    picGraph.FillStyle = old_style
End Sub
