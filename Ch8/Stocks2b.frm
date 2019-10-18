VERSION 5.00
Begin VB.Form frmStocks2b 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stocks2b"
   ClientHeight    =   4140
   ClientLeft      =   1305
   ClientTop       =   810
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   276
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   458
   Begin VB.TextBox txtFramesPerSecond 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "20"
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox txtNumStocks 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "5"
      Top             =   3480
      Width           =   375
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   3540
      Width           =   855
   End
   Begin VB.PictureBox picCourt 
      AutoRedraw      =   -1  'True
      Height          =   3375
      Left            =   0
      ScaleHeight     =   221
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
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Number of stocks:"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   3480
      Width           =   1455
   End
End
Attribute VB_Name = "frmStocks2b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NumStocks As Integer
Private StockValue() As Integer
Private StockTrend() As Integer
Private CourtWid As Single
Private CourtHgt As Single
Private BigValue As Single
Private SmallValue As Single

Private Playing As Boolean
Private NumPlayed As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

' Generate some random data.
Private Sub InitData()
Dim stock As Integer

    ' See how many stocks there should be.
    If Not IsNumeric(txtNumStocks.Text) Then _
        txtNumStocks.Text = "10"
    NumStocks = CInt(txtNumStocks.Text)
    ReDim StockValue(1 To NumStocks)
    ReDim StockTrend(1 To NumStocks)

    ' Set the initial stock data.
    For stock = 1 To NumStocks
        StockValue(stock) = Int(CourtHgt * 0.3 + Rnd * CourtHgt * 0.4)
        StockTrend(stock) = Int(Rnd * 6 - 3)
    Next stock
End Sub


' Return a new stock value for this stock.
Private Function NewStockValue(ByVal stock_number As Integer) As Integer
Dim new_value As Integer

    ' Set the new value.
    new_value = StockValue(stock_number) + StockTrend(stock_number)

    ' Update the trend value.
    If new_value > BigValue Then
        StockTrend(stock_number) = StockTrend(stock_number) + Int(Rnd * 5 - 3)
    ElseIf new_value < SmallValue Then
        StockTrend(stock_number) = StockTrend(stock_number) + Int(Rnd * 5 - 1)
    Else
        StockTrend(stock_number) = StockTrend(stock_number) + Int(Rnd * 5 - 2)
    End If

    ' Keep the trend under control.
    If StockTrend(stock_number) > 10 Then StockTrend(stock_number) = 10
    If StockTrend(stock_number) < -10 Then StockTrend(stock_number) = -10

    NewStockValue = new_value
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

    ' See how fast we should go.
    If Not IsNumeric(txtFramesPerSecond.Text) Then _
        txtFramesPerSecond.Text = "10"
    ms_per_frame = 1000 \ CLng(txtFramesPerSecond.Text)

    ' Clear the drawing area.
    picCourt.Line (0, 0)-(CourtWid, CourtHgt), picCourt.BackColor, BF
    picCourt.Picture = picCourt.Image

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
Dim stock As Integer
Dim next_time As Long
Dim new_value As Integer

    ' Get the current time.
    next_time = GetTickCount()

    ' Start the animation.
    Do While Playing
        NumPlayed = NumPlayed + 1

        ' Move the background 5 pixels left.
        BitBlt picCourt.hDC, _
            0, 0, CourtWid - 5, CourtHgt, _
            picCourt.hDC, 5, 0, vbSrcCopy

        ' Clear the area for the new data.
        picCourt.Line (CourtWid - 5, 0)-Step(5, CourtHgt), picCourt.BackColor, BF

        ' Draw the new stock data.
        For stock = 1 To NumStocks
            ' Get the stock's new value.
            new_value = NewStockValue(stock)

            ' Draw the new segment.
            picCourt.Line (CourtWid - 5, StockValue(stock))-(CourtWid, new_value), QBColor(stock Mod 15)

            ' Update the saved data.
            StockValue(stock) = new_value
        Next stock
        picCourt.Picture = picCourt.Image

        ' Wait until it's time for the next frame.
        next_time = next_time + ms_per_frame
        WaitTill next_time

        If Not Playing Then Exit Do
    Loop
End Sub
Private Sub Form_Load()
    Randomize

    ' Get the drawing area size.
    CourtWid = picCourt.ScaleWidth
    CourtHgt = picCourt.ScaleHeight
    BigValue = CourtHgt * 0.7
    SmallValue = CourtHgt * 0.3

    ' Make a permanent background image.
    picCourt.Picture = picCourt.Image
End Sub
