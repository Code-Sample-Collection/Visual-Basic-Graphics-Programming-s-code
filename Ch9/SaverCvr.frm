VERSION 5.00
Begin VB.Form frmCover 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2910
   ClientLeft      =   2010
   ClientTop       =   2430
   ClientWidth     =   3480
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   194
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   232
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrMoveBalls 
      Interval        =   50
      Left            =   1800
      Top             =   960
   End
End
Attribute VB_Name = "frmCover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Click()
    If RunMode = rmScreenSaver Then Unload Me
End Sub

Private Sub Form_DblClick()
    If RunMode = rmScreenSaver Then Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If RunMode = rmScreenSaver Then Unload Me
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If RunMode = rmScreenSaver Then Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If RunMode = rmScreenSaver Then Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static x0 As Integer
Static y0 As Integer

    ' Do nothing except in screen saver mode.
    If RunMode <> rmScreenSaver Then Exit Sub

    ' Unload on large mouse movements.
    If ((x0 = 0) And (y0 = 0)) Or _
        ((Abs(x0 - X) < 5) And (Abs(y0 - Y) < 5)) _
        Then
            ' It's a small movement.
            x0 = X
            y0 = Y
            Exit Sub
    End If

    Unload Me
End Sub


Private Sub Form_Resize()
    ' Load configuration information.
    LoadConfig

    ' Initialize the balls.
    InitializeBalls
End Sub

' Redisplay the cursor if we hid it in Sub Main.
Private Sub Form_Unload(Cancel As Integer)
    If RunMode = rmScreenSaver Then ShowCursor True
End Sub

' Move the balls.
Private Sub tmrMoveBalls_Timer()
Dim i As Integer
Dim wid As Single
Dim hgt As Single

    ' Erase the balls.
    For i = 1 To NumBalls
        With Balls(i)
            FillColor = BackColor
            Circle (.BallX, .BallY), .BallR, BackColor
        End With
    Next i

    ' Move and redraw the balls.
    wid = ScaleWidth
    hgt = ScaleHeight
    For i = 1 To NumBalls
        With Balls(i)
            .BallX = .BallX + .BallVx
            If .BallX < .BallR Then
                .BallX = 2 * .BallR - .BallX
                .BallVx = -.BallVx
            ElseIf .BallX > wid - .BallR Then
                .BallX = 2 * (wid - .BallR) - .BallX
                .BallVx = -.BallVx
            End If

            .BallY = .BallY + .BallVy
            If .BallY < .BallR Then
                .BallY = 2 * .BallR - .BallY
                .BallVy = -.BallVy
            ElseIf .BallY > hgt - .BallR Then
                .BallY = 2 * (hgt - .BallR) - .BallY
                .BallVy = -.BallVy
            End If

            FillColor = .BallClr
            Circle (.BallX, .BallY), .BallR, .BallClr
        End With
    Next i
End Sub

