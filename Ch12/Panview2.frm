VERSION 5.00
Begin VB.Form frmPanview2 
   Caption         =   "Panview2"
   ClientHeight    =   3165
   ClientLeft      =   2550
   ClientTop       =   1800
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3165
   ScaleWidth      =   3165
   Begin VB.HScrollBar HScrollBar 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2880
      Width           =   2895
   End
   Begin VB.VScrollBar VScrollBar 
      Height          =   2895
      Left            =   2880
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picViewport 
      Height          =   2880
      Left            =   0
      ScaleHeight     =   2820
      ScaleWidth      =   2820
      TabIndex        =   0
      Top             =   0
      Width           =   2880
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuScale 
      Caption         =   "&Scale"
      Begin VB.Menu mnuScaleZoom 
         Caption         =   "&Zoom"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuScaleMag 
         Caption         =   "Full  Scale"
         Index           =   1
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuScaleMag 
         Caption         =   "Magnify 1/2"
         Index           =   20
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuScaleMag 
         Caption         =   "Magnify 1/4"
         Index           =   40
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmPanview2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Global max and min world coordinates
' (including margins).
Private Const DataXmin = 0
Private Const DataXmax = 10
Private Const DataYmin = 0
Private Const DataYmax = 10

' Set the min and max allowed width and height.
Private Const DataMinWid = 1
Private Const DataMinHgt = 1
Private Const DataMaxWid = DataXmax - DataXmin
Private Const DataMaxHgt = DataYmax - DataYmin

' The aspect ratio of the viewport.
Private VAspect As Single

' Current world window bounds.
Private Wxmin As Single
Private Wxmax As Single
Private Wymin As Single
Private Wymax As Single

' Prevent change events when we are adjusting the
' scroll bars.
Private IgnoreSbarChange As Boolean

' Variables used for zooming.
Private Enum DrawingModes
    mode_None
    mode_StartZoom
    mode_Zooming
End Enum
Private DrawingMode As DrawingModes

Private StartX As Single
Private StartY As Single
Private LastX As Single
Private LastY As Single
Private OldMode As Integer
' Draw a smiley face in the viewport centered
' around the point (5, 5).
Private Sub DrawSmiley(ByVal pic As PictureBox)
Const PI = 3.14159265

Dim i As Single

    ' Head.
    pic.FillColor = vbYellow
    pic.FillStyle = vbSolid
    pic.Circle (5, 5), 4

    ' Nose.
    pic.FillColor = RGB(0, &HFF, &H80)
    pic.Circle (5, 4.5), 1, , , , 1.5

    ' Eye whites.
    pic.FillColor = vbWhite
    pic.Circle (3.5, 6), 0.75, , , , 1.25
    pic.Circle (6.5, 6), 0.75, , , , 1.25

    ' Pupils.
    pic.FillColor = vbBlack
    pic.Circle (3.7, 6), 0.5, , , , 1.25
    pic.Circle (6.7, 6), 0.5, , , , 1.25

    ' Smile.
    pic.Circle (5, 5), 2.75, , 1.15 * PI, 1.8 * PI

    ' Draw some grid lines to make small scales
    ' easier to understand.
    i = DataXmin + 0.5
    Do While i < DataXmax
        picViewport.Line (i, DataYmin)-(i, DataYmax)
        i = i + 0.5
    Loop
    i = DataYmin + 0.5
    Do While i < DataYmax
        picViewport.Line (DataXmin, i)-(DataXmax, i)
        i = i + 0.5
    Loop
End Sub
' End a zoom operation early. This happens if the
' user starts a zoom and the selects another menu
' item instead of doing the zoom.
Private Sub StopZoom()
    If DrawingMode <> mode_StartZoom Then Exit Sub
    DrawingMode = mode_None

    picViewport.DrawMode = OldMode
    picViewport.MousePointer = vbDefault
End Sub
' Change the level of magnification.
Private Sub SetScaleFactor(fact As Single)
Dim wid As Single
Dim hgt As Single
Dim mid As Single

    fact = 1 / fact

    ' Compute the new world window size.
    wid = fact * (Wxmax - Wxmin)
    hgt = fact * (Wymax - Wymin)

    ' Center the new world window over the old.
    mid = (Wxmax + Wxmin) / 2
    Wxmin = mid - wid / 2
    Wxmax = mid + wid / 2

    mid = (Wymax + Wymin) / 2
    Wymin = mid - hgt / 2
    Wymax = mid + hgt / 2

    ' Set the new world window bounds.
    SetWorldWindow
End Sub
' Adjust the world window so it is not too big,
' too small, off to one side, or of the wrong
' aspect ratio. Then map the world window to the
' viewport and force the viewport to repaint.
Private Sub SetWorldWindow()
Dim wid As Single
Dim hgt As Single
Dim xmid As Single
Dim ymid As Single
Dim aspect As Single

    ' Find the size and center of the world window.
    wid = Wxmax - Wxmin
    hgt = Wymax - Wymin
    xmid = (Wxmax + Wxmin) / 2
    ymid = (Wymax + Wymin) / 2

    ' Make sure we're not too big or too small.
    If wid > DataMaxWid Then
        wid = DataMaxWid
    ElseIf wid < DataMinWid Then
        wid = DataMinWid
    End If
    If hgt > DataMaxHgt Then
        hgt = DataMaxHgt
    ElseIf hgt < DataMinHgt Then
        hgt = DataMinHgt
    End If

    ' Make the aspect ratio match the viewport
    ' aspect ratio, VAspect (set in Form_Resize).
    aspect = hgt / wid
    If aspect > VAspect Then
        ' Too tall and thin. Make it wider.
        wid = hgt / VAspect
    Else
        ' Too short and wide. Make it taller.
        hgt = wid * VAspect
    End If

    ' Compute the new coordinates
    Wxmin = xmid - wid / 2
    Wxmax = xmid + wid / 2
    Wymin = ymid - hgt / 2
    Wymax = ymid + hgt / 2

    ' See if we're off to one side.
    If wid > DataMaxWid Then
        ' We're wider than the picture. Center.
        xmid = (DataXmax + DataXmin) / 2
        Wxmin = xmid - wid / 2
        Wxmax = xmid + wid / 2
    Else
        ' Else see if we're too far to one side.
        If Wxmin < DataXmin And Wxmax < DataXmax Then
            ' Adjust to the right.
            Wxmax = Wxmax + DataXmin - Wxmin
            Wxmin = DataXmin
        End If
        If Wxmax > DataXmax And Wxmin > DataXmin Then
            ' Adjust to the left.
            Wxmin = Wxmin + DataXmax - Wxmax
            Wxmax = DataXmax
        End If
    End If
    If hgt > DataMaxHgt Then
        ' We're taller than the picture. Shrink.
        ymid = (DataYmax + DataYmin) / 2
        Wymin = ymid - hgt / 2
        Wymax = ymid + hgt / 2
    Else
        ' See if we're too far to top or bottom.
        If Wymin < DataYmin And Wymax < DataYmax Then
            ' Adjust downward.
            Wymax = Wymax + DataYmin - Wymin
            Wymin = DataYmin
        End If
        If Wymax > DataYmax And Wymin > DataYmin Then
            ' Adjust upward.
            Wymin = Wymin + DataYmax - Wymax
            Wymax = DataYmax
        End If
    End If

    ' Map the world window to the viewport.
    picViewport.ScaleLeft = Wxmin
    picViewport.ScaleTop = Wymax
    picViewport.ScaleWidth = Wxmax - Wxmin
    picViewport.ScaleHeight = Wymin - Wymax

    ' Force the viewport to repaint.
    picViewport.Refresh

    ' Reset the scroll bars.
    IgnoreSbarChange = True
    HScrollBar.Visible = (wid < DataXmax - DataXmin)
    VScrollBar.Visible = (hgt < DataYmax - DataYmin)

    ' The values of the scroll bars will be where
    ' the top/left of the world window should be.
    VScrollBar.Min = 100 * (DataYmax)
    VScrollBar.Max = 100 * (DataYmin + hgt)
    HScrollBar.Min = 100 * (DataXmin)
    HScrollBar.Max = 100 * (DataXmax - wid)

    ' SmallChange moves the world window 1/10
    ' of its width/height.
    VScrollBar.SmallChange = 100 * (hgt / 10)
    VScrollBar.LargeChange = 100 * hgt
    HScrollBar.SmallChange = 100 * (wid / 10)
    HScrollBar.LargeChange = 100 * wid

    ' Set the current scroll bar values.
    VScrollBar.Value = 100 * Wymax
    HScrollBar.Value = 100 * Wxmin

    IgnoreSbarChange = False
End Sub

' Return to the default magnification scale.
Private Sub SetScaleFull()
    ' Reset the world window coordinates.
    Wxmin = DataXmin
    Wxmax = DataXmax
    Wymin = DataYmin
    Wymax = DataYmax

    ' Set the new world window bounds.
    SetWorldWindow
End Sub

Private Sub Form_Load()
    ' Start at full scale.
    Wxmin = DataXmin
    Wxmax = DataXmax
    Wymin = DataYmin
    Wymax = DataYmax
End Sub

Private Sub Form_Resize()
Dim X As Single
Dim Y As Single
Dim wid As Single
Dim hgt As Single

    ' Fit the viewport to the window.
    X = picViewport.Left
    Y = picViewport.Top
    wid = ScaleWidth - 2 * X - VScrollBar.Width
    hgt = ScaleHeight - 2 * Y - HScrollBar.Height
    picViewport.Move X, Y, wid, hgt
    VAspect = hgt / wid

    ' Place the scroll bars next to the viewport.
    X = picViewport.Left + picViewport.Width + 10
    Y = picViewport.Top
    wid = VScrollBar.Width
    hgt = picViewport.Height
    VScrollBar.Move X, Y, wid, hgt

    X = picViewport.Left
    Y = picViewport.Top + picViewport.Height + 10
    wid = picViewport.Width
    hgt = HScrollBar.Height
    HScrollBar.Move X, Y, wid, hgt

    ' Set the new world window bounds.
    SetWorldWindow
End Sub

' Move the world window.
Private Sub HScrollBar_Change()
    If IgnoreSbarChange Then Exit Sub
    HScrollBarChanged
End Sub

' The vertical scroll bar has been moved.
' Adjust the world window.
Private Sub VScrollBarChanged()
Dim hgt As Single

    hgt = Wymax - Wymin
    Wymax = VScrollBar.Value / 100
    Wymin = Wymax - hgt
    
    ' Remap the world window.
    IgnoreSbarChange = True
    SetWorldWindow
    IgnoreSbarChange = False
End Sub
' The horizontal scroll bar has been moved.
' Adjust the world window.
Private Sub HScrollBarChanged()
Dim wid As Single
    
    wid = Wxmax - Wxmin
    Wxmin = HScrollBar.Value / 100
    Wxmax = Wxmin + wid
    
    ' Remap the world window.
    IgnoreSbarChange = True
    SetWorldWindow
    IgnoreSbarChange = False
End Sub


Private Sub mnuFileExit_Click()
    StopZoom    ' If we're zooming, stop it.
    
    Unload Me
End Sub

' Change the level of magnification.
Private Sub mnuScaleMag_Click(Index As Integer)
    StopZoom    ' If we're zooming, stop it.
    
    If Index = 1 Then
        ' Return to full scale.
        SetScaleFull
    ElseIf Index < 10 Then
        ' Magnify by the indicated amount.
        SetScaleFactor CSng(Index)
    Else
        ' Zoom out by 1/(Index \ 10).
        SetScaleFactor 1 / (Index \ 10)
    End If
End Sub


' Allow the user to select an area to zoom in on.
Private Sub mnuScaleZoom_Click()
    ' Enable zooming.
    picViewport.MousePointer = vbCrosshair
    DrawingMode = mode_StartZoom
End Sub

' If we are zooming, start the rubberband box.
Private Sub picViewport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DrawingMode <> mode_StartZoom Then Exit Sub
    DrawingMode = mode_Zooming

    OldMode = picViewport.DrawMode
    picViewport.DrawMode = vbInvert
    
    StartX = X
    StartY = Y
    LastX = X
    LastY = Y
    picViewport.Line (StartX, StartY)-(LastX, LastY), , B
End Sub

' If we are zooming, continue the rubberband box.
Private Sub picViewport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DrawingMode <> mode_Zooming Then Exit Sub

    ' Erase the old box.
    picViewport.Line (StartX, StartY)-(LastX, LastY), , B
    
    ' Draw the new box.
    LastX = X
    LastY = Y
    picViewport.Line (StartX, StartY)-(LastX, LastY), , B
End Sub

' If we are zooming, finish the rubberband box.
Private Sub picViewport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim wid As Single
Dim hgt As Single
Dim mid As Single

    If DrawingMode <> mode_Zooming Then Exit Sub
    DrawingMode = mode_None

    ' Erase the old box.
    picViewport.Line (StartX, StartY)-(LastX, LastY), , B
    LastX = X
    LastY = Y
    
    ' We're done drawing for this rubberband box.
    picViewport.DrawMode = OldMode
    picViewport.MousePointer = vbDefault

    ' Set the new world window bounds.
    If StartX > LastX Then
        Wxmin = LastX
        Wxmax = StartX
    Else
        Wxmin = StartX
        Wxmax = LastX
    End If
    If StartY > LastY Then
        Wymin = LastY
        Wymax = StartY
    Else
        Wymin = StartY
        Wymax = LastY
    End If

    ' Set the new world window bounds.
    SetWorldWindow
End Sub


Private Sub picViewport_Paint()
    DrawSmiley picViewport
End Sub



' Move the world window.
Private Sub VScrollBar_Change()
    If IgnoreSbarChange Then Exit Sub
    VScrollBarChanged
End Sub

