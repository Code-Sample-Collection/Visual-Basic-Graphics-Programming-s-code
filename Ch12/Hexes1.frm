VERSION 5.00
Begin VB.Form frmHexes1 
   Caption         =   "Hexes1"
   ClientHeight    =   3150
   ClientLeft      =   2550
   ClientTop       =   1800
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3150
   ScaleWidth      =   3150
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
   Begin VB.PictureBox picCanvas 
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
Attribute VB_Name = "frmHexes1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' All of the Hex objects.
Private Hexes As Collection

' Global max and min world coordinates
' (including margins).
Private DataXmin As Single
Private DataXmax As Single
Private DataYmin As Single
Private DataYmax As Single

' Set the min and max allowed width and height.
Private DataMinWid As Single
Private DataMinHgt As Single
Private DataMaxWid As Single
Private DataMaxHgt As Single

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
Private DrawingMode As Integer
Const MODE_NONE = 0
Const MODE_START_ZOOM = 1
Const MODE_ZOOMING = 2

Private StartX As Single
Private StartY As Single
Private LastX As Single
Private LastY As Single
Private OldMode As Integer

' The object that is highlighted.
Private Selectedhex As Object

' Find the object at this point.
Private Function ObjectAt(ByVal X As Single, ByVal Y As Single)
Dim obj As Hex

    Set ObjectAt = Nothing
    For Each obj In Hexes
        With obj
            If obj.IsAt(X, Y) Then
                Set ObjectAt = obj
                Exit For
            End If
        End With
    Next obj
End Function
' End a zoom operation early. This happens if the
' user starts a zoom and the selects another menu
' item instead of doing the zoom.
Private Sub StopZoom()
    If DrawingMode <> MODE_START_ZOOM Then Exit Sub
    DrawingMode = MODE_NONE
    
    picCanvas.DrawMode = OldMode
    picCanvas.MousePointer = vbDefault
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

    wid = Wxmax - Wxmin
    xmid = (Wxmax + Wxmin) / 2
    hgt = Wymax - Wymin
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

    ' Make the aspect ratio match the
    ' viewport aspect ratio.
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
    
    ' Check that we're not off to one side.
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
        ' We're taller than the picture. Center.
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
    picCanvas.Scale (Wxmin, Wymax)-(Wxmax, Wymin)

    ' Force the viewport to repaint.
    picCanvas.Refresh
        
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
    ' of its width/height. Large change moves it
    ' 9/10 of its width/height.
    VScrollBar.SmallChange = 100 * (hgt / 10)
    VScrollBar.LargeChange = 100 * (9 * hgt / 10)
    HScrollBar.SmallChange = 100 * (wid / 10)
    HScrollBar.LargeChange = 100 * (9 * wid / 10)

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
    MakeHexes
End Sub

Private Sub Form_Resize()
Dim X As Single
Dim Y As Single
Dim wid As Single
Dim hgt As Single

    ' Fit the viewport to the window.
    X = picCanvas.Left
    Y = picCanvas.Top
    wid = ScaleWidth - 2 * X - VScrollBar.Width
    hgt = ScaleHeight - 2 * Y - HScrollBar.Height
    picCanvas.Move X, Y, wid, hgt
    VAspect = hgt / wid
    
    ' Place the scroll bars next to the viewport.
    X = picCanvas.Left + picCanvas.Width + 10
    Y = picCanvas.Top
    wid = VScrollBar.Width
    hgt = picCanvas.Height
    VScrollBar.Move X, Y, wid, hgt
    
    X = picCanvas.Left
    Y = picCanvas.Top + picCanvas.Height + 10
    wid = picCanvas.Width
    hgt = HScrollBar.Height
    HScrollBar.Move X, Y, wid, hgt

    ' Start at full scale.
    SetScaleFull
End Sub

' Make the Hexes.
Private Sub MakeHexes()
Const NUM_ROWS = 50
Const NUM_COLS = 50

Dim new_hex As Hex
Dim i As Integer
Dim j As Integer
Dim X As Single
Dim Y As Single
Dim wid As Single
Dim hgt As Single

    MousePointer = vbHourglass
    DoEvents

    Set Hexes = New Collection

    Y = 0
    For i = 1 To NUM_ROWS
        X = 0
        For j = 1 To NUM_COLS
            Set new_hex = New Hex
            Hexes.Add new_hex
            new_hex.Cx = X
            new_hex.Cy = Y
            new_hex.Radius = 0.4
            X = X + 2
        Next j
        Y = Y + 2
    Next i

    wid = 2 * NUM_COLS + 1
    hgt = 2 * NUM_ROWS + 1
    DataXmin = -0.1 * wid   ' 10 % margins.
    DataYmin = -0.1 * hgt
    DataXmax = 1.1 * wid
    DataYmax = 1.1 * hgt

    DataMinWid = 10
    DataMinHgt = 10
    DataMaxWid = DataXmax - DataXmin
    DataMaxHgt = DataYmax - DataYmin

    MousePointer = vbDefault
End Sub

' Move the world window.
Private Sub HScrollBar_Change()
    If IgnoreSbarChange Then Exit Sub
    HScrollBarChanged
End Sub

' The vertical scroll bar has been moved. Adjust
' the world window.
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
' The horizontal scroll bar has been moved. Adjust
' the world window.
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
    picCanvas.MousePointer = vbCrosshair
    DrawingMode = MODE_START_ZOOM
End Sub

' If we are zooming, start the rubberband hex.
Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case DrawingMode
        Case MODE_START_ZOOM
            ' Start a zooming rubberband hex.
            DrawingMode = MODE_ZOOMING
        
            OldMode = picCanvas.DrawMode
            picCanvas.DrawMode = vbInvert
            
            StartX = X
            StartY = Y
            LastX = X
            LastY = Y
            picCanvas.Line (StartX, StartY)-(LastX, LastY), , B
        
        Case MODE_NONE
            ' Select a hex.
            Dim oldcolor As Long

            ' Unhighlight the previous hex.
            If Not Selectedhex Is Nothing Then
                Selectedhex.Highlighted = False
                Selectedhex.Draw picCanvas
            End If

            ' Find the selected hex.
            Set Selectedhex = ObjectAt(X, Y)

            ' Highlight the selected hex.
            If Not Selectedhex Is Nothing Then
                Selectedhex.Highlighted = True
                Selectedhex.Draw picCanvas
            End If
    End Select
End Sub

' If we are zooming, continue the rubberband hex.
Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DrawingMode <> MODE_ZOOMING Then Exit Sub

    ' Erase the old hex.
    picCanvas.Line (StartX, StartY)-(LastX, LastY), , B
    
    ' Draw the new hex.
    LastX = X
    LastY = Y
    picCanvas.Line (StartX, StartY)-(LastX, LastY), , B
End Sub

' If we are zooming, finish the rubberband hex.
Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim wid As Single
Dim hgt As Single
Dim mid As Single
    
    If DrawingMode <> MODE_ZOOMING Then Exit Sub
    DrawingMode = MODE_NONE
    
    ' Erase the old hex.
    picCanvas.Line (StartX, StartY)-(LastX, LastY), , B
    LastX = X
    LastY = Y
    
    ' We're done drawing for this rubberband hex.
    picCanvas.DrawMode = OldMode
    picCanvas.MousePointer = vbDefault
    
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


Private Sub picCanvas_Paint()
Dim obj As Hex

    MousePointer = vbHourglass
    DoEvents

    ' Make the Hexes draw themselves.
    For Each obj In Hexes
        obj.Draw picCanvas
    Next obj

    MousePointer = vbDefault
End Sub



' Move the world window.
Private Sub VScrollBar_Change()
    If IgnoreSbarChange Then Exit Sub
    VScrollBarChanged
End Sub
