VERSION 5.00
Begin VB.Form frmCurveTxt 
   AutoRedraw      =   -1  'True
   Caption         =   "CurveTxt"
   ClientHeight    =   5325
   ClientLeft      =   1815
   ClientTop       =   870
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   266.25
   ScaleMode       =   2  'Point
   ScaleWidth      =   285.75
End
Attribute VB_Name = "frmCurveTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PI = 3.14159
Private Const PI_OVER_2 = PI / 2

' Font weight constants.
Private Const FW_DONTCARE = 0
Private Const FW_THIN = 100
Private Const FW_EXTRALIGHT = 200
Private Const FW_LIGHT = 300
Private Const FW_NORMAL = 400
Private Const FW_MEDIUM = 500
Private Const FW_SEMIBOLD = 600
Private Const FW_BOLD = 700
Private Const FW_EXTRABOLD = 800
Private Const FW_HEAVY = 900
Private Const FW_ULTRALIGHT = FW_EXTRALIGHT
Private Const FW_REGULAR = FW_NORMAL
Private Const FW_DEMIBOLD = FW_SEMIBOLD
Private Const FW_ULTRABOLD = FW_EXTRABOLD
Private Const FW_BLACK = FW_HEAVY

' Character set constants.
Private Const ANSI_CHARSET = 0
Private Const DEFAULT_CHARSET = 1
Private Const SYMBOL_CHARSET = 2
Private Const SHIFTJIS_CHARSET = 128
Private Const OEM_CHARSET = 255

' Output precision constants.
Private Const OUT_CHARACTER_PRECIS = 2
Private Const OUT_DEFAULT_PRECIS = 0
Private Const OUT_DEVICE_PRECIS = 5
Private Const OUT_RASTER_PRECIS = 6
Private Const OUT_STRING_PRECIS = 1
Private Const OUT_STROKE_PRECIS = 3
Private Const OUT_TT_ONLY_PRECIS = 7
Private Const OUT_TT_PRECIS = 4

' Clipping precision constants.
Private Const CLIP_CHARACTER_PRECIS = 1
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const CLIP_EMBEDDED = &H80
Private Const CLIP_LH_ANGLES = &H10
Private Const CLIP_STROKE_PRECIS = 2
Private Const CLIP_TO_PATH = 4097
Private Const CLIP_TT_ALWAYS = &H20

' Character quality constants.
Private Const DEFAULT_QUALITY = 0
Private Const DRAFT_QUALITY = 1
Private Const PROOF_QUALITY = 2

' Pitch and family constants.
Private Const DEFAULT_PITCH = 0
Private Const FIXED_PITCH = 1
Private Const VARIABLE_PITCH = 2
Private Const TRUETYPE_FONTTYPE = &H4
Private Const FF_DECORATIVE = 80  '  Old English, etc.
Private Const FF_DONTCARE = 0     '  Don't care or don't know.
Private Const FF_MODERN = 48      '  Constant stroke width, serifed or sans-serifed.
Private Const FF_ROMAN = 16       '  Variable stroke width, serifed.
Private Const FF_SCRIPT = 64      '  Cursive, etc.
Private Const FF_SWISS = 32       '  Variable stroke width, sans-serifed.

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W2 As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long

' Draw a text string along a path specified by a
' series of points (ptx(i), pty(i)). The text is
' placed above the curve if parameter above is
' true. The font uses the given font metrics.
Private Sub CurveText(txt As String, numpts As Integer, ptx() As Single, pty() As Single, above As Boolean, nHeight As Long, nWidth As Long, fnWeight As Long, fbItalic As Long, fbUnderline As Long, fbStrikeOut As Long, fbCharSet As Long, fbOutputPrecision As Long, fbClipPrecision As Long, fbQuality As Long, fbPitchAndFamily As Long, lpszFace As String)
Dim newfont As Long
Dim oldfont As Long
Dim theta As Single
Dim escapement As Long
Dim ch As String
Dim chnum As Integer
Dim needed As Single
Dim avail As Single
Dim newavail As Single
Dim pt As Integer
Dim x1 As Single
Dim y1 As Single
Dim x2 As Single
Dim y2 As Single
Dim dx As Single
Dim dy As Single

    avail = 0
    chnum = 1
    
    x1 = ptx(1)
    y1 = pty(1)
    For pt = 2 To numpts
        ' See how long the new segment is.
        x2 = ptx(pt)
        y2 = pty(pt)
        dx = x2 - x1
        dy = y2 - y1
        newavail = Sqr(dx * dx + dy * dy)
        avail = avail + newavail
        
        ' Create a font along the segment.
        If dx > -0.1 And dx < 0.1 Then
            If dy > 0 Then
                theta = PI_OVER_2
            Else
                theta = -PI_OVER_2
            End If
        Else
            theta = Atn(dy / dx)
            If dx < 0 Then theta = theta - PI
        End If
        escapement = -theta * 180# / PI * 10#
        If escapement = 0 Then escapement = 3600
        newfont = CreateFont(nHeight, nWidth, escapement, 0, fnWeight, fbItalic, fbUnderline, fbStrikeOut, fbCharSet, fbOutputPrecision, fbClipPrecision, fbQuality, fbPitchAndFamily, lpszFace)
        oldfont = SelectObject(hdc, newfont)
    
        ' Output characters until no more fit.
        Do
            ' See how big the next character is.
            ' (Add a little to prevent characters
            ' from becoming too close together.)
            ch = Mid$(txt, chnum, 1)
            needed = TextWidth(ch) * 1.2
            
            ' If it's too big, get another segment.
            If needed > avail Then Exit Do
            
            ' See where the character belongs
            ' along the segment.
            CurrentX = x2 - dx / newavail * avail
            CurrentY = y2 - dy / newavail * avail
            If above Then
                ' Place text above the segment.
                CurrentX = CurrentX + dy * nHeight / newavail
                CurrentY = CurrentY - dx * nHeight / newavail
            End If
            
            ' Display the character.
            Print ch;
            
            ' Move on to the next character.
            avail = avail - needed
            chnum = chnum + 1
            If chnum > Len(txt) Then Exit Do
        Loop
        
        ' Free the font.
        newfont = SelectObject(hdc, oldfont)
        DeleteObject newfont

        If chnum > Len(txt) Then Exit For
        x1 = x2
        y1 = y2
    Next pt
End Sub

' Draw a text string along a circle centered at
' (X, Y) with radius R, centered around the angle
' theta in radians measured counterclockwise from
' the X axis.
Private Sub CircleText(txt As String, X As Single, Y As Single, R As Single, ByVal theta As Single, inside As Boolean, nHeight As Long, nWidth As Long, fnWeight As Long, fbItalic As Long, fbUnderline As Long, fbStrikeOut As Long, fbCharSet As Long, fbOutputPrecision As Long, fbClipPrecision As Long, fbQuality As Long, fbPitchAndFamily As Long, lpszFace As String)
Dim newfont As Long
Dim oldfont As Long
Dim escapement As Long
Dim ch As String
Dim I As Integer
Dim wid As Single
Dim R2 As Single

    If inside Then
        R2 = R
    Else
        R2 = R + 0.8 * nHeight
    End If
    
    ' See how long the string is.
    newfont = CreateFont(nHeight, nWidth, 3600, 0, fnWeight, fbItalic, fbUnderline, fbStrikeOut, fbCharSet, fbOutputPrecision, fbClipPrecision, fbQuality, fbPitchAndFamily, lpszFace)
    oldfont = SelectObject(hdc, newfont)
    wid = TextWidth(txt)
    newfont = SelectObject(hdc, oldfont)
    DeleteObject newfont

    ' The minus sign is needed because Sin and Cos
    ' measure angles clockwise while the input
    ' parameter theta is measured counterclockwise.
    theta = -(theta + wid / R / 2)

    ' Start printing letters.
    For I = 1 To Len(txt)
        CurrentX = X + R2 * Cos(theta)
        CurrentY = Y + R2 * Sin(theta)
    
        escapement = (-PI_OVER_2 - theta) * 180# / PI * 10#
        If escapement = 0 Then escapement = 3600
        newfont = CreateFont(nHeight, nWidth, escapement, 0, fnWeight, fbItalic, fbUnderline, fbStrikeOut, fbCharSet, fbOutputPrecision, fbClipPrecision, fbQuality, fbPitchAndFamily, lpszFace)
        oldfont = SelectObject(hdc, newfont)
        
        ch = Mid$(txt, I, 1)
        Print ch
        theta = theta + TextWidth(ch) / R
        
        newfont = SelectObject(hdc, oldfont)
        DeleteObject newfont
    Next I
End Sub

' Draw an assortment of text samples.
Private Sub Form_Load()
Const NUM_PTS = 22

Dim X As Single
Dim Y As Single
Dim R As Single
Dim pt As Long
Dim fnt As String
Dim ang As Single
Dim I As Integer
Dim ptx(1 To NUM_PTS) As Single
Dim pty(1 To NUM_PTS) As Single
Dim dx As Single
Dim dy As Single

    AutoRedraw = True

    ' *************************
    ' * Text along a polyline *
    ' *************************
    pt = 23
    fnt = "Times New Roman"

    ' Build the points in the path.
    dx = ScaleWidth / (NUM_PTS + 1)
    For I = 1 To NUM_PTS
        ptx(I) = I * dx
        pty(I) = 10 + pt + 10 * Sin(3 * I * PI / NUM_PTS)
    Next I

    ' Display the path.
    Line (ptx(1), pty(1))-(ptx(2), pty(2))
    For I = 3 To NUM_PTS
        Line -(ptx(I), pty(I))
    Next I

    ' Place text along the path.
    CurveText "Text looks best on smooth curves.", NUM_PTS, ptx, pty, True, pt, 0, 700, False, False, False, DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_LH_ANGLES Or CLIP_DEFAULT_PRECIS, PROOF_QUALITY, TRUETYPE_FONTTYPE, fnt
    CurveText "Text looks best on smooth curves.", NUM_PTS, ptx, pty, False, pt, 0, 700, False, False, False, DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_LH_ANGLES Or CLIP_DEFAULT_PRECIS, PROOF_QUALITY, TRUETYPE_FONTTYPE, fnt
    
    ' *****************
    ' * Circular text *
    ' *****************
    pt = 23
    R = 90
    X = ScaleWidth / 2
    Y = R + 20 + 2 * pt
    Circle (X, Y), R
    
    ' Text outside the circle.
    ang = PI_OVER_2
    CircleText "Round and round the mulberry bush", X, Y, R, ang, False, pt, 0, 700, False, False, False, DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_LH_ANGLES Or CLIP_DEFAULT_PRECIS, PROOF_QUALITY, TRUETYPE_FONTTYPE, "Times New Roman"
    ang = -PI_OVER_2
    CircleText "The programmer chased the weasel", X, Y, R, ang, False, pt, 0, 700, False, False, False, DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_LH_ANGLES Or CLIP_DEFAULT_PRECIS, PROOF_QUALITY, TRUETYPE_FONTTYPE, "Times New Roman"

    ' Text inside the circle.
    pt = 20
    ang = 0
    CircleText "CircleText can display text", X, Y, R, ang, True, pt, 0, 700, False, False, False, DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_LH_ANGLES Or CLIP_DEFAULT_PRECIS, PROOF_QUALITY, TRUETYPE_FONTTYPE, "Courier New"
    ang = PI
    CircleText "Inside or outside the circle", X, Y, R, ang, True, pt, 0, 700, False, False, False, DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_LH_ANGLES Or CLIP_DEFAULT_PRECIS, PROOF_QUALITY, TRUETYPE_FONTTYPE, "Courier New"

    ' ************************
    ' * Text along a diamond *
    ' ************************
    pt = 15
    
    ' Build the points in the path.
    dx = 61
    dy = 61
    ptx(1) = X - dx: pty(1) = Y
    ptx(2) = X: pty(2) = Y - dy
    ptx(3) = X + dx: pty(3) = Y
    ptx(4) = X: pty(4) = Y + dy
    ptx(5) = X - dx: pty(5) = Y
    
    ' Display the path.
    Line (ptx(1), pty(1))-(ptx(2), pty(2))
    For I = 3 To 5
        Line -(ptx(I), pty(I))
    Next I
    
    ' Place text along the path.
    CurveText "Sharp corners can cause gaps or overlap when text follows a path.", 5, ptx, pty, True, pt, 0, 700, False, False, False, DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_LH_ANGLES Or CLIP_DEFAULT_PRECIS, PROOF_QUALITY, TRUETYPE_FONTTYPE, fnt
End Sub
