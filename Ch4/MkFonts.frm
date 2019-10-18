VERSION 5.00
Begin VB.Form frmMkFonts 
   AutoRedraw      =   -1  'True
   Caption         =   "MkFonts"
   ClientHeight    =   3585
   ClientLeft      =   2040
   ClientTop       =   645
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   179.25
   ScaleMode       =   2  'Point
   ScaleWidth      =   360
End
Attribute VB_Name = "frmMkFonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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


' Draw a text string at the indicated position
' using the indicated font parameters.
Private Sub DrawText(ByVal txt As String, ByVal X As Single, ByVal Y As Single, ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal fnWeight As Long, ByVal fbItalic As Long, ByVal fbUnderline As Long, ByVal fbStrikeOut As Long, ByVal fbCharSet As Long, ByVal fbOutputPrecision As Long, ByVal fbClipPrecision As Long, ByVal fbQuality As Long, ByVal fbPitchAndFamily As Long, ByVal lpszFace As String)
Dim newfont As Long
Dim oldfont As Long

    newfont = CreateFont(nHeight, nWidth, nEscapement, 0, fnWeight, fbItalic, fbUnderline, fbStrikeOut, fbCharSet, fbOutputPrecision, fbClipPrecision, fbQuality, fbPitchAndFamily, lpszFace)
    oldfont = SelectObject(hdc, newfont)

    CurrentX = X
    CurrentY = Y
    Print txt
    
    newfont = SelectObject(hdc, oldfont)
    If DeleteObject(newfont) = 0 Then
        Beep
        MsgBox "Error deleting font object.", vbExclamation
    End If
End Sub
' Draw an assortment of text samples.
Private Sub Form_Load()
Dim X As Single
Dim Y As Single
Dim R As Single
Dim I As Long
Dim theta As Long
Dim pt As Long
Dim fnt As String
Dim ang As Single

    AutoRedraw = True

    ' Different weights.
    X = 10
    CurrentY = 0
    pt = 15
    fnt = "Times New Roman"
    For I = 0 To 900 Step 100
        DrawText "Weight" & Str$(I), X, CurrentY, pt, 0, 0, I, False, False, False, DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, TRUETYPE_FONTTYPE, fnt
    Next I

    ' Tall, thin characters.
    X = 85
    Y = 0
    I = 5
    For pt = 15 To 55 Step 10
        DrawText Format$(pt) & "x" & Format$(I), X, Y, pt, I, 0, 0, False, False, False, DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, TRUETYPE_FONTTYPE, fnt
        Y = Y + pt * 0.5
    Next pt

    ' Short, wide characters.
    X = 135
    pt = 15
    CurrentY = 0
    For I = 3 To 20 Step 3
        DrawText Format$(pt) & "x" & Format$(I), X, CurrentY, pt, I, 0, 0, False, False, False, DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, TRUETYPE_FONTTYPE, fnt
    Next I

    ' Rotated characters.
    pt = 15
    X = 280
    Y = 90
    For theta = 360 To 3600 Step 360
        DrawText "     Escapement" & Str$(theta), X, Y, pt, 0, theta, 0, False, False, False, DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, CLIP_LH_ANGLES Or CLIP_DEFAULT_PRECIS, PROOF_QUALITY, TRUETYPE_FONTTYPE, fnt
    Next theta
End Sub
