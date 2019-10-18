VERSION 5.00
Begin VB.Form frmCentText 
   Caption         =   "CentText"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkShowBoundingBox 
      Caption         =   "Show Bounding Box"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chkMarkCenters 
      Caption         =   "Mark Centers"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.PictureBox picText 
      AutoRedraw      =   -1  'True
      Height          =   3735
      Left            =   2040
      ScaleHeight     =   3675
      ScaleWidth      =   3675
      TabIndex        =   4
      Top             =   240
      Width           =   3735
   End
   Begin VB.TextBox txtText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   840
      TabIndex        =   3
      Text            =   "Msg"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtAngle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   840
      TabIndex        =   0
      Text            =   "30"
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Text"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Angle"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmCentText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long

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

' Draw a rotated string centered at the indicated
' position using the indicated font parameters.
Private Sub CenterText(ByVal pic As PictureBox, ByVal xmid As Single, ByVal ymid As Single, ByVal txt As String, ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal fnWeight As Long, ByVal fbItalic As Long, ByVal fbUnderline As Long, ByVal fbStrikeOut As Long, ByVal fbCharSet As Long, ByVal fbOutputPrecision As Long, ByVal fbClipPrecision As Long, ByVal fbQuality As Long, ByVal fbPitchAndFamily As Long, ByVal lpszFace As String)
Const PI = 3.14159265

Dim newfont As Long
Dim oldfont As Long
Dim text_metrics As TEXTMETRIC
Dim internal_leading As Single
Dim total_hgt As Single
Dim text_wid As Single
Dim text_hgt As Single
Dim text_bound_wid As Single
Dim text_bound_hgt As Single
Dim total_bound_wid As Single
Dim total_bound_hgt As Single
Dim theta As Single
Dim phi As Single
Dim x1 As Single
Dim y1 As Single
Dim x2 As Single
Dim y2 As Single
Dim x3 As Single
Dim y3 As Single
Dim x4 As Single
Dim y4 As Single

    ' Create the font.
    newfont = CreateFont(nHeight, nWidth, nEscapement, 0, fnWeight, fbItalic, fbUnderline, fbStrikeOut, fbCharSet, fbOutputPrecision, fbClipPrecision, fbQuality, fbPitchAndFamily, lpszFace)
    oldfont = SelectObject(pic.hdc, newfont)

    ' Get the font metrics.
    GetTextMetrics pic.hdc, text_metrics
    internal_leading = pic.ScaleY(text_metrics.tmInternalLeading, vbPixels, pic.ScaleMode)
    total_hgt = pic.ScaleY(text_metrics.tmHeight, vbPixels, pic.ScaleMode)
    text_hgt = total_hgt - internal_leading
    text_wid = pic.TextWidth(txt)

    ' Get the bounding box geometry.
    theta = nEscapement / 10 / 180 * PI
    phi = PI / 2 - theta
    text_bound_wid = text_hgt * Cos(phi) + text_wid * Cos(theta)
    text_bound_hgt = text_hgt * Sin(phi) + text_wid * Sin(theta)
    total_bound_wid = total_hgt * Cos(phi) + text_wid * Cos(theta)
    total_bound_hgt = total_hgt * Sin(phi) + text_wid * Sin(theta)

    ' Find the desired center point.
    x1 = xmid
    y1 = ymid

    ' Subtract half the height and width of the text
    ' bounding box. This puts (x1, y2) in the upper
    ' left corner of the text bounding box.
    x1 = x1 - text_bound_wid / 2
    y1 = y1 - text_bound_hgt / 2

    ' The start position's X coordinate belongs at
    ' the left edge of the text bounding box, so
    ' x1 is correct. Move the Y coordinate down to
    ' its start position.
    y1 = y1 + text_wid * Sin(theta)

    ' Find the other points on the text bounding box.
    x2 = x1 + text_wid * Cos(theta)
    y2 = y1 - text_wid * Sin(theta)
    x3 = x2 + text_hgt * Cos(phi)
    y3 = y2 + text_hgt * Sin(phi)
    x4 = x3 + -text_wid * Cos(theta)
    y4 = y3 + text_wid * Sin(theta)

    ' See if we should draw the bounding box.
    If chkShowBoundingBox.Value = vbChecked Then
        ' Draw the text bounding box.
        pic.Line (x1, y1)-(x2, y2)
        pic.Line -(x3, y3)
        pic.Line -(x4, y4)
        pic.Line -(x1, y1)
    End If

    ' See if we should mark the text and PictureBox
    ' center positions.
    If chkMarkCenters.Value = vbChecked Then
        ' Draw lines to mark the center of the PictureBox.
        pic.Line (0, 0)-(pic.ScaleWidth, pic.ScaleHeight)
        pic.Line (0, pic.ScaleHeight)-(pic.ScaleWidth, 0)
    
        ' Draw lines to mark the center of the text rectangle.
        pic.Line (x1, y1)-(x3, y3)
        pic.Line (x2, y2)-(x4, y4)
    End If

    ' Move (x1, y1) to the start corner of the
    ' outer bounding box.
    x1 = x1 - (total_bound_wid - text_bound_wid)
    y1 = y1 - (total_bound_hgt - text_bound_hgt)

    ' Display the text.
    pic.CurrentX = x1
    pic.CurrentY = y1
    pic.Print txt

    ' Reselect the old font and delete the new one.
    newfont = SelectObject(pic.hdc, oldfont)
    If DeleteObject(newfont) = 0 Then
        Beep
        MsgBox "Error deleting font object.", vbExclamation
    End If
End Sub

' Draw the rotated text centered in the PictureBox.
Private Sub DrawText()
Dim escapement As Long

    ' Clear the display.
    picText.Line (0, 0)-(picText.ScaleWidth, picText.ScaleHeight), vbWhite, BF

    ' Get the text and angle.
    ' Watch for non-numeric values.
    On Error Resume Next
    escapement = 10 * CInt(txtAngle.Text)
    On Error GoTo 0

    CenterText picText, _
        picText.ScaleWidth / 2, picText.ScaleHeight / 2, _
        txtText.Text, 120, 0, escapement, _
        FW_NORMAL, False, False, False, _
        DEFAULT_CHARSET, OUT_TT_ONLY_PRECIS, _
        CLIP_DEFAULT_PRECIS, PROOF_QUALITY, _
        TRUETYPE_FONTTYPE, "Times New Roman"
End Sub

' Draw the text.
Private Sub chkMarkCenters_Click()
    DrawText
End Sub

' Draw the text.
Private Sub chkShowBoundingBox_Click()
    DrawText
End Sub


' Display the text.
Private Sub Form_Load()
    DrawText
End Sub
' Display the text at the new angle.
Private Sub txtAngle_Change()
    DrawText
End Sub


' Display the new text.
Private Sub txtText_Change()
    DrawText
End Sub


