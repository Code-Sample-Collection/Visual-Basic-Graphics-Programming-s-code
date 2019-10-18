VERSION 5.00
Begin VB.Form frmDevInfo 
   Caption         =   "DevInfo"
   ClientHeight    =   3630
   ClientLeft      =   1320
   ClientTop       =   1035
   ClientWidth     =   5055
   LinkTopic       =   "PalInfo"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   242
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   Begin VB.TextBox txtInfo 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmDevInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Const TECHNOLOGY = 2     ' Device type.
Private Const RASTERCAPS = 38    ' Raster capabilities.
Private Const NUMRESERVED = 106  ' # reserved entries in palette.
Private Const SIZEPALETTE = 104  ' Size of system palette.
Private Const HORZSIZE = 4       ' Horizontal size in millimeters.
Private Const VERTSIZE = 6       ' Vertical size in millimeters.
Private Const HORZRES = 8        ' Horizontal width in pixels.
Private Const VERTRES = 10       ' Vertical width in pixels.
Private Const LOGPIXELSX = 88    ' Logical pixels/inch horizontally.
Private Const LOGPIXELSY = 90    ' Logical pixels/inch horizontally.
Private Const BITSPIXEL = 12     ' # bits per pixel.
Private Const PLANES = 14        ' # color planes.
Private Const NUMBRUSHES = 16    ' # brushes.
Private Const NUMCOLORS = 24     ' # colors in device color table.
Private Const NUMFONTS = 22      ' # fonts.
Private Const NUMMARKERS = 20    ' # markers.
Private Const NUMPENS = 18       ' # pens.
Private Const COLORRES = 108     ' Color resolution.
Private Const CURVECAPS = 28     ' Curve capabilities.
Private Const LINECAPS = 30      ' Line capabilities.
Private Const POLYGONALCAPS = 32 ' Polygon capabilities.
Private Const TEXTCAPS = 34      ' Text capabilities.

' TECHNOLOGY values.
Private Const DT_PLOTTER = 0     ' Vector plotter.
Private Const DT_RASDISPLAY = 1  ' Raster display.
Private Const DT_RASPRINTER = 2  ' Raster printer.
Private Const DT_RASCAMERA = 3   ' Raster camera.
Private Const DT_CHARSTREAM = 4  ' Character-stream, PLP.
Private Const DT_METAFILE = 5    ' Metafile, VDM.
Private Const DT_DISPFILE = 6    ' Display-file.

' RASTERCAPS values.
Private Const RC_BITBLT = 1          ' Can BLT.
Private Const RC_BANDING = 2         ' Supports banding support.
Private Const RC_SCALING = 4         ' Supports scaling support.
Private Const RC_BITMAP64 = 8        ' Supports >64K bitmaps.
Private Const RC_GDI20_OUTPUT = &H10 ' Has 2.0 output calls.
Private Const RC_DI_BITMAP = &H80    ' Supports DIB to memory.
Private Const RC_PALETTE = &H100     ' Supports palettes.
Private Const RC_DIBTODEV = &H200    ' Supports DIBitsToDevice.
Private Const RC_BIGFONT = &H400     ' Supports >64K fonts.
Private Const RC_STRETCHBLT = &H800  ' Supports StretchBlt.
Private Const RC_FLOODFILL = &H1000  ' Supports FloodFill.
Private Const RC_STRETCHDIB = &H2000 ' Supports StretchDIBits.

' CURVECAP values.
Private Const CC_CHORD = 4       ' Chords.
Private Const CC_CIRCLES = 1     ' Circles.
Private Const CC_ELLIPSES = 8    ' Ellipses.
Private Const CC_INTERIORS = 128 ' Can do interiors.
Private Const CC_PIE = 2         ' Pie slices.
Private Const CC_STYLED = 32     ' Styled lines.
Private Const CC_WIDE = 16       ' Wide lines.
Private Const CC_WIDESTYLED = 64 ' Wide styled lines.

' LINECAPS values.
Private Const LC_INTERIORS = 128 ' Interiors.
Private Const LC_MARKER = 4      ' Markers.
Private Const LC_POLYLINE = 2    ' Polylines.
Private Const LC_POLYMARKER = 8  ' Polymarkers.
Private Const LC_STYLED = 32     ' Styled lines.
Private Const LC_WIDE = 16       ' Wide lines.
Private Const LC_WIDESTYLED = 64 ' Wide styled lines.

' POLYGONCAPS values.
Private Const PC_INTERIORS = 128 ' Interiors.
Private Const PC_POLYGON = 1     ' Alternate filled polygons.
Private Const PC_RECTANGLE = 2   ' Rectangles.
Private Const PC_SCANLINE = 8    ' Scanlines.
Private Const PC_STYLED = 32     ' Styled borders.
Private Const PC_WIDE = 16       ' Wide borders.
Private Const PC_WIDESTYLED = 64 ' Wide styled borders.
Private Const PC_WINDPOLYGON = 4 ' Winding number filled polygons.

' TEXTCAPS values.
Private Const TC_CP_STROKE = &H4     ' Stroke clip precision.
Private Const TC_CR_90 = &H8         ' Characters rotated 90 degrees.
Private Const TC_CR_ANY = &H10       ' Characters rotated by any angle.
Private Const TC_EA_DOUBLE = &H200   ' Bold.
Private Const TC_IA_ABLE = &H400     ' Italics.
Private Const TC_OP_CHARACTER = &H1  ' Character output precision.
Private Const TC_OP_STROKE = &H2     ' Stroke output precision.
Private Const TC_RA_ABLE = &H2000    ' Raster fonts.
Private Const TC_SA_CONTIN = &H100   ' Continuously scaled fonts.
Private Const TC_SA_DOUBLE = &H40    ' Fonts scaled by a double.
Private Const TC_SA_INTEGER = &H80   ' Fonts scaled by an integer.
Private Const TC_SF_X_YINDEP = &H20  ' Fonts scaled in the X and Y directions independently.
Private Const TC_SO_ABLE = &H1000    ' Strikeout.
Private Const TC_UA_ABLE = &H800     ' Underline.
Private Const TC_VA_ABLE = &H4000    ' Vector fonts.

' Get the device information.
Private Sub Form_Load()
Dim txt As String
Dim sys_pal_size As Integer
Dim num_static As Integer
Dim clrres As Integer
Dim rascaps As Integer
Dim curves As Integer
Dim lines As Integer
Dim poly As Integer
Dim text As Integer

    ' Get the device type.
    txt = "This device is a "
    Select Case GetDeviceCaps(hdc, TECHNOLOGY)
        Case DT_PLOTTER
            txt = txt & "vector plotter"
        Case DT_RASDISPLAY
            txt = txt & "raster display"
        Case DT_RASPRINTER
            txt = txt & "raster printer"
        Case DT_RASCAMERA
            txt = txt & "raster camera"
        Case DT_CHARSTREAM
            txt = txt & "character-stream, PLP"
        Case DT_METAFILE
            txt = txt & "metafile, VDM"
        Case DT_DISPFILE
            txt = txt & "display-file"
    End Select
    txt = txt & "." & vbCrLf
    
    ' Get the display size in millimeters.
    txt = txt & "The display is" & _
        Str$(GetDeviceCaps(hdc, HORZSIZE)) & "x" & _
        Format$(GetDeviceCaps(hdc, VERTSIZE))
    
    ' Get the display size in pixels.
    txt = txt & " millimeters or" & _
        Str$(GetDeviceCaps(hdc, HORZRES)) & "x" & _
        Format$(GetDeviceCaps(hdc, VERTRES)) & _
        " pixels." & vbCrLf
    
    ' Get logical pixels per inch.
    txt = txt & "Horizontal pixels per inch:" & _
        Str$(GetDeviceCaps(hdc, LOGPIXELSX)) & _
        vbCrLf
    txt = txt & "Vertical pixels per inch:" & _
        Str$(GetDeviceCaps(hdc, LOGPIXELSY)) & _
        vbCrLf
        
    ' Get color and tool information.
    txt = txt & "Bits per pixel:" & _
        Str$(GetDeviceCaps(hdc, BITSPIXEL)) & _
        "." & vbCrLf
    txt = txt & "Color planes:" & _
        Str$(GetDeviceCaps(hdc, PLANES)) & _
        "." & vbCrLf
    txt = txt & "Device brushes:" & _
        Str$(GetDeviceCaps(hdc, NUMBRUSHES)) & _
        "." & vbCrLf
    txt = txt & "Device colors:" & _
        Str$(GetDeviceCaps(hdc, NUMCOLORS)) & _
        "." & vbCrLf
    txt = txt & "Device fonts:" & _
        Str$(GetDeviceCaps(hdc, NUMFONTS)) & _
        "." & vbCrLf
    txt = txt & "Device markers:" & _
        Str$(GetDeviceCaps(hdc, NUMMARKERS)) & _
        "." & vbCrLf
    txt = txt & "Device pens:" & _
        Str$(GetDeviceCaps(hdc, NUMPENS)) & _
        "." & vbCrLf
    
    ' See if the screen supports palettes.
    rascaps = GetDeviceCaps(hdc, RASTERCAPS)
    If rascaps And RC_PALETTE Then
        txt = txt & "This device supports palettes." & vbCrLf
        
        ' See how big the system palette is.
        sys_pal_size = GetDeviceCaps(hdc, SIZEPALETTE)
        txt = txt & "The system palette holds" & _
            Str$(sys_pal_size) & " entries." & _
            vbCrLf
        
        ' See how many static colors there are.
        num_static = GetDeviceCaps(hdc, NUMRESERVED)
        txt = txt & "There are" & Str$(num_static) & _
            " static colors." & vbCrLf
        
        ' Give the indexes of the static colors.
        txt = txt & "The static colors are in system palette entries: 0-" & _
            Format$(num_static \ 2 - 1) & " and " & _
            Format$(sys_pal_size - num_static \ 2) & _
            "-" & Format$(sys_pal_size - 1) & _
            "." & vbCrLf
    
        ' Get the color resolution.
        clrres = GetDeviceCaps(hdc, COLORRES)
        txt = txt & "The color resolution is" & _
            Str$(clrres) & " bits per pixel (" & _
            Format$(2 ^ clrres) & _
            " possible values)." & vbCrLf
    
        ' Get RASTERCAPS values.
        txt = txt & "This device supports the following raster features:" & _
            vbCrLf
        If rascaps And RC_BANDING Then _
            txt = txt & "    Banding." & vbCrLf
        If rascaps And RC_BIGFONT Then _
            txt = txt & "    Fonts bigger than 64K." & vbCrLf
        If rascaps And RC_BITBLT Then _
            txt = txt & "    Bitmap transfer." & vbCrLf
        If rascaps And RC_BITMAP64 Then _
            txt = txt & "    Bitmaps bigger than 64K." & vbCrLf
        If rascaps And RC_DI_BITMAP Then _
            txt = txt & "    The SetDIBits and GetDIBits functions." & vbCrLf
        If rascaps And RC_DIBTODEV Then _
            txt = txt & "    The SetDIBitsToDevice function." & vbCrLf
        If rascaps And RC_FLOODFILL Then _
            txt = txt & "    Flood fills." & vbCrLf
        If rascaps And RC_GDI20_OUTPUT Then _
            txt = txt & "    Windows 2.0 features." & vbCrLf
        If rascaps And RC_PALETTE Then _
            txt = txt & "    Palettes." & vbCrLf
        If rascaps And RC_SCALING Then _
            txt = txt & "    Scaling." & vbCrLf
        If rascaps And RC_STRETCHBLT Then _
            txt = txt & "    The StretchBlt function." & vbCrLf
        If rascaps And RC_STRETCHDIB Then _
            txt = txt & "    The StretchDIBits function." & vbCrLf
            
        ' Get CURVECAPS values.
        curves = GetDeviceCaps(hdc, CURVECAPS)
        txt = txt & "This device supports the following curve features:" & _
            vbCrLf
        If curves And CC_CHORD Then _
            txt = txt & "    Chords." & vbCrLf
        If curves And CC_CIRCLES Then _
            txt = txt & "    Circles." & vbCrLf
        If curves And CC_ELLIPSES Then _
            txt = txt & "    Ellipses." & vbCrLf
        If curves And CC_INTERIORS Then _
            txt = txt & "    Interiors." & vbCrLf
        If curves And CC_PIE Then _
            txt = txt & "    Pie slices." & vbCrLf
        If curves And CC_STYLED Then _
            txt = txt & "    Line styles." & vbCrLf
        If curves And CC_WIDE Then _
            txt = txt & "    Wide lines." & vbCrLf
        If curves And CC_WIDESTYLED Then _
            txt = txt & "    Wide styled lines." & vbCrLf

        ' Get LINECAPS values.
        lines = GetDeviceCaps(hdc, LINECAPS)
        txt = txt & "This device supports the following line features:" & _
            vbCrLf
        If lines And LC_INTERIORS Then _
            txt = txt & "    Interiors." & vbCrLf
        If lines And LC_MARKER Then _
            txt = txt & "    Markers." & vbCrLf
        If lines And LC_POLYLINE Then _
            txt = txt & "    Polyline." & vbCrLf
        If lines And LC_POLYMARKER Then _
            txt = txt & "    Polymarkers." & vbCrLf
        If lines And LC_STYLED Then _
            txt = txt & "    Styled lines." & vbCrLf
        If lines And LC_WIDE Then _
            txt = txt & "    Wide lines." & vbCrLf
        If lines And LC_WIDESTYLED Then _
            txt = txt & "    Wide styled lines." & vbCrLf

        ' Get POLYGONALCAPS values.
        poly = GetDeviceCaps(hdc, POLYGONALCAPS)
        txt = txt & "This device supports the following polygon features:" & _
            vbCrLf
        If lines And PC_INTERIORS Then _
            txt = txt & "    Interiors." & vbCrLf
        If lines And PC_POLYGON Then _
            txt = txt & "    Alternate filled polygons." & vbCrLf
        If lines And PC_RECTANGLE Then _
            txt = txt & "    Rectangles." & vbCrLf
        If lines And PC_SCANLINE Then _
            txt = txt & "    Scan lines." & vbCrLf
        If lines And PC_STYLED Then _
            txt = txt & "    Styled borders." & vbCrLf
        If lines And PC_WIDE Then _
            txt = txt & "    Wide borders." & vbCrLf
        If lines And PC_WIDESTYLED Then _
            txt = txt & "    Wide styled borders." & vbCrLf
        If lines And PC_WINDPOLYGON Then _
            txt = txt & "    Winding number filled polygons." & vbCrLf

        ' Get TEXTCAPS values.
        text = GetDeviceCaps(hdc, TEXTCAPS)
        txt = txt & "This device supports the following text features:" & _
            vbCrLf
        If lines And TC_CP_STROKE Then _
            txt = txt & "    Stroke clip precision." & vbCrLf
        If lines And TC_CR_90 Then _
            txt = txt & "    Characters rotated 90 degrees." & vbCrLf
        If lines And TC_CR_ANY Then _
            txt = txt & "    Characters rotated through any angle." & vbCrLf
        If lines And TC_EA_DOUBLE Then _
            txt = txt & "    Double weight fonts (bold)." & vbCrLf
        If lines And TC_IA_ABLE Then _
            txt = txt & "    Italics." & vbCrLf
        If lines And TC_OP_CHARACTER Then _
            txt = txt & "    Character output precision." & vbCrLf
        If lines And TC_OP_STROKE Then _
            txt = txt & "    Stroke output precision." & vbCrLf
        If lines And TC_RA_ABLE Then _
            txt = txt & "    Raster fonts." & vbCrLf
        If lines And TC_SA_CONTIN Then _
            txt = txt & "    Fonts scaled by any factor." & vbCrLf
        If lines And TC_SA_DOUBLE Then _
            txt = txt & "    Font scaled by a factor of 2." & vbCrLf
        If lines And TC_SA_INTEGER Then _
            txt = txt & "    Fonts scaled by integer multiples." & vbCrLf
        If lines And TC_SF_X_YINDEP Then _
            txt = txt & "    Fonts scaled in the X and Y directions independently." & vbCrLf
        If lines And TC_SO_ABLE Then _
            txt = txt & "    Strikeout." & vbCrLf
        If lines And TC_UA_ABLE Then _
            txt = txt & "    Underline." & vbCrLf
        If lines And TC_VA_ABLE Then _
            txt = txt & "    Vector fonts." & vbCrLf

    Else
        txt = txt & "This device is not using palettes." & vbCrLf
    End If

    txtInfo.text = txt
End Sub

' Make the text box as large as possible.
Private Sub Form_Resize()
    If WindowState = vbMinimized Then Exit Sub

    txtInfo.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
