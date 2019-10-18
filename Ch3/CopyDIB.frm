VERSION 5.00
Begin VB.Form frmCopyDIB 
   Caption         =   "CopyDIB"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   Palette         =   "CopyDIB.frx":0000
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3135
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboBitsPerPixel 
      Height          =   315
      ItemData        =   "CopyDIB.frx":20E32
      Left            =   4440
      List            =   "CopyDIB.frx":20E42
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy -->"
      Height          =   495
      Left            =   3660
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.PictureBox picTo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3120
      Left            =   5160
      ScaleHeight     =   3060
      ScaleWidth      =   3300
      TabIndex        =   2
      Top             =   0
      Width           =   3360
   End
   Begin VB.PictureBox picFrom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3120
      Left            =   0
      Picture         =   "CopyDIB.frx":20E53
      ScaleHeight     =   3060
      ScaleWidth      =   3300
      TabIndex        =   1
      Top             =   0
      Width           =   3360
   End
   Begin VB.Label Label1 
      Caption         =   "Color Depth"
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "frmCopyDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Copy picFrom's picture into picTo.
Private Sub cmdCopy_Click()
Dim pixels() As RGBTriplet
Dim color_index() As Byte
Dim bits_per_pixel As Integer
Dim X As Integer
Dim Y As Integer
Dim bitmap_info As BITMAPINFO
Dim clr As Byte
Dim i As Integer
Dim j As Integer

    MousePointer = vbHourglass
    DoEvents

    ' Make the PictureBoxes the same size.
    picTo.Width = picFrom.Width
    picTo.Height = picFrom.Height

    ' Make picFrom's image permanent.
    picFrom.AutoRedraw = True

    ' Get the desired number of bits per pixel.
    bits_per_pixel = CInt(cboBitsPerPixel.Text)

    ' Get picFrom's pixels.
    Select Case bits_per_pixel
        Case 1
            ' Work with color a 2-color palette.
            GetDIBPixelsWithPalette picFrom, bitmap_info, color_index, bits_per_pixel

            ' Fill some checkered boxes to verify
            ' that we can set pixel values correctly.
            For i = 0 To 3
                For j = 0 To 3
                    clr = (i + j) Mod 2
                    For Y = 0 To 20
                        For X = 0 To 20
                            color_index(i * 20 + X, j * 20 + Y) = clr
                        Next X
                    Next Y
                Next j
            Next i

            ' Set picTo's pixels.
            SetDIBPixelsWithPalette picTo, bitmap_info, color_index

        Case 4, 8
            ' Work with color palettes.
            GetDIBPixelsWithPalette picFrom, bitmap_info, color_index, bits_per_pixel
    
            ' Fill some colored boxes to verify that we
            ' can set pixel values correctly.
            clr = FindColorIndex(255, 0, 0, 2 ^ bits_per_pixel, bitmap_info.bmiColors)
            For Y = 0 To 20
                For X = 0 To 20
                    color_index(X, Y) = clr
                Next X
            Next Y
            clr = FindColorIndex(0, 255, 0, 2 ^ bits_per_pixel, bitmap_info.bmiColors)
            For Y = 21 To 40
                For X = 21 To 40
                    color_index(X, Y) = clr
                Next X
            Next Y
            clr = FindColorIndex(0, 0, 255, 2 ^ bits_per_pixel, bitmap_info.bmiColors)
            For Y = 41 To 60
                For X = 41 To 60
                    color_index(X, Y) = clr
                Next X
            Next Y
    
            ' Set picTo's pixels.
            SetDIBPixelsWithPalette picTo, bitmap_info, color_index

        Case 24
            ' Work with 24-bit color.
            GetDIBPixels24Bit picFrom, bitmap_info, pixels
    
            ' Fill some colored boxes to verify that we
            ' can set pixel values correctly.
            For Y = 0 To 20
                For X = 0 To 20
                    With pixels(X, Y)
                        .rgbRed = 255
                        .rgbGreen = 0
                        .rgbBlue = 0
                    End With
                Next X
            Next Y
            For Y = 21 To 40
                For X = 21 To 40
                    With pixels(X, Y)
                        .rgbRed = 0
                        .rgbGreen = 255
                        .rgbBlue = 0
                    End With
                Next X
            Next Y
            For Y = 41 To 60
                For X = 41 To 60
                    With pixels(X, Y)
                        .rgbRed = 0
                        .rgbGreen = 0
                        .rgbBlue = 255
                    End With
                Next X
            Next Y
    
            ' Set picTo's pixels.
            SetDIBPixels24Bit picTo, bitmap_info, pixels
    End Select

    MousePointer = vbDefault
End Sub
Private Sub Form_Load()
    picTo.Width = picFrom.Width
    picTo.Height = picFrom.Height
    Width = picTo.Left + picTo.Width + Width - ScaleWidth
    Height = picTo.Top + picTo.Height + Height - ScaleHeight

    cboBitsPerPixel.Text = "24"
End Sub


