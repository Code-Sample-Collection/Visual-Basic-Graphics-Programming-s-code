VERSION 5.00
Begin VB.Form frmSetPixs 
   Caption         =   "SetPixs"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   Palette         =   "SetPixs.frx":0000
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   290
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   622
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSetBitmapBits 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3060
      Left            =   6240
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   7
      Top             =   960
      Width           =   3060
   End
   Begin VB.PictureBox picTriplet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3060
      Left            =   3120
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   2
      Top             =   960
      Width           =   3060
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox picPSet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3060
      Left            =   0
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   960
      Width           =   3060
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Using SetBitmapBits Directly"
      Height          =   255
      Index           =   2
      Left            =   6240
      TabIndex        =   9
      Top             =   720
      Width           =   3060
   End
   Begin VB.Label lblSetbitmapBits 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   6240
      TabIndex        =   8
      Top             =   4080
      Width           =   3060
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Using SetPixels"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   6
      Top             =   720
      Width           =   3060
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Using PSet"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   3060
   End
   Begin VB.Label lblDDBTime 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   4080
      Width           =   3060
   End
   Begin VB.Label lblPSetTime 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4080
      Width           =   3060
   End
End
Attribute VB_Name = "frmSetPixs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Set the pixel values in picSetBitmapBits directly.
Private Sub SetPixelsDirectly(ByVal pic As PictureBox)
Dim r(0 To 15) As Byte
Dim g(0 To 15) As Byte
Dim b(0 To 15) As Byte
Dim clr As Long
Dim bits_per_pixel As Integer
Dim hbm As Long
Dim bm As BITMAP
Dim i As Integer
Dim bytes() As Byte
Dim wid As Integer
Dim hgt As Integer
Dim X As Integer
Dim Y As Integer

    ' Initialize the color component lists.
    For i = 0 To 15
        clr = QBColor(i)
        r(i) = clr Mod 256
        g(i) = (clr \ 256) Mod 256
        b(i) = (clr \ 256 \ 256)
    Next i

    ' Get the bitmap information.
    hbm = pic.Image
    GetObject hbm, Len(bm), bm
    bits_per_pixel = bm.bmBitsPixel

    ' Make sure this is a 24-bit image.
    If bits_per_pixel <> 24 Then Exit Sub

    ' Get the bits.
    ReDim bytes(0 To bm.bmWidthBytes - 1, 0 To bm.bmHeight - 1)

    ' Create the pixels array.
    wid = bm.bmWidth
    hgt = bm.bmHeight
    ReDim pixels(0 To wid - 1, 0 To hgt - 1)

    For Y = 0 To hgt - 1
        For X = 0 To wid - 1
            i = ((X \ 10) + (Y \ 10)) Mod 16
            bytes(X * 3, Y) = b(i)
            bytes(X * 3 + 1, Y) = g(i)
            bytes(X * 3 + 2, Y) = r(i)
        Next X
    Next Y

    ' Set the pixel values.
    SetBitmapBits pic.Image, bm.bmWidthBytes * hgt, _
        bytes(0, 0)
    pic.Refresh
End Sub
' Set each pixel in the pictures.
Private Sub cmdGo_Click()
Dim r(0 To 15) As Byte
Dim g(0 To 15) As Byte
Dim b(0 To 15) As Byte
Dim clr As Long
Dim i As Integer
Dim pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim X As Integer
Dim Y As Integer
Dim start_time As Single

    ' Blank previous results.
    cmdGo.Enabled = False
    MousePointer = vbHourglass
    picPSet.Cls
    picTriplet.Cls
    picSetBitmapBits.Cls
    lblPSetTime.Caption = ""
    lblDDBTime.Caption = ""
    lblSetbitmapBits.Caption = ""
    DoEvents

    ' Use Point and PSet.
    start_time = Timer
    For Y = 0 To picPSet.ScaleHeight - 1
        For X = 0 To picPSet.ScaleWidth - 1
            i = ((X \ 10) + (Y \ 10)) Mod 16
            picPSet.PSet (X, Y), QBColor(i)
        Next X
    Next Y
    lblPSetTime.Caption = _
        Format$(Timer - start_time, "0.00") & _
        " seconds"
    DoEvents

    ' Use the RGBTriplet array.
    start_time = Timer

    ' Initialize the color component lists.
    For i = 0 To 15
        clr = QBColor(i)
        r(i) = clr Mod 256
        g(i) = (clr \ 256) Mod 256
        b(i) = (clr \ 256 \ 256)
    Next i

    ' Make room for picTriplet's pixel values.
    ReDim pixels(0 To picTriplet.ScaleWidth - 1, 0 To picTriplet.ScaleHeight - 1)

    ' Set the pixel colors.
    For Y = 0 To picTriplet.ScaleHeight - 1
        For X = 0 To picTriplet.ScaleWidth - 1
            i = ((X \ 10) + (Y \ 10)) Mod 16
            With pixels(X, Y)
                .rgbRed = r(i)
                .rgbGreen = g(i)
                .rgbBlue = b(i)
            End With
        Next X
    Next Y

    ' Set picTriplet's pixels.
    SetBitmapPixels picTriplet, 24, pixels

    lblDDBTime.Caption = _
        Format$(Timer - start_time, "0.00") & _
        " seconds"
    DoEvents

    ' Use SetBitmapBits directly.
    start_time = Timer
    SetPixelsDirectly picSetBitmapBits
    lblSetbitmapBits.Caption = _
        Format$(Timer - start_time, "0.00") & _
        " seconds"

    MousePointer = vbDefault
    cmdGo.Enabled = True
End Sub
