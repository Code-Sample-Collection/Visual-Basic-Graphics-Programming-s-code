VERSION 5.00
Begin VB.Form frmCopyDDB 
   Caption         =   "CopyDDB"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   Palette         =   "CopyDDB.frx":0000
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5400
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTriplet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2355
      Left            =   2640
      ScaleHeight     =   153
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   3
      Top             =   2760
      Width           =   2535
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2355
      Left            =   720
      Picture         =   "CopyDDB.frx":20E32
      ScaleHeight     =   153
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Default         =   -1  'True
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.PictureBox picPSet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2355
      Left            =   0
      ScaleHeight     =   153
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   1
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Using an RGBTriplet array"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   7
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Using Point and PSet"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label lblDDBTime 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label lblPSetTime 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   5160
      Width           =   2535
   End
End
Attribute VB_Name = "frmCopyDDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Copy picSource's picture into the result pictures.
Private Sub cmdCopy_Click()
Dim pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim X As Integer
Dim Y As Integer
Dim start_time As Single

    ' Blank previous results.
    cmdCopy.Enabled = False
    MousePointer = vbHourglass
    picPSet.Cls
    picTriplet.Cls
    lblPSetTime.Caption = ""
    lblDDBTime.Caption = ""
    DoEvents

    ' Make the PictureBoxes the same size.
    picPSet.Width = picSource.Width
    picPSet.Height = picSource.Height
    picTriplet.Width = picSource.Width
    picTriplet.Height = picSource.Height

    ' Use Point and PSet.
    start_time = Timer
    For Y = 0 To picSource.ScaleHeight
        For X = 0 To picSource.ScaleWidth
            picPSet.PSet (X, Y), picSource.Point(X, Y)
        Next X
    Next Y
    ' Fill some colored boxes to verify that we
    ' can set pixel values correctly.
    For Y = 0 To 20
        For X = 0 To 20
            picPSet.PSet (X, Y), vbRed
        Next X
    Next Y
    For Y = 21 To 40
        For X = 21 To 40
            picPSet.PSet (X, Y), vbGreen
        Next X
    Next Y
    For Y = 41 To 60
        For X = 41 To 60
            picPSet.PSet (X, Y), vbBlue
        Next X
    Next Y
    lblPSetTime.Caption = _
        Format$(Timer - start_time, "0.00") & _
        " seconds"
    DoEvents

    ' Use the RGBTriplet array.
    start_time = Timer

    ' Get picSource's pixels.
    GetBitmapPixels picSource, pixels, bits_per_pixel

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

    ' If this is 8 bit color, make picTriplet use the
    ' same palette as picSource.
    If bits_per_pixel = 8 Then
        picTriplet.Picture = picSource.Picture
    End If

    ' Set picTriplet's pixels.
    SetBitmapPixels picTriplet, bits_per_pixel, pixels

    lblDDBTime.Caption = _
        Format$(Timer - start_time, "0.00") & _
        " seconds"
    MousePointer = vbDefault
    cmdCopy.Enabled = True
End Sub

Private Sub Form_Load()
    Show

    picSource.ZOrder
    picSource.SetFocus
End Sub


