VERSION 5.00
Begin VB.Form frmStretch 
   Caption         =   "Stretch"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResult 
      AutoRedraw      =   -1  'True
      Height          =   2310
      Index           =   2
      Left            =   2520
      ScaleHeight     =   2250
      ScaleWidth      =   2250
      TabIndex        =   3
      Top             =   2520
      Width           =   2310
   End
   Begin VB.PictureBox picResult 
      AutoRedraw      =   -1  'True
      Height          =   2310
      Index           =   1
      Left            =   120
      ScaleHeight     =   2250
      ScaleWidth      =   2250
      TabIndex        =   2
      Top             =   2520
      Width           =   2310
   End
   Begin VB.PictureBox picResult 
      AutoRedraw      =   -1  'True
      Height          =   2310
      Index           =   0
      Left            =   2520
      ScaleHeight     =   2250
      ScaleWidth      =   2250
      TabIndex        =   1
      Top             =   120
      Width           =   2310
   End
   Begin VB.PictureBox picOriginal 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2310
      Left            =   120
      Picture         =   "Stretch.frx":0000
      ScaleHeight     =   2250
      ScaleWidth      =   2250
      TabIndex        =   0
      Top             =   120
      Width           =   2310
   End
End
Attribute VB_Name = "frmStretch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Make the enlarged images.
Private Sub Form_Load()
Dim i As Integer
Dim scale_factor As Single
Dim X As Single
Dim Y As Single
Dim orig_wid As Single
Dim orig_hgt As Single
Dim wid As Single
Dim hgt As Single

    orig_wid = picOriginal.ScaleWidth
    orig_hgt = picOriginal.ScaleHeight

    scale_factor = 4
    For i = 0 To 2
        wid = orig_wid / scale_factor
        hgt = orig_hgt / scale_factor
        X = (orig_wid - wid) / 2
        Y = (orig_hgt - hgt) / 2
        picResult(i).PaintPicture _
            picOriginal.Picture, _
            0, 0, orig_wid, orig_hgt, _
            X, Y, wid, hgt
        scale_factor = scale_factor * 2
    Next i
End Sub
