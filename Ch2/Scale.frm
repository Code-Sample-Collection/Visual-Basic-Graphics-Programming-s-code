VERSION 5.00
Begin VB.Form frmScale 
   AutoRedraw      =   -1  'True
   Caption         =   "Scale"
   ClientHeight    =   3735
   ClientLeft      =   2235
   ClientTop       =   1305
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   -10.826
   ScaleLeft       =   -1
   ScaleMode       =   0  'User
   ScaleTop        =   11
   ScaleWidth      =   12.217
End
Attribute VB_Name = "frmScale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Sub Form_Load()
Dim X1 As Long
Dim Y1 As Long
Dim X2 As Long
Dim Y2 As Long
Dim scale_left As Single
Dim scale_top As Single
Dim pix_per_unit_x As Single
Dim pix_per_unit_y As Single
    
    ' Custom coordinates for the points.
    X1 = 0
    Y1 = 2
    X2 = 10
    Y2 = 8

    ' Draw an X.
    Line (X1, Y1)-(X2, Y2)
    Line (X1, Y2)-(X2, Y1)

    ' Prepare for translation.
    scale_left = ScaleX(ScaleLeft, ScaleMode, vbPixels)
    scale_top = ScaleY(ScaleTop, ScaleMode, vbPixels)
    pix_per_unit_x = ScaleX(1, ScaleMode, vbPixels)
    pix_per_unit_y = ScaleY(1, ScaleMode, vbPixels)

    ' Translate coordinates into pixels.
    X1 = X1 * pix_per_unit_x - scale_left
    Y1 = Y1 * pix_per_unit_y - scale_top
    X2 = X2 * pix_per_unit_x - scale_left
    Y2 = Y2 * pix_per_unit_y - scale_top

    ' Draw a box around the X.
    If Rectangle(hdc, X1, Y1, X2, Y2) = 0 Then Exit Sub
End Sub
