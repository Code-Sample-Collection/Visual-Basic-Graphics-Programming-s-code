VERSION 5.00
Begin VB.Form frmRect 
   AutoRedraw      =   -1  'True
   Caption         =   "Rect"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Private Sub Form_Load()
Dim hgt As Integer
Dim dh As Integer
Dim wid As Integer
Dim dw As Integer
Dim ymid As Integer
Dim xmid As Integer
Dim i As Integer
Dim round As Boolean

    ScaleMode = vbPixels
    Cls
    hgt = ScaleHeight * 0.1
    dh = ScaleHeight * 0.4 / 5
    wid = ScaleWidth * 0.5
    dw = -ScaleWidth * 0.4 / 5
    ymid = ScaleHeight / 2
    xmid = ScaleWidth / 2
    
    For i = 1 To 5
        If round Then
            RoundRect hdc, xmid - wid, ymid - hgt, xmid + wid, ymid + hgt, 50, 30
        Else
            Rectangle hdc, xmid - wid, ymid - hgt, xmid + wid, ymid + hgt
        End If
        round = Not round
        
        wid = wid + dw
        hgt = hgt + dh
    Next i
    Refresh
End Sub


