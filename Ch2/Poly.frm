VERSION 5.00
Begin VB.Form frmPoly 
   Caption         =   "Poly"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPolyPolygon 
      AutoRedraw      =   -1  'True
      Height          =   1695
      Left            =   2520
      ScaleHeight     =   1635
      ScaleWidth      =   2355
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
   End
   Begin VB.PictureBox picPolyPolyline 
      AutoRedraw      =   -1  'True
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1635
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   2400
      Width           =   2415
   End
   Begin VB.PictureBox picPolygon 
      AutoRedraw      =   -1  'True
      Height          =   1695
      Left            =   2520
      ScaleHeight     =   1635
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.PictureBox picPolyline 
      AutoRedraw      =   -1  'True
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1635
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PolyPolygon"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   7
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PolyPolyline"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Polygon"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   5
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Polyline"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "frmPoly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function PolyPolyline Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, lpdwPolyPoints As Long, ByVal cCount As Long) As Long
Private Declare Function PolyPolygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long

Private Sub Form_Load()
Dim pt(1 To 100) As POINTAPI
Dim counts(1 To 10) As Long
Dim i As Integer
Dim j As Integer
Dim start_i As Integer

    picPolygon.FillStyle = vbDiagonalCross
    picPolyPolygon.FillStyle = vbDiagonalCross

    ' ---------------
    ' Draw a polyline
    ' ---------------
    i = 1
    pt(i).X = 96:    pt(i).Y = 60:    i = i + 1
    pt(i).X = 84:    pt(i).Y = 71:    i = i + 1
    pt(i).X = 66:    pt(i).Y = 71:    i = i + 1
    pt(i).X = 60:    pt(i).Y = 48:    i = i + 1
    pt(i).X = 82:    pt(i).Y = 35:    i = i + 1
    pt(i).X = 112:    pt(i).Y = 42:    i = i + 1
    pt(i).X = 114:    pt(i).Y = 63:    i = i + 1
    pt(i).X = 106:    pt(i).Y = 78:    i = i + 1
    pt(i).X = 85:    pt(i).Y = 86:    i = i + 1
    pt(i).X = 51:    pt(i).Y = 86:    i = i + 1
    pt(i).X = 36:    pt(i).Y = 64:    i = i + 1
    pt(i).X = 44:    pt(i).Y = 35:    i = i + 1
    pt(i).X = 70:    pt(i).Y = 17:    i = i + 1
    pt(i).X = 108:    pt(i).Y = 17:    i = i + 1
    pt(i).X = 126:    pt(i).Y = 32:    i = i + 1
    pt(i).X = 139:    pt(i).Y = 60:    i = i + 1
    pt(i).X = 134:    pt(i).Y = 87:    i = i + 1
    pt(i).X = 115:    pt(i).Y = 99:    i = i + 1
    pt(i).X = 86:    pt(i).Y = 104:    i = i + 1
    pt(i).X = 40:    pt(i).Y = 102:    i = i + 1
    pt(i).X = 19:    pt(i).Y = 79:    i = i + 1
    pt(i).X = 13:    pt(i).Y = 46:    i = i + 1
    pt(i).X = 25:    pt(i).Y = 16:    i = i + 1
    Polyline picPolyline.hdc, pt(1), i - 1

    ' --------------
    ' Draw a polygon
    ' --------------
    i = 1
    pt(i).X = 66:    pt(i).Y = 20:    i = i + 1
    pt(i).X = 53:    pt(i).Y = 50:    i = i + 1
    pt(i).X = 110:   pt(i).Y = 52:    i = i + 1
    pt(i).X = 105:   pt(i).Y = 22:    i = i + 1
    pt(i).X = 144:   pt(i).Y = 26:    i = i + 1
    pt(i).X = 123:   pt(i).Y = 81:    i = i + 1
    pt(i).X = 38:    pt(i).Y = 83:    i = i + 1
    pt(i).X = 11:    pt(i).Y = 13:    i = i + 1
    Polygon picPolygon.hdc, pt(1), i - 1

    ' ------------------
    ' Draw a PolyPolygon
    ' ------------------
    j = 1
    i = 1
    ' Polygon 1.
    start_i = i
    pt(i).X = 15:    pt(i).Y = 33:    i = i + 1
    pt(i).X = 21:    pt(i).Y = 47:    i = i + 1
    pt(i).X = 51:    pt(i).Y = 48:    i = i + 1
    pt(i).X = 64:    pt(i).Y = 19:    i = i + 1
    pt(i).X = 46:    pt(i).Y = 7:     i = i + 1
    counts(j) = i - start_i
    j = j + 1

    ' Polygon 2.
    start_i = i
    pt(i).X = 80:    pt(i).Y = 29:   i = i + 1
    pt(i).X = 75:    pt(i).Y = 17:   i = i + 1
    pt(i).X = 95:    pt(i).Y = 6:    i = i + 1
    pt(i).X = 144:   pt(i).Y = 11:   i = i + 1
    pt(i).X = 152:   pt(i).Y = 82:   i = i + 1
    pt(i).X = 138:   pt(i).Y = 100:  i = i + 1
    pt(i).X = 63:    pt(i).Y = 103:  i = i + 1
    pt(i).X = 49:    pt(i).Y = 91:   i = i + 1
    pt(i).X = 59:    pt(i).Y = 80:   i = i + 1
    pt(i).X = 72:    pt(i).Y = 88:   i = i + 1
    pt(i).X = 127:   pt(i).Y = 84:   i = i + 1
    pt(i).X = 139:   pt(i).Y = 72:   i = i + 1
    pt(i).X = 131:   pt(i).Y = 20:   i = i + 1
    pt(i).X = 101:   pt(i).Y = 16:   i = i + 1
    counts(j) = i - start_i
    j = j + 1

    ' Polygon 3.
    start_i = i
    pt(i).X = 36:    pt(i).Y = 93:    i = i + 1
    pt(i).X = 23:    pt(i).Y = 103:   i = i + 1
    pt(i).X = 7:     pt(i).Y = 92:    i = i + 1
    pt(i).X = 9:     pt(i).Y = 72:    i = i + 1
    pt(i).X = 32:    pt(i).Y = 57:    i = i + 1
    pt(i).X = 78:    pt(i).Y = 57:    i = i + 1
    pt(i).X = 103:   pt(i).Y = 49:    i = i + 1
    pt(i).X = 102:   pt(i).Y = 37:    i = i + 1
    pt(i).X = 108:   pt(i).Y = 28:    i = i + 1
    pt(i).X = 121:   pt(i).Y = 28:    i = i + 1
    pt(i).X = 128:   pt(i).Y = 42:    i = i + 1
    pt(i).X = 126:   pt(i).Y = 58:    i = i + 1
    pt(i).X = 110:   pt(i).Y = 66:    i = i + 1
    pt(i).X = 86:    pt(i).Y = 70:    i = i + 1
    pt(i).X = 43:    pt(i).Y = 70:    i = i + 1
    counts(j) = i - start_i
    j = j + 1
    ' Draw the PolyPolygon.
    PolyPolygon picPolyPolygon.hdc, _
        pt(1), counts(1), j - 1

    ' -------------------
    ' Draw a PolyPolyline
    ' -------------------
    j = 1
    i = 1
    ' Polyline 1.
    start_i = i
    pt(i).X = 14:    pt(i).Y = 31:    i = i + 1
    pt(i).X = 26:    pt(i).Y = 42:    i = i + 1
    pt(i).X = 16:    pt(i).Y = 58:    i = i + 1
    pt(i).X = 29:    pt(i).Y = 73:    i = i + 1
    pt(i).X = 19:    pt(i).Y = 96:    i = i + 1
    counts(j) = i - start_i
    j = j + 1

    ' Polyline 2.
    start_i = i
    pt(i).X = 34:    pt(i).Y = 28:    i = i + 1
    pt(i).X = 51:    pt(i).Y = 40:    i = i + 1
    pt(i).X = 39:    pt(i).Y = 56:    i = i + 1
    pt(i).X = 52:    pt(i).Y = 75:    i = i + 1
    pt(i).X = 43:    pt(i).Y = 93:    i = i + 1
    counts(j) = i - start_i
    j = j + 1

    ' Polyline 3.
    start_i = i
    pt(i).X = 50:    pt(i).Y = 26:    i = i + 1
    pt(i).X = 74:    pt(i).Y = 40:    i = i + 1
    pt(i).X = 59:    pt(i).Y = 55:    i = i + 1
    pt(i).X = 68:    pt(i).Y = 73:    i = i + 1
    pt(i).X = 60:    pt(i).Y = 98:    i = i + 1
    counts(j) = i - start_i
    j = j + 1

    ' Polyline 4.
    start_i = i
    pt(i).X = 71:    pt(i).Y = 24:    i = i + 1
    pt(i).X = 82:    pt(i).Y = 42:    i = i + 1
    pt(i).X = 74:    pt(i).Y = 60:    i = i + 1
    pt(i).X = 81:    pt(i).Y = 78:    i = i + 1
    pt(i).X = 72:    pt(i).Y = 97:    i = i + 1
    counts(j) = i - start_i
    j = j + 1

    ' Polyline 5.
    start_i = i
    pt(i).X = 87:    pt(i).Y = 25:    i = i + 1
    pt(i).X = 99:    pt(i).Y = 41:    i = i + 1
    pt(i).X = 93:    pt(i).Y = 56:    i = i + 1
    pt(i).X = 98:    pt(i).Y = 75:    i = i + 1
    pt(i).X = 87:    pt(i).Y = 95:    i = i + 1
    counts(j) = i - start_i
    j = j + 1

    ' Polyline 6.
    start_i = i
    pt(i).X = 101:    pt(i).Y = 25:    i = i + 1
    pt(i).X = 112:    pt(i).Y = 42:    i = i + 1
    pt(i).X = 104:    pt(i).Y = 58:    i = i + 1
    pt(i).X = 108:    pt(i).Y = 77:    i = i + 1
    pt(i).X = 100:    pt(i).Y = 97:    i = i + 1
    counts(j) = i - start_i
    j = j + 1

    ' Polyline 7.
    start_i = i
    pt(i).X = 115:    pt(i).Y = 24:    i = i + 1
    pt(i).X = 125:    pt(i).Y = 44:    i = i + 1
    pt(i).X = 118:    pt(i).Y = 59:    i = i + 1
    pt(i).X = 123:    pt(i).Y = 81:    i = i + 1
    pt(i).X = 114:    pt(i).Y = 95:    i = i + 1
    counts(j) = i - start_i
    j = j + 1

    ' Polyline 8.
    start_i = i
    pt(i).X = 127:    pt(i).Y = 25:    i = i + 1
    pt(i).X = 142:    pt(i).Y = 43:    i = i + 1
    pt(i).X = 131:    pt(i).Y = 58:    i = i + 1
    pt(i).X = 133:    pt(i).Y = 77:    i = i + 1
    pt(i).X = 126:    pt(i).Y = 94:    i = i + 1
    ' Draw the PolyPolyline.
    PolyPolyline picPolyPolyline.hdc, _
        pt(1), counts(1), j - 1
End Sub
