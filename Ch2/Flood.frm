VERSION 5.00
Begin VB.Form frmFlood 
   Caption         =   "Flood"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmFlood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function PolyPolygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long

Private Sub Form_Load()
Dim Point(1 To 100) As POINTAPI
Dim NumPoints(1 To 10) As Long

    AutoRedraw = True
    ScaleMode = vbPixels

    ' Initialize the point data.
    Point(1).X = 25:  Point(1).Y = 248
    Point(2).X = 16:  Point(2).Y = 163
    Point(3).X = 71:  Point(3).Y = 152
    Point(4).X = 65:  Point(4).Y = 171
    Point(5).X = 32:  Point(5).Y = 173
    Point(6).X = 35:  Point(6).Y = 190
    Point(7).X = 59:  Point(7).Y = 186
    Point(8).X = 62:  Point(8).Y = 206
    Point(9).X = 38:  Point(9).Y = 210
    Point(10).X = 46: Point(10).Y = 243
    NumPoints(1) = 10
    Point(11).X = 83:  Point(11).Y = 169
    Point(12).X = 82:  Point(12).Y = 239
    Point(13).X = 126: Point(13).Y = 241
    Point(14).X = 127: Point(14).Y = 214
    Point(15).X = 96:  Point(15).Y = 222
    Point(16).X = 105: Point(16).Y = 166
    NumPoints(2) = 6
    Point(17).X = 153: Point(17).Y = 168
    Point(18).X = 144: Point(18).Y = 185
    Point(19).X = 143: Point(19).Y = 212
    Point(20).X = 156: Point(20).Y = 236
    Point(21).X = 175: Point(21).Y = 241
    Point(22).X = 190: Point(22).Y = 222
    Point(23).X = 192: Point(23).Y = 186
    Point(24).X = 177: Point(24).Y = 165
    NumPoints(3) = 8
    Point(25).X = 166: Point(25).Y = 182
    Point(26).X = 155: Point(26).Y = 198
    Point(27).X = 164: Point(27).Y = 221
    Point(28).X = 176: Point(28).Y = 219
    Point(29).X = 179: Point(29).Y = 195
    NumPoints(4) = 5
    Point(30).X = 213: Point(30).Y = 165
    Point(31).X = 206: Point(31).Y = 184
    Point(32).X = 204: Point(32).Y = 215
    Point(33).X = 219: Point(33).Y = 235
    Point(34).X = 237: Point(34).Y = 236
    Point(35).X = 248: Point(35).Y = 211
    Point(36).X = 246: Point(36).Y = 176
    Point(37).X = 231: Point(37).Y = 164
    NumPoints(5) = 8
    Point(38).X = 225: Point(38).Y = 175
    Point(39).X = 217: Point(39).Y = 192
    Point(40).X = 219: Point(40).Y = 215
    Point(41).X = 230: Point(41).Y = 220
    Point(42).X = 239: Point(42).Y = 198
    Point(43).X = 234: Point(43).Y = 182
    NumPoints(6) = 6
    Point(44).X = 262: Point(44).Y = 166
    Point(45).X = 264: Point(45).Y = 236
    Point(46).X = 287: Point(46).Y = 238
    Point(47).X = 303: Point(47).Y = 227
    Point(48).X = 310: Point(48).Y = 201
    Point(49).X = 303: Point(49).Y = 174
    Point(50).X = 282: Point(50).Y = 160
    NumPoints(7) = 7
    Point(51).X = 280: Point(51).Y = 182
    Point(52).X = 279: Point(52).Y = 217
    Point(53).X = 291: Point(53).Y = 213
    Point(54).X = 295: Point(54).Y = 197
    Point(55).X = 290: Point(55).Y = 184
    NumPoints(8) = 5
    Point(56).X = 158: Point(56).Y = 32
    Point(57).X = 142: Point(57).Y = 63
    Point(58).X = 105: Point(58).Y = 57
    Point(59).X = 131: Point(59).Y = 91
    Point(60).X = 121: Point(60).Y = 127
    Point(61).X = 160: Point(61).Y = 101
    Point(62).X = 200: Point(62).Y = 124
    Point(63).X = 190: Point(63).Y = 81
    Point(64).X = 213: Point(64).Y = 49
    Point(65).X = 174: Point(65).Y = 58
    NumPoints(9) = 10

    ' Draw the polygons.
    PolyPolygon hdc, Point(1), NumPoints(1), 9

    FillStyle = vbSolid
End Sub
' Flood the clicked area with a random color.
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FillColor = QBColor(Int(1 + Rnd * 15))
    FloodFill hdc, X, Y, vbBlack
    Refresh
End Sub


