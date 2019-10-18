Attribute VB_Name = "Geometry"
Option Explicit

Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Const ALTERNATE = 1
Private Const WINDING = 2

' Find the distance from the point (x1, y1) to the
' line passing through (x1, y1) and (x2, y2).
Public Function DistPointToLine(ByVal A As Single, ByVal B As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Single
Dim vx As Single
Dim vy As Single
Dim t As Single
Dim dx As Single
Dim dy As Single
Dim close_x As Single
Dim close_y As Single

    ' Get the vector component for the segment.
    ' The segment is given by:
    '       x(t) = x1 + t * vx
    '       y(t) = y1 + t * vy
    ' where 0.0 <= t <= 1.0
    vx = x2 - x1
    vy = y2 - y1

    ' Find the best t value.
    If (vx = 0) And (vy = 0) Then
        ' The points are the same. There is no segment.
        t = 0
    Else
        ' Calculate the minimal value for t.
        t = -((x1 - A) * vx + (y1 - B) * vy) / (vx * vx + vy * vy)
    End If

    ' Keep the point on the segment.
    If t < 0# Then
        t = 0#
    ElseIf t > 1# Then
        t = 1#
    End If

    ' Set the return values.
    close_x = x1 + t * vx
    close_y = y1 + t * vy
    dx = A - close_x
    dy = B - close_y
    DistPointToLine = Sqr(dx * dx + dy * dy)
End Function
' Return True if the polygon is at this location.
Public Function PolygonIsAt(ByVal is_closed As Boolean, ByVal X As Single, ByVal Y As Single, points() As POINTAPI) As Boolean
Const HIT_DIST = 3
Dim start_i As Integer
Dim i As Integer
Dim num_points As Integer
Dim x1 As Single
Dim y1 As Single
Dim x2 As Single
Dim y2 As Single
Dim dist As Single

    PolygonIsAt = False

    num_points = UBound(points)
    If is_closed Then
        x2 = points(num_points).X
        y2 = points(num_points).Y
        start_i = 1
    Else
        x2 = points(1).X
        y2 = points(1).Y
        start_i = 2
    End If

    ' Check each segment in the Polyline.
    For i = start_i To num_points
        With points(i)
            x1 = .X
            y1 = .Y
        End With
        dist = DistPointToLine(X, Y, x1, y1, x2, y2)
        If dist <= HIT_DIST Then
            PolygonIsAt = True
            Exit For
        End If
        x2 = x1
        y2 = y1
    Next i
End Function
' Return True if the point is inside the polygon.
Public Function PointIsInPolygon(ByVal X As Single, ByVal Y As Single, points() As POINTAPI) As Boolean
Dim polygon_region As Long

    polygon_region = CreatePolygonRgn(points(1), UBound(points), ALTERNATE)
    PointIsInPolygon = PtInRegion(polygon_region, X, Y)
    DeleteObject polygon_region
End Function

