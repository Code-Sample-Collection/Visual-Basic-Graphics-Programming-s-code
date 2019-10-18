VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RayCylinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A cylinder object.

Implements RayTraceable

' Geometry.
Private Point2 As Point3D
Private Point1 As Point3D
Private Radius As Single

Private Const NUM_THETA = 10
Private WireFrame(1 To NUM_THETA, 1 To 2) As Point3D

' The value of u at the hit location.
Private HitU As Single

' Ambient light parameters.
Private AmbientKr As Single
Private AmbientKg As Single
Private AmbientKb As Single

' Diffuse light parameters.
Private DiffuseKr As Single
Private DiffuseKg As Single
Private DiffuseKb As Single

' Specular reflection parameters.
Private SpecularN As Single
Private SpecularK As Single

' Reflected light parameters.
Private ReflectedKr As Single
Private ReflectedKg As Single
Private ReflectedKb As Single

' Refracted light parameters.
Private TransN As Single
Private n1 As Single   ' Index of refraction outside the object.
Private n2 As Single   ' Index of refraction inside the object.
Private TransmittedKr As Single
Private TransmittedKg As Single
Private TransmittedKb As Single

Private IsReflective As Boolean
Private IsTransparent As Boolean
Private DoneOnThisScanline As Boolean

' We had a hit on this scanline.
Private HadHit As Boolean

' We have had a hit on a previous scanline.
Private HadHitOnPreviousScanline As Boolean

' We will not be visible on later scanlines.
Private ForeverCulled As Boolean
' Return the right shade for this polygon.
Private Function GetShade(ByVal pgon As SimplePolygon) As Long
Dim i As Integer
Dim px As Single
Dim py As Single
Dim pz As Single
Dim light_source As LightSource
Dim total_r As Single
Dim total_g As Single
Dim total_b As Single
Dim R1 As Integer
Dim g1 As Integer
Dim b1 As Integer
Dim empty_objects As Collection

    With pgon
        ' Find a central point on this polygon.
        For i = 1 To .PointX.Count
            px = px + .PointX(i)
            py = py + .PointY(i)
            pz = pz + .PointZ(i)
        Next i
        px = px / .PointX.Count
        py = py / .PointX.Count
        pz = pz / .PointX.Count

        ' Add up the light components.
        Set empty_objects = New Collection
        For Each light_source In LightSources
            CalculateHitColorDSA _
                1, empty_objects, Nothing, _
                EyeX, EyeY, EyeZ, _
                px, py, pz, .Nx, .Ny, .Nz, _
                DiffuseKr, DiffuseKg, DiffuseKb, AmbientKr, AmbientKg, AmbientKb, _
                SpecularK, SpecularN, R1, g1, b1
            total_r = total_r + R1
            total_g = total_g + g1
            total_b = total_b + b1
        Next light_source
    End With

    If total_r > 255 Then total_r = 255
    If total_g > 255 Then total_g = 255
    If total_b > 255 Then total_b = 255

    GetShade = RGB(total_r, total_g, total_b)
End Function
' Draw a face if it is not a backface.
Private Sub DrawFace(ByVal pic As PictureBox, X() As Single, Y() As Single, Z() As Single)
Dim pgon As SimplePolygon
Dim i As Integer

    ' Make a polygon.
    Set pgon = New SimplePolygon
    For i = 1 To 4
        pgon.AddPoint X(i), Y(i), Z(i)
    Next i
    pgon.Finish

    ' If it is not a backface, draw it.
    If Not pgon.IsBackface() Then
        pgon.ForeColor = GetColor()
        pgon.DrawPolygon pic
    End If
End Sub
' Add non-backface polygons to this collection.
Public Sub RayTraceable_GetPolygons(ByRef num_polygons As Integer, polygons() As SimplePolygon, ByVal shaded As Boolean)
Dim t As Integer
Dim last_t As Integer
Dim pgon As SimplePolygon
Dim i As Integer
Dim color As Long

    ' If all polygons are the same color,
    ' get an appropriate color.
    If Not shaded Then
        color = GetColor()
    End If

    last_t = NUM_THETA
    For t = 1 To NUM_THETA
        ' Make a polygon.
        Set pgon = New SimplePolygon
        With WireFrame(last_t, 1)
            pgon.AddPoint .Trans(1), .Trans(2), .Trans(3)
        End With
        With WireFrame(t, 1)
            pgon.AddPoint .Trans(1), .Trans(2), .Trans(3)
        End With
        With WireFrame(t, 2)
            pgon.AddPoint .Trans(1), .Trans(2), .Trans(3)
        End With
        With WireFrame(last_t, 2)
            pgon.AddPoint .Trans(1), .Trans(2), .Trans(3)
        End With
        pgon.Finish

        ' See if this is a backface.
        If Not pgon.IsBackface() Then
            ' This is not a backface. Add it to
            ' the list.
            With pgon
                ' See if we are shaded.
                If shaded Then
                    ' We are shaded. Get the
                    ' right color.
                    .ForeColor = GetShade(pgon)
                    .FillColor = .ForeColor
                Else
                    ' We are not shaded. Use the
                    ' normal colors.
                    .ForeColor = vbBlack
                    .FillColor = color
                End If
                num_polygons = num_polygons + 1
                ReDim Preserve polygons(1 To num_polygons)
                Set polygons(num_polygons) = pgon
            End With
        End If
        last_t = t
    Next t
End Sub
' Make a wire frame.
Private Sub MakeWireFrame()
Const PI = 3.14159265

Dim v1x As Single
Dim v1y As Single
Dim v1z As Single
Dim v2x As Single
Dim v2y As Single
Dim v2z As Single
Dim i As Integer
Dim theta As Single
Dim dtheta As Single
Dim v1_scale As Single
Dim v2_scale As Single
Dim X As Single
Dim Y As Single
Dim Z As Single

    ' Get two vectors perpendicular to the axis.
    GetLineNormals _
        Point2.Coord(1) - Point1.Coord(1), _
        Point2.Coord(2) - Point1.Coord(2), _
        Point2.Coord(3) - Point1.Coord(3), _
        v1x, v1y, v1z, v2x, v2y, v2z

    ' Generate the wire frame points.
    theta = 0
    dtheta = 2 * PI / NUM_THETA
    For i = 1 To NUM_THETA
        v1_scale = Sin(theta) * Radius
        v2_scale = Cos(theta) * Radius
        With WireFrame(i, 1)
            .Coord(1) = Point1.Coord(1) + v2_scale * v1x + v1_scale * v2x
            .Coord(2) = Point1.Coord(2) + v2_scale * v1y + v1_scale * v2y
            .Coord(3) = Point1.Coord(3) + v2_scale * v1z + v1_scale * v2z
            .Coord(4) = 1#
        End With
        With WireFrame(i, 2)
            .Coord(1) = Point2.Coord(1) + v2_scale * v1x + v1_scale * v2x
            .Coord(2) = Point2.Coord(2) + v2_scale * v1y + v1_scale * v2y
            .Coord(3) = Point2.Coord(3) + v2_scale * v1z + v1_scale * v2z
            .Coord(4) = 1#
        End With
        theta = theta + dtheta
    Next i
End Sub

' Return an appropriate color for this object.
Private Function GetColor() As Long
Dim R As Integer
Dim G As Integer
Dim B As Integer

    R = 255 * (DiffuseKr + AmbientKr): If R > 255 Then R = 255
    G = 255 * (DiffuseKg + AmbientKg): If G > 255 Then G = 255
    B = 255 * (DiffuseKb + AmbientKb): If B > 255 Then B = 255
    GetColor = RGB(R, G, B)
End Function

' Initialize the object using text parameters in
' a comma-delimited list.
Public Sub SetParameters(ByVal txt As String)
    On Error GoTo CylinderParamError

    ' Read the parameters and initialize the object.
    ' Geometry.
    Radius = CSng(GetDelimitedToken(txt, ","))
    Point1.Coord(1) = CSng(GetDelimitedToken(txt, ","))
    Point1.Coord(2) = CSng(GetDelimitedToken(txt, ","))
    Point1.Coord(3) = CSng(GetDelimitedToken(txt, ","))
    Point1.Coord(4) = 1
    Point2.Coord(1) = CSng(GetDelimitedToken(txt, ","))
    Point2.Coord(2) = CSng(GetDelimitedToken(txt, ","))
    Point2.Coord(3) = CSng(GetDelimitedToken(txt, ","))
    Point2.Coord(4) = 1

    ' Ambient light.
    AmbientKr = CSng(GetDelimitedToken(txt, ","))
    AmbientKg = CSng(GetDelimitedToken(txt, ","))
    AmbientKb = CSng(GetDelimitedToken(txt, ","))

    ' Diffuse reflection.
    DiffuseKr = CSng(GetDelimitedToken(txt, ","))
    DiffuseKg = CSng(GetDelimitedToken(txt, ","))
    DiffuseKb = CSng(GetDelimitedToken(txt, ","))

    ' Specular reflection.
    SpecularN = CSng(GetDelimitedToken(txt, ","))
    SpecularK = CSng(GetDelimitedToken(txt, ","))

    ' Reflected light.
    ReflectedKr = CSng(GetDelimitedToken(txt, ","))
    ReflectedKg = CSng(GetDelimitedToken(txt, ","))
    ReflectedKb = CSng(GetDelimitedToken(txt, ","))
    IsReflective = (ReflectedKr > 0) Or (ReflectedKg > 0) Or (ReflectedKb > 0)

    ' Transmitted light.
    TransN = CSng(GetDelimitedToken(txt, ","))
    n1 = CSng(GetDelimitedToken(txt, ","))
    n2 = CSng(GetDelimitedToken(txt, ","))
    TransmittedKr = CSng(GetDelimitedToken(txt, ","))
    TransmittedKg = CSng(GetDelimitedToken(txt, ","))
    TransmittedKb = CSng(GetDelimitedToken(txt, ","))
    IsTransparent = (TransmittedKr > 0) Or (TransmittedKg > 0) Or (TransmittedKb > 0)

    ' Make a wire frame.
    MakeWireFrame

    Exit Sub

CylinderParamError:
    MsgBox "Error initializing cylinder parameters."
End Sub

' Draw a wireframe for this object.
Public Sub RayTraceable_DrawWireFrame(ByVal pic As PictureBox)
Dim t As Integer
Dim last_t As Integer

    ' Use an appropriate color.
    pic.ForeColor = GetColor()

    last_t = NUM_THETA
    For t = 1 To NUM_THETA
        pic.Line (WireFrame(last_t, 1).Trans(1), WireFrame(last_t, 1).Trans(2))- _
            (WireFrame(t, 1).Trans(1), WireFrame(t, 1).Trans(2))
        pic.Line -(WireFrame(t, 2).Trans(1), WireFrame(t, 2).Trans(2))
        pic.Line -(WireFrame(last_t, 2).Trans(1), WireFrame(last_t, 2).Trans(2))
        last_t = t
    Next t
End Sub
' Draw the object with backfaces removed.
Public Sub RayTraceable_DrawBackfacesRemoved(ByVal pic As PictureBox)
Dim t As Integer
Dim last_t As Integer
Dim X(1 To 4) As Single
Dim Y(1 To 4) As Single
Dim Z(1 To 4) As Single

    ' Use an appropriate color.
    pic.ForeColor = GetColor()

    last_t = NUM_THETA
    For t = 1 To NUM_THETA
        With WireFrame(last_t, 1)
            X(1) = .Trans(1)
            Y(1) = .Trans(2)
            Z(1) = .Trans(3)
        End With
        With WireFrame(t, 1)
            X(2) = .Trans(1)
            Y(2) = .Trans(2)
            Z(2) = .Trans(3)
        End With
        With WireFrame(t, 2)
            X(3) = .Trans(1)
            Y(3) = .Trans(2)
            Z(3) = .Trans(3)
        End With
        With WireFrame(last_t, 2)
            X(4) = .Trans(1)
            Y(4) = .Trans(2)
            Z(4) = .Trans(3)
        End With

        DrawFace pic, X, Y, Z
        last_t = t
    Next t
End Sub
' Apply a transformation matrix to the object.
Public Sub RayTraceable_Apply(M() As Single)
Dim i_theta As Integer

    ' Transform the wire frame.
    For i_theta = 1 To NUM_THETA
        m3Apply WireFrame(i_theta, 1).Coord, _
             M, WireFrame(i_theta, 1).Trans
        m3Apply WireFrame(i_theta, 2).Coord, _
             M, WireFrame(i_theta, 2).Trans
    Next i_theta

    ' Transform the points.
    m3Apply Point1.Coord, M, Point1.Trans
    m3Apply Point2.Coord, M, Point2.Trans
End Sub
' Apply a transformation matrix to the object.
Public Sub RayTraceable_ApplyFull(M() As Single)
Dim i_theta As Integer

    ' Transform the wire frame.
    For i_theta = 1 To NUM_THETA
        m3ApplyFull WireFrame(i_theta, 1).Coord, _
                 M, WireFrame(i_theta, 1).Trans
        m3ApplyFull WireFrame(i_theta, 2).Coord, _
                 M, WireFrame(i_theta, 2).Trans
    Next i_theta

    ' Transform the points.
    m3ApplyFull Point1.Coord, M, Point1.Trans
    m3ApplyFull Point2.Coord, M, Point2.Trans
End Sub

' Return the red, green, and blue components of
' the surface at the hit position.
Public Sub RayTraceable_FindHitColor( _
    ByVal depth As Integer, Objects As Collection, _
    ByVal eye_x As Single, ByVal eye_y As Single, ByVal eye_z As Single, _
    ByVal px As Single, ByVal py As Single, ByVal pz As Single, _
    ByRef R As Integer, ByRef G As Integer, ByRef B As Integer)
Dim Vx As Single
Dim Vy As Single
Dim Vz As Single
Dim Nx As Single
Dim Ny As Single
Dim Nz As Single
Dim n_len  As Single

    ' Find the unit normal at this point.
    Vx = Point2.Trans(1) - Point1.Trans(1)
    Vy = Point2.Trans(2) - Point1.Trans(2)
    Vz = Point2.Trans(3) - Point1.Trans(3)
    Nx = px - (Point1.Trans(1) + HitU * Vx)
    Ny = py - (Point1.Trans(2) + HitU * Vy)
    Nz = pz - (Point1.Trans(3) + HitU * Vz)
    n_len = Sqr(Nx * Nx + Ny * Ny + Nz * Nz)
    Nx = Nx / n_len
    Ny = Ny / n_len
    Nz = Nz / n_len

    ' Get the hit color.
    CalculateHitColor depth, Objects, Me, _
        eye_x, eye_y, eye_z, _
        px, py, pz, _
        Nx, Ny, Nz, _
        DiffuseKr, DiffuseKg, DiffuseKb, _
        AmbientKr, AmbientKg, AmbientKb, _
        SpecularK, SpecularN, _
        ReflectedKr, ReflectedKg, ReflectedKb, IsReflective, _
        TransmittedKr, TransmittedKg, TransmittedKb, TransN, n1, n2, IsTransparent, _
        R, G, B
End Sub
' See if the scanline plane with the indicated
' point and normal intersects this object. Set
' the object's DoneOnThisScanline flag appropriately.
Public Sub RayTraceable_CullScanline(ByVal px As Single, ByVal py As Single, ByVal pz As Single, ByVal Nx As Single, ByVal Ny As Single, ByVal Nz As Single)
    ' See if we will ever be visible again.
    If ForeverCulled Then
        DoneOnThisScanline = True
        Exit Sub
    End If

    ' We have not yet had a hit on this scanline.
    HadHit = False

    ' Do not scanline cull.
    DoneOnThisScanline = False
End Sub
' Return the value T for the point of intersection
' between the vector from point (px, py, pz) in
' the direction <wx, wy, wz>.
'
' direct_calculation is true if we are finding the
' intersection from a viewing position ray. It is
' false if we are finding an reflected intersection
' or a shadow feeler.
Public Function RayTraceable_FindT(ByVal direct_calculation As Boolean, ByVal px As Single, ByVal py As Single, ByVal pz As Single, ByVal Vx As Single, ByVal Vy As Single, ByVal Vz As Single) As Single
Dim x1 As Single
Dim y1 As Single
Dim z1 As Single
Dim Wx As Single
Dim Wy As Single
Dim Wz As Single
Dim w_len_squared As Single
Dim WdotV As Single
Dim A As Single
Dim B As Single
Dim Cx As Single
Dim Cy As Single
Dim Cz As Single
Dim dx As Single
Dim dy As Single
Dim dz As Single
Dim A1 As Single
Dim b1 As Single
Dim C1 As Single
Dim B24AC As Single
Dim t1 As Single
Dim t2 As Single

    ' Find the axis vector.
    Wx = Point2.Trans(1) - Point1.Trans(1)
    Wy = Point2.Trans(2) - Point1.Trans(2)
    Wz = Point2.Trans(3) - Point1.Trans(3)
    
    ' Find A and B for u = A * t + B.
    w_len_squared = Wx * Wx + Wy * Wy + Wz * Wz
    WdotV = Wx * Vx + Wy * Vy + Wz * Vz
    A = WdotV / w_len_squared

    x1 = Point1.Trans(1)
    y1 = Point1.Trans(2)
    z1 = Point1.Trans(3)
    B = (Wx * (px - x1) + _
         Wy * (py - y1) + _
         Wz * (pz - z1)) / w_len_squared

    ' Solve for t.
    Cx = Vx - Wx * A
    Cy = Vy - Wy * A
    Cz = Vz - Wz * A
    dx = px - x1 - Wx * B
    dy = py - y1 - Wy * B
    dz = pz - z1 - Wz * B
    A1 = Cx * Cx + Cy * Cy + Cz * Cz
    b1 = 2 * (Cx * dx + Cy * dy + Cz * dz)
    C1 = dx * dx + dy * dy + dz * dz - Radius * Radius

    ' Solve the quadratic A1*t^2 + B1*t + C1 = 0.
    B24AC = b1 * b1 - 4 * A1 * C1
    If B24AC < 0 Then
        RayTraceable_FindT = -1
        Exit Function
    ElseIf B24AC = 0 Then
        t1 = -b1 / 2 / A1
    Else
        B24AC = Sqr(B24AC)
        t1 = (-b1 + B24AC) / 2 / A1
        t2 = (-b1 - B24AC) / 2 / A1
        ' Use only positive t values.
        If t1 < 0.02 Then t1 = t2
        If t2 < 0.02 Then t2 = t1
        ' Use the smaller t value.
        If t1 > t2 Then t1 = t2
    End If

    ' If there is no positive t value, there's no
    ' intersection in this direction.
    If t1 < 0.02 Then
        RayTraceable_FindT = -1
        Exit Function
    End If

    ' See where on the cylinder this point is.
    HitU = t1 * A + B

    ' If this is not between Point1 and Point2,
    ' ignore it.
    If HitU < 0 Or HitU > 1 Then
        RayTraceable_FindT = -1
    Else
        RayTraceable_FindT = t1
    End If
End Function
' Return the minimum and maximum distances from
' this point.
Private Sub RayTraceable_GetRminRmax(new_min As Single, new_max As Single, ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
Dim dx As Single
Dim dy As Single
Dim dz As Single
Dim dist As Single

    ' Find the distance to the first point.
    dx = X - Point1.Trans(1)
    dy = Y - Point1.Trans(2)
    dz = Z - Point1.Trans(3)
    dist = Sqr(dx * dx + dy * dy + dz * dz)
    new_max = dist + Radius
    new_min = dist - Radius

    ' Find the distance to the second point.
    dx = X - Point2.Trans(1)
    dy = Y - Point2.Trans(2)
    dz = Z - Point2.Trans(3)
    dist = Sqr(dx * dx + dy * dy + dz * dz)
    If new_max < dist + Radius Then new_max = dist + Radius
    If new_min > dist - Radius Then new_min = dist - Radius

    If new_min < 0 Then new_min = 0
End Sub
' Reset the ForeverCulled flag.
Private Sub RayTraceable_ResetCulling()
    ForeverCulled = False
    HadHitOnPreviousScanline = False
End Sub


