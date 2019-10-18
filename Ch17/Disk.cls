VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RayDisk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A disk in a plane.

Implements RayTraceable

Private Point1 As Point3D   ' Point on plane.
Private Point2 As Point3D   ' Normal = P2 - P1.
Private Radius As Single    ' Radius.

' Wire frame variables.
Private Const WIRE_POINTS = 30
Private WireFrame(1 To WIRE_POINTS) As Point3D

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

' Return the unit surface normal.
Private Sub GetUnitNormal(ByRef Nx As Single, ByRef Ny As Single, ByRef Nz As Single)
Dim n_len As Single

    Nx = Point2.Trans(1) - Point1.Trans(1)
    Ny = Point2.Trans(2) - Point1.Trans(2)
    Nz = Point2.Trans(3) - Point1.Trans(3)
    n_len = Sqr(Nx * Nx + Ny * Ny + Nz * Nz)
    Nx = Nx / n_len
    Ny = Ny / n_len
    Nz = Nz / n_len
End Sub

' Add non-backface polygons to this collection.
Public Sub RayTraceable_GetPolygons(ByRef num_polygons As Integer, polygons() As SimplePolygon, ByVal shaded As Boolean)
Dim i As Integer
Dim pgon As SimplePolygon

    ' Make a polygon.
    Set pgon = New SimplePolygon

    ' Add points to the polygon.
    For i = 1 To WIRE_POINTS
        With WireFrame(i)
            pgon.AddPoint .Trans(1), .Trans(2), .Trans(3)
        End With
    Next i

    ' See if we are shaded.
    If shaded Then
        ' We are shaded. Get the right color.
        pgon.ForeColor = GetShade(pgon)
        pgon.FillColor = pgon.ForeColor
    Else
        ' We are not shaded. Use the normal colors.
        pgon.ForeColor = vbBlack
        pgon.FillColor = GetColor()
    End If

    ' Add the polygon to the list.
    num_polygons = num_polygons + 1
    ReDim Preserve polygons(1 To num_polygons)
    Set polygons(num_polygons) = pgon
End Sub
' Draw a wireframe for this object.
Public Sub RayTraceable_DrawWireFrame(ByVal pic As PictureBox)
Dim i As Integer

    ' Use an appropriate color.
    pic.ForeColor = GetColor()

    ' Draw the disk.
    With WireFrame(WIRE_POINTS)
        pic.CurrentX = .Trans(1)
        pic.CurrentY = .Trans(2)
    End With
    For i = 1 To WIRE_POINTS
        With WireFrame(i)
            pic.Line -(.Trans(1), .Trans(2))
        End With
    Next i
End Sub

' Initialize the object using text parameters in
' a comma-delimited list.
Public Sub SetParameters(ByVal txt As String)
    On Error GoTo DiskParamError

    ' Read the parameters and initialize the object.
    ' Geometry.
    Radius = CSng(GetDelimitedToken(txt, ","))
    Point1.Coord(1) = CSng(GetDelimitedToken(txt, ","))
    Point1.Coord(2) = CSng(GetDelimitedToken(txt, ","))
    Point1.Coord(3) = CSng(GetDelimitedToken(txt, ","))
    Point1.Coord(4) = 1
    Point2.Coord(1) = Point1.Coord(1) + CSng(GetDelimitedToken(txt, ","))
    Point2.Coord(2) = Point1.Coord(2) + CSng(GetDelimitedToken(txt, ","))
    Point2.Coord(3) = Point1.Coord(3) + CSng(GetDelimitedToken(txt, ","))
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

DiskParamError:
    MsgBox "Error initializing disk parameters."
End Sub

' Make a wire frame.
Private Sub MakeWireFrame()
Dim i As Integer
Dim X As Single
Dim Y As Single
Dim Z As Single
Dim v1x As Single
Dim v1y As Single
Dim v1z As Single
Dim v2x As Single
Dim v2y As Single
Dim v2z As Single
Dim theta As Single
Dim dtheta As Single
Dim v1_scale As Single
Dim v2_scale As Single

    GetLineNormals _
        Point2.Coord(1) - Point1.Coord(1), _
        Point2.Coord(2) - Point1.Coord(2), _
        Point2.Coord(3) - Point1.Coord(3), _
        v1x, v1y, v1z, v2x, v2y, v2z

    ' Make a polygon around the edge of the disk.
    theta = 0
    dtheta = 2 * PI / WIRE_POINTS
    For i = 1 To WIRE_POINTS
        v1_scale = Sin(theta) * Radius
        v2_scale = Cos(theta) * Radius
        With WireFrame(i)
            .Coord(1) = Point1.Coord(1) + v2_scale * v1x + v1_scale * v2x
            .Coord(2) = Point1.Coord(2) + v2_scale * v1y + v1_scale * v2y
            .Coord(3) = Point1.Coord(3) + v2_scale * v1z + v1_scale * v2z
            .Coord(4) = 1#
        End With
        theta = theta + dtheta
    Next i
End Sub
' Apply a transformation matrix to the object.
Public Sub RayTraceable_Apply(M() As Single)
Dim i As Integer

    ' Transform the wire frame.
    For i = 1 To WIRE_POINTS
        m3Apply WireFrame(i).Coord, _
             M, WireFrame(i).Trans
    Next i

    ' Transform the plane's points.
    m3Apply Point1.Coord, M, Point1.Trans
    m3Apply Point2.Coord, M, Point2.Trans
End Sub
' Apply a transformation matrix to the object.
Public Sub RayTraceable_ApplyFull(M() As Single)
Dim i As Integer

    ' Transform the wire frame.
    For i = 1 To WIRE_POINTS
        m3ApplyFull WireFrame(i).Coord, _
                 M, WireFrame(i).Trans
    Next i

    ' Transform the plane's points.
    m3ApplyFull Point1.Coord, M, Point1.Trans
    m3ApplyFull Point2.Coord, M, Point2.Trans
End Sub

' Draw the object with backfaces removed.
' Draw the whole wire frame for planes.
Public Sub RayTraceable_DrawBackfacesRemoved(ByVal pic As PictureBox)
    RayTraceable_DrawWireFrame pic
End Sub
' Return the red, green, and blue components of
' the surface at the hit position.
Public Sub RayTraceable_FindHitColor(ByVal depth As Integer, Objects As Collection, ByVal eye_x As Single, ByVal eye_y As Single, ByVal eye_z As Single, ByVal px As Single, ByVal py As Single, ByVal pz As Single, ByRef R As Integer, ByRef G As Integer, ByRef B As Integer)
Dim Nx As Single
Dim Ny As Single
Dim Nz As Single
Dim Vx As Single
Dim Vy As Single
Dim Vz As Single
Dim NdotV As Single

    ' Find the unit normal at this point.
    GetUnitNormal Nx, Ny, Nz

    ' Make sure the normal points towards the
    ' center of projection.
    Vx = EyeX - px
    Vy = EyeY - py
    Vz = EyeZ - pz
    NdotV = Nx * Vx + Ny * Vy + Nz * Vz
    If NdotV < 0 Then
        Nx = -Nx
        Ny = -Ny
        Nz = -Nz
    End If

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
' point and normal intersects this object.
'
' Just see if the plane passes within distance
' Radius of the center.
Public Sub RayTraceable_CullScanline(ByVal px As Single, ByVal py As Single, ByVal pz As Single, ByVal Nx As Single, ByVal Ny As Single, ByVal Nz As Single)
Dim dx As Single
Dim dy As Single
Dim dz As Single
Dim dist As Single

    ' See if we will ever be visible again.
    If ForeverCulled Then
        DoneOnThisScanline = True
        Exit Sub
    End If

    ' We have not yet had a hit on this scanline.
    HadHit = False

    ' Find the distance from the center of the
    ' disk to the scanline plane.

    ' Get the vector from our center to the point.
    With Point1
        dx = .Trans(1) - px
        dy = .Trans(2) - py
        dz = .Trans(3) - pz
    End With

    ' Take the dot product of this and the normal.
    ' If the resulting distance > Radius, cull.
    DoneOnThisScanline = (Abs(dx * Nx + dy * Ny + dz * Nz) > Radius)

    ' See if we will be culled in the future.
    If DoneOnThisScanline Then
        ' We were not culled on a previous scanline
        ' but we are now. We will be culled on
        ' all later scanlines.
        If HadHitOnPreviousScanline Then ForeverCulled = True
    Else
        ' We are not culled. Remember that.
        HadHitOnPreviousScanline = True
    End If
End Sub
' Return the value T for the point of intersection
' between the vector from point (px, py, pz) in
' the direction <vx, vy, vz>.
'
' direct_calculation is true if we are finding the
' intersection from a viewing position ray. It is
' false if we are finding an reflected intersection
' or a shadow feeler.
Public Function RayTraceable_FindT(ByVal direct_calculation As Boolean, ByVal px As Single, ByVal py As Single, ByVal pz As Single, ByVal Vx As Single, ByVal Vy As Single, ByVal Vz As Single) As Single
Dim A As Single
Dim B As Single
Dim C As Single
Dim D As Single
Dim Nx As Single
Dim Ny As Single
Dim Nz As Single
Dim denom As Single
Dim t As Single
Dim Cx As Single
Dim Cy As Single
Dim Cz As Single
Dim dx As Single
Dim dy As Single
Dim dz As Single
Dim X As Single
Dim Y As Single
Dim Z As Single

    ' See if we have been culled.
    If direct_calculation And DoneOnThisScanline Then
        RayTraceable_FindT = -1
        Exit Function
    End If

    ' Find the unit normal at this point.
    GetUnitNormal Nx, Ny, Nz

    ' Compute the plane's parameters.
    A = Nx
    B = Ny
    C = Nz
    D = -(Nx * Point1.Trans(1) + _
          Ny * Point1.Trans(2) + _
          Nz * Point1.Trans(3))

    ' If the denominator = 0, the ray is parallel
    ' to the plane so there's no intersection.
    denom = A * Vx + B * Vy + C * Vz
    If denom = 0 Then
        RayTraceable_FindT = -1
        Exit Function
    End If

    ' Solve for t.
    t = -(A * px + B * py + C * pz + D) / denom

    ' If there is no positive t value, there's no
    ' intersection in this direction.
    If t < 0.01 Then
        RayTraceable_FindT = -1
        Exit Function
    End If

    ' Get the coordinates of the disk's center.
    Cx = Point1.Trans(1)
    Cy = Point1.Trans(2)
    Cz = Point1.Trans(3)

    ' Get the point of intersection with the plane.
    X = px + t * Vx
    Y = py + t * Vy
    Z = pz + t * Vz

    ' See if the point is within distance
    ' Radius of the center.
    dx = Cx - X
    dy = Cy - Y
    dz = Cz - Z
    If dx * dx + dy * dy + dz * dz > Radius * Radius Then
        ' We are not within distance Radius.
        RayTraceable_FindT = -1
        Exit Function
    End If

    ' We had a hit.
    If direct_calculation Then HadHit = True

    RayTraceable_FindT = t
End Function
' Return the minimum and maximum distances from
' this point.
' Use the wireframe points.
Private Sub RayTraceable_GetRminRmax(new_min As Single, new_max As Single, ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
Dim i As Integer
Dim dx As Single
Dim dy As Single
Dim dz As Single
Dim dist As Single

    new_min = 1E+30
    new_max = -1E+30

    For i = 1 To WIRE_POINTS
        With WireFrame(i)
            dx = X - .Trans(1)
            dy = Y - .Trans(2)
            dz = Z - .Trans(3)
        End With
        dist = Sqr(dx * dx + dy * dy + dz * dz)
        If new_min > dist Then new_min = dist
        If new_max < dist Then new_max = dist
    Next i
End Sub
' Reset the ForeverCulled flag.
Private Sub RayTraceable_ResetCulling()
    ForeverCulled = False
    HadHitOnPreviousScanline = False
End Sub


