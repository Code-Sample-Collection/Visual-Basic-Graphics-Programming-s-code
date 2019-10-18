VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RayPlane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A plane object.

Implements RayTraceable

Private Point1 As Point3D   ' Point on plane.
Private Point2 As Point3D   ' Normal = P2 - P1.

' Wire frame variables.
Private Const WIRE_MAX = 3
Private Const WIRE_DX = 50
Private WireFrame(-WIRE_MAX To WIRE_MAX, -WIRE_MAX To WIRE_MAX) As Point3D

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
Dim j As Integer
Dim pgon As SimplePolygon
Dim color As Long

    ' If all polygons are the same color,
    ' get an appropriate color.
    If Not shaded Then
        color = GetColor()
    End If

    For i = -WIRE_MAX To WIRE_MAX - 1
        For j = -WIRE_MAX To WIRE_MAX - 1
            ' Make a polygon.
            Set pgon = New SimplePolygon
            With WireFrame(i, j)
                pgon.AddPoint .Trans(1), .Trans(2), .Trans(3)
            End With
            With WireFrame(i, j + 1)
                pgon.AddPoint .Trans(1), .Trans(2), .Trans(3)
            End With
            With WireFrame(i + 1, j + 1)
                pgon.AddPoint .Trans(1), .Trans(2), .Trans(3)
            End With
            With WireFrame(i + 1, j)
                pgon.AddPoint .Trans(1), .Trans(2), .Trans(3)
            End With
            pgon.Finish

            ' Add it to the list.
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
        Next j
    Next i
End Sub
' Draw a wireframe for this object.
Public Sub RayTraceable_DrawWireFrame(ByVal pic As PictureBox)
Dim i As Integer
Dim j As Integer

    ' Use an appropriate color.
    pic.ForeColor = GetColor()

    ' Draw i lines.
    For i = -WIRE_MAX To WIRE_MAX
        With WireFrame(i, -WIRE_MAX)
            pic.CurrentX = .Trans(1)
            pic.CurrentY = .Trans(2)
        End With

        For j = -WIRE_MAX + 1 To WIRE_MAX
            With WireFrame(i, j)
                pic.Line -(.Trans(1), .Trans(2))
            End With
        Next j
    Next i

    ' Draw j lines.
    For j = -WIRE_MAX To WIRE_MAX
        With WireFrame(-WIRE_MAX, j)
            pic.CurrentX = .Trans(1)
            pic.CurrentY = .Trans(2)
        End With

        For i = -WIRE_MAX + 1 To WIRE_MAX
            With WireFrame(i, j)
                pic.Line -(.Trans(1), .Trans(2))
            End With
        Next i
    Next j
End Sub

' Initialize the object using text parameters in
' a comma-delimited list.
Public Sub SetParameters(ByVal txt As String)
    On Error GoTo PlaneParamError

    ' Read the parameters and initialize the object.
    ' Geometry.
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

PlaneParamError:
    MsgBox "Error initializing plane parameters."
End Sub

' Make a wire frame.
Private Sub MakeWireFrame()
Dim i As Integer
Dim j As Integer
Dim X As Single
Dim Y As Single
Dim Z As Single
Dim Nx As Single
Dim Ny As Single
Dim Nz As Single
Dim v1x As Single
Dim v1y As Single
Dim v1z As Single
Dim v2x As Single
Dim v2y As Single
Dim v2z As Single
Dim length As Single

    ' Get two vectors in the plane.
    Nx = Point2.Coord(1) - Point1.Coord(1)
    Ny = Point2.Coord(2) - Point1.Coord(2)
    Nz = Point2.Coord(3) - Point1.Coord(3)
    If Nx <> 0 Then
        m3Cross v1x, v1y, v1z, Nx, Ny, Nz, 0, 1, 0
    ElseIf Ny <> 0 Then
        m3Cross v1x, v1y, v1z, Nx, Ny, Nz, 0, 0, 1
    Else
        m3Cross v1x, v1y, v1z, Nx, Ny, Nz, 1, 0, 0
    End If
    m3Cross v2x, v2y, v2z, Nx, Ny, Nz, v1x, v1y, v1z

    ' Normalize the vectors.
    length = Sqr(v1x * v1x + v1y * v1y + v1z * v1z)
    v1x = v1x / length
    v1y = v1y / length
    v1z = v1z / length

    length = Sqr(v2x * v2x + v2y * v2y + v2z * v2z)
    v2x = v2x / length
    v2y = v2y / length
    v2z = v2z / length

    ' Make some rectangles.
    For i = -WIRE_MAX To WIRE_MAX
        For j = -WIRE_MAX To WIRE_MAX
            WireFrame(i, j).Coord(1) = Point1.Coord(1) + i * WIRE_DX * v1x + j * WIRE_DX * v2x
            WireFrame(i, j).Coord(2) = Point1.Coord(2) + i * WIRE_DX * v1y + j * WIRE_DX * v2y
            WireFrame(i, j).Coord(3) = Point1.Coord(3) + i * WIRE_DX * v1z + j * WIRE_DX * v2z
            WireFrame(i, j).Coord(4) = 1
        Next j
    Next i
End Sub
' Apply a transformation matrix to the object.
Public Sub RayTraceable_Apply(M() As Single)
Dim i As Integer
Dim j As Integer

    ' Transform the wire frame.
    For i = -WIRE_MAX To WIRE_MAX
        For j = -WIRE_MAX To WIRE_MAX
            m3Apply WireFrame(i, j).Coord, _
                 M, WireFrame(i, j).Trans
        Next j
    Next i

    ' Transform the plane's points.
    m3Apply Point1.Coord, M, Point1.Trans
    m3Apply Point2.Coord, M, Point2.Trans
End Sub
' Apply a transformation matrix to the object.
Public Sub RayTraceable_ApplyFull(M() As Single)
Dim i As Integer
Dim j As Integer

    ' Transform the wire frame.
    For i = -WIRE_MAX To WIRE_MAX
        For j = -WIRE_MAX To WIRE_MAX
            m3ApplyFull WireFrame(i, j).Coord, _
                     M, WireFrame(i, j).Trans
        Next j
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
Public Sub RayTraceable_CullScanline(ByVal px As Single, ByVal py As Single, ByVal pz As Single, ByVal Nx As Single, ByVal Ny As Single, ByVal Nz As Single)
    ' Do not scanline cull planes.
    DoneOnThisScanline = False
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

    ' Do not cull planes.

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

    ' We had a hit.
    If direct_calculation Then HadHit = True

    RayTraceable_FindT = t
End Function
' Return the minimum and maximum distances from
' this point.
'
' Because planes extend infinitely, this is not
' terribly useful for this object. Instead return
' the point used to define the plane.
Private Sub RayTraceable_GetRminRmax(new_min As Single, new_max As Single, ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
Dim dx As Single
Dim dy As Single
Dim dz As Single
Dim dist As Single

    dx = X - Point1.Trans(1)
    dy = Y - Point1.Trans(2)
    dz = Z - Point1.Trans(3)
    dist = Sqr(dx * dx + dy * dy + dz * dz)
    new_max = dist
    new_min = dist
End Sub

' Reset the ForeverCulled flag.
Private Sub RayTraceable_ResetCulling()
    ForeverCulled = False
    HadHitOnPreviousScanline = False
End Sub


