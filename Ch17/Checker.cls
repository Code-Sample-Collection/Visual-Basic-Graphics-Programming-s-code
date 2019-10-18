VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RayCheckerboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A checkerboard in a plane.

Implements RayTraceable

Private Point1 As Point3D       ' Point on corner.
Private Point2 As Point3D       ' Point on 1st corner of first rectangle.
Private Point3 As Point3D       ' Point on 2nd corner of first rectangle.
Private NumSquares1 As Integer  ' # squares in 1st direction.
Private NumSquares2 As Integer  ' # squares in 2nd direction.

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
Dim v1x As Single
Dim v1y As Single
Dim v1z As Single
Dim v2x As Single
Dim v2y As Single
Dim v2z As Single
Dim n_len As Single

    ' Get the square vectors.
    v1x = Point2.Trans(1) - Point1.Trans(1)
    v1y = Point2.Trans(2) - Point1.Trans(2)
    v1z = Point2.Trans(3) - Point1.Trans(3)
    v2x = Point3.Trans(1) - Point1.Trans(1)
    v2y = Point3.Trans(2) - Point1.Trans(2)
    v2z = Point3.Trans(3) - Point1.Trans(3)

    ' Calculate the normal.
    m3Cross Nx, Ny, Nz, v1x, v1y, v1z, v2x, v2y, v2z
    n_len = Sqr(Nx * Nx + Ny * Ny + Nz * Nz)
    Nx = Nx / n_len
    Ny = Ny / n_len
    Nz = Nz / n_len
End Sub
' Add non-backface polygons to this collection.
Public Sub RayTraceable_GetPolygons(ByRef num_polygons As Integer, polygons() As SimplePolygon, ByVal shaded As Boolean)
Dim pgon As SimplePolygon
Dim px As Single
Dim py As Single
Dim pz As Single
Dim v1x As Single
Dim v1y As Single
Dim v1z As Single
Dim v2x As Single
Dim v2y As Single
Dim v2z As Single
Dim i As Integer
Dim j As Integer

    ' Get the square vectors.
    px = Point1.Trans(1)
    py = Point1.Trans(2)
    pz = Point1.Trans(3)
    v1x = Point2.Trans(1) - px
    v1y = Point2.Trans(2) - py
    v1z = Point2.Trans(3) - pz
    v2x = Point3.Trans(1) - px
    v2y = Point3.Trans(2) - py
    v2z = Point3.Trans(3) - pz

    ' Make the squares.
    For i = 0 To NumSquares1 - 1
        For j = 0 To NumSquares2 - 1
            If (i + j) Mod 2 = 0 Then
                ' Make a polygon.
                Set pgon = New SimplePolygon
                pgon.AddPoint px + i * v1x + j * v2x, py + i * v1y + j * v2y, pz + i * v1z + j * v2z
                pgon.AddPoint px + (i + 1) * v1x + j * v2x, py + (i + 1) * v1y + j * v2y, pz + (i + 1) * v1z + j * v2z
                pgon.AddPoint px + (i + 1) * v1x + (j + 1) * v2x, py + (i + 1) * v1y + (j + 1) * v2y, pz + (i + 1) * v1z + (j + 1) * v2z
                pgon.AddPoint px + i * v1x + (j + 1) * v2x, py + i * v1y + (j + 1) * v2y, pz + i * v1z + (j + 1) * v2z

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
            End If
        Next j
    Next i
End Sub
' Draw a wireframe for this object.
Public Sub RayTraceable_DrawWireFrame(ByVal pic As PictureBox)
Dim px As Single
Dim py As Single
Dim v1x As Single
Dim v1y As Single
Dim v2x As Single
Dim v2y As Single
Dim i As Integer
Dim j As Integer

    ' Use an appropriate color.
    pic.ForeColor = GetColor()

    ' Get the square vectors.
    px = Point1.Trans(1)
    py = Point1.Trans(2)
    v1x = Point2.Trans(1) - px
    v1y = Point2.Trans(2) - py
    v2x = Point3.Trans(1) - px
    v2y = Point3.Trans(2) - py

    ' Draw the squares.
    For i = 0 To NumSquares1 - 1
        For j = 0 To NumSquares2 - 1
            If (i + j) Mod 2 = 0 Then
                pic.Line (px + i * v1x + j * v2x, py + i * v1y + j * v2y)-(px + (i + 1) * v1x + j * v2x, py + (i + 1) * v1y + j * v2y)
                pic.Line -(px + (i + 1) * v1x + (j + 1) * v2x, py + (i + 1) * v1y + (j + 1) * v2y)
                pic.Line -(px + i * v1x + (j + 1) * v2x, py + i * v1y + (j + 1) * v2y)
                pic.Line -(px + i * v1x + j * v2x, py + i * v1y + j * v2y)
            End If
        Next j
    Next i
End Sub

' Initialize the object using text parameters in
' a comma-delimited list.
Public Sub SetParameters(ByVal txt As String)
    On Error GoTo CheckerboardParamError

    ' Read the parameters and initialize the object.
    ' Geometry.
    NumSquares1 = CInt(GetDelimitedToken(txt, ","))
    NumSquares2 = CInt(GetDelimitedToken(txt, ","))
    Point1.Coord(1) = CSng(GetDelimitedToken(txt, ","))
    Point1.Coord(2) = CSng(GetDelimitedToken(txt, ","))
    Point1.Coord(3) = CSng(GetDelimitedToken(txt, ","))
    Point1.Coord(4) = 1
    Point2.Coord(1) = CSng(GetDelimitedToken(txt, ","))
    Point2.Coord(2) = CSng(GetDelimitedToken(txt, ","))
    Point2.Coord(3) = CSng(GetDelimitedToken(txt, ","))
    Point2.Coord(4) = 1
    Point3.Coord(1) = CSng(GetDelimitedToken(txt, ","))
    Point3.Coord(2) = CSng(GetDelimitedToken(txt, ","))
    Point3.Coord(3) = CSng(GetDelimitedToken(txt, ","))
    Point3.Coord(4) = 1

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

    Exit Sub

CheckerboardParamError:
    MsgBox "Error initializing checkerboard parameters."
End Sub

' Apply a transformation matrix to the object.
Public Sub RayTraceable_Apply(M() As Single)
    ' Transform the points.
    m3Apply Point1.Coord, M, Point1.Trans
    m3Apply Point2.Coord, M, Point2.Trans
    m3Apply Point3.Coord, M, Point3.Trans
End Sub
' Apply a transformation matrix to the object.
Public Sub RayTraceable_ApplyFull(M() As Single)
    ' Transform the points.
    m3ApplyFull Point1.Coord, M, Point1.Trans
    m3ApplyFull Point2.Coord, M, Point2.Trans
    m3ApplyFull Point3.Coord, M, Point3.Trans
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
    ' Do not scanline cull.
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
Dim Cx As Single
Dim Cy As Single
Dim Cz As Single
Dim dx As Single
Dim dy As Single
Dim dz As Single
Dim X As Single
Dim Y As Single
Dim Z As Single
Dim v1x As Single
Dim v1y As Single
Dim v1z As Single
Dim v2x As Single
Dim v2y As Single
Dim v2z As Single
Dim i As Single
Dim j As Single

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

    ' Get the point of intersection with the plane.
    X = px + t * Vx
    Y = py + t * Vy
    Z = pz + t * Vz

    ' Get the square vectors.
    px = Point1.Trans(1)
    py = Point1.Trans(2)
    pz = Point1.Trans(3)
    v1x = Point2.Trans(1) - px
    v1y = Point2.Trans(2) - py
    v1z = Point2.Trans(3) - pz
    v2x = Point3.Trans(1) - px
    v2y = Point3.Trans(2) - py
    v2z = Point3.Trans(3) - pz

    ' Get the i and j values for this point.
    If (Abs(v1x) > 0.001) And (Abs(v1y * v2x - v2y * v1x) > 0.001) Then
        j = (v1y * (X - px) + v1x * (py - Y)) / (v1y * v2x - v2y * v1x)
        i = (X - px - v2x * j) / v1x
    ElseIf (Abs(v1y) > 0.001) And (Abs(v1z * v2y - v2z * v1y) > 0.001) Then
        j = (v1z * (Y - py) + v1y * (pz - Z)) / (v1z * v2y - v2z * v1y)
        i = (Y - py - v2y * j) / v1y
    Else
        j = (v1x * (Z - pz) + v1z * (px - X)) / (v1x * v2z - v2x * v1z)
        i = (Z - pz - v2z * j) / v1z
    End If

    ' See if the point is ok.
    If (i < 0) Or (j < 0) Or (i > NumSquares1) Or (j > NumSquares2) Then
        ' Not on the area of interest.
        RayTraceable_FindT = -1
        Exit Function
    ElseIf (Int(i) + Int(j)) Mod 2 <> 0 Then
        ' Not on a drawn square.
        RayTraceable_FindT = -1
        Exit Function
    Else
        ' We had a hit.
        If direct_calculation Then HadHit = True

        RayTraceable_FindT = t
    End If
End Function
' Return the minimum and maximum distances from
' this point.
' Use the wireframe points.
Private Sub RayTraceable_GetRminRmax(new_min As Single, new_max As Single, ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
Dim dx As Single
Dim dy As Single
Dim dz As Single
Dim px As Single
Dim py As Single
Dim pz As Single
Dim v1x As Single
Dim v1y As Single
Dim v1z As Single
Dim v2x As Single
Dim v2y As Single
Dim v2z As Single
Dim i As Integer
Dim j As Integer
Dim dist As Single

    new_min = 1E+30
    new_max = -1E+30

    ' Get the square vectors.
    px = Point1.Trans(1)
    py = Point1.Trans(2)
    pz = Point1.Trans(3)
    v1x = Point2.Trans(1) - px
    v1y = Point2.Trans(2) - py
    v1z = Point2.Trans(3) - pz
    v2x = Point3.Trans(1) - px
    v2y = Point3.Trans(2) - py
    v2z = Point3.Trans(3) - pz

    For i = 0 To NumSquares1 Step NumSquares1
        For j = 0 To NumSquares2 Step NumSquares2
            dx = X - (px + i * v1x + j * v2x)
            dy = Y - (py + i * v1y + j * v2y)
            dz = Z - (pz + i * v1z + j * v2z)
            dist = Sqr(dx * dx + dy * dy + dz * dz)
            If new_min > dist Then new_min = dist
            If new_max < dist Then new_max = dist
        Next j
    Next i
End Sub
' Reset the ForeverCulled flag.
Private Sub RayTraceable_ResetCulling()
    ForeverCulled = False
    HadHitOnPreviousScanline = False
End Sub


