VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RayBumpySphere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A sphere object with bump texturing.

Implements RayTraceable

' Geometry.
Private Radius As Single
Private Center As Point3D

Private Const NUM_THETA = 10
Private Const NUM_PHI = 10
Private WireFrame(1 To NUM_THETA, 1 To NUM_PHI) As Point3D

' Bumpiness.
Private Bumpiness As Single

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
Dim P As Integer
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
        For P = 1 To NUM_PHI - 1
            ' Make a polygon.
            Set pgon = New SimplePolygon
            With WireFrame(last_t, P)
                pgon.AddPoint .Trans(1), .Trans(2), .Trans(3)
            End With
            With WireFrame(t, P)
                pgon.AddPoint .Trans(1), .Trans(2), .Trans(3)
            End With
            With WireFrame(t, P + 1)
                pgon.AddPoint .Trans(1), .Trans(2), .Trans(3)
            End With
            With WireFrame(last_t, P + 1)
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
        Next P
        last_t = t
    Next t
End Sub
' Make a wire frame.
Private Sub MakeWireFrame()
Const PI = 3.14159265

Dim i_theta As Integer
Dim i_phi As Integer
Dim theta As Single
Dim phi As Single
Dim dtheta As Single
Dim dphi As Single
Dim X As Single
Dim Y As Single
Dim Z As Single
Dim rad As Single

    dtheta = 2 * PI / NUM_THETA
    dphi = PI / (NUM_PHI - 1)
    theta = 0
    For i_theta = 1 To NUM_THETA
        phi = -PI / 2
        For i_phi = 1 To NUM_PHI
            Z = Center.Coord(3) + Radius * Sin(phi)
            rad = Radius * Cos(phi)
            X = Center.Coord(1) + rad * Cos(theta)
            Y = Center.Coord(2) + rad * Sin(theta)

            WireFrame(i_theta, i_phi).Coord(1) = X
            WireFrame(i_theta, i_phi).Coord(2) = Y
            WireFrame(i_theta, i_phi).Coord(3) = Z
            WireFrame(i_theta, i_phi).Coord(4) = 1

            phi = phi + dphi
        Next i_phi
        theta = theta + dtheta
    Next i_theta
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
    On Error GoTo SphereParamError

    ' Read the parameters and initialize the object.
    ' Geometry.
    Radius = CSng(GetDelimitedToken(txt, ","))
    Center.Coord(1) = CSng(GetDelimitedToken(txt, ","))
    Center.Coord(2) = CSng(GetDelimitedToken(txt, ","))
    Center.Coord(3) = CSng(GetDelimitedToken(txt, ","))
    Center.Coord(4) = 1

    ' Bumpiness.
    Bumpiness = CSng(GetDelimitedToken(txt, ","))

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

SphereParamError:
    MsgBox "Error initializing sphere parameters."
End Sub

' Draw a wireframe for this object.
Public Sub RayTraceable_DrawWireFrame(ByVal pic As PictureBox)
Dim t As Integer
Dim P As Integer
Dim last_t As Integer

    ' Use an appropriate color.
    pic.ForeColor = GetColor()

    last_t = NUM_THETA
    For t = 1 To NUM_THETA
        For P = 1 To NUM_PHI
            With WireFrame(last_t, P)
                pic.CurrentX = .Trans(1)
                pic.CurrentY = .Trans(2)
            End With
            With WireFrame(t, P)
                pic.Line -(.Trans(1), .Trans(2))
            End With
            If P < NUM_PHI Then
                With WireFrame(t, P + 1)
                    pic.Line -(.Trans(1), .Trans(2))
                End With
            End If
        Next P
        last_t = t
    Next t
End Sub
' Draw the object with backfaces removed.
Public Sub RayTraceable_DrawBackfacesRemoved(ByVal pic As PictureBox)
Dim t As Integer
Dim P As Integer
Dim last_t As Integer
Dim X(1 To 4) As Single
Dim Y(1 To 4) As Single
Dim Z(1 To 4) As Single

    ' Use an appropriate color.
    pic.ForeColor = GetColor()

    last_t = NUM_THETA
    For t = 1 To NUM_THETA
        For P = 1 To NUM_PHI - 1
            With WireFrame(last_t, P)
                X(1) = .Trans(1)
                Y(1) = .Trans(2)
                Z(1) = .Trans(3)
            End With
            With WireFrame(t, P)
                X(2) = .Trans(1)
                Y(2) = .Trans(2)
                Z(2) = .Trans(3)
            End With
            With WireFrame(t, P + 1)
                X(3) = .Trans(1)
                Y(3) = .Trans(2)
                Z(3) = .Trans(3)
            End With
            With WireFrame(last_t, P + 1)
                X(4) = .Trans(1)
                Y(4) = .Trans(2)
                Z(4) = .Trans(3)
            End With

            DrawFace pic, X, Y, Z
        Next P
        last_t = t
    Next t
End Sub
' Apply a transformation matrix to the object.
Public Sub RayTraceable_Apply(M() As Single)
Dim i_theta As Integer
Dim i_phi As Integer

    ' Transform the wire frame.
    For i_theta = 1 To NUM_THETA
        For i_phi = 1 To NUM_PHI
            m3Apply WireFrame(i_theta, i_phi).Coord, _
                 M, WireFrame(i_theta, i_phi).Trans
        Next i_phi
    Next i_theta

    ' Transform the center.
    m3Apply Center.Coord, M, Center.Trans
End Sub
' Apply a transformation matrix to the object.
Public Sub RayTraceable_ApplyFull(M() As Single)
Dim i_theta As Integer
Dim i_phi As Integer

    ' Transform the wire frame.
    For i_theta = 1 To NUM_THETA
        For i_phi = 1 To NUM_PHI
            m3ApplyFull WireFrame(i_theta, i_phi).Coord, _
                     M, WireFrame(i_theta, i_phi).Trans
        Next i_phi
    Next i_theta

    ' Transform the center.
    m3ApplyFull Center.Coord, M, Center.Trans
End Sub

' Return the red, green, and blue components of
' the surface at the hit position.
Public Sub RayTraceable_FindHitColor( _
    ByVal depth As Integer, Objects As Collection, _
    ByVal eye_x As Single, ByVal eye_y As Single, ByVal eye_z As Single, _
    ByVal px As Single, ByVal py As Single, ByVal pz As Single, _
    ByRef R As Integer, ByRef G As Integer, ByRef B As Integer)
Dim Nx As Single
Dim Ny As Single
Dim Nz As Single
Dim n_len  As Single

    ' Find the unit normal at this point.
    Nx = px - Center.Trans(1)
    Ny = py - Center.Trans(2)
    Nz = pz - Center.Trans(3)
    n_len = Sqr(Nx * Nx + Ny * Ny + Nz * Nz)
    Nx = Nx / n_len
    Ny = Ny / n_len
    Nz = Nz / n_len

    ' Randommize the normal a little bit and
    ' renormalize.
    Nx = Nx + Rnd * Bumpiness
    Ny = Ny + Rnd * Bumpiness
    Nz = Nz + Rnd * Bumpiness
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
    ' sphere to the scanline plane.

    ' Get the vector from our center to the point.
    With Center
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
Dim Cx As Single
Dim Cy As Single
Dim Cz As Single
Dim B24AC As Single
Dim t1 As Single
Dim t2 As Single
Dim dx As Single
Dim dy As Single
Dim dz As Single

    ' See if we have been culled.
    If direct_calculation And DoneOnThisScanline Then
        RayTraceable_FindT = -1
        Exit Function
    End If

    Cx = Center.Trans(1)
    Cy = Center.Trans(2)
    Cz = Center.Trans(3)

    ' Get the coefficients for the quadratic.
    A = Vx * Vx + Vy * Vy + Vz * Vz
    B = 2 * Vx * (px - Cx) + _
        2 * Vy * (py - Cy) + _
        2 * Vz * (pz - Cz)
    C = Cx * Cx + Cy * Cy + Cz * Cz + _
        px * px + py * py + pz * pz - _
        2 * (Cx * px + Cy * py + Cz * pz) - _
        Radius * Radius

    ' Solve the quadratic A*t^2 + B*t + C = 0.
    B24AC = B * B - 4 * A * C
    If B24AC < 0 Then
        ' There is no real intersection.
        RayTraceable_FindT = -1

        ' If we had a hit before on this scanline
        ' but we don't have one now. We are done
        ' for this scanline.
        If HadHit And direct_calculation Then DoneOnThisScanline = True

        Exit Function
    ElseIf B24AC = 0 Then
        ' There is one intersection.
        t1 = -B / 2 / A
    Else
        ' There are two intersections.
        B24AC = Sqr(B24AC)
        t1 = (-B + B24AC) / 2 / A
        t2 = (-B - B24AC) / 2 / A
        ' Use only positive t values.
        If t1 < 0.01 Then t1 = t2
        If t2 < 0.01 Then t2 = t1
        ' Use the smaller t value.
        If t1 > t2 Then t1 = t2
    End If

    ' If there is no positive t value, there's no
    ' intersection in this direction.
    If t1 < 0.01 Then
        RayTraceable_FindT = -1

        ' If we had a hit before on this scanline
        ' but we don't have one now. We are done
        ' for this scanline.
        If HadHit And direct_calculation Then DoneOnThisScanline = True

        Exit Function
    End If

    ' We had a hit.
    If direct_calculation Then HadHit = True

    RayTraceable_FindT = t1
End Function

' Return the minimum and maximum distances from
' this point.
Private Sub RayTraceable_GetRminRmax(new_min As Single, new_max As Single, ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
Dim dx As Single
Dim dy As Single
Dim dz As Single
Dim dist As Single

    dx = X - Center.Trans(1)
    dy = Y - Center.Trans(2)
    dz = Z - Center.Trans(3)
    dist = Sqr(dx * dx + dy * dy + dz * dz)
    new_max = dist + Radius
    new_min = dist - Radius
    If new_min < 0 Then new_min = 0
End Sub
' Reset the ForeverCulled flag.
Private Sub RayTraceable_ResetCulling()
    ForeverCulled = False
    HadHitOnPreviousScanline = False
End Sub


