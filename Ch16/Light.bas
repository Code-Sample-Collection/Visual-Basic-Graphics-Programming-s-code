Attribute VB_Name = "Light"
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Const ALTERNATE = 1
Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
' Draw a Gouraud shaded quadrilateral.
Public Sub GouraudQuadrilateral(ByVal pic As PictureBox, _
    ByVal light_sources As Collection, _
    ByVal ambient_light As Integer, _
    ByVal eye_x As Single, ByVal eye_y As Single, ByVal eye_z As Single, _
    ByVal DiffuseKr As Single, ByVal DiffuseKg As Single, ByVal DiffuseKb As Single, _
    ByVal AmbientKr As Single, ByVal AmbientKg As Single, ByVal AmbientKb As Single, _
    ByVal SpecularK As Single, ByVal SpecularN As Integer, _
    ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, _
    ByVal x2 As Single, ByVal y2 As Single, ByVal z2 As Single, _
    ByVal x3 As Single, ByVal y3 As Single, ByVal z3 As Single, _
    ByVal x4 As Single, ByVal y4 As Single, ByVal z4 As Single, _
    ByVal Nx1 As Single, ByVal Ny1 As Single, ByVal Nz1 As Single, _
    ByVal Nx2 As Single, ByVal Ny2 As Single, ByVal Nz2 As Single, _
    ByVal Nx3 As Single, ByVal Ny3 As Single, ByVal Nz3 As Single, _
    ByVal Nx4 As Single, ByVal Ny4 As Single, ByVal Nz4 As Single, _
    ByVal Tx1 As Single, ByVal Ty1 As Single, _
    ByVal Tx2 As Single, ByVal Ty2 As Single, _
    ByVal Tx3 As Single, ByVal Ty3 As Single, _
    ByVal Tx4 As Single, ByVal Ty4 As Single _
    )
Dim pts(1 To 4) As POINTAPI
Dim vertex_r(1 To 4) As Integer
Dim vertex_g(1 To 4) As Integer
Dim vertex_b(1 To 4) As Integer
Dim i As Integer
Dim hRgn As Long
Dim R As RECT
Dim X As Long
Dim Y As Long
Dim clr As Long

    ' Calculate the colors at the corners.
    clr = CalculateSurfaceColor( _
        x1, y1, z1, Nx1, Ny1, Nz1, _
        light_sources, ambient_light, _
        eye_x, eye_y, eye_z, _
        DiffuseKr, DiffuseKg, DiffuseKb, _
        AmbientKr, AmbientKg, AmbientKb, _
        SpecularK, SpecularN)
    vertex_r(1) = clr And &HFF&
    vertex_g(1) = (clr And &HFF00&) \ &H100&
    vertex_b(1) = (clr And &HFF0000) \ &H10000

    clr = CalculateSurfaceColor( _
        x2, y2, z2, Nx2, Ny2, Nz2, _
        light_sources, ambient_light, _
        eye_x, eye_y, eye_z, _
        DiffuseKr, DiffuseKg, DiffuseKb, _
        AmbientKr, AmbientKg, AmbientKb, _
        SpecularK, SpecularN)
    vertex_r(2) = clr And &HFF&
    vertex_g(2) = (clr And &HFF00&) \ &H100&
    vertex_b(2) = (clr And &HFF0000) \ &H10000

    clr = CalculateSurfaceColor( _
        x3, y3, z3, Nx3, Ny3, Nz3, _
        light_sources, ambient_light, _
        eye_x, eye_y, eye_z, _
        DiffuseKr, DiffuseKg, DiffuseKb, _
        AmbientKr, AmbientKg, AmbientKb, _
        SpecularK, SpecularN)
    vertex_r(3) = clr And &HFF&
    vertex_g(3) = (clr And &HFF00&) \ &H100&
    vertex_b(3) = (clr And &HFF0000) \ &H10000

    clr = CalculateSurfaceColor( _
        x4, y4, z4, Nx4, Ny4, Nz4, _
        light_sources, ambient_light, _
        eye_x, eye_y, eye_z, _
        DiffuseKr, DiffuseKg, DiffuseKb, _
        AmbientKr, AmbientKg, AmbientKb, _
        SpecularK, SpecularN)
    vertex_r(4) = clr And &HFF&
    vertex_g(4) = (clr And &HFF00&) \ &H100&
    vertex_b(4) = (clr And &HFF0000) \ &H10000

    ' Create a region using the transformed vertices.
    pts(1).X = Tx1
    pts(1).Y = Ty1
    pts(2).X = Tx2
    pts(2).Y = Ty2
    pts(3).X = Tx3
    pts(3).Y = Ty3
    pts(4).X = Tx4
    pts(4).Y = Ty4

    hRgn = CreatePolygonRgn(pts(1), 4&, ALTERNATE)

    ' Get the region's bounding box.
    GetRgnBox hRgn, R

    ' Examine the points within the bounding box.
    For X = R.Left To R.Right
        For Y = R.Top To R.Bottom
            ' See if this point is in the region.
            If PtInRegion(hRgn, X, Y) Then
                ' Get the weighted average of the
                ' vertex colors.
                clr = GouraudColor(X, Y, pts, _
                    vertex_r, vertex_g, vertex_b)

                ' Draw the point.
                If clr >= 0 Then pic.PSet (X, Y), clr
            End If
        Next Y
    Next X

    ' Destroy the region to free its resources.
    DeleteObject hRgn
End Sub
' Draw a Phong shaded quadrilateral.
Public Sub PhongQuadrilateral(ByVal pic As PictureBox, _
    ByVal light_sources As Collection, _
    ByVal ambient_light As Integer, _
    ByVal eye_x As Single, ByVal eye_y As Single, ByVal eye_z As Single, _
    ByVal DiffuseKr As Single, ByVal DiffuseKg As Single, ByVal DiffuseKb As Single, _
    ByVal AmbientKr As Single, ByVal AmbientKg As Single, ByVal AmbientKb As Single, _
    ByVal SpecularK As Single, ByVal SpecularN As Integer, _
    ByVal x1 As Single, ByVal y1 As Single, ByVal z1 As Single, _
    ByVal x2 As Single, ByVal y2 As Single, ByVal z2 As Single, _
    ByVal x3 As Single, ByVal y3 As Single, ByVal z3 As Single, _
    ByVal x4 As Single, ByVal y4 As Single, ByVal z4 As Single, _
    ByVal Nx1 As Single, ByVal Ny1 As Single, ByVal Nz1 As Single, _
    ByVal Nx2 As Single, ByVal Ny2 As Single, ByVal Nz2 As Single, _
    ByVal Nx3 As Single, ByVal Ny3 As Single, ByVal Nz3 As Single, _
    ByVal Nx4 As Single, ByVal Ny4 As Single, ByVal Nz4 As Single, _
    ByVal Tx1 As Single, ByVal Ty1 As Single, _
    ByVal Tx2 As Single, ByVal Ty2 As Single, _
    ByVal Tx3 As Single, ByVal Ty3 As Single, _
    ByVal Tx4 As Single, ByVal Ty4 As Single _
    )
Dim pts(1 To 4) As POINTAPI
Dim vertex_r(1 To 4) As Integer
Dim vertex_g(1 To 4) As Integer
Dim vertex_b(1 To 4) As Integer
Dim i As Integer
Dim hRgn As Long
Dim R As RECT
Dim X As Long
Dim Y As Long
Dim Z As Long
Dim ptx As Single
Dim pty As Single
Dim ptz As Single
Dim Nx As Single
Dim Ny As Single
Dim Nz As Single
Dim clr As Long
Dim S As Single
Dim T As Single
Dim N_len As Single

    ' Create a region using the transformed vertices.
    pts(1).X = Tx1
    pts(1).Y = Ty1
    pts(2).X = Tx2
    pts(2).Y = Ty2
    pts(3).X = Tx3
    pts(3).Y = Ty3
    pts(4).X = Tx4
    pts(4).Y = Ty4

    hRgn = CreatePolygonRgn(pts(1), 4&, ALTERNATE)

    ' Get the region's bounding box.
    GetRgnBox hRgn, R

    ' Examine the points within the bounding box.
    For X = R.Left To R.Right
        For Y = R.Top To R.Bottom
            ' See if this point is in the region.
            If PtInRegion(hRgn, X, Y) Then

                ' Find the parameters s and t that
                ' map the point in its quadrilateral.
                PointsToST X, Y, S, T, _
                    pts(3).X, pts(3).Y, pts(2).X, pts(2).Y, _
                    pts(4).X, pts(4).Y, pts(1).X, pts(1).Y

                ' Get the weighted average of the
                ' vertex normals.
                Nx = STInterpolate(S, T, Nx3, Nx2, Nx4, Nx1)
                Ny = STInterpolate(S, T, Ny3, Ny2, Ny4, Ny1)
                Nz = STInterpolate(S, T, Nz3, Nz2, Nz4, Nz1)

                ' Normalize the vector.
                N_len = Sqr(Nx * Nx + Ny * Ny + Nz * Nz)
                Nx = Nx / N_len
                Ny = Ny / N_len
                Nz = Nz / N_len

                ' Get the weighted average of the
                ' vertex locations.
                ptx = STInterpolate(S, T, x3, x2, x4, x1)
                pty = STInterpolate(S, T, y3, y2, y4, y1)
                ptz = STInterpolate(S, T, z3, z2, z4, z1)

                ' Calculate the point's color
                ' using the normal and location.
                clr = CalculateSurfaceColor( _
                    ptx, pty, ptz, Nx, Ny, Nz, _
                    light_sources, ambient_light, _
                    eye_x, eye_y, eye_z, _
                    DiffuseKr, DiffuseKg, DiffuseKb, _
                    AmbientKr, AmbientKg, AmbientKb, _
                    SpecularK, SpecularN)

                ' Draw the point.
                If clr >= 0 Then pic.PSet (X, Y), clr
            End If
        Next Y
    Next X

    ' Destroy the region to free its resources.
    DeleteObject hRgn
End Sub


' Return a weighted average for this point's color.
Private Function GouraudColor(ByVal X As Single, ByVal Y As Single, pts() As POINTAPI, vertex_r() As Integer, vertex_g() As Integer, vertex_b() As Integer) As Long
Dim S As Single
Dim T As Single
Dim R As Single
Dim G As Single
Dim B As Single

    ' Find the parameters s and t that map the
    ' point in its quadrilateral.
    If PointsToST(X, Y, S, T, _
        pts(3).X, pts(3).Y, pts(2).X, pts(2).Y, _
        pts(4).X, pts(4).Y, pts(1).X, pts(1).Y) _
    Then
        ' Use s and t to interpolate the color value.
        R = STInterpolate(S, T, vertex_r(3), vertex_r(2), vertex_r(4), vertex_r(1))
        G = STInterpolate(S, T, vertex_g(3), vertex_g(2), vertex_g(4), vertex_g(1))
        B = STInterpolate(S, T, vertex_b(3), vertex_b(2), vertex_b(4), vertex_b(1))

        GouraudColor = RGB(R, G, B)
    Else
        GouraudColor = RGB(vertex_r(1), vertex_g(1), vertex_b(1))
    End If
End Function
' Using s and t values, return a weighted average
' of the colors RGB(r1, g1, b1), etc.
Private Function STToColor(ByVal S As Single, ByVal T As Single, _
    ByVal R1 As Integer, ByVal g1 As Integer, ByVal b1 As Integer, _
    ByVal r2 As Integer, ByVal g2 As Integer, ByVal b2 As Integer, _
    ByVal r3 As Integer, ByVal g3 As Integer, ByVal b3 As Integer, _
    ByVal r4 As Integer, ByVal g4 As Integer, ByVal b4 As Integer) As Long
Dim ra As Single
Dim ga As Single
Dim ba As Single
Dim rb As Single
Dim gb As Single
Dim bb As Single
Dim new_r As Single
Dim new_g As Single
Dim new_b As Single

    ra = R1 + T * (r2 - R1)
    ga = g1 + T * (g2 - g1)
    ba = b1 + T * (b2 - b1)
    rb = r3 + T * (r4 - r3)
    gb = g3 + T * (g4 - g3)
    bb = b3 + T * (b4 - b3)
    new_r = ra + S * (rb - ra)
    new_g = ga + S * (gb - ga)
    new_b = ba + S * (bb - ba)

    If new_r < 0 Then new_r = 0
    If new_g < 0 Then new_g = 0
    If new_b < 0 Then new_b = 0

    STToColor = RGB(new_r, new_g, new_b)
End Function
' Using s and t values to return a weighted average
' of the values v1, v2, v3, v4.
Private Function STInterpolate(ByVal S As Single, ByVal T As Single, _
    ByVal v1 As Single, ByVal v2 As Single, ByVal v3 As Single, ByVal v4 As Single) As Single
Dim va As Single
Dim vb As Single

    va = v1 + T * (v2 - v1)
    vb = v3 + T * (v4 - v3)
    STInterpolate = va + S * (vb - va)
End Function
' Find S and T for the point (X, Y) in the
' quadrilateral with points (x1, y1), (x2, y2),
' (x3, y3), and (x4, y4). Return True if the point
' lies within the quadrilateral and False otherwise.
Private Function PointsToST(ByVal X As Single, ByVal Y As Single, ByRef S As Single, ByRef T As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single) As Boolean
Dim Ax As Single
Dim Bx As Single
Dim Cx As Single
Dim dx As Single
Dim Ex As Single
Dim Ay As Single
Dim By As Single
Dim Cy As Single
Dim dy As Single
Dim Ey As Single
Dim A As Single
Dim B As Single
Dim C As Single
Dim det As Single
Dim denom As Single

    Ax = x2 - x1: Ay = y2 - y1
    Bx = x4 - x3: By = y4 - y3
    Cx = x3 - x1: Cy = y3 - y1
    dx = X - x1: dy = Y - y1
    Ex = Bx - Ax: Ey = By - Ay

    A = -Ax * Ey + Ay * Ex
    B = Ey * dx - dy * Ex + Ay * Cx - Ax * Cy
    C = dx * Cy - dy * Cx

    det = B * B - 4 * A * C
    If (det >= 0) And (Abs(B) > 0.001) Then
        If Abs(A) < 0.001 Then
            T = -C / B
        Else
            T = (-B - Sqr(det)) / (2 * A)
        End If
        denom = (Cx + Ex * T)
        If Abs(denom) > 0.001 Then
            S = (dx - Ax * T) / denom
        Else
            denom = (Cy + Ey * T)
            If Abs(denom) > 0.001 Then
                S = (dy - Ay * T) / denom
            Else
                S = -1
            End If
        End If

        PointsToST = _
            (T >= -0.00001 And T <= 1.00001 And _
             S >= -0.00001 And S <= 1.00001)
    Else
        PointsToST = False
    End If
End Function



' Return the proper shade for this face
' due to the indicated light source.
Public Function CalculateSurfaceColor( _
    ByVal hit_x As Single, ByVal hit_y As Single, ByVal hit_z As Single, _
    ByVal Nx As Single, ByVal Ny As Single, ByVal Nz As Single, _
    ByVal light_sources As Collection, _
    ByVal ambient_light As Integer, _
    ByVal eye_x As Single, ByVal eye_y As Single, ByVal eye_z As Single, _
    ByVal DiffuseKr As Single, ByVal DiffuseKg As Single, ByVal DiffuseKb As Single, _
    ByVal AmbientKr As Single, ByVal AmbientKg As Single, ByVal AmbientKb As Single, _
    ByVal SpecularK As Single, ByVal SpecularN As Integer _
) As Long

Dim Light As LightSource
Dim Lx As Single
Dim Ly As Single
Dim Lz As Single
Dim L_len As Single
Dim NdotL As Single
Dim R As Integer
Dim G As Integer
Dim B As Integer
Dim distance_factor As Single
Dim diffuse_factor As Single
Dim Vx As Single
Dim Vy As Single
Dim Vz As Single
Dim V_len As Single
Dim Rx As Single
Dim Ry As Single
Dim Rz As Single
Dim RdotV As Single
Dim specular_factor As Single

    For Each Light In light_sources
        ' **********************
        ' * Diffuse Reflection *
        ' **********************
        ' Find the unit vector pointing towards the light.
        Lx = Light.X - hit_x
        Ly = Light.Y - hit_y
        Lz = Light.Z - hit_z
        L_len = Sqr(Lx * Lx + Ly * Ly + Lz * Lz)
        Lx = Lx / L_len
        Ly = Ly / L_len
        Lz = Lz / L_len

        ' See how intense to make the color.
        NdotL = Nx * Lx + Ny * Ly + Nz * Lz

        ' The light does not hit the top of the
        ' surface if NdotL <= 0.
        If (NdotL > 0) Then
            ' Don't use distance shading.
            'distance_factor = (Light.Rmin + Light.Kdist) / (L_len + Light.Kdist)
            distance_factor = 1

            diffuse_factor = NdotL * distance_factor
            R = R + Light.Ir * DiffuseKr * diffuse_factor
            G = G + Light.Ig * DiffuseKg * diffuse_factor
            B = B + Light.Ib * DiffuseKb * diffuse_factor

            ' ***********************
            ' * Specular Reflection *
            ' ***********************
            ' Find the unit vector V from the surface
            ' to the viewing position.
            Vx = eye_x - hit_x
            Vy = eye_y - hit_y
            Vz = eye_z - hit_z
            V_len = Sqr(Vx * Vx + Vy * Vy + Vz * Vz)
            Vx = Vx / V_len
            Vy = Vy / V_len
            Vz = Vz / V_len

            ' Find the mirror vector R.
            Rx = 2 * Nx * NdotL - Lx
            Ry = 2 * Ny * NdotL - Ly
            Rz = 2 * Nz * NdotL - Lz

            ' Calculate the specular component.
            RdotV = Rx * Vx + Ry * Vy + Rz * Vz
            specular_factor = SpecularK * (RdotV ^ SpecularN)
            R = R + Light.Ir * specular_factor
            G = G + Light.Ig * specular_factor
            B = B + Light.Ib * specular_factor
        End If ' End if NdotL > 0 ...
    Next Light

    ' Add the ambient term.
    R = R + ambient_light * AmbientKr
    G = G + ambient_light * AmbientKg
    B = B + ambient_light * AmbientKb

    ' Keep the color components <= 255.
    If R > 255 Then R = 255
    If G > 255 Then G = 255
    If B > 255 Then B = 255

    ' Return the color.
    CalculateSurfaceColor = RGB(R, G, B)
End Function
