Attribute VB_Name = "RayTracing"
Option Explicit

Public Running As Boolean

' The collection of objects in the scene.
Public Objects As Collection

' Viewing position.
Public EyeR As Single
Public EyeTheta As Single
Public EyePhi As Single
Public EyeX As Single
Public EyeY As Single
Public EyeZ As Single

' Focus point.
Public FocusX As Single
Public FocusY As Single
Public FocusZ As Single

' Collection of light sources.
Public LightSources As Collection

' Ambient light.
Public AmbientIr As Single
Public AmbientIg As Single
Public AmbientIb As Single

' The background color.
Public BackR As Long
Public BackG As Long
Public BackB As Long
' Return an interpolated pixel color value.
Public Sub InterpolateColor(ByRef R As Single, ByRef G As Single, ByRef B As Single, pixels() As RGBTriplet, ByVal X As Single, ByVal Y As Single)
Dim ix As Integer
Dim iy As Integer
Dim dx1 As Single
Dim dx2 As Single
Dim dy1 As Single
Dim dy2 As Single
Dim v11 As Integer
Dim v12 As Integer
Dim v21 As Integer
Dim v22 As Integer

    ' Find the nearest integral position.
    ix = Int(X)
    iy = Int(Y)

    ' See if this is out of bounds.
    If (ix < 0) Or (ix >= UBound(pixels, 1)) Or _
       (iy < 0) Or (iy >= UBound(pixels, 2)) _
    Then
        ' The point is outside the image. Use black.
        R = 0
        G = 0
        B = 0
    Else
        ' The point lies within the image.
        ' Calculate its value.
        dx1 = X - ix
        dy1 = Y - iy
        dx2 = 1# - dx1
        dy2 = 1# - dy1

        ' Calculate the red value.
        v11 = pixels(ix, iy).rgbRed
        v12 = pixels(ix, iy + 1).rgbRed
        v21 = pixels(ix + 1, iy).rgbRed
        v22 = pixels(ix + 1, iy + 1).rgbRed
        R = v11 * dx2 * dy2 + v12 * dx2 * dy1 + _
            v21 * dx1 * dy2 + v22 * dx1 * dy1

        ' Calculate the green value.
        v11 = pixels(ix, iy).rgbGreen
        v12 = pixels(ix, iy + 1).rgbGreen
        v21 = pixels(ix + 1, iy).rgbGreen
        v22 = pixels(ix + 1, iy + 1).rgbGreen
        G = v11 * dx2 * dy2 + v12 * dx2 * dy1 + _
            v21 * dx1 * dy2 + v22 * dx1 * dy1

        ' Calculate the blue value.
        v11 = pixels(ix, iy).rgbBlue
        v12 = pixels(ix, iy + 1).rgbBlue
        v21 = pixels(ix + 1, iy).rgbBlue
        v22 = pixels(ix + 1, iy + 1).rgbBlue
        B = v11 * dx2 * dy2 + v12 * dx2 * dy1 + _
            v21 * dx1 * dy2 + v22 * dx1 * dy1
    End If
End Sub
' Return the red, green, and blue components of
' an object at a hit position (px, py, pz) with
' normal vector (nx, ny, nz).
Public Sub CalculateHitColor(ByVal depth As Integer, _
    Objects As Collection, ByVal target_object As RayTraceable, _
    ByVal eye_x As Single, ByVal eye_y As Single, ByVal eye_z As Single, _
    ByVal px As Single, ByVal py As Single, ByVal pz As Single, _
    ByVal Nx As Single, ByVal Ny As Single, ByVal Nz As Single, _
    ByVal DiffuseKr As Single, ByVal DiffuseKg As Single, ByVal DiffuseKb As Single, _
    ByVal AmbientKr As Single, ByVal AmbientKg As Single, ByVal AmbientKb As Single, _
    ByVal SpecularK As Single, ByVal SpecularN As Single, _
    ByVal Krr As Single, ByVal Krg As Single, ByVal Krb As Single, _
    ByVal is_reflective As Boolean, _
    ByVal Ktr As Single, ByVal Ktg As Single, ByVal Ktb As Single, _
    ByVal TransN As Single, ByVal n1 As Single, ByVal n2 As Single, _
    ByVal is_transparent As Boolean, _
    ByRef R As Integer, ByRef G As Integer, ByRef B As Integer)

' Vectors.
Dim Vx As Single        ' V: P to viewpoint.
Dim Vy As Single
Dim Vz As Single
Dim v_len As Single
Dim lx As Single        ' L: P to light source.
Dim ly As Single
Dim lz As Single
Dim nlx As Single       ' Unit length NL: P to light source.
Dim nly As Single
Dim nlz As Single
Dim l_len As Single
Dim lmx As Single       ' LM: Light source mirror vector.
Dim lmy As Single
Dim lmz As Single
Dim vmx As Single       ' VM: View direction mirror.
Dim vmy As Single
Dim vmz As Single
Dim ltx As Single       ' LT: Light transmission vector.
Dim lty As Single
Dim ltz As Single
Dim vtx As Single       ' VT: Viewing transmission vector.
Dim vty As Single
Dim vtz As Single

' Dot products.
Dim LdotN As Single
Dim VdotN As Single
Dim LMdotV As Single
Dim VdotLT As Single

' Colors
Dim total_r As Single
Dim total_g As Single
Dim total_b As Single
Dim r_refl As Integer
Dim g_refl As Integer
Dim b_refl As Integer
Dim r_tran As Integer
Dim g_tran As Integer
Dim b_tran As Integer

Dim light_source As LightSource
Dim shadowed As Boolean
Dim shadow_object As RayTraceable
Dim shadow_t As Single
Dim cos1 As Single
Dim cos2 As Single
Dim cos2_squared As Single
Dim n1_over_n2 As Single
Dim normal_factor As Single
Dim trans_d As Single
Dim distance_factor As Single
Dim diffuse_factor As Single
Dim specular_factor As Single
Dim n_ratio As Single
Dim cos_factor As Single
Dim transmitted_factor As Single

    ' Get vector V.
    Vx = eye_x - px
    Vy = eye_y - py
    Vz = eye_z - pz
    v_len = Sqr(Vx * Vx + Vy * Vy + Vz * Vz)
    Vx = Vx / v_len
    Vy = Vy / v_len
    Vz = Vz / v_len

    ' Calculate V dot N.
    VdotN = Vx * Nx + Vy * Ny + Vz * Nz

    ' ***********
    ' * Ambient *
    ' ***********
    total_r = AmbientIr * AmbientKr
    total_g = AmbientIg * AmbientKg
    total_b = AmbientIb * AmbientKb

    ' Consider each light source.
    For Each light_source In LightSources
        ' Find vector L not normalized.
        lx = light_source.TransX - px
        ly = light_source.TransY - py
        lz = light_source.TransZ - pz

        ' Get the distance factor for depth queueing.
        l_len = Sqr(lx * lx + ly * ly + lz * lz)
        distance_factor = (light_source.Rmin + light_source.Kdist) / (l_len + light_source.Kdist)

        ' Normalize vector L.
        nlx = lx / l_len
        nly = ly / l_len
        nlz = lz / l_len

        ' See if the light is on the same side of the
        ' surface as the normal.
        LdotN = nlx * Nx + nly * Ny + nlz * Nz

        ' See if the light and viewpoint are on
        ' opposite sides of the surface.
        shadowed = (LdotN * VdotN < 0)

        ' See if we are shadowed.
        If Not shadowed Then
            For Each shadow_object In Objects
                If Not (shadow_object Is target_object) Then
                    ' See where vector L intersects
                    ' the shadow object.
                    shadow_t = shadow_object.FindT( _
                        False, _
                        light_source.TransX, _
                        light_source.TransY, _
                        light_source.TransZ, _
                        -lx, -ly, -lz)

                    ' If shadow_t < 1, we're shadowed.
                    If (shadow_t > 0.00001) And (shadow_t < 0.99999) Then
                        shadowed = True
                        Exit For
                    End If
                End If
            Next shadow_object
        End If

        ' We have diffuse and specular components if
        ' the light and viewpoint are on the same
        ' side of the surface normal, and if
        ' we are not shadowed.
        If (LdotN > 0) And (VdotN > 0) And (Not shadowed) Then
            ' The light is shining on the surface.
            ' ***********
            ' * Diffuse *
            ' ***********
            ' There is a diffuse component.
            diffuse_factor = LdotN * distance_factor
            total_r = total_r + light_source.Ir * DiffuseKr * diffuse_factor
            total_g = total_g + light_source.Ig * DiffuseKg * diffuse_factor
            total_b = total_b + light_source.Ib * DiffuseKb * diffuse_factor

            ' ************
            ' * Specular *
            ' ************
            ' Find the light mirror vector LM.
            lmx = 2 * Nx * LdotN - nlx
            lmy = 2 * Ny * LdotN - nly
            lmz = 2 * Nz * LdotN - nlz

            LMdotV = lmx * Vx + lmy * Vy + lmz * Vz
            If LMdotV > 0 Then
                specular_factor = SpecularK * (LMdotV ^ SpecularN)
                total_r = total_r + light_source.Ir * specular_factor
                total_g = total_g + light_source.Ig * specular_factor
                total_b = total_b + light_source.Ib * specular_factor
            End If
        End If ' End if the light shines on the surface.

        ' **********************
        ' * Direct Transmitted *
        ' **********************
        ' See if the light and viewpoint are on
        ' opposite sides of the surface and if we
        ' are not in shadow.
        If is_transparent Then
            ' Find LT, the light transmission vector.
            If LdotN < 0 Then
                ' L and N point in opposite directions.
                ' The ray is leaving the object.
                n1_over_n2 = n2 / n1
            Else
                ' L and N point in the same direction.
                ' The ray is entering the object.
                n1_over_n2 = n1 / n2
            End If

            cos1 = LdotN
            cos2_squared = 1 - (1 - cos1 * cos1) * n1_over_n2 * n1_over_n2
            If cos2_squared > 0 Then
                cos2 = Sqr(cos2_squared)
                ' Note that the incident vector I = -L.
                normal_factor = cos2 - n1_over_n2 * cos1
                ltx = -n1_over_n2 * nlx - normal_factor * Nx
                lty = -n1_over_n2 * nly - normal_factor * Ny
                ltz = -n1_over_n2 * nlz - normal_factor * Nz

                ' Calculate V dot LT.
                VdotLT = Vx * ltx + Vy * lty + Vz * ltz

                ' See if V and LT point in generally
                ' the same direction.
                If VdotLT > 0 Then
                    ' Calculate V dot LT to the TransN.
                    transmitted_factor = VdotLT ^ TransN

                    ' Add the direct transmitted component.
                    total_r = total_r + Ktr * light_source.Ir * transmitted_factor
                    total_g = total_g + Ktg * light_source.Ig * transmitted_factor
                    total_b = total_b + Ktb * light_source.Ib * transmitted_factor
                End If
            End If
        End If
    Next light_source

    ' *************
    ' * Reflected *
    ' *************
    If (depth > 1) And is_reflective Then
        ' Find the view mirror vector VM.
        vmx = 2 * Nx * VdotN - Vx
        vmy = 2 * Ny * VdotN - Vy
        vmz = 2 * Nz * VdotN - Vz

        ' Trace a ray from (px, py, pz) in the
        ' direction VM.
        TraceRay False, depth - 1, target_object, _
            px, py, pz, vmx, vmy, vmz, _
            r_refl, g_refl, b_refl

        ' Multiply by the reflection coefficients.
        total_r = total_r + Krr * r_refl
        total_g = total_g + Krg * g_refl
        total_b = total_b + Krb * b_refl
    End If

    ' **************************
    ' * Indirectly Transmitted *
    ' **************************
    ' See if the surface is transparent.
    If (depth > 1) And is_transparent Then
        ' Find VT, the viewing transmission vector.
        cos1 = Abs(VdotN)
        n1_over_n2 = n1 / n2
        cos2_squared = 1 - (1 - cos1 * cos1) * n1_over_n2 * n1_over_n2
        If cos2_squared > 0 Then
            cos2 = Sqr(1 - (1 - cos1 * cos1) * n1_over_n2 * n1_over_n2)
            ' Note that the incident vector I = -V.
            normal_factor = cos2 - n1_over_n2 * cos1
            vtx = -n1_over_n2 * Vx - normal_factor * Nx
            vty = -n1_over_n2 * Vy - normal_factor * Ny
            vtz = -n1_over_n2 * Vz - normal_factor * Nz

            TraceRay False, depth - 1, target_object, _
                px, py, pz, vtx, vty, vtz, _
                r_tran, g_tran, b_tran

            ' Add the indirectly transmitted components.
            total_r = total_r + Ktr * r_tran
            total_g = total_g + Ktg * g_tran
            total_b = total_b + Ktb * b_tran
        End If
    End If

    ' For points close to a light source, these
    ' values can be big. Keep them <= 255.
    If total_r > 255 Then total_r = 255
    If total_g > 255 Then total_g = 255
    If total_b > 255 Then total_b = 255
    If total_r < 0 Then total_r = 0
    If total_g < 0 Then total_g = 0
    If total_b < 0 Then total_b = 0
    R = total_r
    G = total_g
    B = total_b
End Sub
' Return the red, green, and blue components of
' an object at a hit position (px, py, pz) with
' normal vector (nx, ny, nz). Consider diffuse,
' specular, and ambient light components.
Public Sub CalculateHitColorDSA(ByVal depth As Integer, _
    Objects As Collection, ByVal target_object As RayTraceable, _
    ByVal eye_x As Single, ByVal eye_y As Single, ByVal eye_z As Single, _
    ByVal px As Single, ByVal py As Single, ByVal pz As Single, _
    ByVal Nx As Single, ByVal Ny As Single, ByVal Nz As Single, _
    ByVal DiffuseKr As Single, ByVal DiffuseKg As Single, ByVal DiffuseKb As Single, _
    ByVal AmbientKr As Single, ByVal AmbientKg As Single, ByVal AmbientKb As Single, _
    ByVal SpecularK As Single, ByVal SpecularN As Single, _
    ByRef R As Integer, ByRef G As Integer, ByRef B As Integer)

' Vectors.
Dim Vx As Single        ' V: P to viewpoint.
Dim Vy As Single
Dim Vz As Single
Dim v_len As Single
Dim lx As Single        ' L: P to light source.
Dim ly As Single
Dim lz As Single
Dim l_len As Single
Dim lmx As Single       ' LM: Light source mirror vector.
Dim lmy As Single
Dim lmz As Single

' Dot products.
Dim LdotN As Single
Dim VdotN As Single
Dim LMdotV As Single

' Colors
Dim total_r As Single
Dim total_g As Single
Dim total_b As Single

Dim light_source As LightSource
Dim shadowed As Boolean
Dim shadow_object As RayTraceable
Dim shadow_t As Single
Dim spec As Single

    ' Get vector V.
    Vx = eye_x - px
    Vy = eye_y - py
    Vz = eye_z - pz
    v_len = Sqr(Vx * Vx + Vy * Vy + Vz * Vz)
    Vx = Vx / v_len
    Vy = Vy / v_len
    Vz = Vz / v_len

    ' Consider each light source.
    For Each light_source In LightSources
        ' Find vector L not normalized.
        lx = light_source.TransX - px
        ly = light_source.TransY - py
        lz = light_source.TransZ - pz

        ' See if we are shadowed.
        shadowed = False
        For Each shadow_object In Objects
            If Not (shadow_object Is target_object) Then
                ' See where vector L intersects
                ' ths shadow object.
                shadow_t = shadow_object.FindT( _
                    False, _
                    light_source.TransX, _
                    light_source.TransY, _
                    light_source.TransZ, _
                    -lx, -ly, -lz)

                ' If shadow_t < 1, we're shadowed.
                If (shadow_t > 0) And (shadow_t < 1) Then
                    shadowed = True
                    Exit For
                End If
            End If
        Next shadow_object

        ' Normalize vector L.
        l_len = Sqr(lx * lx + ly * ly + lz * lz)
        lx = lx / l_len
        ly = ly / l_len
        lz = lz / l_len

        ' See if the viewpoint is on the same
        ' side of the surface as the normal.
        VdotN = Vx * Nx + Vy * Ny + Vz * Nz

        ' See if the light is on the same side of the
        ' surface as the normal.
        LdotN = lx * Nx + ly * Ny + lz * Nz

        ' We have diffuse and specular components if
        ' the light and viewpoint are on the same
        ' side of the surface as the normal, and if
        ' we are not shadowed.
        If (VdotN >= 0) And (LdotN >= 0) And (Not shadowed) Then
            ' The light is shining on the surface.
            ' ***********
            ' * Diffuse *
            ' ***********
            ' There is a diffuse component.
            total_r = total_r + light_source.Ir * DiffuseKr * LdotN
            total_g = total_g + light_source.Ig * DiffuseKg * LdotN
            total_b = total_b + light_source.Ib * DiffuseKb * LdotN

            ' ************
            ' * Specular *
            ' ************
            ' Find the light mirror vector LM.
            lmx = 2 * Nx * LdotN - lx
            lmy = 2 * Ny * LdotN - ly
            lmz = 2 * Nz * LdotN - lz

            LMdotV = lmx * Vx + lmy * Vy + lmz * Vz
            If LMdotV > 0 Then
                spec = SpecularK * (LMdotV ^ SpecularN)
                total_r = total_r + light_source.Ir * spec
                total_g = total_g + light_source.Ig * spec
                total_b = total_b + light_source.Ib * spec
            End If
        End If ' End if the light shines on the surface.
    Next light_source

    ' ***********
    ' * Ambient *
    ' ***********
    total_r = total_r + AmbientIr * AmbientKr
    total_g = total_g + AmbientIg * AmbientKg
    total_b = total_b + AmbientIb * AmbientKb

    ' For points close to a light source, these
    ' values can be big. Keep them <= 255.
    If total_r > 255 Then total_r = 255
    If total_g > 255 Then total_g = 255
    If total_b > 255 Then total_b = 255
    R = total_r
    G = total_g
    B = total_b
End Sub

' Sort the polygons by Zmin value.
Private Sub QuickSortPolygons(ByVal min As Integer, ByVal max As Integer, polygons() As SimplePolygon)
Dim mid_polygon As SimplePolygon
Dim mid_z As Single
Dim lo As Integer
Dim hi As Integer

    ' See if we're done.
    If min >= max Then Exit Sub

    Set mid_polygon = polygons(min)
    mid_z = mid_polygon.Zmin
    lo = min
    hi = max
    Do
        ' Look down from hi for a value <= med_z.
        Do While (polygons(hi).Zmin >= mid_z)
            hi = hi - 1
            If hi <= lo Then Exit Do
        Loop
        If hi <= lo Then
            Set polygons(lo) = mid_polygon
            Exit Do
        End If

        ' Swap the low and high values.
        Set polygons(lo) = polygons(hi)

        ' Look up from lo for a value >= mid_z.
        lo = lo + 1
        Do While (polygons(lo).Zmin < mid_z)
            lo = lo + 1
            If lo >= hi Then Exit Do
        Loop
        If lo >= hi Then
            lo = hi
            Set polygons(hi) = mid_polygon
            Exit Do
        End If

        ' Swap the low and high values.
        Set polygons(hi) = polygons(lo)
    Loop

    ' Recursively sort the sublists.
    QuickSortPolygons min, lo - 1, polygons
    QuickSortPolygons lo + 1, max, polygons
End Sub
' Sort the polygons so those with the smallest Z
' values come first.
Public Sub OrderPolygons(ByVal num_polygons As Integer, polygons() As SimplePolygon)
Dim min_unfixed As Integer
Dim max_z As Single
Dim i As Integer
Dim j As Integer
Dim pgon As SimplePolygon
Dim obscured As Boolean

    ' Use QuickSort to sort by minimum Z value.
    ' This gets them mostly in order.
    QuickSortPolygons 1, num_polygons, polygons

    ' Fix any small differences.
    min_unfixed = 1
    Do While min_unfixed < num_polygons
        ' Get this polygon's maximum Z value.
        Set pgon = polygons(min_unfixed)
        max_z = pgon.Zmax

        ' Examine the following polygons until we come
        ' to one where Zmin >= this polygon's Zmax.
        ' See pgon obscures them.
        obscured = False
        i = min_unfixed + 1
        Do While i <= num_polygons
            ' See if we have checked far enough.
            If (polygons(i).Zmin >= max_z) Then Exit Do

            ' See if pgon belongs above polygons(i).
            ' This is true if pgon obscures it.
            If pgon.Obscures(polygons(i)) Then
                obscured = True
                Exit Do
            End If
            i = i + 1
        Loop

        ' See if we are obscured.
        If obscured Then
            ' We obscure polygons(i).
            ' Move polygons(i) into position
            ' min_unfixed, keeping the other
            ' polygons in order.
            For j = min_unfixed + 1 To i
                Set polygons(j - 1) = polygons(j)
            Next j
            Set polygons(i) = pgon
        Else
            ' We do not obscure polygons(i).
            ' Pgon is in its correct position.
            min_unfixed = min_unfixed + 1
        End If
    Loop
End Sub
' Create a projection transformation for a non-ray
' traced rendering.
Private Sub TransformForNonRayTracing(M() As Single, ByVal pic As PictureBox)
Dim t1(1 To 4, 1 To 4) As Single
Dim t2(1 To 4, 1 To 4) As Single
Dim T3(1 To 4, 1 To 4) As Single
Dim T12(1 To 4, 1 To 4) As Single

    ' Rotate the eye onto the Z axis.
    m3PProject t1, project_Parallel, _
        EyeR, EyePhi, EyeTheta, _
        FocusX, FocusY, FocusZ, _
        0, 1, 0

    ' Transform the viewing location.
    EyeX = 0
    EyeY = 0
    EyeZ = EyeR

    ' Project as if we were ray tracing.
    project_PerspectiveXY t2, EyeR

    ' Translate the origin to the center
    ' of the PictureBox.
    m3Translate T3, pic.ScaleWidth / 2, pic.ScaleHeight / 2, 0

    ' Combine the transformations.
    m3MatMultiplyFull T12, t1, t2
    m3MatMultiplyFull M, T12, T3
End Sub
' Return the pixel color given by tracing from
' point (px, py, pz) in direction <vx, vy, vz>.
Public Sub TraceRay(ByVal direct_calculation As Boolean, ByVal depth As Integer, ByVal skip_object As RayTraceable, ByVal px As Single, ByVal py As Single, ByVal pz As Single, ByVal Vx As Single, ByVal Vy As Single, ByVal Vz As Single, ByRef R As Integer, ByRef G As Integer, ByRef B As Integer)
Dim obj As RayTraceable
Dim best_obj As RayTraceable
Dim best_t As Single
Dim t As Single

    ' Find the object that's closest.
    best_t = INFINITY
    For Each obj In Objects
        ' Skip the object skip_object. We use this
        ' to avoid erroneously hitting the object
        ' casting out a ray.
        If Not (obj Is skip_object) Then
            t = obj.FindT(direct_calculation, px, py, pz, Vx, Vy, Vz)
            If (t > 0) And (best_t > t) Then
                best_t = t
                Set best_obj = obj
            End If
        End If
    Next obj

    ' See if we hit anything.
    If best_obj Is Nothing Then
        ' We hit nothing. Return the background color.
        R = BackR
        G = BackG
        B = BackB
    Else
        ' Compute the color at that point.
        best_obj.FindHitColor depth, Objects, _
            px, py, pz, _
            px + best_t * Vx, _
            py + best_t * Vy, _
            pz + best_t * Vz, _
            R, G, B

        ' This is a problem for some values of Kdist.
        If R < 0 Then R = 0
        If G < 0 Then G = 0
        If B < 0 Then B = 0
    End If
End Sub
' Ray trace on this picture box.
Public Sub TraceAllRays(ByVal pic As PictureBox, ByVal skip As Integer, ByVal depth As Integer)
Dim pixels() As RGBTriplet
Dim bits_per_pixel As Integer
Dim pix_x As Long
Dim pix_y As Long
Dim real_x As Long
Dim real_y As Long
Dim Xmin As Integer
Dim Xmax As Integer
Dim Ymin As Integer
Dim Ymax As Integer
Dim xoff As Integer
Dim yoff As Integer
Dim R As Integer
Dim G As Integer
Dim B As Integer
Dim obj As RayTraceable
Dim Nx As Single
Dim Ny As Single
Dim Nz As Single
Dim dist As Single

    If skip < 2 Then
        ' Get the picture box's pixels.
        GetBitmapPixels pic, pixels, bits_per_pixel
    End If

    ' Get the transformed coordinates of the eye.
    xoff = pic.ScaleWidth / 2
    yoff = pic.ScaleHeight / 2
    Xmin = pic.ScaleLeft
    Xmax = Xmin + pic.ScaleWidth - 1
    Ymin = pic.ScaleTop
    Ymax = Ymin + pic.ScaleHeight - 1
    For pix_y = Ymin To Ymax Step skip
        real_y = pix_y - yoff

        ' The points in this scanline are on the
        ' plane determined by the points:
        '   A: (0, 0, EyeR)
        '   B: (1, Y, 0)
        '   C: (0, Y, 0)
        ' The cross product AB x AC gives a
        ' normal to this plane as:
        '     <1, Y, -EyeR>
        '   x <0, Y, -EyeR>
        '   = <0, EyeR, Y>
        ' Find the unit normal.
        dist = Sqr(EyeR * EyeR + real_y * real_y)
        Nx = 0
        Ny = EyeR / dist
        Nz = real_y / dist

        ' Prepare the objects for this scanline.
        For Each obj In Objects
            obj.CullScanline 0, 0, EyeR, Nx, Ny, Nz
        Next obj

        For pix_x = Xmin To Xmax Step skip
            real_x = pix_x - xoff
            ' Calculate the value of pixel (x, y).
            ' After transformation the eye is
            ' at (0, 0, EyeR) and the plane of
            ' projection lies in the X-Y plane.
            TraceRay True, depth, Nothing, _
                0, 0, EyeR, _
                CSng(real_x), CSng(real_y), -EyeR, _
                R, G, B

            ' Draw the pixel.
            If skip < 2 Then
                ' Save the pixel value.
                With pixels(pix_x, pix_y)
                    .rgbRed = R
                    .rgbGreen = G
                    .rgbBlue = B
                End With
            Else
                pic.Line (pix_x, pix_y)- _
                    Step(skip - 1, skip - 1), _
                    RGB(R, G, B), BF
            End If
        Next pix_x

        ' Let the user see what's going on.
        If skip < 2 Then
            pic.Line (pic.ScaleLeft, pix_y)-(Xmax, pix_y), vbWhite
        Else
            pic.Refresh
        End If

        ' If the Stop button was pressed, stop.
        DoEvents
        If Not Running Then Exit For
    Next pix_y

    If skip < 2 Then
        SetBitmapPixels pic, bits_per_pixel, pixels
    End If
End Sub
' Perform a ray tracing.
Public Sub RenderRayTracing(ByVal pic As Object, ByVal skip As Integer, ByVal depth As Integer)
Dim M(1 To 4, 1 To 4) As Single
Dim obj As RayTraceable
Dim light_source As LightSource

    ' Rotate the eye onto the Z axis.
    m3PProject M, project_Parallel, _
        EyeR, EyePhi, EyeTheta, _
        FocusX, FocusY, FocusZ, _
        0, 1, 0

    ' Transform the viewing location.
    EyeX = 0
    EyeY = 0
    EyeZ = EyeR

    ' Transform the objects.
    For Each obj In Objects
        obj.Apply M
    Next obj

    ' Transform the light sources.
    For Each light_source In LightSources
        light_source.Apply M
    Next light_source

    ' Scale the light intensities for depth queueing.
    ScaleLightSourcesForDepth

    ' Trace all the rays.
    TraceAllRays pic, skip, depth
End Sub


' Project and draw all the objects.
Public Sub RenderWireFrame(ByVal pic As Object)
Dim M(1 To 4, 1 To 4) As Single
Dim obj As RayTraceable
Dim light_source As LightSource

    ' Get the projection transformation.
    TransformForNonRayTracing M, pic

    ' Transform the objects.
    For Each obj In Objects
        obj.ApplyFull M
    Next obj

    ' Transform the light sources.
    For Each light_source In LightSources
        light_source.ApplyFull M
    Next light_source

    ' Draw the wireframes.
    For Each obj In Objects
        obj.DrawWireFrame pic
    Next obj
End Sub
' Project and draw all the objects with backfaces
' removed.
Public Sub RenderBackfacesRemoved(ByVal pic As Object)
Dim M(1 To 4, 1 To 4) As Single
Dim obj As RayTraceable

    ' Get the projection transformation.
    TransformForNonRayTracing M, pic

    ' Transform the objects.
    For Each obj In Objects
        obj.ApplyFull M
    Next obj

    ' Draw the wireframes.
    pic.FillStyle = vbFSTransparent
    For Each obj In Objects
        obj.DrawBackfacesRemoved pic
    Next obj
End Sub
' Project and draw all the objects with hidden
' surfaces removed.
Public Sub RenderHiddenSurfacesRemoved(ByVal pic As Object, ByVal lblPolygons As Label)
Dim M(1 To 4, 1 To 4) As Single
Dim obj As RayTraceable
Dim num_polygons As Integer
Dim polygons() As SimplePolygon
Dim i As Integer

    ' Get the projection transformation.
    TransformForNonRayTracing M, pic

    ' Transform the objects.
    For Each obj In Objects
        obj.ApplyFull M
    Next obj

    ' Get polygons from the objects.
    For Each obj In Objects
        obj.GetPolygons num_polygons, polygons, False
    Next obj
    lblPolygons.Caption = Format$(num_polygons) & " polygons"
    lblPolygons.Refresh

    ' Sort the polygons.
    OrderPolygons num_polygons, polygons

    ' Draw the polygons in order.
    pic.FillStyle = vbFSSolid
    For i = 1 To num_polygons
        polygons(i).DrawPolygon pic
    Next i
    pic.Refresh
End Sub
' Project and draw all the objects with visible
' shaded surfaces.
Public Sub RenderShaded(ByVal pic As Object, ByVal lblPolygons As Label)
Dim M(1 To 4, 1 To 4) As Single
Dim obj As RayTraceable
Dim light_source As LightSource
Dim num_polygons As Integer
Dim polygons() As SimplePolygon
Dim i As Integer

    ' Get the projection transformation.
    TransformForNonRayTracing M, pic

    ' Transform the objects.
    For Each obj In Objects
        obj.ApplyFull M
    Next obj

    ' Transform the light sources.
    For Each light_source In LightSources
        light_source.ApplyFull M
    Next light_source

    ' Scale the light intensities for depth queueing.
    ScaleLightSourcesForDepth

    ' Get polygons from the objects.
    For Each obj In Objects
        obj.GetPolygons num_polygons, polygons, True
    Next obj
    lblPolygons.Caption = Format$(num_polygons) & " polygons"
    lblPolygons.Refresh

    ' Sort the polygons.
    OrderPolygons num_polygons, polygons

    ' Draw the polygons in order.
    pic.FillStyle = vbFSSolid
    For i = 1 To num_polygons
        polygons(i).DrawPolygon pic
    Next i
    pic.Refresh
End Sub

' Set the light sources' Kdist and Rmin values.
Private Sub ScaleLightSourcesForDepth()
Dim light As LightSource

    For Each light In LightSources
        ScaleIntensityForDepth light
    Next light
End Sub

' Set this light source's Kdist and Rmin values.
Private Sub ScaleIntensityForDepth(ByVal light As LightSource)
Dim solid As RayTraceable
Dim Rmin As Single
Dim Rmax As Single
Dim new_rmin As Single
Dim new_rmax As Single

    Rmin = 1E+30
    Rmax = -1E+30

    For Each solid In Objects
        solid.GetRminRmax new_rmin, new_rmax, _
            light.TransX, light.TransY, light.TransZ
        If Rmin > new_rmin Then Rmin = new_rmin
        If Rmax < new_rmax Then Rmax = new_rmax
    Next solid

    light.Rmin = Rmin
'    light.Kdist = (Rmax - 5 * Rmin) / 4 ' Fade to 1/5.
    light.Kdist = Rmax - 2 * Rmin ' Fade to 1/2.
End Sub
