Attribute VB_Name = "Geometry"
Option Explicit

' Return two perpendicular vectors normal
' to the vector <Vx, Vy, Vz>.
Public Sub GetLineNormals(ByVal Vx As Single, ByVal Vy As Single, ByVal Vz As Single, ByRef v1x As Single, ByRef v1y As Single, ByRef v1z As Single, ByRef v2x As Single, ByRef v2y As Single, ByRef v2z As Single)
Dim length As Single

    ' Normalize the input vector.
    length = Sqr(Vx * Vx + Vy * Vy + Vz * Vz)
    Vx = Vx / length
    Vy = Vy / length
    Vz = Vz / length

    ' Get the two vectors.
    If Vx <> 0 Then
        m3Cross v1x, v1y, v1z, Vx, Vy, Vz, 0, 1, 0
    ElseIf Vy <> 0 Then
        m3Cross v1x, v1y, v1z, Vx, Vy, Vz, 0, 0, 1
    Else
        m3Cross v1x, v1y, v1z, Vx, Vy, Vz, 1, 0, 0
    End If
    m3Cross v2x, v2y, v2z, Vx, Vy, Vz, v1x, v1y, v1z
End Sub
