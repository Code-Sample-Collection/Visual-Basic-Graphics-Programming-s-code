VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Solid3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' These Face3d objects are the oriented faces.
Public Faces As Collection

Public zmax As Single

Public IsConvex As Boolean
Public HideSurfaces As Boolean
' Set the diffuse reflection coefficients for
' the faces.
Public Sub SetDiffuseCoefficients(ByVal kr As Single, ByVal kg As Single, ByVal kb As Single)
Dim face As Face3d

    For Each face In Faces
        face.DiffuseKr = kr
        face.DiffuseKg = kg
        face.DiffuseKb = kb
    Next face
End Sub
' Set the ambient coefficients for the faces.
Public Sub SetAmbientCoefficients(ByVal kr As Single, ByVal kg As Single, ByVal kb As Single)
Dim face As Face3d

    For Each face In Faces
        face.AmbientKr = kr
        face.AmbientKg = kg
        face.AmbientKb = kb
    Next face
End Sub

' Sort the faces so those with the largest
' transformed Z coordinates come first.
'
' As we switch faces around, we keep track of the
' number of switches we have made. If it clear we
' are stuck in an infinite loop, just move the
' first face to the ordered_faces collection so we
' can continue.
Public Sub SortFaces()
Dim ordered_faces As Collection
Dim face_1 As Face3d
Dim face_i As Face3d
Dim i As Integer
Dim Xmin As Single
Dim xmax As Single
Dim ymin As Single
Dim ymax As Single
Dim zmin As Single
Dim zmax As Single
Dim xmini As Single
Dim xmaxi As Single
Dim ymini As Single
Dim ymaxi As Single
Dim zmini As Single
Dim zmaxi As Single
Dim overlap As Boolean
Dim switches As Integer
Dim max_switches As Integer

    Set ordered_faces = New Collection

    ' Pull out any that are culled. These are not
    ' drawn so we can put them at the front of
    ' the ordered_faces collection.
    For i = Faces.Count To 1 Step -1
        If Faces(i).IsCulled Then
            ordered_faces.Add Faces(i)
            Faces.Remove i
        End If
    Next i

    ' Order the remaining faces.
    max_switches = Faces.Count
    Do While Faces.Count > 0
        ' Get the first item's extent.
        Set face_1 = Faces(1)
        face_1.GetExtent Xmin, xmax, ymin, ymax, zmin, zmax

        ' Compare this face to the others.
        overlap = False     ' In case Face.Count = 0.
        For i = 2 To Faces.Count
            Set face_i = Faces(i)

            ' Get item i's extent.
            face_i.GetExtent xmini, xmaxi, ymini, ymaxi, zmini, zmaxi
            overlap = True
            If xmaxi <= Xmin Or xmini >= xmax Or _
               ymaxi <= ymin Or ymini >= ymax Or _
               zmini >= zmax _
            Then
                ' The extents do not overlap.
                overlap = False
            ElseIf face_i.IsAbove(face_1) Then
                ' Face i is all above the plane
                ' of face 1.
                overlap = False
            ElseIf face_1.IsBelow(face_i) Then
                ' Face 1 is all beneath the plane
                ' of face i.
                overlap = False
            ElseIf Not face_1.Obscures(face_i) Then
                ' If face_1 does not lie partly above
                ' face_i, then there is no problem.
                overlap = False
            End If

            If overlap Then Exit For
        Next i

        If overlap And switches < max_switches Then
            ' There's overlap, move face i to the
            ' top of the list.
            Faces.Remove i
            Faces.Add face_i, , 1 ' Before position 1.
            switches = switches + 1
        Else
            ' There's no overlap. Move face 1 to
            ' the ordered_faces collection.
            ordered_faces.Add face_1
            Faces.Remove 1
            max_switches = Faces.Count
            switches = 0
        End If
    Loop ' Loop until we've ordered all the faces.

    ' Replace the Faces collection with the
    ' ordered_faces collection.
    Set Faces = ordered_faces
End Sub
' Set the ZMax value for the solid.
Public Sub SetZmax()
Dim face As Face3d
Dim z_max As Single

    zmax = -1E+30

    For Each face In Faces
        z_max = face.zmax()
        If zmax < z_max Then zmax = z_max
    Next face
End Sub
' Create a pyramid with height L and base given
' by the points in the coord array. Add the
' faces that make up the pyramid to this solid.
Public Sub Stellate(L As Single, ParamArray coord() As Variant)
Dim x0 As Single
Dim y0 As Single
Dim z0 As Single
Dim x1 As Single
Dim y1 As Single
Dim z1 As Single
Dim x2 As Single
Dim y2 As Single
Dim z2 As Single
Dim x3 As Single
Dim y3 As Single
Dim z3 As Single
Dim Ax As Single
Dim Ay As Single
Dim Az As Single
Dim Bx As Single
Dim By As Single
Dim Bz As Single
Dim Nx As Single
Dim Ny As Single
Dim Nz As Single
Dim num As Integer
Dim i As Integer
Dim pt As Integer

    num = (UBound(coord) + 1) \ 3
    If num < 3 Then
        MsgBox "Must have at least 3 points to stellate.", , vbExclamation
        Exit Sub
    End If

    ' (x0, y0, z0) is the center of the polygon.
    x0 = 0
    y0 = 0
    z0 = 0
    pt = 0
    For i = 1 To num
        x0 = x0 + coord(pt)
        y0 = y0 + coord(pt + 1)
        z0 = z0 + coord(pt + 2)
        pt = pt + 3
    Next i
    x0 = x0 / num
    y0 = y0 / num
    z0 = z0 / num

    ' Find the normal.
    x1 = coord(0)
    y1 = coord(1)
    z1 = coord(2)
    x2 = coord(3)
    y2 = coord(4)
    z2 = coord(5)
    x3 = coord(6)
    y3 = coord(7)
    z3 = coord(8)
    Ax = x2 - x1
    Ay = y2 - y1
    Az = z2 - z1
    Bx = x3 - x2
    By = y3 - y2
    Bz = z3 - z2
    m3Cross Nx, Ny, Nz, Ax, Ay, Az, Bx, By, Bz

    ' Give the normal length L.
    m3SizeVector L, Nx, Ny, Nz

    ' The normal + <x0, y0, z0> gives the point.
    x0 = x0 + Nx
    y0 = y0 + Ny
    z0 = z0 + Nz

    ' Build the faces.
    x1 = coord(3 * num - 3)
    y1 = coord(3 * num - 2)
    z1 = coord(3 * num - 1)
    pt = 0
    For i = 1 To num
        x2 = coord(pt)
        y2 = coord(pt + 1)
        z2 = coord(pt + 2)
        AddFace x1, y1, z1, x2, y2, z2, x0, y0, z0
        x1 = x2
        y1 = y2
        z1 = z2
        pt = pt + 3
    Next i
End Sub


' Clip faces.
Public Sub ClipEye(ByVal R As Single)
Dim obj As Face3d

    For Each obj In Faces
        obj.ClipEye R
    Next obj
End Sub

' Add an oriented face to the solid.
Public Sub AddFace(ParamArray coord() As Variant)
Dim face As Face3d
Dim num As Integer
Dim pt As Integer
Dim i As Integer

    num = (UBound(coord) + 1) \ 3
    If num < 3 Then
        MsgBox "Faces in a Solid must contain at least 3 points.", , vbExclamation
        Exit Sub
    End If

    Set face = New Face3d
    Faces.Add face

    pt = 0
    For i = 1 To num
        face.AddPoints (coord(pt)), (coord(pt + 1)), (coord(pt + 2))
        pt = pt + 3
    Next i
End Sub
' Perform backface removal on the faces for
' center of projection at (X, Y, Z).
Public Sub Cull(ByVal X As Single, ByVal Y As Single, ByVal Z As Single)
Dim obj As Face3d

    For Each obj In Faces
        obj.Cull X, Y, Z
    Next obj
End Sub
' Set or clear the Culled property for all faces.
Property Let Culled(ByVal new_value As Boolean)
Dim obj As Face3d

    For Each obj In Faces
        obj.IsCulled = new_value
    Next obj
End Property



' Apply a transformation matrix which may not
' contain 0, 0, 0, 1 in the last column to the
' object.
Public Sub ApplyFull(M() As Single)
Dim obj As Face3d

    For Each obj In Faces
        obj.ApplyFull M
    Next obj
End Sub

' Apply a transformation matrix to the object.
Public Sub Apply(M() As Single)
Dim obj As Face3d

    For Each obj In Faces
        obj.Apply M
    Next obj
End Sub


' Draw the transformed solid on a PictureBox.
Public Sub Draw(ByVal pic As PictureBox, ByVal light_sources As Collection, ByVal ambient_light As Integer)
Dim face As Face3d

    ' If we do not know this is a convex solid,
    ' order the faces.
    If HideSurfaces And (Not IsConvex) Then SortFaces

    ' Draw the faces.
    For Each face In Faces
        face.Draw pic, light_sources, ambient_light
    Next face
End Sub

Private Sub Class_Initialize()
    Set Faces = New Collection
End Sub


