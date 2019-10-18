VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QtreeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' A quadtree node.

' If this is a leaf node, its Objects
' collection contains the objects to draw.
'
' Otherwise the object's children contain other
' QtreeNode objects.

' The maximum number of objects the node can hold.
Private Const MAX_OBJECTS = 100

' The bounds this quadtree node represents.
Public xmin As Single
Public ymin As Single
Public xmid As Single
Public ymid As Single
Public xmax As Single
Public ymax As Single

' The objects, if this is a leaf node.
Private Objects As Collection

' The quadtree children otherwise.
Public NWchild As QtreeNode
Public NEchild As QtreeNode
Public SWchild As QtreeNode
Public SEchild As QtreeNode
' Set the Drawn properties of the objects.
Public Sub SetDrawn(ByVal new_value As Boolean)
Dim obj As Object

    If Objects Is Nothing Then
        ' We are not a leaf. Make our children
        ' set Drawn for their objects.
        NWchild.SetDrawn new_value
        NEchild.SetDrawn new_value
        SWchild.SetDrawn new_value
        SEchild.SetDrawn new_value
    Else
        ' We are a leaf. Set Drawn for our objects.
        For Each obj In Objects
            obj.Drawn = new_value
        Next obj
    End If
End Sub
' Find an object that contains this point.
Public Function ObjectAt(ByVal X As Single, ByVal Y As Single) As Object
Dim obj As Object

    Set ObjectAt = Nothing

    ' Stop if we don't contain the point.
    If X < xmin Or X > xmax Or _
       Y < ymin Or Y > ymax _
       Then Exit Function

    ' Find the object.
    If Objects Is Nothing Then
        ' This is not a leaf node.
        ' Search our children.
        If Y > ymid Then
            If X < xmid Then
                ' Search the northwest child.
                Set ObjectAt = NWchild.ObjectAt(X, Y)
            Else
                ' Search the northeast child.
                Set ObjectAt = NEchild.ObjectAt(X, Y)
            End If
        Else
            If X < xmid Then
                ' Search the southwest child.
                Set ObjectAt = SWchild.ObjectAt(X, Y)
            Else
                ' Search the southeast child.
                Set ObjectAt = SEchild.ObjectAt(X, Y)
            End If
        End If
    Else
        ' This is a leaf node.
        ' Search the objects it contains.
        For Each obj In Objects
            If obj.IsAt(X, Y) Then
                Set ObjectAt = obj
                Exit Function
            End If
        Next obj
    End If
End Function
' Add an object to the Objects collection.
'
' If this gives us too many objects, create
' child nodes and subdivide.
Public Sub Add(obj As Object)
    If Objects Is Nothing Then
        ' We are not a leaf node. Put the
        ' object in the appropriate child.
        PlaceObject obj
    Else
        ' We are a leaf node. Add the object
        ' to the Objects collection.
        Objects.Add obj

        ' See if need to subdivide.
        If Objects.Count > MAX_OBJECTS Then Divide
    End If
End Sub

' Create the children and divide the object.
Private Sub Divide()
Dim obj As Object

    ' Create the children.
    Set NWchild = New QtreeNode
    NWchild.SetBounds xmin, xmid, ymid, ymax

    Set NEchild = New QtreeNode
    NEchild.SetBounds xmid, xmax, ymid, ymax

    Set SWchild = New QtreeNode
    SWchild.SetBounds xmin, xmid, ymin, ymid

    Set SEchild = New QtreeNode
    SEchild.SetBounds xmid, xmax, ymin, ymid

    ' Move the objects into the proper children.
    For Each obj In Objects
        PlaceObject obj
    Next obj

    ' Remove the Objects collection.
    Set Objects = Nothing
End Sub
' Set the bounds for this quadtree node.
Public Sub SetBounds(ByVal x1 As Single, ByVal x2 As Single, ByVal y1 As Single, ByVal y2 As Single)
    xmin = x1
    ymin = y1
    xmax = x2
    ymax = y2
    xmid = (xmin + xmax) / 2
    ymid = (ymin + ymax) / 2
End Sub
' Place this object in the proper child(ren).
Private Sub PlaceObject(ByVal obj As Object)
Dim x1 As Single
Dim x2 As Single
Dim y1 As Single
Dim y2 As Single

    obj.Bound x1, y1, x2, y2
    If y2 > ymid And x1 < xmid Then NWchild.Add obj
    If y2 > ymid And x2 > xmid Then NEchild.Add obj
    If y1 < ymid And x1 < xmid Then SWchild.Add obj
    If y1 < ymid And x2 > xmid Then SEchild.Add obj
End Sub
' Draw the objects in this node on a PictureBox.
Public Sub Draw(ByVal pic As PictureBox, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
Dim obj As Object

    ' Stop if we don't intersect the region
    ' we're trying to draw.
    If x2 < xmin Or x1 > xmax Or _
       y2 < ymin Or y1 > ymax _
       Then Exit Sub

    ' Draw a red box around our display region.
    pic.Line (xmin, ymin)-(xmax, ymax), vbRed, B

    ' Draw the objects.
    If Objects Is Nothing Then
        ' We are not a leaf. Make our children
        ' draw themselves.
        NWchild.Draw pic, x1, y1, x2, y2
        NEchild.Draw pic, x1, y1, x2, y2
        SWchild.Draw pic, x1, y1, x2, y2
        SEchild.Draw pic, x1, y1, x2, y2
    Else
        ' We are a leaf. Make the objects
        ' draw themselves.
        For Each obj In Objects
            obj.Draw pic
        Next obj
    End If
End Sub
' Start with an empty Objects collection.
Private Sub Class_Initialize()
    Set Objects = New Collection
End Sub
