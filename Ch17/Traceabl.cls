VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RayTraceable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Apply a transformation matrix to the object.
Public Sub Apply(M() As Single)

End Sub
' Apply a transformation matrix to the object.
Public Sub ApplyFull(M() As Single)

End Sub

' Draw the object with backfaces removed.
Public Sub DrawBackfacesRemoved(ByVal pic As PictureBox)

End Sub
' Draw a wireframe for this object.
Public Sub DrawWireFrame(ByVal pic As PictureBox)

End Sub
' Return the red, green, and blue components of
' the surface at the hit position.
Public Sub FindHitColor(ByVal depth As Integer, Objects As Collection, ByVal eye_x As Single, ByVal eye_y As Single, ByVal eye_z As Single, ByVal px As Single, ByVal py As Single, ByVal pz As Single, ByRef R As Integer, ByRef G As Integer, ByRef B As Integer)

End Sub
' Return the value T for the point of intersection
' between the vector from point (px, py, pz) in
' the direction <vx, vy, vz>.
'
' direct_calculation is true if we are finding the
' intersection from a viewing position ray. It is
' false if we are finding an reflected intersection
' or a shadow feeler.
Public Function FindT(ByVal direct_calculation As Boolean, ByVal px As Single, ByVal py As Single, ByVal pz As Single, ByVal Vx As Single, ByVal Vy As Single, ByVal Vz As Single) As Single

End Function
' See if the scanline plane with the indicated
' point and normal intersects this object. Set
' the object's DoneOnThisScanline flag appropriately.
Public Sub CullScanline(ByVal px As Single, ByVal py As Single, ByVal pz As Single, ByVal Nx As Single, ByVal Ny As Single, ByVal Nz As Single)

End Sub
' Add non-backface polygons to this collection.
Public Sub GetPolygons(ByRef num_polygons As Integer, polygons() As SimplePolygon, ByVal shaded As Boolean)

End Sub

' Return the minimum and maximum distances from
' this point.
Public Sub GetRminRmax(ByRef new_min As Single, ByRef new_max As Single, ByVal X As Single, ByVal Y As Single, ByVal Z As Single)

End Sub

' Reset the ForeverCulled flag.
Public Sub ResetCulling()

End Sub


