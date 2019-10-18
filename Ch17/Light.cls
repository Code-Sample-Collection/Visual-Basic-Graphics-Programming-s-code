VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LightSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Light source.

Private Location As Point3D

Public Ir As Single
Public Ig As Single
Public Ib As Single

Public Rmin As Single
Public Kdist As Single

' Apply a transformation matrix to the object.
Public Sub Apply(M() As Single)
    ' Transform the center.
    m3Apply Location.Coord, M, Location.Trans
End Sub
' Apply a transformation matrix to the object.
Public Sub ApplyFull(M() As Single)
    ' Transform the center.
    m3ApplyFull Location.Coord, M, Location.Trans
End Sub

' Initialize the object using text parameters in
' a comma-delimited list.
Public Sub SetParameters(ByVal txt As String)
    On Error GoTo ParamError

    ' Read the parameters and initialize the object.
    Location.Coord(1) = CSng(GetDelimitedToken(txt, ","))
    Location.Coord(2) = CSng(GetDelimitedToken(txt, ","))
    Location.Coord(3) = CSng(GetDelimitedToken(txt, ","))
    Location.Coord(4) = 1

    ' Light intensities.
    Ir = CSng(GetDelimitedToken(txt, ","))
    Ig = CSng(GetDelimitedToken(txt, ","))
    Ib = CSng(txt)

    Exit Sub

ParamError:
    MsgBox "Error initializing light source parameters."
End Sub

' Return the light source's transformed coordinate.
Property Get TransX() As Single
    TransX = Location.Trans(1)
End Property
' Return the light source's original coordinate.
Property Get CoordX() As Single
    CoordX = Location.Coord(1)
End Property

' Return the light source's original coordinate.
Property Get CoordY() As Single
    CoordY = Location.Coord(1)
End Property
' Return the light source's original coordinate.
Property Get CoordZ() As Single
    CoordZ = Location.Coord(1)
End Property
' Return the light source's transformed coordinate.
Property Get TransY() As Single
    TransY = Location.Trans(2)
End Property
' Return the light source's transformed coordinate.
Property Get TransZ() As Single
    TransZ = Location.Trans(3)
End Property

