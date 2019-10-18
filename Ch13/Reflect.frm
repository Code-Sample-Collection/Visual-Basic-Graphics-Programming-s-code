VERSION 5.00
Begin VB.Form frmReflect 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Reflect"
   ClientHeight    =   5415
   ClientLeft      =   300
   ClientTop       =   735
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5415
   ScaleWidth      =   9015
   Begin VB.PictureBox Pict 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Index           =   7
      Left            =   6840
      ScaleHeight     =   -10
      ScaleLeft       =   -5
      ScaleMode       =   0  'User
      ScaleTop        =   5
      ScaleWidth      =   10
      TabIndex        =   7
      Top             =   2760
      Width           =   2175
   End
   Begin VB.PictureBox Pict 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Index           =   6
      Left            =   4560
      ScaleHeight     =   -10
      ScaleLeft       =   -5
      ScaleMode       =   0  'User
      ScaleTop        =   5
      ScaleWidth      =   10
      TabIndex        =   6
      Top             =   2760
      Width           =   2175
   End
   Begin VB.PictureBox Pict 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Index           =   5
      Left            =   2280
      ScaleHeight     =   -10
      ScaleLeft       =   -5
      ScaleMode       =   0  'User
      ScaleTop        =   5
      ScaleWidth      =   10
      TabIndex        =   5
      Top             =   2760
      Width           =   2175
   End
   Begin VB.PictureBox Pict 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Index           =   4
      Left            =   0
      ScaleHeight     =   -10
      ScaleLeft       =   -5
      ScaleMode       =   0  'User
      ScaleTop        =   5
      ScaleWidth      =   10
      TabIndex        =   4
      Top             =   2760
      Width           =   2175
   End
   Begin VB.PictureBox Pict 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Index           =   3
      Left            =   6840
      ScaleHeight     =   -10
      ScaleLeft       =   -5
      ScaleMode       =   0  'User
      ScaleTop        =   5
      ScaleWidth      =   10
      TabIndex        =   3
      Top             =   0
      Width           =   2175
   End
   Begin VB.PictureBox Pict 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Index           =   2
      Left            =   4560
      ScaleHeight     =   -10
      ScaleLeft       =   -5
      ScaleMode       =   0  'User
      ScaleTop        =   5
      ScaleWidth      =   10
      TabIndex        =   2
      Top             =   0
      Width           =   2175
   End
   Begin VB.PictureBox Pict 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Index           =   1
      Left            =   2280
      ScaleHeight     =   -10
      ScaleLeft       =   -5
      ScaleMode       =   0  'User
      ScaleTop        =   5
      ScaleWidth      =   10
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
   Begin VB.PictureBox Pict 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Index           =   0
      Left            =   0
      ScaleHeight     =   -10
      ScaleLeft       =   -5
      ScaleMode       =   0  'User
      ScaleTop        =   5
      ScaleWidth      =   10
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Reverse translation"
      Height          =   255
      Index           =   7
      Left            =   6840
      TabIndex        =   15
      Top             =   5040
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Reverse 1st rotation"
      Height          =   255
      Index           =   6
      Left            =   4560
      TabIndex        =   14
      Top             =   5040
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Reverse 2nd rotation"
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   13
      Top             =   5040
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Reflect"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   12
      Top             =   5040
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Rotate into Y axis"
      Height          =   255
      Index           =   3
      Left            =   6840
      TabIndex        =   11
      Top             =   2280
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Rotate into Y-Z plane"
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   10
      Top             =   2280
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Translate to origin"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   9
      Top             =   2280
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Original picture"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   2280
      Width           =   2175
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmReflect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The plane across which to reflect.
Private Const Px = 1
Private Const Py = 0
Private Const Pz = 0
Private Const Dx = 1
Private Const Dy = 1
Private Const Dz = 1

' Location of viewing eye.
Private EyeR As Single
Private EyeTheta As Single
Private EyePhi As Single

' Location of focus point.
Private Const FocusX = 0#
Private Const FocusY = 0#
Private Const FocusZ = 0#

Private Projector(1 To 4, 1 To 4) As Single

' Matrices used for the reflection.
Private trans(1 To 4, 1 To 4) As Single
Private Rot1(1 To 4, 1 To 4) As Single
Private Rot2(1 To 4, 1 To 4) As Single
Private Ref(1 To 4, 1 To 4) As Single
Private Rot2_Inv(1 To 4, 1 To 4) As Single
Private Rot1_Inv(1 To 4, 1 To 4) As Single
Private Trans_Inv(1 To 4, 1 To 4) As Single

Private M(0 To 7) As Transformation
' Create the matrices used when reflecting across
' a plane passing through (p1, p2, p3) with
' normal <n1, n2, n3>.
Private Sub CreateMatrices(ByVal p1 As Single, ByVal p2 As Single, ByVal p3 As Single, ByVal n1 As Single, ByVal n2 As Single, ByVal n3 As Single)
Dim D As Single
Dim L As Single

    ' Translate the plane to the origin.
    m3Translate trans, -p1, -p2, -p3
    m3Translate Trans_Inv, p1, p2, p3

    ' Rotate around Z-axis until the normal is in
    ' the Y-Z plane.
    m3Identity Rot1
    D = Sqr(n1 * n1 + n2 * n2)
    Rot1(1, 1) = n2 / D
    Rot1(1, 2) = n1 / D
    Rot1(2, 1) = -Rot1(1, 2)
    Rot1(2, 2) = Rot1(1, 1)
    
    m3Identity Rot1_Inv
    Rot1_Inv(1, 1) = Rot1(1, 1)
    Rot1_Inv(1, 2) = -Rot1(1, 2)
    Rot1_Inv(2, 1) = -Rot1(2, 1)
    Rot1_Inv(2, 2) = Rot1(2, 2)
    
    ' Rotate around the X-axis until the normal
    ' lies along the Y axis.
    m3Identity Rot2
    L = Sqr(n1 * n1 + n2 * n2 + n3 * n3)
    Rot2(2, 2) = D / L
    Rot2(2, 3) = -n3 / L
    Rot2(3, 2) = -Rot2(2, 3)
    Rot2(3, 3) = Rot2(2, 2)
    
    m3Identity Rot2_Inv
    Rot2_Inv(2, 2) = Rot2(2, 2)
    Rot2_Inv(2, 3) = -Rot2(2, 3)
    Rot2_Inv(3, 2) = -Rot2(3, 2)
    Rot2_Inv(3, 3) = Rot2(3, 3)

    ' Reflect across the X-Z plane.
    m3Identity Ref
    Ref(2, 2) = -1

    ' Put the matrices in the M array.
    m3Identity M(0).M
    m3MatCopy M(1).M, trans
    m3MatCopy M(2).M, Rot1
    m3MatCopy M(3).M, Rot2
    m3MatCopy M(4).M, Ref
    m3MatCopy M(5).M, Rot2_Inv
    m3MatCopy M(6).M, Rot1_Inv
    m3MatCopy M(7).M, Trans_Inv
End Sub
' Let the user change the location of the eye.
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Const Dtheta = PI / 20
    
    Select Case KeyCode
        Case vbKeyLeft
            EyeTheta = EyeTheta - Dtheta
            
        Case vbKeyRight
            EyeTheta = EyeTheta + Dtheta
        
        Case vbKeyUp
            EyePhi = EyePhi - Dtheta
        
        Case vbKeyDown
            EyePhi = EyePhi + Dtheta
        
        Case Else
            Exit Sub
    End Select

    ' Redraw the pictures.
    DrawTheData
End Sub


' Rotate the points in the cube and draw the cube.
Private Sub DrawData(ByVal pic As Object)
Dim i As Integer
Dim x1 As Single
Dim y1 As Single
Dim x2 As Single
Dim y2 As Single
Dim oldwidth As Integer

    ' Compute the projection matrix.
    m3PProject Projector, m3Parallel, EyeR, EyePhi, EyeTheta, FocusX, FocusY, FocusZ, 0, 1, 0

    ' Transform the points.
    TransformAllData Projector

    ' Draw the points.
    pic.Cls
    
    oldwidth = pic.DrawWidth
    For i = 1 To NumSegments
        x1 = Segments(i).fr_tr(1)
        y1 = Segments(i).fr_tr(2)
        x2 = Segments(i).to_tr(1)
        y2 = Segments(i).to_tr(2)
        
        ' Draw the plane's normal in bold.
        If i = 4 Then pic.DrawWidth = 3
        pic.Line (x1, y1)-(x2, y2)
        If i = 4 Then pic.DrawWidth = oldwidth
    Next i

    pic.Refresh
End Sub

Private Sub Form_Load()
    ' Initialize the eye position.
    EyeR = 3
    EyeTheta = PI * 0.4
    EyePhi = PI * 0.1
    
    ' Create the reflection matrices.
    CreateMatrices Px, Py, Pz, Dx, Dy, Dz

    ' Create, project, and draw the data.
    DrawTheData
End Sub
' Draw all the objects.
Private Sub DrawTheData()
Dim i As Integer

    ' Start with fresh data.
    CreateData

    For i = 0 To 7
        ' Apply the next transformation.
        TransformData M(i).M, 4, NumSegments
        SetPoints 4, NumSegments
        
        ' Display the data.
        DrawData Pict(i)
    Next i
End Sub


' Create the data.
Private Sub CreateData()
    ' Start with no data.
    NumSegments = 0
    
    ' Create the axes.
    MakeSegment 0, 0, 0, 5, 0, 0    ' X axis.
    MakeSegment 0, 0, 0, 0, 5, 0    ' Y axis.
    MakeSegment 0, 0, 0, 0, 0, 5    ' Z axis.
    
    ' Create the reflection plane's normal.
    MakeSegment Px, Py, Pz, Px + Dx, Py + Dy, Pz + Dz

    ' Create the edges of the reflection plane.
    MakeSegment 0.5, -2, 2.5, 3.5, -2, -0.5
    MakeSegment 3.5, -2, -0.5, 1.5, 2, -2.5
    MakeSegment 1.5, 2, -2.5, -1.5, 2, 0.5
    MakeSegment -1.5, 2, 0.5, 0.5, -2, 2.5

    ' Create the object to reflect.
    MakeSegment 3, 1, 2, 3, 3, 2
    MakeSegment 3, 3, 2, 4, 3, 2
    MakeSegment 4, 3, 2, 4, 2, 2
    MakeSegment 4, 2, 2, 5, 2, 2
    MakeSegment 5, 2, 2, 5, 1, 2
    MakeSegment 5, 1, 2, 3, 1, 2
    
    MakeSegment 3, 1, 1, 3, 3, 1
    MakeSegment 3, 3, 1, 4, 3, 1
    MakeSegment 4, 3, 1, 4, 2, 1
    MakeSegment 4, 2, 1, 5, 2, 1
    MakeSegment 5, 2, 1, 5, 1, 1
    MakeSegment 5, 1, 1, 3, 1, 1
    
    MakeSegment 3, 1, 1, 3, 1, 2
    MakeSegment 3, 3, 1, 3, 3, 2
    MakeSegment 4, 3, 1, 4, 3, 2
    MakeSegment 4, 2, 1, 4, 2, 2
    MakeSegment 5, 2, 1, 5, 2, 2
    MakeSegment 5, 1, 1, 5, 1, 2
End Sub
