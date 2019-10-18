VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVBRay 
   Appearance      =   0  'Flat
   Caption         =   "VBRay []"
   ClientHeight    =   4215
   ClientLeft      =   1830
   ClientTop       =   1260
   ClientWidth     =   12210
   DrawMode        =   14  'Copy Pen
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4215
   ScaleWidth      =   12210
   Begin VB.PictureBox picTexture 
      Height          =   375
      Index           =   0
      Left            =   2880
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   19
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtObjects 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   13
      Text            =   "VBRay.frx":0000
      Top             =   2280
      Width           =   4395
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   2160
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rendering"
      Height          =   1095
      Index           =   2
      Left            =   2520
      TabIndex        =   9
      Top             =   0
      Width           =   1875
      Begin VB.TextBox txtStep 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   11
         Text            =   "4"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdRender 
         Caption         =   "Render"
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Step"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Display Method"
      Height          =   1935
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2475
      Begin VB.OptionButton optMethod 
         Caption         =   "Ray Tracing"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "Surface Shading"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "Hidden Surface Removal"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   2145
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "Backface Removal"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton optMethod 
         Caption         =   "Wire Frame"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtDepth 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Text            =   "1"
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Depth"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   3
         Top             =   1560
         Width           =   495
      End
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   4440
      ScaleHeight     =   277
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   505
      TabIndex        =   0
      Top             =   0
      Width           =   7635
   End
   Begin VB.Label Label1 
      Caption         =   "Polygons"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   18
      Top             =   1560
      Width           =   700
   End
   Begin VB.Label Label1 
      Caption         =   "Time"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   17
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblPolygons 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblTime 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Objects"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   14
      Top             =   2040
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpenScene 
         Caption         =   "&Open Scene..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSaveScene 
         Caption         =   "&Save Scene..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileSaveSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveBitmap 
         Caption         =   "Save Bitmap..."
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu mnuObj 
      Caption         =   "&Objects"
      Begin VB.Menu mnuObjViewpoint 
         Caption         =   "&Viewpoint"
      End
      Begin VB.Menu mnuObjAmbientLight 
         Caption         =   "&Ambient Light"
      End
      Begin VB.Menu mnuObjLightSource 
         Caption         =   "&Light Source"
      End
      Begin VB.Menu mnuObjSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuObjNormal 
         Caption         =   "&Normal Objects"
         Begin VB.Menu mnuObjSphere 
            Caption         =   "&Sphere"
         End
         Begin VB.Menu mnuObjPlane 
            Caption         =   "&Plane"
         End
         Begin VB.Menu mnuObjDisk 
            Caption         =   "&Disk"
         End
         Begin VB.Menu mnuObjPolygon 
            Caption         =   "Poly&gon"
         End
         Begin VB.Menu mnuObjCylinder 
            Caption         =   "&Cylinder"
         End
         Begin VB.Menu mnuObjCheckerboard 
            Caption         =   "Checker&board"
         End
         Begin VB.Menu mnuObjFace 
            Caption         =   "&Face"
         End
      End
      Begin VB.Menu mnuObjTextured 
         Caption         =   "&Textured Objects"
         Begin VB.Menu mnuObjBumpyShere 
            Caption         =   "&Bumpy Sphere"
         End
         Begin VB.Menu mnuObjMappedCheckerboard 
            Caption         =   "Mapped &Checkerboard"
         End
         Begin VB.Menu mnuObjHoledCheckerboard 
            Caption         =   "&Holed Checkerboard"
         End
      End
   End
   Begin VB.Menu mnuRender 
      Caption         =   "&Render"
      Begin VB.Menu mnuRenderRender 
         Caption         =   "&Render"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuRenderAnimate 
         Caption         =   "&Animate"
         Shortcut        =   {F6}
      End
   End
End
Attribute VB_Name = "frmVBRay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum RenderingMethodTypes
    render_WireFrame = 0
    render_BackfacesRemoved = 1
    render_HiddenSurfacesRemoved = 2
    render_Shaded = 3
    render_RayTracing = 4
End Enum

Private RenderingMethod As RenderingMethodTypes

' Create the objects in the scene.
Private Sub CreateData()
Dim all_objects As String
Dim obj_type As String
Dim obj_parameters As String
Dim light_source As LightSource
Dim Sphere As RaySphere
Dim bumpy_sphere As RayBumpySphere
Dim plane As RayPlane
Dim disk As RayDisk
Dim pgon As RayPolygon
Dim checker As RayCheckerboard
Dim cyl As RayCylinder
Dim face As RayFace
Dim textured_checker As RayMappedCheckerboard
Dim holed_checker As RayHoledCheckerboard

    ' Initialize the ambient light.
    AmbientIr = 0
    AmbientIg = 0
    AmbientIb = 0

    ' Initialize the eye position.
    EyeR = 1000
    EyeTheta = 1.3
    EyePhi = -0.3

    ' Start with new collections.
    Set Objects = New Collection
    Set LightSources = New Collection

    ' Get the objects string. Remove comments and
    ' non-printing characters. Trim.
    all_objects = RemoveComments(txtObjects.Text)
    all_objects = NonPrintingToSpace(all_objects)
    all_objects = Trim$(all_objects)

    ' Parse the objects.
    Do While Len(all_objects) > 0
        obj_type = LCase$(GetDelimitedToken(all_objects, "("))
        obj_parameters = GetDelimitedToken(all_objects, ")")

        Select Case obj_type
            Case "viewpoint"
                ' Get the current eye location.
                EyeR = CSng(GetDelimitedToken(obj_parameters, ","))
                EyeTheta = CSng(GetDelimitedToken(obj_parameters, ","))
                EyePhi = CSng(obj_parameters)

            Case "ambientlight"
                ' Get the ambient light values.
                AmbientIr = CSng(GetDelimitedToken(obj_parameters, ","))
                AmbientIg = CSng(GetDelimitedToken(obj_parameters, ","))
                AmbientIb = CSng(obj_parameters)

            Case "lightsource"
                ' Make a lighht source.
                Set light_source = New LightSource
                light_source.SetParameters obj_parameters
                LightSources.Add light_source

            Case "sphere"
                ' Make a sphere.
                Set Sphere = New RaySphere
                Sphere.SetParameters obj_parameters
                Objects.Add Sphere

            Case "plane"
                ' Make a plane.
                Set plane = New RayPlane
                plane.SetParameters obj_parameters
                Objects.Add plane

            Case "disk"
                ' Make a disk.
                Set disk = New RayDisk
                disk.SetParameters obj_parameters
                Objects.Add disk

            Case "polygon"
                ' Make a polygon.
                Set pgon = New RayPolygon
                pgon.SetParameters obj_parameters
                Objects.Add pgon

            Case "checkerboard"
                Set checker = New RayCheckerboard
                checker.SetParameters obj_parameters
                Objects.Add checker

            Case "cylinder"
                Set cyl = New RayCylinder
                cyl.SetParameters obj_parameters
                Objects.Add cyl

            Case "face"
                Set face = New RayFace
                face.SetParameters obj_parameters
                Objects.Add face

            Case "bumpysphere"
                ' Make a bumy sphere.
                Set bumpy_sphere = New RayBumpySphere
                bumpy_sphere.SetParameters obj_parameters
                Objects.Add bumpy_sphere

            Case "mappedcheckerboard"
                Set textured_checker = New RayMappedCheckerboard
                Load picTexture(picTexture.UBound + 1)
                textured_checker.SetParameters picTexture(picTexture.UBound), obj_parameters
                Objects.Add textured_checker

            Case "holedcheckerboard"
                Set holed_checker = New RayHoledCheckerboard
                Load picTexture(picTexture.UBound + 1)
                holed_checker.SetParameters picTexture(picTexture.UBound), obj_parameters
                Objects.Add holed_checker

            Case Else
                MsgBox "Unknown object type " & obj_type
        End Select
    Loop
End Sub
' Project and draw all the objects.
Private Sub RenderObjects(ByVal pic As Object, ByVal lblPolygons As Label)
Dim start_time As Single
Dim ellapsed As Single

    lblPolygons.Caption = ""
    lblTime.Caption = ""

    ' Create the data.
    CreateData
    If Objects.Count < 1 Then Exit Sub

    ' Focus on the origin.
    FocusX = 0#
    FocusY = 0#
    FocusZ = 0#

    ' Create a background color.
    BackR = 0
    BackG = 0
    BackB = 0

    ' Fill with another color so we can see progress.
    pic.Line (pic.ScaleLeft, pic.ScaleTop)- _
        Step(pic.ScaleWidth, pic.ScaleHeight), _
        RGB(0, 0, &H80), BF

    ' Display the data.
    start_time = Timer
    Select Case RenderingMethod
        Case render_WireFrame
            RenderWireFrame pic
        Case render_BackfacesRemoved
            RenderBackfacesRemoved pic
        Case render_HiddenSurfacesRemoved
            RenderHiddenSurfacesRemoved pic, lblPolygons
        Case render_Shaded
            RenderShaded pic, lblPolygons
        Case render_RayTracing
            RenderRayTracing pic, CInt(txtStep.Text), CInt(txtDepth.Text)
    End Select
    ellapsed = Timer - start_time
    lblTime.Caption = Format$(ellapsed \ 60) & _
        ":" & Format$(ellapsed Mod 60, "00")
End Sub
' Animate the objects.
Private Sub AnimateObjects(ByVal pic As PictureBox, ByVal lblPolygons As Label)
Const NUM_FRAMES = 50

Dim start_time As Single
Dim ellapsed As Single
Dim theta As Single
Dim dtheta As Single
Dim i As Integer
Dim file_base As String

    lblPolygons.Caption = ""
    lblTime.Caption = ""

    file_base = "C:\Temp\frame"

    start_time = Timer
    dtheta = 2 * PI / NUM_FRAMES
    Do While theta < 2 * PI - 0.1
        ' Create the data.
        CreateData
        If Objects.Count < 1 Then Exit Sub

        ' Focus on the origin.
        FocusX = 0#
        FocusY = 0#
        FocusZ = 0#

        ' Create a background color.
        BackR = 0
        BackG = 0
        BackB = 0

        EyeTheta = theta

        ' Fill with another color so we can see progress.
        pic.Line (pic.ScaleLeft, pic.ScaleTop)- _
            Step(pic.ScaleWidth, pic.ScaleHeight), _
            RGB(0, 0, &H80), BF

        ' Display the data.
        Select Case RenderingMethod
            Case render_WireFrame
                RenderWireFrame pic
            Case render_BackfacesRemoved
                RenderBackfacesRemoved pic
            Case render_HiddenSurfacesRemoved
                RenderHiddenSurfacesRemoved pic, lblPolygons
            Case render_Shaded
                RenderShaded pic, lblPolygons
            Case render_RayTracing
                RenderRayTracing pic, CInt(txtStep.Text), CInt(txtDepth.Text)
        End Select

        Set pic.Picture = pic.Image
        SavePicture pic.Picture, file_base & Format$(i, "00") & ".bmp"
        i = i + 1

        theta = theta + dtheta
        DoEvents
        If Not Running Then Exit Do
    Loop

    ellapsed = Timer - start_time
    lblTime.Caption = Format$(ellapsed \ 60) & _
        ":" & Format$(ellapsed Mod 60, "00")
End Sub
' Make the common dialog control's file name have the
' indicated extension.
Private Sub SetDialogExtension(ByVal dlg As Control, ByVal extension As String)
Dim pos As Integer

    pos = InStrRev(dlg.FileName, ".")
    If pos > 0 Then
        dlg.FileName = Left$(dlg.FileName, pos) & extension
    End If
End Sub

' Render the objects.
Private Sub cmdRender_Click()
    If Running Then
        Running = False
        cmdRender.Caption = "Stopped"
        cmdRender.Enabled = False
        DoEvents
    Else
        Running = True
        cmdRender.Caption = "Stop"
        MousePointer = vbHourglass
        DoEvents

        ' Render the objects.
        RenderObjects picCanvas, lblPolygons

        MousePointer = vbDefault
        cmdRender.Enabled = True
        cmdRender.Caption = "Render"
        Running = False
        Beep
    End If
End Sub
Private Sub Form_Load()
    dlgFile.InitDir = App.Path

    optMethod(0).value = True
End Sub

Private Sub Form_Resize()
Dim hgt As Single

#If False Then
Dim wid As Single
    wid = ScaleWidth - picCanvas.Left
    If wid < 120 Then wid = 120
    picCanvas.Width = wid
    picCanvas.Height = ScaleHeight
#End If

    hgt = ScaleHeight - txtObjects.Top
    If hgt < 120 Then hgt = 120
    txtObjects.Height = hgt
End Sub


' Halt immediately in case we're in the middle of
' ray tracing.
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnuFileNew_Click()
    txtObjects.Text = ""
End Sub

Private Sub mnuFileSaveBitmap_Click()
    ' Allow the user to pick a file.
    On Error Resume Next
    dlgFile.Filter = "Bitmaps (*.bmp)|*.bmp|" & _
        "All Files (*.*)|*.*"
    dlgFile.Flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly
    SetDialogExtension dlgFile, "bmp"
    dlgFile.ShowSave
    If Err.Number = cdlCancel Then
        Unload dlgFile
        Exit Sub
    ElseIf Err.Number <> 0 Then
        Unload dlgFile
        Beep
        MsgBox "Error selecting file.", , vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    SavePicture picCanvas.Image, dlgFile.FileName
End Sub
Private Sub mnuFileOpenScene_Click()
Dim fnum As Integer
Dim file_name As String

    ' Allow the user to pick a file.
    On Error Resume Next
    dlgFile.Filter = "Ray Scenes (*.ray)|*.ray|" & _
        "All Files (*.*)|*.*"
    dlgFile.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    SetDialogExtension dlgFile, "ray"
    dlgFile.ShowOpen
    If Err.Number = cdlCancel Then
        Unload dlgFile
        Exit Sub
    ElseIf Err.Number <> 0 Then
        Unload dlgFile
        Beep
        MsgBox "Error selecting file.", , vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    file_name = Trim$(dlgFile.FileName)
    dlgFile.InitDir = Left$(file_name, Len(file_name) _
        - Len(dlgFile.FileTitle) - 1)

    fnum = FreeFile
    Open file_name For Input As fnum
    txtObjects.Text = Input$(LOF(fnum), fnum)
    Caption = "VBRay [" & dlgFile.FileTitle & "]"
    Close fnum
End Sub
Private Sub mnuFileSaveScene_Click()
Dim fnum As Integer
Dim file_name As String

    ' Allow the user to pick a file.
    On Error Resume Next
    dlgFile.Filter = "Ray Scenes (*.ray)|*.ray|" & _
        "All Files (*.*)|*.*"
    dlgFile.Flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly
    SetDialogExtension dlgFile, "ray"
    dlgFile.ShowSave
    If Err.Number = cdlCancel Then
        Unload dlgFile
        Exit Sub
    ElseIf Err.Number <> 0 Then
        Unload dlgFile
        Beep
        MsgBox "Error selecting file.", , vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    file_name = Trim$(dlgFile.FileName)
    dlgFile.InitDir = Left$(file_name, Len(file_name) _
        - Len(dlgFile.FileTitle) - 1)

    fnum = FreeFile
    Open file_name For Output As fnum
    Print #fnum, txtObjects.Text
    Close fnum
    Caption = "VBRay [" & dlgFile.FileTitle & "]"
End Sub

' Add an ambient light entry to the object list.
Private Sub mnuObjAmbientLight_Click()
    txtObjects.Text = txtObjects.Text & _
        "AmbientLight( Ir, Ig, Ib )" & vbCrLf
    txtObjects.SelStart = Len(txtObjects.Text)
End Sub

Private Sub mnuObjBumpyShere_Click()
    txtObjects.Text = txtObjects.Text & _
        "BumpySphere( Radius, X, Y, Z," & vbCrLf & _
        "  bumpiness,        ' Bumpiness " & vbCrLf & _
        "  ka_r, ka_g, ka_b, ' Ambient" & vbCrLf & _
        "  kd_r, kd_g, kd_b, ' Diffuse" & vbCrLf & _
        "  spec_n, spec_s,   ' Specular" & vbCrLf & _
        "  kr_r, kr_g, kr_b, ' Reflected" & vbCrLf & _
        "  kt_n, n1, n2,     ' TransN, n1, n2" & vbCrLf & _
        "  kt_r, kt_g, kt_b  ' Transmitted" & vbCrLf & _
        ")" & vbCrLf
    txtObjects.SelStart = Len(txtObjects.Text)
End Sub

Private Sub mnuObjCheckerboard_Click()
    txtObjects.Text = txtObjects.Text & _
        "Checkerboard(" & vbCrLf & _
        "  num_squares_1,    ' # squares in 1st direction" & vbCrLf & _
        "  num_squares_2,    ' # squares in 2nd direction" & vbCrLf & _
        "  x1, y1, z1,       ' Point in corner" & vbCrLf & _
        "  x2, y2, z2,       ' Point in 1st corner of square" & vbCrLf & _
        "  x3, y3, z3,       ' Point in 2nd corner of square" & vbCrLf & _
        "  ka_r, ka_g, ka_b, ' Ambient" & vbCrLf & _
        "  kd_r, kd_g, kd_b, ' Diffuse" & vbCrLf & _
        "  spec_n, spec_s,   ' Specular" & vbCrLf & _
        "  kr_r, kr_g, kr_b, ' Reflected" & vbCrLf & _
        "  kt_n, n1, n2,     ' TransN, n1, n2" & vbCrLf & _
        "  kt_r, kt_g, kt_b  ' Transmitted" & vbCrLf & _
        ")" & vbCrLf
    txtObjects.SelStart = Len(txtObjects.Text)
End Sub
Private Sub mnuObjCylinder_Click()
    txtObjects.Text = txtObjects.Text & _
        "Cylinder(radius,    ' Radius" & vbCrLf & _
        "  p1x, p1y, p1z,    ' Point at one end" & vbCrLf & _
        "  p2x, p2y, p2z,    ' Point at other end" & vbCrLf & _
        "  ka_r, ka_g, ka_b, ' Ambient" & vbCrLf & _
        "  kd_r, kd_g, kd_b, ' Diffuse" & vbCrLf & _
        "  spec_n, spec_s,   ' Specular" & vbCrLf & _
        "  kr_r, kr_g, kr_b, ' Reflected" & vbCrLf & _
        "  kt_n, n1, n2,     ' TransN, n1, n2" & vbCrLf & _
        "  kt_r, kt_g, kt_b  ' Transmitted" & vbCrLf & _
        ")" & vbCrLf
    txtObjects.SelStart = Len(txtObjects.Text)
End Sub

Private Sub mnuObjDisk_Click()
    txtObjects.Text = txtObjects.Text & _
        "Disk(radius,        ' Radius" & vbCrLf & _
        "  x, y, z,          ' Point on plane" & vbCrLf & _
        "  Nx, Ny, Nz,       ' Normal vector" & vbCrLf & _
        "  ka_r, ka_g, ka_b, ' Ambient" & vbCrLf & _
        "  kd_r, kd_g, kd_b, ' Diffuse" & vbCrLf & _
        "  spec_n, spec_s,   ' Specular" & vbCrLf & _
        "  kr_r, kr_g, kr_b, ' Reflected" & vbCrLf & _
        "  kt_n, n1, n2,     ' TransN, n1, n2" & vbCrLf & _
        "  kt_r, kt_g, kt_b  ' Transmitted" & vbCrLf & _
        ")" & vbCrLf
    txtObjects.SelStart = Len(txtObjects.Text)
End Sub

Private Sub mnuObjFace_Click()
    txtObjects.Text = txtObjects.Text & _
        "Face(num_points,    ' Number of points" & vbCrLf & _
        "  x1, y1, z1,       ' Point 1" & vbCrLf & _
        "  x2, y2, z2,       ' Point 2" & vbCrLf & _
        "  ...,              ' Other points" & vbCrLf & _
        "  ka_r, ka_g, ka_b, ' Ambient" & vbCrLf & _
        "  kd_r, kd_g, kd_b, ' Diffuse" & vbCrLf & _
        "  spec_n, spec_s,   ' Specular" & vbCrLf & _
        "  kr_r, kr_g, kr_b, ' Reflected" & vbCrLf & _
        "  kt_n, n1, n2,     ' TransN, n1, n2" & vbCrLf & _
        "  kt_r, kt_g, kt_b  ' Transmitted" & vbCrLf & _
        ")" & vbCrLf
    txtObjects.SelStart = Len(txtObjects.Text)
End Sub

Private Sub mnuObjHoledCheckerboard_Click()
Dim file_name As String

    file_name = App.Path
    If Right$(file_name, 1) <> "\" Then file_name = file_name & "\"
    file_name = file_name & "filename.bmp"

    txtObjects.Text = txtObjects.Text & _
        "HoledCheckerboard(" & vbCrLf & _
        "  ' The texture file name (absolute or relative)" & vbCrLf & _
        "  " & file_name & "," & vbCrLf & _
        "  num_squares_1,    ' # squares in 1st direction" & vbCrLf & _
        "  num_squares_2,    ' # squares in 2nd direction" & vbCrLf & _
        "  x1, y1, z1,       ' Point in corner" & vbCrLf & _
        "  x2, y2, z2,       ' Point in 1st corner of square" & vbCrLf & _
        "  x3, y3, z3,       ' Point in 2nd corner of square" & vbCrLf & _
        "  ambient_factor,   ' Scale factor for ambient values" & vbCrLf & _
        "  diffuse_factor,   ' Scale factor for diffuse values" & vbCrLf & _
        "  spec_n, spec_s,   ' Specular" & vbCrLf & _
        "  kr_r, kr_g, kr_b, ' Reflected" & vbCrLf & _
        "  kt_n, n1, n2,     ' TransN, n1, n2" & vbCrLf & _
        "  kt_r, kt_g, kt_b  ' Transmitted" & vbCrLf & _
        ")" & vbCrLf
    txtObjects.SelStart = Len(txtObjects.Text)
End Sub

' Add a light source to the object list.
Private Sub mnuObjLightSource_Click()
    txtObjects.Text = txtObjects.Text & _
        "LightSource( X, Y, Z," & vbCrLf & _
        "  Ir, Ig, Ib )" & vbCrLf
    txtObjects.SelStart = Len(txtObjects.Text)
End Sub

' Add a plane to the object list.
Private Sub mnuObjPlane_Click()
    txtObjects.Text = txtObjects.Text & _
        "Plane( x, y, z,     ' Point on plane" & vbCrLf & _
        "  Nx, Ny, Nz,       ' Normal vector" & vbCrLf & _
        "  ka_r, ka_g, ka_b, ' Ambient" & vbCrLf & _
        "  kd_r, kd_g, kd_b, ' Diffuse" & vbCrLf & _
        "  spec_n, spec_s,   ' Specular" & vbCrLf & _
        "  kr_r, kr_g, kr_b, ' Reflected" & vbCrLf & _
        "  kt_n, n1, n2,     ' TransN, n1, n2" & vbCrLf & _
        "  kt_r, kt_g, kt_b  ' Transmitted" & vbCrLf & _
        ")" & vbCrLf
    txtObjects.SelStart = Len(txtObjects.Text)
End Sub

Private Sub mnuObjPolygon_Click()
    txtObjects.Text = txtObjects.Text & _
        "Polygon(num_points, ' Number of points" & vbCrLf & _
        "  x1, y1, z1,       ' Point 1" & vbCrLf & _
        "  x2, y2, z2,       ' Point 2" & vbCrLf & _
        "  ...,              ' Other points" & vbCrLf & _
        "  ka_r, ka_g, ka_b, ' Ambient" & vbCrLf & _
        "  kd_r, kd_g, kd_b, ' Diffuse" & vbCrLf & _
        "  spec_n, spec_s,   ' Specular" & vbCrLf & _
        "  kr_r, kr_g, kr_b, ' Reflected" & vbCrLf & _
        "  kt_n, n1, n2,     ' TransN, n1, n2" & vbCrLf & _
        "  kt_r, kt_g, kt_b  ' Transmitted" & vbCrLf & _
        ")" & vbCrLf
    txtObjects.SelStart = Len(txtObjects.Text)
End Sub

' Add a sphere to the object list.
Private Sub mnuObjSphere_Click()
    txtObjects.Text = txtObjects.Text & _
        "Sphere( Radius, X, Y, Z," & vbCrLf & _
        "  ka_r, ka_g, ka_b, ' Ambient" & vbCrLf & _
        "  kd_r, kd_g, kd_b, ' Diffuse" & vbCrLf & _
        "  spec_n, spec_s,   ' Specular" & vbCrLf & _
        "  kr_r, kr_g, kr_b, ' Reflected" & vbCrLf & _
        "  kt_n, n1, n2,     ' TransN, n1, n2" & vbCrLf & _
        "  kt_r, kt_g, kt_b  ' Transmitted" & vbCrLf & _
        ")" & vbCrLf
    txtObjects.SelStart = Len(txtObjects.Text)
End Sub

Private Sub mnuObjMappedCheckerboard_Click()
Dim file_name As String

    file_name = App.Path
    If Right$(file_name, 1) <> "\" Then file_name = file_name & "\"
    file_name = file_name & "filename.bmp"

    txtObjects.Text = txtObjects.Text & _
        "MappedCheckerboard(" & vbCrLf & _
        "  ' The texture file name (absolute or relative)" & vbCrLf & _
        "  " & file_name & "," & vbCrLf & _
        "  num_squares_1,    ' # squares in 1st direction" & vbCrLf & _
        "  num_squares_2,    ' # squares in 2nd direction" & vbCrLf & _
        "  x1, y1, z1,       ' Point in corner" & vbCrLf & _
        "  x2, y2, z2,       ' Point in 1st corner of square" & vbCrLf & _
        "  x3, y3, z3,       ' Point in 2nd corner of square" & vbCrLf & _
        "  ambient_factor,   ' Scale factor for ambient values" & vbCrLf & _
        "  diffuse_factor,   ' Scale factor for diffuse values" & vbCrLf & _
        "  spec_n, spec_s,   ' Specular" & vbCrLf & _
        "  kr_r, kr_g, kr_b, ' Reflected" & vbCrLf & _
        "  kt_n, n1, n2,     ' TransN, n1, n2" & vbCrLf & _
        "  kt_r, kt_g, kt_b  ' Transmitted" & vbCrLf & _
        ")" & vbCrLf
    txtObjects.SelStart = Len(txtObjects.Text)
End Sub

' Add a viewpoint entry to the object list.
Private Sub mnuObjViewpoint_Click()
    txtObjects.Text = txtObjects.Text & _
        "Viewpoint( X, Y, Z )" & vbCrLf
    txtObjects.SelStart = Len(txtObjects.Text)
End Sub

' Animate the objects.
Private Sub mnuRenderAnimate_Click()
    If Running Then
        Running = False
        cmdRender.Caption = "Stopped"
        cmdRender.Enabled = False
        DoEvents
    Else
        Running = True
        cmdRender.Caption = "Stop"
        MousePointer = vbHourglass
        DoEvents

        ' Animate the objects.
        AnimateObjects picCanvas, lblPolygons

        MousePointer = vbDefault
        cmdRender.Enabled = True
        cmdRender.Caption = "Render"
        Running = False
        Beep
    End If
End Sub

Private Sub mnuRenderRender_Click()
    cmdRender_Click
End Sub


Private Sub optMethod_Click(Index As Integer)
    RenderingMethod = Index
End Sub

' Print the coordinates of the point clicked.
' This is useful for debugging.
Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "(" & Format$(X) & ", " & Format$(Y) & ")"
End Sub
