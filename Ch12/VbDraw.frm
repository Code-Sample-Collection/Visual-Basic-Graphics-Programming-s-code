VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVbDraw 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VbDraw []"
   ClientHeight    =   6585
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHidden 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3840
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.ImageList imlFillStyles 
      Left            =   1920
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":0212
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":0424
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":0636
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":0848
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":0A5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":0C6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":0E7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlDrawStyles 
      Left            =   1200
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1090
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":12A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":14B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":16C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":18D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1AEA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picColorToolbar 
      Align           =   2  'Align Bottom
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   9315
      TabIndex        =   2
      Top             =   5730
      Width           =   9375
      Begin MSComctlLib.ImageCombo icbDrawStyle 
         Height          =   330
         Left            =   1920
         TabIndex        =   7
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         ImageList       =   "imlDrawStyles"
      End
      Begin VB.PictureBox picBackColorSample 
         Height          =   255
         Index           =   0
         Left            =   840
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox picForeColorSample 
         Height          =   255
         Index           =   0
         Left            =   840
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   5
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picBackColor 
         AutoRedraw      =   -1  'True
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox picForeColor 
         AutoRedraw      =   -1  'True
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   3
         Top             =   120
         Width           =   495
      End
      Begin MSComctlLib.ImageCombo icbFillStyle 
         Height          =   330
         Left            =   1920
         TabIndex        =   8
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         ImageList       =   "imlFillStyles"
      End
      Begin MSComctlLib.ImageCombo icbDrawWidth 
         Height          =   330
         Left            =   3120
         TabIndex        =   9
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         ImageList       =   "imlDrawWidths"
      End
   End
   Begin VB.PictureBox picCanvas 
      BackColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   480
      ScaleHeight     =   2355
      ScaleWidth      =   3075
      TabIndex        =   1
      Top             =   0
      Width           =   3135
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   3840
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbrTools 
      Align           =   3  'Align Left
      Height          =   5730
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   10107
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlTools"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Arrow"
            Object.ToolTipText     =   "Select"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Line"
            Object.ToolTipText     =   "Line"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Scribble"
            Object.ToolTipText     =   "Scribble"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Polyline"
            Object.ToolTipText     =   "Polyline"
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Polygon"
            Object.ToolTipText     =   "Polygon"
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Rectangle"
            Object.ToolTipText     =   "Rectangle"
            ImageIndex      =   7
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlTools 
      Left            =   480
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":2032
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":2144
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":2256
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":2368
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":247A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlDrawWidths 
      Left            =   2640
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":258C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":279E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":29B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":2BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":2DD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":2FE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":31F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":340A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":361C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":382E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileOpenSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileSaveBitmapSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveBitmap 
         Caption         =   "Save &Bitmap..."
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileSaveMetafile 
         Caption         =   "Save &Metafile..."
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuFileExitSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
   End
   Begin VB.Menu mnuArrange 
      Caption         =   "&Arrange"
      Begin VB.Menu mnuArrangeSendToFront 
         Caption         =   "&Bring To Front"
         Enabled         =   0   'False
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuArrangeSendToBack 
         Caption         =   "&Send To Back"
         Enabled         =   0   'False
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu mnuTransform 
      Caption         =   "&Transform"
      Begin VB.Menu mnuTransformClear 
         Caption         =   "&Clear Transformations"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTransformRotate 
         Caption         =   "&Rotate..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTransformScale 
         Caption         =   "&Scale..."
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmVbDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The new object we are building.
Private m_NewObject As vbdObject
Private m_ToolKey As String

' The selected object.
Private m_SelectedObjects As Collection

' Undo variables.
Private Const MAX_UNDO = 50
Private m_Snapshots As Collection
Private m_CurrentSnapshot As Integer

' The scene that holds all objects.
Private m_TheScene As vbdObject

' The currently selected colors.
Private m_ForeColor As Integer
Private m_BackColor As Integer

' The name and title of the current file.
Private m_FileName As String
Private m_FileTitle As String

' MRU list file names.
Private m_MruList As Collection

' Indicates the data has changed since load/save.
Private m_DataModified As Boolean

Private Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Private Declare Function CloseMetaFile Lib "gdi32" (ByVal hmf As Long) As Long
Private Declare Function DeleteMetaFile Lib "gdi32" (ByVal hmf As Long) As Long
Private Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As SIZE) As Long
Private Type SIZE
    cx As Long
    cy As Long
End Type

' Arrange the color toolbar.
Private Sub ArrangeColorToolbar()
Dim i As Integer
Dim tf As Single
Dim tb As Single
Dim X As Single
Dim dx As Single

    tf = (picColorToolbar.ScaleHeight - 2 * picForeColor.Height) / 3
    tb = picForeColor.Height + 2 * tf

    ' Arrange the forecolor and backcolor pictures.
    picForeColor.Move tf, tf
    picBackColor.Move tf, tb
    X = picBackColor.Width + 2 * tf
    dx = picForeColorSample(0).Width + tf / 2

    ' Create the color samples.
    For i = 0 To 15
        If i > 0 Then
            Load picForeColorSample(i)
            Load picBackColorSample(i)
        End If

        picForeColorSample(i).Top = tf
        picForeColorSample(i).Left = X
        picForeColorSample(i).BackColor = QBColor(i)
        picForeColorSample(i).Visible = True
        picBackColorSample(i).Top = tb
        picBackColorSample(i).Left = X
        picBackColorSample(i).BackColor = QBColor(i)
        picBackColorSample(i).Visible = True
        X = X + dx
    Next i

    ' Arrange the DrawStyles ImageCombo.
    X = X + dx + tf
    dx = icbFillStyle.Width + tf / 2
    tf = (picColorToolbar.ScaleHeight - 2 * icbDrawStyle.Height) / 3
    tb = icbDrawStyle.Height + 2 * tf
    icbDrawStyle.Top = tf
    icbDrawStyle.Left = X
    Set icbDrawStyle.ImageList = imlDrawStyles
    For i = 1 To 6
        icbDrawStyle.ComboItems.Add i
        icbDrawStyle.ComboItems(i).Image = i
    Next i

    ' Arrange the FillStyles ImageCombo.
    icbFillStyle.Top = tb
    icbFillStyle.Left = X
    Set icbFillStyle.ImageList = imlFillStyles
    For i = 1 To 8
        icbFillStyle.ComboItems.Add i
        icbFillStyle.ComboItems(i).Image = i
    Next i
    X = X + dx

    ' Arrange the DrawWidth ImageCombo.
    icbDrawWidth.Top = tb
    icbDrawWidth.Left = X
    Set icbDrawWidth.ImageList = imlDrawWidths
    For i = 1 To 10
        icbDrawWidth.ComboItems.Add i
        icbDrawWidth.ComboItems(i).Image = i
    Next i
End Sub

' Return True if it is safe to discard the
' current picture.
Private Function DataSafe() As Boolean
    If Not m_DataModified Then
        DataSafe = True
    Else
        Select Case MsgBox("The data has been modified. Do you want to save the changes?", vbYesNoCancel)
            Case vbYes
                mnuFileSave_Click
                DataSafe = Not m_DataModified
            Case vbNo
                DataSafe = True
            Case vbCancel
                DataSafe = False
        End Select
    End If
End Function
' Save the picture.
Private Sub DataSave(ByVal file_name As String, ByVal file_title As String)
Dim fnum As Integer

    On Error GoTo SaveError

    ' Open the file.
    fnum = FreeFile
    Open file_name For Output As fnum

    ' Write the scene serialization into the file.
    Print #fnum, m_TheScene.Serialization

    ' Close the file.
    Close fnum

    ' Update the caption.
    SetFileName file_name, file_title

    m_DataModified = False
    Exit Sub

SaveError:
    MsgBox "Error " & Format$(Err.Number) & _
        " saving file " & file_name & "." & _
        vbCrLf & Err.Description
    Exit Sub
End Sub
' Load the picture.
Private Sub DataLoad(ByVal file_name As String, ByVal file_title As String)
Dim fnum As Integer
Dim txt As String
Dim token_name As String
Dim token_value As String

    On Error GoTo LoadError

    ' Open the file.
    fnum = FreeFile
    Open file_name For Input As fnum

    ' Read the scene serialization from the file.
    txt = Input$(LOF(fnum), fnum)

    ' Close the file.
    Close fnum

    ' Initialize the scene.
    GetNamedToken txt, token_name, token_value
    If token_name <> "vbdScene" Then
        MsgBox "Error loading file " & file_name & "." & _
            vbCrLf & "This is not a VbDraw file."
    Else
        m_TheScene.Serialization = token_value

        ' Update the caption.
        SetFileName file_name, file_title
        m_DataModified = False

        ' Prepare to edit.
        PrepareToEdit
    End If
    Exit Sub

LoadError:
    MsgBox "Error " & Format$(Err.Number) & _
        " loading file " & file_name & "." & _
        vbCrLf & Err.Description
    Exit Sub
End Sub

' Deselect this object.
Private Sub DeselectVbdObject(ByVal target As vbdObject)
Dim obj As vbdObject
Dim i As Integer

    ' Remove the object from the
    ' m_SelectedObjects collection.
    i = 1
    For Each obj In m_SelectedObjects
        If obj Is target Then
            m_SelectedObjects.Remove i
            Exit For
        End If
        i = i + 1
    Next obj

    ' Mark the object as not selected.
    target.Selected = False
End Sub
' Deselect all objects.
Private Sub DeselectAllVbdObjects()
Dim obj As vbdObject

    ' Deselect all selected objects.
    For Each obj In m_SelectedObjects
        obj.Selected = False
    Next obj

    ' Empty the m_SelectedObjects collection.
    Set m_SelectedObjects = New Collection
End Sub

' Enable the appropriate transformation menus.
Private Sub EnableMenusForSelection()
Dim objects_selected As Boolean

    objects_selected = (m_SelectedObjects.Count > 0)
    mnuArrangeSendToFront.Enabled = objects_selected
    mnuArrangeSendToBack.Enabled = objects_selected
    mnuTransformClear.Enabled = objects_selected
    mnuTransformRotate.Enabled = objects_selected
    mnuTransformScale.Enabled = objects_selected
End Sub

' Select the arrow tool.
Private Sub SelectArrowTool()
    ' Make sure the arrow button is pressed.
    tbrTools.Buttons("Arrow").Value = tbrPressed

    ' Prepare to deal with this tool.
    SelectTool "Arrow"
End Sub

' Create an appropriate object for this tool.
Private Sub SelectTool(ByVal Key As String)
Dim new_pgon As vbdPolygon
Dim new_line As vbdLine

    ' Free any previously started object.
    Set m_NewObject = Nothing

    ' Create the new object.
    m_ToolKey = Key
    Select Case m_ToolKey
        Case "Polyline"
            Set m_NewObject = New vbdPolygon
            Set new_pgon = m_NewObject
            new_pgon.IsClosed = False
        Case "Polygon"
            Set m_NewObject = New vbdPolygon
            Set new_pgon = m_NewObject
            new_pgon.IsClosed = True
        Case "Line"
            Set m_NewObject = New vbdLine
            Set new_line = m_NewObject
            new_line.IsBox = False
        Case "Rectangle"
            Set m_NewObject = New vbdLine
            Set new_line = m_NewObject
            new_line.IsBox = True
        Case "Scribble"
            Set m_NewObject = New vbdScribble
'        Case "Ellipse"
'            Set m_NewObject = New vbdEllipse
    End Select

    ' Let the new object receive picCanvas events.
    If Not (m_NewObject Is Nothing) Then
        Set m_NewObject.Canvas = picCanvas
    End If
End Sub
' Select this object.
Private Sub SelectVbdObject(ByVal target As vbdObject)
    ' See if it is aleady selected.
    If target.Selected Then Exit Sub

    ' Add the object to the
    ' m_SelectedObjects collection.
    m_SelectedObjects.Add target

    ' Mark the object as selected.
    target.Selected = True
End Sub


' Find the object at this position.
Private Function FindObjectAt(ByVal X As Single, ByVal Y As Single) As vbdObject
Dim the_scene As vbdScene

    Set the_scene = m_TheScene
    Set FindObjectAt = the_scene.FindObjectAt(X, Y)
End Function
' Add this file name to the MRU list.
Private Sub MruAddName(ByVal file_name As String)
Dim i As Integer

    ' Remove any duplicates.
    For i = m_MruList.Count To 1 Step -1
        If m_MruList(i) = file_name Then
            m_MruList.Remove i
        End If
    Next i

    ' Add the new name at the front.
    If m_MruList.Count = 0 Then
        m_MruList.Add file_name
    Else
        m_MruList.Add file_name, , 1
    End If

    ' Only keep 4.
    Do While m_MruList.Count > 4
        m_MruList.Remove 5
    Loop

    ' Save the MRU list in the registry.
    For i = 1 To m_MruList.Count
        SaveSetting App.Title, "MRU", _
            Format$(i), m_MruList(i)
    Next i
    For i = m_MruList.Count + 1 To 4
        SaveSetting App.Title, "MRU", _
            Format$(i), ""
    Next i

    ' Display the MRU list.
    MruDisplay
End Sub
' Display the MRU list.
Private Sub MruDisplay()
Dim i As Integer

    mnuFileMRU(0).Visible = (m_MruList.Count > 0)
    For i = 1 To m_MruList.Count
        If i > mnuFileMRU.UBound Then
            Load mnuFileMRU(i)
        End If
        mnuFileMRU(i).Caption = "&" & _
            Format$(i) & " " & m_MruList(i)
        mnuFileMRU(i).Visible = True
    Next i
End Sub
' Load the MRU list.
Private Sub MruLoad()
Dim i As Integer
Dim file_name As String

    Set m_MruList = New Collection
    For i = 1 To 4
        file_name = GetSetting(App.Title, "MRU", _
            Format$(i), "")
        If Len(file_name) > 0 Then
            m_MruList.Add file_name
        End If
    Next i

    ' Display the list.
    MruDisplay
End Sub

' Select default values and prepare to edit.
Private Sub PrepareToEdit()
    ' Select default colors.
    picForeColorSample_Click 0  ' Black
    picbackColorSample_Click 7  ' Gray

    ' Save the initial snapshot.
    Set m_Snapshots = New Collection
    m_CurrentSnapshot = 0
    SaveSnapshot

    ' Start at normal (pixel) scale.
    picCanvas.ScaleMode = vbPixels

    ' Select the arrow tool.
    tbrTools.Buttons("Arrow").Value = tbrPressed

    ' Select the solid DrawStyle.
    icbDrawStyle.SelectedItem = icbDrawStyle.ComboItems(1)

    ' Select the solid FillStyle.
    icbFillStyle.SelectedItem = icbDrawStyle.ComboItems(1)

    ' Select the 1 pixel DrawWidth.
    icbDrawWidth.SelectedItem = icbDrawStyle.ComboItems(1)

    ' Redraw.
    picCanvas.Refresh
End Sub
' Flag the data as modified.
Private Sub SetDirty()
    If Not m_DataModified Then
        Caption = App.Title & "*[" & m_FileTitle & "]"
    End If

    ' Save the current snapshot.
    SaveSnapshot

    m_DataModified = True
End Sub

' Set the file's name.
Private Sub SetFileName(ByVal file_name As String, ByVal file_title As String)
    ' Save the file's name and title.
    m_FileName = file_name
    m_FileTitle = file_title
    mnuFileSave.Enabled = Len(m_FileTitle) > 0

    ' Update the caption.
    Caption = App.Title & " [" & m_FileTitle & "]"

    ' Add the name to the MRU list.
    If Len(m_FileName) > 0 Then MruAddName m_FileName
End Sub
' Enable or disable the undo and redo menus.
Private Sub SetUndoMenus()
    mnuEditUndo.Enabled = (m_CurrentSnapshot > 1)
    mnuEditRedo.Enabled = (m_CurrentSnapshot < m_Snapshots.Count)
End Sub

' Save a snapshot for undo.
Private Sub SaveSnapshot()
    ' Remove any previously undone snapshots.
    Do While m_Snapshots.Count > m_CurrentSnapshot
        m_Snapshots.Remove m_Snapshots.Count
    Loop

    ' Save the current snapshot.
    m_Snapshots.Add m_TheScene.Serialization
    If m_Snapshots.Count > MAX_UNDO + 1 Then
        m_Snapshots.Remove 1
    End If
    m_CurrentSnapshot = m_Snapshots.Count

    ' Enable/disable the undo and redo menus.
    SetUndoMenus
End Sub

' Add this object to the collection.
Public Sub AddObject(ByVal obj As vbdObject)
Dim the_scene As vbdScene

    ' Give the object its drawing properties.
    obj.ForeColor = QBColor(m_ForeColor)
    obj.FillColor = QBColor(m_BackColor)
    obj.DrawStyle = icbDrawStyle.SelectedItem.Index - 1
    obj.FillStyle = icbFillStyle.SelectedItem.Index - 1
    obj.DrawWidth = icbDrawWidth.SelectedItem.Index

    ' Save the new object.
    Set the_scene = m_TheScene
    the_scene.SceneObjects.Add obj
    Set m_NewObject = Nothing

    ' Select the new object only.
    DeselectAllVbdObjects
    SelectVbdObject obj

    ' See if any objects are selected.
    EnableMenusForSelection

    ' Select the arrow tool.
    SelectArrowTool

    ' The data has changed.
    SetDirty

    ' Redraw.
    picCanvas.Refresh
End Sub
' Cancel adding an object to the collection.
Public Sub CancelObject()
    Set m_NewObject = Nothing

    ' Select the arrow tool.
    SelectArrowTool
End Sub

' Restore the previous snapshot.
Private Sub Undo()
Dim token_name As String
Dim token_value As String

    If m_CurrentSnapshot <= 1 Then Exit Sub

    ' Restore the previous snapshot.
    m_CurrentSnapshot = m_CurrentSnapshot - 1
    GetNamedToken m_Snapshots(m_CurrentSnapshot), token_name, token_value
    m_TheScene.Serialization = token_value

    ' Display the scene.
    picCanvas.Refresh

    ' Enable/disable the undo and redo menus.
    SetUndoMenus
End Sub
' Reapply a previously undone snapshot.
Private Sub Redo()
Dim token_name As String
Dim token_value As String

    If m_CurrentSnapshot >= m_Snapshots.Count Then Exit Sub

    ' Restore the previous snapshot.
    m_CurrentSnapshot = m_CurrentSnapshot + 1

    GetNamedToken m_Snapshots(m_CurrentSnapshot), token_name, token_value
    m_TheScene.Serialization = token_value

    ' Display the scene.
    picCanvas.Refresh

    ' Enable/disable the undo and redo menus.
    SetUndoMenus
End Sub

' Process key presses.
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim the_scene As vbdScene

    Select Case KeyCode
        Case vbKeyDelete
            If m_SelectedObjects.Count > 0 Then
                ' Delete the selected objects.
                Set the_scene = m_TheScene
                the_scene.RemoveObjects m_SelectedObjects

                ' The data has changed.
                SetDirty
                picCanvas.Refresh
            End If
    End Select
End Sub

Private Sub Form_Load()
    picHidden.Visible = False
    picHidden.AutoRedraw = True
    picHidden.ScaleMode = vbPixels
    picHidden.BackColor = vbWhite
    picHidden.BorderStyle = vbFixedSingle

    ' Load the MRU list.
    MruLoad

    ' Prepare the dialog.
    dlgFile.CancelError = True
    dlgFile.InitDir = App.Path

    ' Arrange the color toolbar.
    ArrangeColorToolbar

    ' Start a new picture.
    mnuFileNew_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = (Not DataSafe())
End Sub


Private Sub Form_Resize()
Dim wid As Single
Dim hgt As Single

    wid = ScaleWidth - tbrTools.Width
    If wid < 120 Then wid = 120
    hgt = ScaleHeight - picColorToolbar.Height
    If hgt < 120 Then hgt = 120
    picCanvas.Move tbrTools.Width, 0, wid, hgt
End Sub


' Move this object to the front of the scene's
' object list.
Private Sub mnuArrangeSendToBack_Click()
Dim the_scene As vbdScene

    Set the_scene = m_TheScene
    the_scene.MoveToBack m_SelectedObjects

    ' The data has changed.
    SetDirty
    picCanvas.Refresh
End Sub
' Move this object to the front of the scene's
' object list.
Private Sub mnuArrangeSendToFront_Click()
Dim the_scene As vbdScene

    Set the_scene = m_TheScene
    the_scene.MoveToFront m_SelectedObjects

    ' The data has changed.
    SetDirty
    picCanvas.Refresh
End Sub

Private Sub mnuEditRedo_Click()
    Redo
End Sub

Private Sub mnuEditUndo_Click()
    Undo
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub

' Load the selected file.
Private Sub mnuFileMRU_Click(Index As Integer)
Dim pos As Integer
Dim file_title As String

    If Not DataSafe() Then Exit Sub

    pos = InStrRev(m_MruList(Index), "\")
    file_title = Mid$(m_MruList(Index), pos + 1)
    DataLoad m_MruList(Index), file_title
End Sub

' Start a new picture.
Private Sub mnuFileNew_Click()
    If Not DataSafe() Then Exit Sub

    ' Create a new, empty scene object.
    Set m_TheScene = New vbdScene

    ' No objects are selected.
    Set m_SelectedObjects = New Collection

    ' Blank the file name.
    SetFileName "", ""

    ' The data has not been modified.
    m_DataModified = False

    ' Prepare to edit.
    PrepareToEdit
End Sub

' Load a file.
Private Sub mnuFileOpen_Click()
Dim file_name As String

    dlgFile.Flags = cdlOFNExplorer Or _
        cdlOFNHideReadOnly Or _
        cdlOFNLongNames Or _
        cdlOFNFileMustExist
    dlgFile.Filter = "VbDraw Files (*.drw)|*.drw|" & _
        "All Files (*.*)|*.*"
    On Error Resume Next
    dlgFile.ShowOpen
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
        Exit Sub
    End If

    file_name = dlgFile.FileName
    dlgFile.InitDir = Left$(file_name, Len(file_name) _
        - Len(dlgFile.FileTitle) - 1)
    DataLoad file_name, dlgFile.FileTitle
End Sub

' Save the data using the current file name.
Private Sub mnuFileSave_Click()
    If Len(m_FileName) = 0 Then
        ' There is no file name. Use Save As.
        mnuFileSaveAs_Click
    Else
        ' Save the data.
        DataSave m_FileName, m_FileTitle
    End If
End Sub
' Save the picture with a new file name.
Private Sub mnuFileSaveAs_Click()
Dim file_name As String

    dlgFile.Flags = cdlOFNExplorer Or _
        cdlOFNHideReadOnly Or _
        cdlOFNLongNames Or _
        cdlOFNOverwritePrompt
    dlgFile.Filter = "VbDraw Files (*.drw)|*.drw|" & _
        "All Files (*.*)|*.*"
    On Error Resume Next
    dlgFile.ShowSave
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
        Exit Sub
    End If

    file_name = dlgFile.FileName
    dlgFile.InitDir = Left$(file_name, Len(file_name) _
        - Len(dlgFile.FileTitle) - 1)
    DataSave file_name, dlgFile.FileTitle
End Sub

' Save a bitmap image.
Private Sub mnuFileSaveBitmap_Click()
Dim old_file_name As String
Dim pos As Integer
Dim file_name As String

    old_file_name = dlgFile.FileName
    pos = InStrRev(old_file_name, ".")
    If pos > 0 Then dlgFile.FileName = Left$(old_file_name, pos) & "bmp"

    dlgFile.Flags = cdlOFNExplorer Or _
        cdlOFNHideReadOnly Or _
        cdlOFNLongNames Or _
        cdlOFNOverwritePrompt
    dlgFile.Filter = "Bitmap Files (*.bmp)|*.bmp|" & _
        "All Files (*.*)|*.*"
    On Error Resume Next
    dlgFile.ShowSave
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
        Exit Sub
    End If

    file_name = dlgFile.FileName
    dlgFile.InitDir = Left$(file_name, Len(file_name) _
        - Len(dlgFile.FileTitle) - 1)

    ' Make picHidden big enough to hold everything.
    picHidden.Width = picCanvas.Width
    picHidden.Height = picCanvas.Height

    ' Erase the picture.
    picHidden.Line (picHidden.ScaleLeft, picHidden.ScaleTop)-Step(picHidden.ScaleWidth, picHidden.ScaleHeight), vbWhite, BF

    ' Deselect all the objects.
    DeselectAllVbdObjects
    picCanvas.Refresh

    ' Draw the bitmap on picHidden.
    m_TheScene.Draw picHidden
    picHidden.Picture = picHidden.Image

    ' Save the bitmap.
    SavePicture picHidden.Picture, file_name

    dlgFile.FileName = old_file_name
End Sub

' Save the objects in a metafile.
Private Sub mnuFileSaveMetafile_Click()
Dim old_file_name As String
Dim pos As Integer
Dim file_name As String
Dim mf_dc As Long
Dim hmf As Long
Dim old_size As SIZE

    old_file_name = dlgFile.FileName
    pos = InStrRev(old_file_name, ".")
    If pos > 0 Then dlgFile.FileName = Left$(old_file_name, pos) & "wmf"

    dlgFile.Flags = cdlOFNExplorer Or _
        cdlOFNHideReadOnly Or _
        cdlOFNLongNames Or _
        cdlOFNOverwritePrompt
    dlgFile.Filter = "Metafiles (*.wmf)|*.wmf|" & _
        "All Files (*.*)|*.*"
    On Error Resume Next
    dlgFile.ShowSave
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
        Exit Sub
    End If

    file_name = dlgFile.FileName
    dlgFile.InitDir = Left$(file_name, Len(file_name) _
        - Len(dlgFile.FileTitle) - 1)

    ' Create the metafile.
    mf_dc = CreateMetaFile(ByVal file_name)
    If mf_dc = 0 Then
        MsgBox "Error creating the metafile.", vbExclamation
        Exit Sub
    End If

    ' Set the metafile's size to something reasonable.
    SetWindowExtEx mf_dc, picCanvas.ScaleWidth, _
        picCanvas.ScaleHeight, old_size

    ' Draw in the metafile.
    m_TheScene.DrawInMetafile mf_dc

    ' Close the metafile.
    hmf = CloseMetaFile(mf_dc)
    If hmf = 0 Then
        MsgBox "Error closing the metafile.", vbExclamation
    End If

    ' Delete the metafile to free resources.
    If DeleteMetaFile(hmf) = 0 Then
        MsgBox "Error deleting the metafile.", vbExclamation
    End If

    dlgFile.FileName = old_file_name
End Sub


' Clear the selected objects' transformations.
Private Sub mnuTransformClear_Click()
Dim obj As vbdObject

    For Each obj In m_SelectedObjects
        obj.ClearTransformation
    Next obj

    ' The data has changed.
    SetDirty
    picCanvas.Refresh
End Sub

' Rotate the selected objects.
Private Sub mnuTransformRotate_Click()
Const PI = 3.14159265

Dim txt As String
Dim angle As Single
Dim xmin As Single
Dim ymin As Single
Dim xmax As Single
Dim ymax As Single
Dim xmid As Single
Dim ymid As Single
Dim obj As vbdObject
Dim M(1 To 3, 1 To 3) As Single

    ' Get the angle of rotation.
    txt = InputBox("Angle (degrees)", "Angle", "")
    txt = Trim$(txt)
    If Len(txt) = 0 Then Exit Sub
    If Not IsNumeric(txt) Then Exit Sub
    angle = CSng(txt) * PI / 180

    ' Bound the selected objects.
    BoundObjects m_SelectedObjects, xmin, ymin, xmax, ymax

    ' Make the transformation matrix.
    xmid = (xmin + xmax) / 2
    ymid = (ymin + ymax) / 2
    m2RotateAround M, angle, xmid, ymid

    ' Add the transformation to the selected objects.
    For Each obj In m_SelectedObjects
        obj.AddTransformation M
    Next obj

    ' The data has changed.
    SetDirty
    picCanvas.Refresh
End Sub
' Let the user scale the selected objects.
Private Sub mnuTransformScale_Click()
Dim user_canceled As Boolean
Dim x_scale As Single
Dim y_scale As Single
Dim xmin As Single
Dim ymin As Single
Dim xmax As Single
Dim ymax As Single
Dim xmid As Single
Dim ymid As Single
Dim obj As vbdObject
Dim M(1 To 3, 1 To 3) As Single

    user_canceled = dlgScale.ShowForm(x_scale, y_scale)
    Unload dlgScale

    ' If the user canceled, do no more.
    If user_canceled Then Exit Sub

    ' Bound the selected objects.
    BoundObjects m_SelectedObjects, xmin, ymin, xmax, ymax

    ' Make the transformation matrix.
    xmid = (xmin + xmax) / 2
    ymid = (ymin + ymax) / 2
    m2ScaleAt M, x_scale, y_scale, xmid, ymid

    ' Add the transformation to the selected objects.
    For Each obj In m_SelectedObjects
        obj.AddTransformation M
    Next obj

    ' The data has changed.
    SetDirty
    picCanvas.Refresh
End Sub
Private Sub picbackColorSample_Click(Index As Integer)
    m_BackColor = Index
    picBackColor.BackColor = picBackColorSample(Index).BackColor
End Sub

' See if we are clicking on an object.
Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim obj As vbdObject

    If Not (m_NewObject Is Nothing) Then Exit Sub

    ' See where we clicked.
    Set obj = FindObjectAt(X, Y)
    If (obj Is Nothing) Then
        ' Deselect all objects.
        DeselectAllVbdObjects
    Else
        ' See if the Shift key is pressed.
        If (Shift And vbShiftMask) Then
            ' Shift is pressed. Toggle this
            ' object's selection.
            If obj.Selected Then
                DeselectVbdObject obj
            Else
                SelectVbdObject obj
            End If
        Else
            ' Shift is not pressed. Select only
            ' this object.
            DeselectAllVbdObjects
            SelectVbdObject obj
        End If
    End If

    ' See if any objects are selected.
    EnableMenusForSelection

    picCanvas.Refresh
End Sub
Private Sub picCanvas_Paint()
    m_TheScene.Draw picCanvas
End Sub


Private Sub picForeColorSample_Click(Index As Integer)
    m_ForeColor = Index
    picForeColor.BackColor = picForeColorSample(Index).BackColor
End Sub

' The user has pressed a button. Prepare to
' handle this kind of object.
Public Sub tbrTools_ButtonClick(ByVal Button As MSComctlLib.Button)
    SelectTool Button.Key
End Sub
