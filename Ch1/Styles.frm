VERSION 5.00
Begin VB.Form frmStyles 
   Caption         =   "Styles"
   ClientHeight    =   4830
   ClientLeft      =   825
   ClientTop       =   1455
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4830
   ScaleWidth      =   8685
   Begin VB.Frame Frame2 
      Caption         =   "ForeColor"
      Height          =   1575
      Index           =   1
      Left            =   0
      TabIndex        =   31
      Top             =   840
      Width           =   2295
      Begin VB.OptionButton optForeColor 
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optForeColor 
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optForeColor 
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optForeColor 
         Caption         =   "Black"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optForeColor 
         Caption         =   "White"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   32
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "FillColor"
      Height          =   1575
      Index           =   0
      Left            =   2400
      TabIndex        =   25
      Top             =   840
      Width           =   2295
      Begin VB.OptionButton optFillColor 
         Caption         =   "White"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton optFillColor 
         Caption         =   "Black"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optFillColor 
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optFillColor 
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optFillColor 
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "FillStyle"
      Height          =   2295
      Index           =   2
      Left            =   2400
      TabIndex        =   15
      Top             =   2520
      Width           =   2295
      Begin VB.OptionButton optFillStyle 
         Caption         =   "vbDiagonalCross"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   1920
         Width           =   1850
      End
      Begin VB.OptionButton optFillStyle 
         Caption         =   "vbFSSolid"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1850
      End
      Begin VB.OptionButton optFillStyle 
         Caption         =   "vbFSTransparent"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Value           =   -1  'True
         Width           =   1850
      End
      Begin VB.OptionButton optFillStyle 
         Caption         =   "vbHorizontalLine"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1850
      End
      Begin VB.OptionButton optFillStyle 
         Caption         =   "vbVerticalLine"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1850
      End
      Begin VB.OptionButton optFillStyle 
         Caption         =   "vbUpwardDiagonal"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1850
      End
      Begin VB.OptionButton optFillStyle 
         Caption         =   "vbCross"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   1850
      End
      Begin VB.OptionButton optFillStyle 
         Caption         =   "vbDownwardDiagonal"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1910
      End
   End
   Begin VB.TextBox txtDrawWidth 
      Height          =   285
      Left            =   840
      MaxLength       =   1
      TabIndex        =   14
      Text            =   "1"
      Top             =   240
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "DrawStyle"
      Height          =   2295
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   2520
      Width           =   2295
      Begin VB.OptionButton optDrawStyle 
         Caption         =   "vbInsideSolid"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton optDrawStyle 
         Caption         =   "vbTransparent"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1455
      End
      Begin VB.OptionButton optDrawStyle 
         Caption         =   "vbDashDotDot"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton optDrawStyle 
         Caption         =   "vbDashDot"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optDrawStyle 
         Caption         =   "vbDot"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optDrawStyle 
         Caption         =   "vbDash"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optDrawStyle 
         Caption         =   "vbSolid"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Object"
      Height          =   615
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   3375
      Begin VB.OptionButton ObjectChoice 
         Caption         =   "Point"
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton ObjectChoice 
         Caption         =   "Box"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton ObjectChoice 
         Caption         =   "Line"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton ObjectChoice 
         Caption         =   "Circle"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      Height          =   4575
      Left            =   4800
      ScaleHeight     =   4515
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "DrawWidth"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   270
      Width           =   855
   End
End
Attribute VB_Name = "frmStyles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ObjectTypes
    objLine = 0
    objBox = 1
    objCircle = 2
    objPoint = 3
End Enum

Private ObjectType As ObjectTypes
Private Rubberbanding As Boolean
Private OldMode As Integer
Private OldStyle As Integer
Private FirstX As Single
Private FirstY As Single
Private LastX As Single
Private LastY As Single
' Make the picCanvas as big as possible.
Private Sub Form_Resize()
Dim wid As Single

    wid = ScaleWidth - picCanvas.Left
    If wid < 120 Then wid = 120
    picCanvas.Move picCanvas.Left, 0, wid, ScaleHeight
End Sub


' Draw an ellipse bounded by a rectangle.
Private Sub DrawEllipse(ByVal obj As Object, ByVal xmin As Single, ByVal ymin As Single, ByVal xmax As Single, ByVal ymax As Single)
Dim cx As Single
Dim cy As Single
Dim wid As Single
Dim hgt As Single
Dim aspect As Single
Dim radius As Single

    ' Find the center.
    cx = (xmin + xmax) / 2
    cy = (ymin + ymax) / 2

    ' Get the ellipse's size.
    wid = xmax - xmin
    hgt = ymax - ymin

    ' Do nothing if the width or height is zero.
    If (wid = 0) Or (hgt = 0) Then Exit Sub

    aspect = hgt / wid

    ' See which dimension is larger.
    If wid > hgt Then
        ' The major axis is horizontal.
        ' Get the radius in custom coordinates.
        radius = wid / 2
    Else
        ' The major axis is vertical.
        ' Get the radius in custom coordinates.
        radius = hgt / 2
    End If

    ' Draw the circle.
    obj.Circle (cx, cy), radius, , , , aspect
End Sub


' Draw the appropriate object.
Private Sub DrawObject(ByVal xmin As Single, ByVal ymin As Single, ByVal xmax As Single, ByVal ymax As Single)
    Select Case ObjectType
        Case objLine
            picCanvas.Line (xmin, ymin)-(xmax, ymax)
        Case objBox
            picCanvas.Line (xmin, ymin)-(xmax, ymax), , B
        Case objCircle
            DrawEllipse picCanvas, xmin, ymin, xmax, ymax
        Case objPoint
            picCanvas.PSet (xmax, ymax)
    End Select
End Sub
' Set the DrawStyle.
Private Sub optDrawStyle_Click(Index As Integer)
    picCanvas.DrawStyle = Index
End Sub

' Set the FillColor.
Private Sub optFillColor_Click(Index As Integer)
    picCanvas.FillColor = optFillColor(Index).ForeColor
End Sub

' Set the FillStyle.
Private Sub optFillStyle_Click(Index As Integer)
    picCanvas.FillStyle = Index
End Sub


' Start a rubberbanding of some sort.
Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Let MouseMove know we are rubberbanding.
    Rubberbanding = True

    ' Save values so we can restore them later.
    OldMode = picCanvas.DrawMode
    OldStyle = picCanvas.DrawStyle
    picCanvas.DrawMode = vbInvert
    If ObjectType = objLine Then
        picCanvas.DrawStyle = vbSolid
    Else
        picCanvas.DrawStyle = vbDot
    End If

    ' Save the starting coordinates.
    FirstX = X
    FirstY = Y

    ' Save the ending coordinates.
    LastX = X
    LastY = Y

    ' Draw the appropriate rubberband object.
    DrawObject FirstX, FirstY, LastX, LastY
End Sub
' Continue rubberbanding.
Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If we are not rubberbanding, do nothing.
    If Not Rubberbanding Then Exit Sub

    ' Erase the previous rubberband object.
    DrawObject FirstX, FirstY, LastX, LastY

    ' Save the new ending coordinates.
    LastX = X
    LastY = Y

    ' Draw the new rubberband object.
    DrawObject FirstX, FirstY, LastX, LastY
End Sub
' Finish rubberbanding and draw the object.
Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If we are not rubberbanding, do nothing.
    If Not Rubberbanding Then Exit Sub

    ' We are no longer rubberbanding.
    Rubberbanding = False

    ' Erase the previous rubberband object.
    DrawObject FirstX, FirstY, LastX, LastY

    ' Restore the original DrawMode and DrawStyle.
    picCanvas.DrawMode = OldMode
    picCanvas.DrawStyle = OldStyle

    ' Draw the final object.
    DrawObject FirstX, FirstY, LastX, LastY
End Sub
' Select the default options.
Private Sub Form_Load()
    optForeColor(0).Value = True
    optFillColor(0).Value = True
    optDrawStyle(picCanvas.DrawStyle).Value = True
    optFillStyle(picCanvas.FillStyle).Value = True
    ObjectChoice(ObjectType).Value = True
    txtDrawWidth.Text = Format$(picCanvas.DrawWidth)
End Sub

' Record the kind of object to draw next.
Private Sub ObjectChoice_Click(Index As Integer)
    ObjectType = Index
End Sub


' Set the ForeColor.
Private Sub optForeColor_Click(Index As Integer)
    picCanvas.ForeColor = optForeColor(Index).ForeColor
End Sub

' Change set DrawWidth.
Private Sub txtDrawWidth_Change()
Dim wid As Integer

    If Not IsNumeric(txtDrawWidth.Text) Then Exit Sub
    
    wid = CInt(txtDrawWidth.Text)
    If wid < 1 Then Exit Sub
    
    picCanvas.DrawWidth = wid
End Sub

' Only allow 1 through 9.
Private Sub txtDrawWidth_KeyPress(KeyAscii As Integer)
    If KeyAscii < Asc(" ") Or _
       KeyAscii > Asc("~") Then Exit Sub
    If KeyAscii >= Asc("1") And _
       KeyAscii <= Asc("9") Then Exit Sub
    Beep
    KeyAscii = 0
End Sub
