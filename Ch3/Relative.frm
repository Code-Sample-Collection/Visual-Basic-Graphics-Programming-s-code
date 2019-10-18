VERSION 5.00
Begin VB.Form RelativeForm 
   AutoRedraw      =   -1  'True
   Caption         =   "Relative"
   ClientHeight    =   3150
   ClientLeft      =   1950
   ClientTop       =   1620
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   210
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   Begin VB.PictureBox RelativePict 
      AutoRedraw      =   -1  'True
      Height          =   2700
      Left            =   2640
      ScaleHeight     =   2640
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.PictureBox RGBPict 
      AutoRedraw      =   -1  'True
      Height          =   2700
      Left            =   120
      ScaleHeight     =   2640
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Palette Relative RGB"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   3
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "RGB"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   2415
   End
End
Attribute VB_Name = "RelativeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CLOSEST_IN_PALETTE = &H2000000
' Fill picture boxes with shades of color.
Sub FillPictures()
Const NUM_COLS = 16
Const ROWS_PER_COLOR = 4
Const NUM_ROWS = ROWS_PER_COLOR * 3
Const NUM_BOXES = ROWS_PER_COLOR * NUM_COLS

Dim dx As Single
Dim dy As Single
Dim x As Single
Dim y As Single
Dim clr As Integer
Dim dr As Integer
Dim dg As Integer
Dim db As Integer
Dim i As Integer
Dim j As Integer
Dim r As Integer
Dim g As Integer
Dim b As Integer

    dx = RGBPict.ScaleWidth / NUM_COLS
    dy = RGBPict.ScaleHeight / NUM_ROWS
    
    For clr = 1 To 3
        dr = 0
        dg = 0
        db = 0
        Select Case clr
            Case 1  ' Shades of red.
                dr = 255 / NUM_BOXES
            Case 2  ' Shades of green.
                dg = 255 / NUM_BOXES
            Case 3  ' Shades of blue.
                db = 255 / NUM_BOXES
        End Select
        
        r = 0
        g = 0
        b = 0
        For i = 1 To ROWS_PER_COLOR
            x = 0
            For j = 1 To NUM_COLS
                RGBPict.Line (x, y)-Step(dx, dy), _
                    RGB(r, g, b), BF
                RelativePict.Line (x, y)-Step(dx, dy), _
                    RGB(r, g, b) + CLOSEST_IN_PALETTE, BF
                r = r + dr
                g = g + dg
                b = b + db
                x = x + dx
            Next j
            y = y + dy
        Next i
    Next clr
End Sub

Private Sub Form_Load()
    FillPictures
End Sub
