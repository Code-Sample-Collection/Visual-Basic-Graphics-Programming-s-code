VERSION 5.00
Begin VB.Form frmListFont 
   Caption         =   "ListFont"
   ClientHeight    =   4020
   ClientLeft      =   2115
   ClientTop       =   1215
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4020
   ScaleWidth      =   4455
   Begin VB.ListBox lstPrinterFonts 
      Height          =   3765
      Left            =   2280
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.ListBox lstScreenFonts 
      Height          =   3765
      Left            =   0
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Printer Fonts"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Screen Fonts"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmListFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Load the font lists.
Private Sub Form_Load()
Dim i1 As Integer
Dim i2 As Integer
Dim tst As Integer

    ' Prevent errors if the printer is offline.
    On Error GoTo SkipPrinter

    ' List the printer fonts.
    For i1 = 0 To Printer.FontCount - 1
        lstPrinterFonts.AddItem Printer.Fonts(i1)
    Next i1

SkipPrinter:
    On Error GoTo 0

    ' List the screen fonts.
    For i2 = 0 To Screen.FontCount - 1
        lstScreenFonts.AddItem Screen.Fonts(i2)
    Next i2

    ' Compare the items in the lists and
    ' select any that are in one list but
    ' missing in the other
    i1 = 0
    i2 = 0
    Do While i1 < lstPrinterFonts.ListCount And _
             i2 < lstScreenFonts.ListCount
        tst = StrComp(lstPrinterFonts.List(i1), lstScreenFonts.List(i2))
        If tst < 0 Then
            ' Form font < Screen font
            lstPrinterFonts.Selected(i1) = True
            i1 = i1 + 1
        ElseIf tst = 0 Then
            ' They match
            i1 = i1 + 1
            i2 = i2 + 1
        Else
            ' Form font > Screen font
            lstScreenFonts.Selected(i2) = True
            i2 = i2 + 1
        End If
    Loop

    Do While i1 < lstPrinterFonts.ListCount
        lstPrinterFonts.Selected(i1) = True
        i1 = i1 + 1
    Loop

    Do While i2 < lstScreenFonts.ListCount
        lstScreenFonts.Selected(i2) = True
        i2 = i2 + 1
    Loop

    lstPrinterFonts.TopIndex = 0
    lstScreenFonts.TopIndex = 0
End Sub
