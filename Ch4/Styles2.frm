VERSION 5.00
Begin VB.Form frmStyles2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Styles2"
   ClientHeight    =   4140
   ClientLeft      =   1200
   ClientTop       =   1440
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   207
   ScaleMode       =   2  'Point
   ScaleWidth      =   334.5
End
Attribute VB_Name = "frmStyles2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long

' Draw a string on the form using randomly chosen
' ForeColor, size, bold, and italic values. Start
' the text at Y position min_y and keep it
' between the margins min_x and max_x.
Private Sub RandomStyles(txt As String, min_size As Integer, max_size As Integer, min_x As Single, max_x As Single, min_y As Single)
Dim length As Integer
Dim pos1 As Integer
Dim pos2 As Integer
Dim new_word As String
Dim clr As Long
Dim y As Integer
Dim font_names As Collection
Dim text_metrics As TEXTMETRIC
Dim ascent As Single

    ' Erase the form.
    Cls

    CurrentX = min_x
    y = 0

    ' Make the list of font names.
    Set font_names = New Collection
    font_names.Add "Times New Roman"
    font_names.Add "Courier New"
    font_names.Add "Arial"
    font_names.Add "MS Sans Serif"

    ' Break the string into words.
    length = Len(txt)
    pos1 = 1
    Do
        ' Get the next word.
        pos2 = InStr(pos1, txt, " ")
        If pos2 = 0 Then
            new_word = Mid$(txt, pos1)
        Else
            new_word = Mid$(txt, pos1, pos2 - pos1)
        End If
        pos1 = pos2 + 1

        ' Randomly select a ForeColor.
        clr = QBColor(Int(16 * Rnd))
        If clr = BackColor Then clr = vbBlack
        ForeColor = clr

        ' Randomly pick Font properties.
        ' (The Underline and Strikethrough
        ' properties make things too cluttered.)
        Font.Name = font_names(Int(font_names.Count * Rnd + 1))
        Font.Size = Int((max_size - min_size + 1) * Rnd + min_size)
        Font.Bold = (Int(2 * Rnd) = 1)
        Font.Italic = (Int(2 * Rnd) = 1)

        ' If the word won't fit, start a new line.
        If CurrentX + TextWidth(new_word) > max_x Then
            CurrentX = min_x
            y = y + 1.25 * max_size
        End If

        ' Get the font's metrics.
        GetTextMetrics hdc, text_metrics
        ascent = ScaleY(text_metrics.tmAscent, vbPixels, ScaleMode)

        ' Display the text.
        CurrentY = y + max_size - ascent
        Print new_word; " ";
    Loop While pos2 > 0
End Sub

' Call RandomStyles to redraw the text string.
Private Sub Form_Resize()
Const txt = "If you draw some text, modify the Font object, and then draw more text, the two pieces of text will be displayed in different styles. Similarly you can change a form or picture box's ForeColor property to produce text of different colors."

    RandomStyles txt, 10, 20, 0, ScaleWidth, 0
End Sub
