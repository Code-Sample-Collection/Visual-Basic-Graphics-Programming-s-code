VERSION 5.00
Begin VB.Form frmFMetrics 
   Caption         =   "FMetrics"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmFMetrics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SAMPLE_SIZE = 96
Private Const LABEL_SIZE = 12

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
' Draw text with a box around it and showing
' different text metrics.
Private Sub BoxText(ByVal X As Single, ByVal Y As Single, ByVal txt As String)
Dim text_width As Single
Dim text_height As Single
Dim text_metrics As TEXTMETRIC
Dim extra As Single
Dim internal_leading As Single
Dim descent As Single
Dim ascent As Single
Dim hgt As Single

    ' Draw the text.
    Font.Size = SAMPLE_SIZE
    CurrentX = X
    CurrentY = Y
    Print txt

    ' Get the text's size.
    text_width = TextWidth(txt)
    text_height = TextHeight(txt)

    ' Draw a box around the text.
    Line (X, Y)-Step(text_width, text_height), , B

    ' Get the text metrics.
    GetTextMetrics hdc, text_metrics

    extra = X / 2
    Font.Size = LABEL_SIZE

    ' Draw a line at the internal leading.
    internal_leading = ScaleY(text_metrics.tmInternalLeading, vbPixels, ScaleMode)
    Line (X - extra, Y + internal_leading)-Step(text_width + 2 * extra, 0)
    CurrentY = CurrentY - TextHeight("X") / 2
    CurrentX = CurrentX + 30
    Print "Internal leading"

    ' Draw a line at the ascent.
    ascent = ScaleY(text_metrics.tmAscent, vbPixels, ScaleMode)
    Line (X - extra, Y + ascent)-Step(text_width + 2 * extra, 0)
    CurrentY = CurrentY - TextHeight("X") / 2
    CurrentX = CurrentX + 30
    Print "Ascent"

    ' Draw a line at the descent.
    descent = ScaleY(text_metrics.tmDescent, vbPixels, ScaleMode)
    Line (X - extra, Y + ascent + descent)-Step(text_width + 2 * extra, 0)
    CurrentY = CurrentY - TextHeight("X") / 2
    CurrentX = CurrentX + 30 + TextWidth("Height")
    Print ", Ascent + Descent"

    ' Draw a line at the height.
    hgt = ScaleY(text_metrics.tmHeight, vbPixels, ScaleMode)
    Line (X - extra, Y + hgt)-Step(text_width + 2 * extra, 0)
    CurrentY = CurrentY - TextHeight("X") / 2
    CurrentX = CurrentX + 30
    Print "Height"
End Sub

Private Sub Form_Load()
    ' Make the text permanent.
    AutoRedraw = True

    ' Select a big font.
    Font.Name = "Times New Roman"
    Font.Size = 96

    ' Draw the text.
    BoxText 240, 240, "Mg"
End Sub

