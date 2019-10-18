VERSION 5.00
Begin VB.Form FrmPrintForm 
   Caption         =   "FrmPrint"
   ClientHeight    =   3405
   ClientLeft      =   2640
   ClientTop       =   1635
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3405
   ScaleWidth      =   3405
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
      End
   End
End
Attribute VB_Name = "FrmPrintForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Draw a Bowditch curve on the indicated object.
Private Sub DrawPicture(obj As Object)
Const PI = 3.14159265

Dim x As Integer
Dim y As Integer
Dim t As Single
Dim maxt As Single
Dim dt As Single
    
    ' Draw the curve.
    maxt = PI * 8
    dt = maxt / 200
    obj.CurrentX = 0
    obj.CurrentY = 0
    For t = dt To maxt + dt / 2 Step dt
        obj.Line -(Sin(0.75 * t), Sin(t))
    Next t
End Sub


' Set the printer's scale properties so it will
' print the object at the correct size, centered
' in the printable area.
Private Sub SetPrinterScale(obj As Object)
Dim pwid As Single
Dim phgt As Single
Dim xmid As Single
Dim ymid As Single
        
    ' Get the printer's dimensions in twips.
    pwid = Printer.ScaleX(Printer.ScaleWidth, Printer.ScaleMode, vbTwips)
    phgt = Printer.ScaleY(Printer.ScaleHeight, Printer.ScaleMode, vbTwips)
    
    ' Convert the printer's dimensions into the
    ' object's coordinates.
    pwid = obj.ScaleX(pwid, vbTwips, obj.ScaleMode)
    phgt = obj.ScaleY(phgt, vbTwips, obj.ScaleMode)
    
    ' Compute the center of the object.
    xmid = obj.ScaleLeft + obj.ScaleWidth / 2
    ymid = obj.ScaleTop + obj.ScaleHeight / 2
    
    ' Pass the coordinates of the upper left and
    ' lower right corners into the Scale method.
    Printer.Scale _
        (xmid - pwid / 2, ymid - phgt / 2)- _
        (xmid + pwid / 2, ymid + phgt / 2)
End Sub
' Draw the picture on the form.
Private Sub Form_Paint()
    Cls
    DrawPicture Me
End Sub

' Reset the form scale properties so the picture
' fills the form.
Private Sub Form_Resize()
    Me.Scale (-1.1, -1.1)-(1.1, 1.1)
    Me.Refresh
End Sub


' Draw the picture on the Printer object.
Private Sub mnuFilePrint_Click()
    MousePointer = vbHourglass
    DoEvents
    
    ' Set the printer's scale properties.
    SetPrinterScale Me

    ' Draw the picture.
    DrawPicture Printer
    Printer.EndDoc

    MousePointer = vbDefault
End Sub
