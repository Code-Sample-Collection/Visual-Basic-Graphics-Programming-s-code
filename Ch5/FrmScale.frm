VERSION 5.00
Begin VB.Form frmFrmScale 
   Caption         =   "FrmScale"
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
Attribute VB_Name = "frmFrmScale"
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
' print the object as large as possible, centered
' in the printable area.
Private Sub SetLargePrinterScale(obj As Object)
Dim owid As Single
Dim ohgt As Single
Dim pwid As Single
Dim phgt As Single
Dim xmid As Single
Dim ymid As Single
Dim s As Single

    ' Get the object's size in twips.
    owid = obj.ScaleX(obj.ScaleWidth, obj.ScaleMode, vbTwips)
    ohgt = obj.ScaleY(obj.ScaleHeight, obj.ScaleMode, vbTwips)

    ' Get the printer's size in twips.
    pwid = Printer.ScaleX(Printer.ScaleWidth, Printer.ScaleMode, vbTwips)
    phgt = Printer.ScaleY(Printer.ScaleHeight, Printer.ScaleMode, vbTwips)

    ' Compare the object and printer aspect ratios.
    If ohgt / owid > phgt / pwid Then
        ' The object is relatively tall and thin.
        ' Use the printer's whole height.
        s = phgt / ohgt ' This is the scale factor.
    Else
        ' The object is relatively short and wide.
        ' Use the printer's whole width.
        s = pwid / owid ' This is the scale factor.
    End If
    
    ' Convert the printer's dimensions into scaled
    ' object coordinates.
    pwid = obj.ScaleX(pwid, vbTwips, obj.ScaleMode) / s
    phgt = obj.ScaleY(phgt, vbTwips, obj.ScaleMode) / s
    
    ' See where the center should be.
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
    DrawPicture Me
End Sub

' Reset the form scale properties so the picture
' fills the whole form.
Private Sub Form_Resize()
    Me.Scale (-1.1, -1.1)-(1.1, 1.1)
    Me.Refresh
End Sub


' Draw the picture on the Printer object.
Private Sub mnuFilePrint_Click()
    MousePointer = vbHourglass
    DoEvents
    
    SetLargePrinterScale Me ' Set scale properties.
    DrawPicture Printer     ' Draw the picture.
    Printer.EndDoc

    MousePointer = vbDefault
End Sub
