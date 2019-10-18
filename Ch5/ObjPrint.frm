VERSION 5.00
Begin VB.Form frmObjPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ObjPrint"
   ClientHeight    =   3090
   ClientLeft      =   2640
   ClientTop       =   1635
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3090
   ScaleWidth      =   3090
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
      End
   End
End
Attribute VB_Name = "frmObjPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Draw a diamond on the form or printer.
Private Sub DrawPicture(obj As Object)
    obj.CurrentX = 1540
    obj.CurrentY = 100
    obj.Line -Step(1440, 1440)
    obj.Line -Step(-1440, 1440)
    obj.Line -Step(-1440, -1440)
    obj.Line -Step(1440, -1440)
End Sub

' Draw the picture on the form.
Private Sub Form_Paint()
    Cls
    DrawPicture Me
End Sub


' Draw the picture on the Printer object.
Private Sub mnuFilePrint_Click()
    MousePointer = vbHourglass
    DoEvents

    DrawPicture Printer
    Printer.EndDoc

    MousePointer = vbDefault
End Sub
