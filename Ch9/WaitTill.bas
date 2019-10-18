Attribute VB_Name = "WaitFuncs"
Option Explicit

Declare Function GetTickCount Lib "kernel32" () As Long
' Pause until GetTickCount shows the indicated
' time. This is accurate to within one clock tick
' (55 ms), not counting variability in DoEvents
' and Windows itself.
Public Sub WaitTill(next_time As Long)
    Do
        DoEvents
    Loop While GetTickCount() < next_time
End Sub


