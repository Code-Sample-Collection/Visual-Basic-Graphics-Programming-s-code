VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Robot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Control point locations.
Public Cx As Integer    ' Location of top of head.
Public Cy As Integer
Public LShoulderAngle As Single
Public RShoulderAngle As Single
Public LElbowAngle As Single
Public RElbowAngle As Single
Public LHipAngle As Single
Public RHipAngle As Single
Public LKneeAngle As Single
Public RKneeAngle As Single

' Dimensions.
Public UpperArmLength As Single
Public LowerArmLength As Single
Public UpperLegLength As Single
Public LowerLegLength As Single
Public NeckLength As Single
Public TrunkLength As Single
Public BodyLength As Single
Public HeadRadius As Single

' Robot parts.
Public Enum RobotPartEnum
    part_Head
    part_Lelbow
    part_RElbow
    part_LHand
    part_RHand
    part_LKnee
    part_RKnee
    part_LFoot
    part_RFoot
    part_Hips
    part_Neck
    part_Shoulders
    part_MinPart = part_Head
    part_MaxControlPart = part_RFoot
    part_MaxPart = part_Shoulders
End Enum

' Grab handle distances.
Public Enum GrabDistances
    Near = 2
    Near2 = 2 * Near
End Enum
' Create a tween with parameters taken using
' fraction f1 of the first robot and (1 - f1)
' of the second.
Public Sub Tween(ByVal f1 As Single, ByVal robot1 As Robot, ByVal robot2 As Robot)
Dim f2 As Single

    f2 = 1# - f1
    With robot1
        Cx = f1 * .Cx + f2 * robot2.Cx
        Cy = f1 * .Cy + f2 * robot2.Cy
        LShoulderAngle = CombinedAngle(.LShoulderAngle, robot2.LShoulderAngle, f1, f2)
        RShoulderAngle = CombinedAngle(.RShoulderAngle, robot2.RShoulderAngle, f1, f2)
        LElbowAngle = CombinedAngle(.LElbowAngle, robot2.LElbowAngle, f1, f2)
        RElbowAngle = CombinedAngle(.RElbowAngle, robot2.RElbowAngle, f1, f2)
        LHipAngle = CombinedAngle(.LHipAngle, robot2.LHipAngle, f1, f2)
        RHipAngle = CombinedAngle(.RHipAngle, robot2.RHipAngle, f1, f2)
        LKneeAngle = CombinedAngle(.LKneeAngle, robot2.LKneeAngle, f1, f2)
        RKneeAngle = CombinedAngle(.RKneeAngle, robot2.RKneeAngle, f1, f2)
    End With
End Sub
' Convert the angles so they differ by no more
' than PI. Then combine them using fraction f1
' of angle a1 and fraction f2 of a2.
Private Function CombinedAngle(ByVal a1 As Single, ByVal a2 As Single, ByVal f1 As Single, ByVal f2 As Single) As Single
Const PI = 3.14159265
Const PI_TIMES_2 = PI * 2

    If Abs(a1 - a2) > PI Then
        If a1 > a2 Then
            Do
                a1 = a1 - PI_TIMES_2
            Loop While a1 > a2
        Else
            Do
                a2 = a2 - PI_TIMES_2
            Loop While a2 > a1
        End If
    End If

    CombinedAngle = f1 * a1 + f2 * a2
End Function
' Copy another robot's parameters.
Public Sub CopyFrame(from_me As Robot)
    With from_me
        Cx = .Cx
        Cy = .Cy
        LShoulderAngle = .LShoulderAngle
        RShoulderAngle = .RShoulderAngle
        LElbowAngle = .LElbowAngle
        RElbowAngle = .RElbowAngle
        LHipAngle = .LHipAngle
        RHipAngle = .RHipAngle
        LKneeAngle = .LKneeAngle
        RKneeAngle = .RKneeAngle
    End With
End Sub
' Return the position of part of the robot.
Public Sub Position(part As Integer, x As Integer, y As Integer)
    Select Case part
        Case part_Head
            x = Cx
            y = Cy
        Case part_Neck
            x = Cx
            y = Cy + 2 * HeadRadius
        Case part_Shoulders
            x = Cx
            y = Cy + 2 * HeadRadius + NeckLength
        Case part_Lelbow
            x = Cx + _
                UpperArmLength * Cos(LShoulderAngle)
            y = Cy + 2 * HeadRadius + NeckLength - _
                UpperArmLength * Sin(LShoulderAngle)
        Case part_RElbow
            x = Cx + _
                UpperArmLength * Cos(RShoulderAngle)
            y = Cy + 2 * HeadRadius + NeckLength - _
                UpperArmLength * Sin(RShoulderAngle)
        Case part_LHand
            x = Cx + _
                UpperArmLength * Cos(LShoulderAngle) + _
                LowerArmLength * Cos(LElbowAngle)
            y = Cy + 2 * HeadRadius + NeckLength - _
                UpperArmLength * Sin(LShoulderAngle) - _
                LowerArmLength * Sin(LElbowAngle)
        Case part_RHand
            x = Cx + _
                UpperArmLength * Cos(RShoulderAngle) + _
                LowerArmLength * Cos(RElbowAngle)
            y = Cy + 2 * HeadRadius + NeckLength - _
                UpperArmLength * Sin(RShoulderAngle) - _
                LowerArmLength * Sin(RElbowAngle)
        Case part_Hips
            x = Cx
            y = Cy + 2 * HeadRadius + BodyLength
        Case part_LKnee
            x = Cx + _
                UpperLegLength * Cos(LHipAngle)
            y = Cy + 2 * HeadRadius + BodyLength - _
                UpperLegLength * Sin(LHipAngle)
        Case part_RKnee
            x = Cx + _
                UpperLegLength * Cos(RHipAngle)
            y = Cy + 2 * HeadRadius + BodyLength - _
                UpperLegLength * Sin(RHipAngle)
        Case part_LFoot
            x = Cx + _
                UpperLegLength * Cos(LHipAngle) + _
                LowerLegLength * Cos(LKneeAngle)
            y = Cy + 2 * HeadRadius + BodyLength - _
                UpperLegLength * Sin(LHipAngle) - _
                LowerLegLength * Sin(LKneeAngle)
        Case part_RFoot
            x = Cx + _
                UpperLegLength * Cos(RHipAngle) + _
                LowerLegLength * Cos(RKneeAngle)
            y = Cy + 2 * HeadRadius + BodyLength - _
                UpperLegLength * Sin(RHipAngle) - _
                LowerLegLength * Sin(RKneeAngle)
    End Select
End Sub



' Draw the robot.
Public Sub Draw(pic As PictureBox, handles As Boolean)
Dim x1 As Integer
Dim y1 As Integer
Dim x2 As Integer
Dim y2 As Integer
Dim x3 As Integer
Dim y3 As Integer

    ' Draw the head.
    x1 = Cx
    y1 = Cy + HeadRadius
    pic.Circle (x1, y1), HeadRadius
    If handles Then _
        pic.Line (Cx - Near, Cy - Near)- _
                Step(Near2, Near2), , BF

    ' Draw the body.
    y1 = y1 + HeadRadius
    pic.Line (x1, y1)-Step(0, BodyLength)
    
    ' Draw the left arm.
    y1 = y1 + NeckLength
    x2 = x1 + UpperArmLength * Cos(LShoulderAngle)
    y2 = y1 - UpperArmLength * Sin(LShoulderAngle)
    pic.Line (x1, y1)-(x2, y2)
    x3 = x2 + LowerArmLength * Cos(LElbowAngle)
    y3 = y2 - LowerArmLength * Sin(LElbowAngle)
    pic.Line -(x3, y3)
    If handles Then _
        pic.Line (x2 - Near, y2 - Near)- _
                Step(Near2, Near2), , BF
    If handles Then _
        pic.Line (x3 - Near, y3 - Near)- _
                Step(Near2, Near2), , BF

    ' Draw the right arm.
    x2 = x1 + UpperArmLength * Cos(RShoulderAngle)
    y2 = y1 - UpperArmLength * Sin(RShoulderAngle)
    pic.Line (x1, y1)-(x2, y2)
    x3 = x2 + LowerArmLength * Cos(RElbowAngle)
    y3 = y2 - LowerArmLength * Sin(RElbowAngle)
    pic.Line -(x3, y3)
    If handles Then _
        pic.Line (x2 - Near, y2 - Near)- _
                Step(Near2, Near2), , BF
    If handles Then _
        pic.Line (x3 - Near, y3 - Near)- _
                Step(Near2, Near2), , BF

    ' Draw the left leg.
    y1 = y1 + TrunkLength
    x2 = x1 + UpperLegLength * Cos(LHipAngle)
    y2 = y1 - UpperLegLength * Sin(LHipAngle)
    pic.Line (x1, y1)-(x2, y2)
    x3 = x2 + LowerLegLength * Cos(LKneeAngle)
    y3 = y2 - LowerLegLength * Sin(LKneeAngle)
    pic.Line -(x3, y3)
    If handles Then _
        pic.Line (x2 - Near, y2 - Near)- _
                Step(Near2, Near2), , BF
    If handles Then _
        pic.Line (x3 - Near, y3 - Near)- _
                Step(Near2, Near2), , BF

    ' Draw the right leg.
    x2 = x1 + UpperLegLength * Cos(RHipAngle)
    y2 = y1 - UpperLegLength * Sin(RHipAngle)
    pic.Line (x1, y1)-(x2, y2)
    x3 = x2 + LowerLegLength * Cos(RKneeAngle)
    y3 = y2 - LowerLegLength * Sin(RKneeAngle)
    pic.Line -(x3, y3)
    If handles Then _
        pic.Line (x2 - Near, y2 - Near)- _
                Step(Near2, Near2), , BF
    If handles Then _
        pic.Line (x3 - Near, y3 - Near)- _
                Step(Near2, Near2), , BF
End Sub



' Move the control point to this location.
Public Sub MoveControlPoint(part As Integer, Ax As Integer, Ay As Integer, x As Integer, y As Integer)
    Select Case part
        Case part_Head
            Cx = x
            Cy = y
        Case part_Lelbow
            LShoulderAngle = Arctan2(x - Ax, Ay - y)
        Case part_RElbow
            RShoulderAngle = Arctan2(x - Ax, Ay - y)
        Case part_LHand
            LElbowAngle = Arctan2(x - Ax, Ay - y)
        Case part_RHand
            RElbowAngle = Arctan2(x - Ax, Ay - y)
        Case part_LKnee
            LHipAngle = Arctan2(x - Ax, Ay - y)
        Case part_RKnee
            RHipAngle = Arctan2(x - Ax, Ay - y)
        Case part_LFoot
            LKneeAngle = Arctan2(x - Ax, Ay - y)
        Case part_RFoot
            RKneeAngle = Arctan2(x - Ax, Ay - y)
    End Select
End Sub

' Initialize the robot's parameters.
Public Sub SetParameters(x As Integer, y As Integer, ls As Single, rs As Single, le As Single, re As Single, lh As Single, rh As Single, lk As Single, rk As Single)
Const PI = 3.14159265
Const DEG_TO_RAD = PI / 180#

    Cx = x
    Cy = y
    LShoulderAngle = ls * DEG_TO_RAD
    RShoulderAngle = rs * DEG_TO_RAD
    LElbowAngle = le * DEG_TO_RAD
    RElbowAngle = re * DEG_TO_RAD
    LHipAngle = lh * DEG_TO_RAD
    RHipAngle = rh * DEG_TO_RAD
    LKneeAngle = lk * DEG_TO_RAD
    RKneeAngle = rk * DEG_TO_RAD
End Sub

' Return the distance between the top of the head
' and the top of the robot when its height is as
' large as possible.
Property Get HeadRoom()
    HeadRoom = (UpperArmLength + LowerArmLength) - _
        (2 * HeadRadius + NeckLength)
End Property
' Return the maximum possible height the robot
' can have.
Property Get MaxHeight()
    MaxHeight = TrunkLength + UpperArmLength + LowerArmLength + _
                UpperLegLength + LowerLegLength
End Property

' Return the maximum possible width the robot
' can have.
Property Get MaxWidth()
    MaxWidth = 2 * (UpperArmLength + LowerArmLength)
End Property


' Read the configuration data from a file.
Sub FileInput(fnum)
    Input #fnum, Cx, Cy, _
        LShoulderAngle, RShoulderAngle, _
        LElbowAngle, RElbowAngle, _
        LHipAngle, RHipAngle, _
        LKneeAngle, RKneeAngle, _
        UpperArmLength, LowerArmLength, _
        UpperLegLength, LowerLegLength, _
        NeckLength, TrunkLength, _
        BodyLength, HeadRadius
End Sub
' Write the configuration data to a file.
Sub FileWrite(fnum)
    Write #fnum, Cx, Cy, _
        LShoulderAngle, RShoulderAngle, _
        LElbowAngle, RElbowAngle, _
        LHipAngle, RHipAngle, _
        LKneeAngle, RKneeAngle, _
        UpperArmLength, LowerArmLength, _
        UpperLegLength, LowerLegLength, _
        NeckLength, TrunkLength, _
        BodyLength, HeadRadius
End Sub
' Set default dimensions.
Private Sub Class_Initialize()
    UpperArmLength = 40
    LowerArmLength = 40
    UpperLegLength = 40
    LowerLegLength = 40
    NeckLength = 10
    TrunkLength = UpperArmLength * 1.1
    BodyLength = NeckLength + TrunkLength
    HeadRadius = (UpperArmLength - NeckLength) / 2
End Sub


