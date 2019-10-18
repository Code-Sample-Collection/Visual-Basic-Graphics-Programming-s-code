Attribute VB_Name = "HiRes"
Option Explicit

Public Enum hiresConstants
    hires_Normal
    hires_StretchToFit
    hires_ResizePrinter
End Enum

' The last font we tried.
Private LastFont As StdFont

Private orig_draw_style As Integer
Private orig_draw_width As Integer
Private orig_fill_style As Integer
Private orig_fore_color As Long

' Compute the offsets required by this control due
' to its containers.
Private Sub ComputeOffsets(ctl As Control, prn As Object, l As Single, t As Single, w As Single, h As Single)
Dim p As Object
Dim c As Object

    Set p = ctl.Parent
    Set c = ctl.Container
    If TypeOf ctl Is Line Then
        l = 0
        t = 0
        w = 0
        h = 0
    Else
        l = ctl.Left
        t = ctl.Top
        w = ctl.Width
        h = ctl.Height
    End If
    Do While Not (c Is p)
        Select Case TypeName(c)
            Case "PictureBox"
                If l + w > c.ScaleWidth Then w = c.ScaleWidth - l
                If t + h > c.ScaleHeight Then h = c.ScaleHeight - t
            Case "Frame"
                If l + w > c.Width Then w = c.Width - l
                If t + h > c.Height Then h = c.Height - t
        End Select
        l = l + c.Left
        t = t + c.Top
        Set c = c.Container
    Loop
End Sub

' Draw this control's border if it has one.
Private Sub DrawBorder(ctl As Control, prn As Object, l As Single, t As Single, w As Single, h As Single)
Dim has_border As Boolean

    ' Draw a border for buttons.
    If (TypeOf ctl Is CommandButton) Or _
       (TypeOf ctl Is ComboBox) _
    Then
        has_border = True
    Else
        On Error Resume Next
        has_border = False
        has_border = (ctl.BorderStyle = vbFixedSingle)
        On Error GoTo 0
    End If

    ' Blank the control's area.
    prn.Line (l, t)-Step(w, h), vbWhite, BF

    ' Draw the border.
    If has_border Then prn.Line (l, t)-Step(w, h), , B
End Sub
' Draw a control.
Private Sub DrawControl(ctl As Control, prn As Object)
Dim l As Single
Dim t As Single
Dim w As Single
Dim h As Single

    ' See where the control must fit.
    ComputeOffsets ctl, prn, l, t, w, h

    ' See if the control is completely hidden.
    If w <= 0 Or h <= 0 Then Exit Sub

    Select Case TypeName(ctl)
        Case "CheckBox"
            DrawCheckBox ctl, prn, l, t, w, h
        Case "ComboBox"
            DrawComboBox ctl, prn, l, t, w, h
        Case "CommandButton"
            DrawCommandButton ctl, prn, l, t, w, h
        Case "Frame"
            DrawFrame ctl, prn, l, t, w, h
        Case "Image"
            DrawImage ctl, prn, l, t, w, h
        Case "Label"
            DrawLabel ctl, prn, l, t, w, h
        Case "Line"
            DrawLine ctl, prn, l, t, w, h
        Case "ListBox"
            DrawListBox ctl, prn, l, t, w, h
        Case "OptionButton"
            DrawOptionButton ctl, prn, l, t, w, h
        Case "PictureBox"
            DrawPictureBox ctl, prn, l, t, w, h
        Case "Shape"
            DrawShape ctl, prn, l, t, w, h
        Case "TextBox"
            DrawTextBox ctl, prn, l, t, w, h
        Case Else
            DrawUnknown ctl, prn, l, t, w, h
    End Select
End Sub
' Draw text properly aligned and wrapped.
Private Sub PrintText(txt As String, l As Single, _
    t As Single, w As Single, h As Single, _
    prn As Object, wrap As Boolean, align As Integer)
Dim CR As String
Dim LF As String
Dim gap As Single
Dim len_txt As Integer
Dim pos As Integer
Dim pos2 As Integer
Dim line_start As Integer
Dim word_start As Integer
Dim new_line As String
Dim len_line As Integer
Dim fits As String
Dim ch As String
Dim last_space As Integer

    CR = Chr$(13)
    LF = Chr$(10)

    ' Leave a little extra room.
    gap = prn.ScaleX(2, vbPoints, prn.ScaleMode)
    l = l + gap
    w = w - 2 * gap
    gap = prn.ScaleY(2, vbPoints, prn.ScaleMode)
    t = t + gap
    h = h - 2 * gap

    prn.CurrentY = t
    len_txt = Len(txt)
    line_start = 1
    ' Display lines.
    Do While line_start <= len_txt
        ' Get the next line.
        pos = InStr(line_start, txt, vbCrLf)
        If pos < 1 Then pos = len_txt + 1
        new_line = Mid$(txt, line_start, pos - line_start)
        len_line = Len(new_line)
        line_start = pos + 2 ' 2 = Len(vbCrLf)

        ' Display the line.
        word_start = 1
        Do While word_start <= len_line
            ' See how much text will fit.
            last_space = 0
            fits = ""
            For pos = word_start To len_line
                ch = Mid$(new_line, pos, 1)
                If ch = " " Then last_space = pos
                fits = fits & ch
                If prn.TextWidth(fits) > w Then
                    ' Make it fit.
                    If last_space > 0 Then
                        ' Remove the last word.
                        fits = Mid$(new_line, word_start, last_space - word_start)
                        word_start = last_space + 1
                    Else
                        ' Break the word here.
                        If Len(fits) > 1 Then _
                            fits = Left$(fits, Len(fits) - 1)
                        word_start = pos
                    End If
                    Exit For
                End If
            Next pos
            If pos > len_line Then word_start = len_line + 1

            ' Display this piece.
            Select Case align
                Case vbLeftJustify
                    prn.CurrentX = l
                Case vbRightJustify
                    prn.CurrentX = l + w - prn.TextWidth(fits)
                Case vbCenter
                    prn.CurrentX = l + (w - prn.TextWidth(fits)) / 2
            End Select
            prn.Print fits
            If Not wrap Then Exit Sub
            
            ' See if we've gone too far vertically.
            If prn.CurrentY + _
                prn.ScaleY(prn.Font.Size, _
                    vbPoints, prn.ScaleMode) > _
                    t + h Then Exit Sub
        Loop
    Loop
End Sub

' Draw a TextBox control.
Private Sub DrawTextBox(ctl As TextBox, prn As Object, l As Single, t As Single, w As Single, h As Single)
Dim wrap As Boolean

    DrawBorder ctl, prn, l, t, w, h
    SetFont ctl, prn

    ' If MultiLine and there's no horizontal scroll
    ' bar, wrap words.
    ' If Alignment is not vbLeftJustify, ignore
    ' the horizontal scroll bar.
    wrap = ctl.MultiLine And _
        (ctl.Alignment <> vbLeftJustify Or _
         ctl.ScrollBars = vbSBNone Or _
         ctl.ScrollBars = vbVertical)

    PrintText ctl.Text, l, t, w, h, _
        prn, wrap, ctl.Alignment
End Sub
' Draw a Label control.
Private Sub DrawLabel(ctl As Label, prn As Object, l As Single, t As Single, w As Single, h As Single)
    DrawBorder ctl, prn, l, t, w, h
    SetFont ctl, prn

    PrintText ctl.Caption, l, t, w, h, _
        prn, True, ctl.Alignment
End Sub
' Draw a PictureBox control. If AutoSize is False,
' this does not draw the control's  picture.
Private Sub DrawPictureBox(ctl As PictureBox, prn As Object, l As Single, t As Single, w As Single, h As Single)
    DrawBorder ctl, prn, l, t, w, h

    ' Be careful in case there's no Picture.
    On Error Resume Next
    prn.PaintPicture ctl.Picture, l, t, w, h
End Sub
' Draw a Frame control.
Private Sub DrawFrame(ctl As Frame, prn As Object, l As Single, t As Single, w As Single, h As Single)
Const lead_in_POINTS = 3
Dim lead_in As Single
Dim x1 As Single
Dim x2 As Single
Dim y1 As Single
Dim y2 As Single
Dim text_x1 As Single
Dim text_x2 As Single
Dim text_y1 As Single
Dim cap As String

    SetFont ctl, prn

    x1 = l
    x2 = x1 + w
    text_y1 = t
    y1 = text_y1 + prn.TextHeight(ctl.Caption) / 2
    y2 = text_y1 + h

    lead_in = prn.ScaleX(lead_in_POINTS, vbPoints, prn.ScaleMode)
    text_x1 = x1 + lead_in

    ' See if the text will fit.
    If x2 - text_x1 <= 0 Then
        cap = ""
    Else
        cap = " " & ctl.Caption & " "
        Do While prn.TextWidth(cap) > x2 - text_x1
            cap = Left$(cap, Len(cap) - 1)
        Loop
    End If
    text_x2 = text_x1 + prn.TextWidth(cap)

    prn.Line (text_x1, y1)-(x1, y1)
    prn.Line -(x1, y2)
    prn.Line -(x2, y2)
    prn.Line -(x2, y1)
    prn.Line -(text_x2, y1)

    prn.CurrentX = text_x1
    prn.CurrentY = text_y1
    prn.Print cap
End Sub
' Draw an Image control.
Private Sub DrawImage(ctl As Image, prn As Object, l As Single, t As Single, w As Single, h As Single)
    DrawBorder ctl, prn, l, t, w, h
    On Error Resume Next ' If there's no Picture.
    prn.PaintPicture ctl.Picture, l, t, w, h
End Sub
' Draw a ListBox control.
Private Sub DrawListBox(ctl As ListBox, prn As Object, l As Single, t As Single, w As Single, h As Single)
Dim i As Integer
Dim gap As Single
Dim fits As String
Dim choice As String
Dim pos As Integer
Dim line_hgt As Single
Dim x As Single
Dim y As Single

    DrawBorder ctl, prn, l, t, w, h
    SetFont ctl, prn

    line_hgt = prn.ScaleY(prn.Font.Size, _
        vbPoints, prn.ScaleMode)

    ' Leave a little extra room.
    gap = prn.ScaleX(2, vbPoints, prn.ScaleMode)
    l = l + gap
    w = w - 2 * gap
    gap = prn.ScaleY(2, vbPoints, prn.ScaleMode)
    t = t + gap
    h = h - 2 * gap

    ' Display the items starting with the first
    ' visible item.
    prn.CurrentY = t
    For i = ctl.TopIndex To ctl.ListCount
        ' See how much text will fit.
        choice = ctl.List(i)
        For pos = Len(choice) To 1 Step -1
            fits = Left$(choice, pos)
            If prn.TextWidth(fits) <= w Then Exit For
        Next pos

        ' Display this piece.
        prn.CurrentX = l
        If ctl.Selected(i) Then
            x = prn.CurrentX
            y = prn.CurrentY
            prn.Line (l, prn.CurrentY)- _
                Step(w, line_hgt + gap), , B
            prn.CurrentX = x
            prn.CurrentY = y
        End If
        prn.Print fits

        ' See if we've gone too far vertically.
        If prn.CurrentY + line_hgt > t + h _
            Then Exit For
    Next i
End Sub
' Draw a ComboBox control.
Private Sub DrawComboBox(ctl As ComboBox, prn As Object, l As Single, t As Single, w As Single, h As Single)
    DrawBorder ctl, prn, l, t, w, h
    SetFont ctl, prn

    PrintText ctl.Text, l, t, w, h, _
        prn, True, vbLeftJustify
End Sub

' Draw a CommandButton control.
Private Sub DrawCommandButton(ctl As CommandButton, prn As Object, l As Single, t As Single, w As Single, h As Single)
    DrawBorder ctl, prn, l, t, w, h
    SetFont ctl, prn

    PrintText ctl.Caption, l, t, w, h, _
        prn, True, vbCenter
End Sub

' Draw a box for an unknown control.
Private Sub DrawUnknown(ctl As Control, prn As Object, l As Single, t As Single, w As Single, h As Single)
Const UNK_TEXT = "???"

    DrawBorder ctl, prn, l, t, w, h

    On Error Resume Next
    SetFont ctl, prn
    On Error GoTo 0
    
    prn.CurrentX = l + (w - prn.TextWidth(UNK_TEXT)) / 2
    prn.CurrentY = t + (h - prn.TextHeight(UNK_TEXT)) / 2
    prn.Print UNK_TEXT
End Sub
' Draw a Shape control.
Private Sub DrawShape(ctl As Shape, prn As Object, l As Single, t As Single, w As Single, h As Single)
Const PI = 3.14159265
Dim x1 As Single
Dim y1 As Single
Dim x2 As Single
Dim y2 As Single
Dim pixx As Single
Dim pixy As Single
Dim cx As Single
Dim cy As Single
Dim radius As Single

    SaveStyles ctl, prn

    cx = l + w / 2
    cy = t + h / 2
    If ctl.Shape = vbShapeRoundedRectangle Or _
       ctl.Shape = vbShapeRoundedSquare _
    Then
        If w > h Then
            radius = h / 10
        Else
            radius = w / 10
        End If
    End If
    If ctl.Shape = vbShapeSquare Or _
       ctl.Shape = vbShapeRoundedSquare Or _
       ctl.Shape = vbShapeCircle _
    Then
        If w > h Then
            w = h
        Else
            h = w
        End If
    End If
    x1 = cx - w / 2
    y1 = cy - h / 2
    x2 = cx + w / 2
    y2 = cy + h / 2

    Select Case ctl.Shape
        Case vbShapeRectangle, vbShapeSquare
            prn.Line (x1, y1)-(x2, y2), , B
        Case vbShapeOval, vbShapeCircle
            If w > h Then
                prn.Circle (cx, cy), w / 2, , , , h / w
            Else
                prn.Circle (cx, cy), h / 2, , , , h / w
            End If
        Case vbShapeRoundedRectangle, vbShapeRoundedSquare
            pixx = prn.ScaleX(1, vbPixels, prn.ScaleMode)
            pixy = prn.ScaleY(1, vbPixels, prn.ScaleMode)
            prn.Line (x1 + radius, y1)-Step(w - 2 * radius + pixx, 0)
            prn.Line (x1 + radius, y1 + h)-Step(w - 2 * radius + pixx, 0)
            prn.Line (x1, y1 + radius)-Step(0, h - 2 * radius + pixy)
            prn.Line (x2, y1 + radius)-Step(0, h - 2 * radius + pixy)
            prn.Circle (x1 + radius, y1 + radius), radius, , PI / 2, PI
            prn.Circle (x1 + radius, y2 - radius), radius, , PI, 1.5 * PI
            prn.Circle (x2 - radius, y1 + radius), radius, , 0, PI / 2
            prn.Circle (x2 - radius, y2 - radius), radius, , 1.5 * PI, 2 * PI
    End Select

    RestoreStyles prn
End Sub
' Draw a CheckBox control.
Private Sub DrawCheckBox(ctl As CheckBox, prn As Object, l As Single, t As Single, w As Single, h As Single)
Const BOX_WID_POINTS = 8
Dim box_wid As Single
Dim box_hgt As Single
Dim box_x1 As Single
Dim box_y1 As Single
Dim box_x2 As Single
Dim box_y2 As Single
Dim text_x1 As Single
Dim text_y1 As Single
Dim text_x2 As Single
Dim text_y2 As Single

    SetFont ctl, prn
    
    box_wid = prn.ScaleX(BOX_WID_POINTS, vbPoints, prn.ScaleMode)
    box_hgt = prn.ScaleY(BOX_WID_POINTS, vbPoints, prn.ScaleMode)
    box_y1 = t + (h - box_hgt) / 2
    box_y2 = t + (h + box_hgt) / 2
    text_y1 = t
    text_y2 = t + h
    If ctl.Alignment = vbLeftJustify Then
        box_x1 = l
        box_x2 = box_x1 + box_wid
        text_x1 = box_x2 + box_wid / 2
        text_x2 = l + w
    Else
        box_x2 = l + w
        box_x1 = box_x2 - box_wid
        text_x1 = l
        text_x2 = box_x1 - box_wid / 2
    End If

    ' Draw the text.
    PrintText ctl.Caption, text_x1, text_y1, _
        text_x2 - text_x1, text_y2 - text_y1, prn, _
        True, vbLeftJustify

    ' Draw the box.
    prn.Line (box_x1, box_y1)-(box_x2, box_y2), , B
    If ctl.Value = vbChecked Then
        prn.Line (box_x1, box_y1)-(box_x2, box_y2)
        prn.Line (box_x2, box_y1)-(box_x1, box_y2)
    End If
End Sub
' Draw a OptionButton control.
Private Sub DrawOptionButton(ctl As OptionButton, prn As Object, l As Single, t As Single, w As Single, h As Single)
Const RADIUS_POINTS = 3.75
Dim radius As Single
Dim cx As Single
Dim cy As Single
Dim text_x1 As Single
Dim text_y1 As Single
Dim text_x2 As Single
Dim text_y2 As Single

    SetFont ctl, prn
    
    radius = prn.ScaleX(RADIUS_POINTS, vbPoints, prn.ScaleMode)
    cy = t + h / 2
    text_y1 = t
    text_y2 = t + h
    If ctl.Alignment = vbLeftJustify Then
        cx = l + radius
        text_x1 = cx + 2 * radius / 2
        text_x2 = l + w
    Else
        cx = l + w - radius
        text_x1 = l
        text_x2 = cx - 2 * radius
    End If

    ' Draw the text.
    PrintText ctl.Caption, text_x1, text_y1, _
        text_x2 - text_x1, text_y2 - text_y1, prn, _
        True, vbLeftJustify

    ' Draw the circle.
    prn.Circle (cx, cy), radius
    If ctl.Value Then
        prn.FillStyle = vbSolid
        prn.Circle (cx, cy), radius * 0.4
        prn.FillStyle = vbFSTransparent
    End If
End Sub
' Draw a Line control.
Private Sub DrawLine(ctl As Line, prn As Object, l As Single, t As Single, w As Single, h As Single)
    SaveStyles ctl, prn
    prn.Line (l + ctl.x1, t + ctl.y1)-(l + ctl.x2, t + ctl.y2)
    RestoreStyles prn
End Sub
' Restore styles saved by SaveStyles.
Private Sub RestoreStyles(prn As Object)
    prn.DrawStyle = orig_draw_style
    prn.DrawWidth = orig_draw_width
    prn.FillStyle = orig_fill_style
    prn.ForeColor = orig_fore_color
End Sub

' Save styles modified by Shape and Line controls.
Private Sub SaveStyles(ctl As Control, prn As Object)
    orig_draw_style = prn.DrawStyle
    orig_draw_width = prn.DrawWidth
    orig_fill_style = prn.FillStyle
    orig_fore_color = prn.ForeColor

    Select Case ctl.BorderStyle
        Case vbTransparent
            prn.DrawStyle = vbInvisible
        Case vbBSSolid
            prn.DrawStyle = vbSolid
        Case vbBSDash
            prn.DrawStyle = vbDash
        Case vbBSDot
            prn.DrawStyle = vbDot
        Case vbBSDashDot
            prn.DrawStyle = vbDashDot
        Case vbBSDashDotDot
            prn.DrawStyle = vbDashDotDot
        Case vbBSInsideSolid
            prn.DrawStyle = vbInsideSolid
    End Select

    prn.ForeColor = ctl.BorderColor
    prn.DrawWidth = ctl.BorderWidth
    If Not (TypeOf ctl Is Line) Then _
        prn.FillStyle = ctl.FillStyle
End Sub
' Produce a high resolution printout of the form.
Public Sub HiResPrint(frm As Form, prn As Object, size_mode As Integer)
Dim i As Integer
Dim ctl As Control
Dim is_visible As Boolean

    ' Set Printer scale properties.
    Select Case size_mode
        Case hires_Normal
            ScaleNormal frm, prn
        Case hires_StretchToFit
            ScaleToFit frm, prn
        Case hires_ResizePrinter
            ' Make the "Printer" the right size.
            ' This is used for print preview.
            prn.Width = frm.ScaleWidth + prn.Width - prn.ScaleWidth
            prn.Height = frm.ScaleHeight + prn.Height - prn.ScaleHeight
            ScaleNormal frm, prn
    End Select

    ' We have not tried any fonts yet.
    Set LastFont = Nothing

    ' Draw non-menu controls with Visible True
    ' in back to front order.
    For i = frm.Controls.Count - 1 To 0 Step -1
        Set ctl = frm.Controls(i)

        ' Watch for controls w/o Visible property.
        is_visible = False
        On Error Resume Next
        If Not (TypeOf ctl Is Menu) Then _
            is_visible = ctl.Visible
        On Error GoTo 0

        ' Draw if a visible control.
        If is_visible Then DrawControl ctl, prn
    Next i

    ' Draw a box around the form.
    prn.Line (0, 0)-Step(frm.ScaleWidth, frm.ScaleHeight), , B
End Sub
' Set the printer's coordinates for full scale
' printing.
Private Sub ScaleNormal(frm As Form, prn As Object)
Dim pwid As Single
Dim phgt As Single
Dim xmid As Single
Dim ymid As Single

    ' Get printer dimensions.
    pwid = prn.ScaleX(prn.ScaleWidth, prn.ScaleMode, vbTwips)
    phgt = prn.ScaleY(prn.ScaleHeight, prn.ScaleMode, vbTwips)

    ' Convert into form dimensions.
    pwid = frm.ScaleX(pwid, vbTwips, frm.ScaleMode)
    phgt = frm.ScaleY(phgt, vbTwips, frm.ScaleMode)

    ' Compute the form's center.
    xmid = frm.ScaleLeft + frm.ScaleWidth / 2
    ymid = frm.ScaleTop + frm.ScaleHeight / 2

    ' Set the printer's scale properties.
    prn.Scale _
        (xmid - pwid / 2, ymid - phgt / 2)- _
        (xmid + pwid / 2, ymid + phgt / 2)
End Sub
' Set the printer's coordinates to print the form
' as large as possible.
Private Sub ScaleToFit(frm As Form, prn As Object)
Dim fwid As Single
Dim fhgt As Single
Dim pwid As Single
Dim phgt As Single
Dim xmid As Single
Dim ymid As Single
Dim s As Single

    ' Get printer dimensions.
    pwid = prn.ScaleX(prn.ScaleWidth, prn.ScaleMode, vbTwips)
    phgt = prn.ScaleY(prn.ScaleHeight, prn.ScaleMode, vbTwips)

    ' Get the form's dimensions.
    fwid = frm.ScaleX(frm.ScaleWidth, frm.ScaleMode, vbTwips)
    fhgt = frm.ScaleY(frm.ScaleHeight, frm.ScaleMode, vbTwips)

    ' Compare the aspect ratios.
    If (fhgt / fwid) > (phgt / pwid) Then
        ' Form is tall. Use the printer's height.
        s = fhgt / phgt
    Else
        ' Form is wide. Use the printer's width.
        s = fwid / pwid
    End If

    ' Convert printer dimensions.
    pwid = frm.ScaleX(pwid, vbTwips, frm.ScaleMode) * s
    phgt = frm.ScaleY(phgt, vbTwips, frm.ScaleMode) * s

    ' See where the center of the form should be.
    xmid = frm.ScaleLeft + frm.ScaleWidth / 2
    ymid = frm.ScaleTop + frm.ScaleHeight / 2

    ' Set the printer's scale properties.
    prn.Scale _
        (xmid - pwid / 2, ymid - phgt / 2)- _
        (xmid + pwid / 2, ymid + phgt / 2)
End Sub
' Make the printer use a Printer object compatible
' font similar to the one used by this control.
Private Sub SetFont(ctl As Control, prn As Object)
    With ctl.Font
        ' If this is the same font we tried last,
        ' we've already selected the proper font.
        If LastFont Is Nothing Then
            Set LastFont = New StdFont
        ElseIf .Bold = LastFont.Bold And _
           .Italic = LastFont.Italic And _
           .Name = LastFont.Name And _
           .Size = LastFont.Size And _
           .Strikethrough = LastFont.Strikethrough And _
           .Underline = LastFont.Underline And _
           .Weight = LastFont.Weight _
        Then
            Exit Sub
        End If

        ' Save the new font properties.
        LastFont.Bold = .Bold
        LastFont.Italic = .Italic
        LastFont.Name = .Name
        LastFont.Size = .Size
        LastFont.Strikethrough = .Strikethrough
        LastFont.Underline = .Underline
        LastFont.Weight = .Weight

        ' Select the font into the Printer object.
        Printer.Font.Name = .Name
        Printer.Font.Size = .Size
        Printer.Font.Name = .Name
        Printer.Font.Size = .Size
        Printer.Font.Bold = .Bold
        Printer.Font.Italic = .Italic
        Printer.Font.Strikethrough = .Strikethrough
        Printer.Font.Underline = .Underline
        Printer.Font.Weight = .Weight
    End With

    ' Copy the font we got into prn.
    prn.Font.Name = Printer.Font.Name
    prn.Font.Size = Printer.Font.Size
    prn.Font.Name = Printer.Font.Name
    prn.Font.Size = Printer.Font.Size
    prn.Font.Bold = Printer.Font.Bold
    prn.Font.Italic = Printer.Font.Italic
    prn.Font.Strikethrough = Printer.Font.Strikethrough
    prn.Font.Underline = Printer.Font.Underline
    prn.Font.Weight = Printer.Font.Weight
End Sub

