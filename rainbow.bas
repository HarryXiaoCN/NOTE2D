Attribute VB_Name = "rainbow"
Public dAngle As Single, dColor As Single
Public Function Rainbow_Line(ByRef formObj As Form, ByRef x1 As Single, ByRef y1 As Single, ByRef x2 As Single, ByRef y2 As Single)
Dim i As Single: Dim dx As Single, dy As Single
Dim color As Single
dx = (x2 - x1) / 70
dy = (y2 - y1) / 70
For i = 0 To 70 Step 0.1
    color = Rainbow_BackEnd(i)
    formObj.Line (x1 + dx * i, y1 + dy * i)-(x1 + dx * (i + 1), y1 + dy * (i + 1)), color
Next
End Function
Public Function Rainbow_Crcle(ByRef formObj As Form, ByRef size As Single, ByRef mX As Single, ByRef mY As Single)
Dim i As Single: Dim x As Single, y As Single
Dim x2 As Single, y2 As Single
Dim color As Single
Dim angle As Single
'formObj.Cls
dAngle = dAngle - 0.1
If dAngle > PI * 2 Then dAngle = 0
For i = 0 To 70
    formObj.DrawWidth = (3 - i / 35)
    angle = PI * 2 / 75 * i + dAngle
    color = Rainbow_BackEnd(i)
    x = Sin(angle) * size + mX
    y = Cos(angle) * size + mY
    angle = PI * 2 / 75 * (i + 1) + dAngle
    x2 = Sin(angle) * size + mX
    y2 = Cos(angle) * size + mY
'    formObj.PSet (, Cos(angle) * size + mY), color
    formObj.Line (x, y)-(x2, y2), color
Next
End Function
Public Function Rainbow_BackEnd(ByRef i As Single) As Single
Rainbow_BackEnd = RGB(255, 0, 0)
If i > 0 And i <= 10 Then
    Rainbow_BackEnd = RGB(255, 165 / 10 * i, 0)
ElseIf i > 10 And i <= 20 Then
    Rainbow_BackEnd = RGB(255, 165 + 90 / 10 * (i - 10), 0)
ElseIf i > 20 And i <= 30 Then
    Rainbow_BackEnd = RGB(255 - 255 / 10 * (i - 20), 255, 0)
ElseIf i > 30 And i <= 40 Then
    Rainbow_BackEnd = RGB(0, 255 - 128 / 10 * (i - 30), 255 / 10 * (i - 30))
ElseIf i > 40 And i <= 50 Then
    Rainbow_BackEnd = RGB(0, 127 - 127 / 10 * (i - 40), 255)
ElseIf i > 50 And i <= 60 Then
    Rainbow_BackEnd = RGB(139 / 10 * (i - 50), 0, 255)
ElseIf i > 60 And i <= 70 Then
'    Rainbow_BackEnd = RGB(139 + 116 / 10 * (i - 60), 0, 255 - 255 / 10 * (i - 60))
    Rainbow_BackEnd = RGB(139 - 139 / 10 * (i - 60), 0, 255 - 255 / 10 * (i - 60))
End If
End Function
Public Function Rainbow_RedEnd(ByRef i As Single) As Single
Rainbow_RedEnd = RGB(255, 0, 0)
If i > 0 And i <= 10 Then
    Rainbow_RedEnd = RGB(255, 165 / 10 * i, 0)
ElseIf i > 10 And i <= 20 Then
    Rainbow_RedEnd = RGB(255, 165 + 90 / 10 * (i - 10), 0)
ElseIf i > 20 And i <= 30 Then
    Rainbow_RedEnd = RGB(255 - 255 / 10 * (i - 20), 255, 0)
ElseIf i > 30 And i <= 40 Then
    Rainbow_RedEnd = RGB(0, 255 - 128 / 10 * (i - 30), 255 / 10 * (i - 30))
ElseIf i > 40 And i <= 50 Then
    Rainbow_RedEnd = RGB(0, 127 - 127 / 10 * (i - 40), 255)
ElseIf i > 50 And i <= 60 Then
    Rainbow_RedEnd = RGB(139 / 10 * (i - 50), 0, 255)
ElseIf i > 60 And i <= 70 Then
    Rainbow_RedEnd = RGB(139 + 116 / 10 * (i - 60), 0, 255 - 255 / 10 * (i - 60))
'    Rainbow_RedEnd = RGB(139 - 139 / 10 * (i - 60), 0, 255 - 255 / 10 * (i - 60))
End If
End Function
Public Function DynamicRainbow_Line(ByRef formObj As Form, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
Dim i, c As Single: Dim dx As Single, dy As Single
Dim color As Single
dx = (x2 - x1) / 70
dy = (y2 - y1) / 70
If dColor > 70 Then
    dColor = 0
End If
For i = 0 To 70 Step 0.1
    c = i - dColor
    If c < 0 Then
        c = 70 + c
    End If
    color = Rainbow_RedEnd(c)
    formObj.Line (x1 + dx * i, y1 + dy * i)-(x1 + dx * (i + 1), y1 + dy * (i + 1)), color
Next
dColor = dColor + 1
End Function
