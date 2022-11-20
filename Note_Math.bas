Attribute VB_Name = "Note_Math"
Public Sub NodeArray(nid As Long, r As Single)
    Dim i As Long, aN As Long, ave As Single
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                If .source = nid Then
                    aN = aN + 1
                End If
            End If
        End With
    Next
    If aN > 0 Then
        ave = PI * 2 / aN
        aN = 0
        For i = 0 To lSum
            With nodeLine(i)
                If .b Then
                    If .source = nid Then
                        aN = aN + 1
                        node(.target).X = Cos(ave * aN) * r + node(nid).X
                        node(.target).Y = Sin(ave * aN) * r + node(nid).Y
                    End If
                End If
            End With
        Next
    End If
End Sub
Public Function MToZF(ByRef mag As Single) As Single
    MToZF = 2 ^ mag
End Function
Public Function StrToBool(ByVal str As String) As Boolean
    If InStr(1, UCase(str), "TRUE") Then StrToBool = True
End Function
Public Function GetLineEquation(ByRef x1 As Single, ByRef y1 As Single, ByRef x2 As Single, ByRef y2 As Single) As 二维坐标
Dim lineRoot As 二维坐标
    lineRoot.X = (y1 - y2) / (x1 - x2)
    lineRoot.Y = (y1 * x2 - y2 * x1) / (x2 - x1)
    GetLineEquation = lineRoot
End Function
Public Function LineIntersectionArea(ByRef x1 As Single, ByRef y1 As Single, ByRef x2 As Single, ByRef y2 As Single, _
ByRef startPos As 二维坐标, ByRef endPos As 二维坐标) As Boolean
    Dim lineRoot As 二维坐标
    Dim checkPos As 二维坐标
    Dim mX, mi As Long
    If x1 <> x2 Then
        lineRoot = GetLineEquation(x1, y1, x2, y2)
        If x1 > x2 Then mX = x1: mi = x2 Else mX = x2: mi = x1
        For checkPos.X = mi To mX Step 10
            checkPos.Y = checkPos.X * lineRoot.X + lineRoot.Y
            If RectangleOverlapCheck(startPos, endPos, checkPos) = True Then
                LineIntersectionArea = True: Exit Function
            End If
        Next
    Else
        If y1 > y2 Then mX = y1: mi = y2 Else mX = y2: mi = y1
        For checkPos.Y = mi To mX Step 10
            checkPos.X = x1
            If RectangleOverlapCheck(startPos, endPos, checkPos) = True Then
                LineIntersectionArea = True: Exit Function
            End If
        Next
    End If
End Function
Public Function TwoPointDistance(ByRef x1 As Single, ByRef y1 As Single, ByRef x2 As Single, ByRef y2 As Single) As Single
    TwoPointDistance = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
End Function
Public Function OverlappingJudgment(ByRef distance As Single, ByRef x1 As Single, ByRef y1 As Single, ByRef x2 As Single, ByRef y2 As Single) As Boolean
    If TwoPointDistance(x1, y1, x2, y2) <= distance Then OverlappingJudgment = True
End Function
Public Function DistanceBetweenLinePoint(lxS As Single, lyS As Single, lxE As Single, lyE As Single, pX As Single, pY As Single, d As Single) As Boolean
    If Sqr((lxS - lxE) ^ 2 + (lyS - lyE) ^ 2) + d >= Sqr((lxS - pX) ^ 2 + (lyS - pY) ^ 2) + Sqr((lxE - pX) ^ 2 + (lyE - pY) ^ 2) Then
        DistanceBetweenLinePoint = True
    End If
End Function
'Public Function 点在线上(lxS As Single, lyS As Single, lxE As Single, lyE As Single, pX As Single, pY As Single, d As Single) As Boolean
'    If (lxE - lxS) * (pY - lyS) = (lyE - lyS) * (pX - lxS) Then
'        点在线上 = True
'    End If
'End Function
Public Function V3ToV2Pos(ByRef x3 As Single, ByRef y3 As Single, ByRef z3 As Single) As 二维坐标
    Dim output As 二维坐标
    output.X = x3 * 500 + Note.width / 2
    output.Y = y3 * 500 + Note.height / 2
    V3ToV2Pos = output
End Function
Public Function StrToV3(ByVal v3str As String) As 三维坐标
    Dim output As 三维坐标: Dim strTmp
    v3str = Replace(v3str, "(", "")
    v3str = Replace(v3str, ")", "")
    strTmp = Split(v3str, ",")
    output.X = Val(strTmp(0))
    output.Y = Val(strTmp(1))
    output.z = Val(strTmp(2))
    StrToV3 = output
End Function
Public Function RectangleOverlapCheck(ByRef startPos As 二维坐标, ByRef endPos As 二维坐标, ByRef checkPos As 二维坐标) As Boolean
    If (checkPos.X > startPos.X And checkPos.X < endPos.X And checkPos.Y > startPos.Y And checkPos.Y < endPos.Y) _
    Or (checkPos.X > endPos.X And checkPos.X < startPos.X And checkPos.Y > endPos.Y And checkPos.Y < startPos.Y) _
    Or (checkPos.X > startPos.X And checkPos.X < endPos.X And checkPos.Y < startPos.Y And checkPos.Y > endPos.Y) _
    Or (checkPos.X > endPos.X And checkPos.X < startPos.X And checkPos.Y < endPos.Y And checkPos.Y > startPos.Y) Then
        RectangleOverlapCheck = True
    End If
End Function
Public Function RectangleRightStartCheck(ByRef startPos As 二维坐标, ByRef endPos As 二维坐标) As Boolean
    If startPos.X > endPos.X Then RectangleRightStartCheck = True
End Function
