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
                        node(.target).x = Cos(ave * aN) * r + node(nid).x
                        node(.target).y = Sin(ave * aN) * r + node(nid).y
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
lineRoot.x = (y1 - y2) / (x1 - x2)
lineRoot.y = (y1 * x2 - y2 * x1) / (x2 - x1)
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
    For checkPos.x = mi To mX Step 10
        checkPos.y = checkPos.x * lineRoot.x + lineRoot.y
        If RectangleOverlapCheck(startPos, endPos, checkPos) = True Then
            LineIntersectionArea = True: Exit Function
        End If
    Next
Else
    If y1 > y2 Then mX = y1: mi = y2 Else mX = y2: mi = y1
    For checkPos.y = mi To mX Step 10
        checkPos.x = x1
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
Public Function V3ToV2Pos(ByRef x3 As Single, ByRef y3 As Single, ByRef z3 As Single) As 二维坐标
Dim output As 二维坐标
output.x = x3 * 500 + Note.Width / 2
output.y = y3 * 500 + Note.Height / 2
V3ToV2Pos = output
End Function
Public Function StrToV3(ByVal v3str As String) As 三维坐标
Dim output As 三维坐标: Dim strTmp
v3str = Replace(v3str, "(", "")
v3str = Replace(v3str, ")", "")
strTmp = Split(v3str, ",")
output.x = Val(strTmp(0))
output.y = Val(strTmp(1))
output.z = Val(strTmp(2))
StrToV3 = output
End Function
Public Function RectangleOverlapCheck(ByRef startPos As 二维坐标, ByRef endPos As 二维坐标, ByRef checkPos As 二维坐标) As Boolean
If (checkPos.x > startPos.x And checkPos.x < endPos.x And checkPos.y > startPos.y And checkPos.y < endPos.y) _
Or (checkPos.x > endPos.x And checkPos.x < startPos.x And checkPos.y > endPos.y And checkPos.y < startPos.y) _
Or (checkPos.x > startPos.x And checkPos.x < endPos.x And checkPos.y < startPos.y And checkPos.y > endPos.y) _
Or (checkPos.x > endPos.x And checkPos.x < startPos.x And checkPos.y < endPos.y And checkPos.y > startPos.y) Then
    RectangleOverlapCheck = True
End If
End Function
Public Function RectangleRightStartCheck(ByRef startPos As 二维坐标, ByRef endPos As 二维坐标) As Boolean
If startPos.x > endPos.x Then RectangleRightStartCheck = True
End Function
