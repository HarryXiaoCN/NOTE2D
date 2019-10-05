Attribute VB_Name = "Node_ContentFunction"
Public Function NodePixel()
    Dim tempX As Long, tempY As Long, tempC As Long, i As Long, firstLock As Boolean, secondLock As Boolean, dX As Single, dY As Single
    Dim firstX As Long, firstY As Long, firstNid As Long
    For i = 0 To nSum
        With Node(i)
            If .b Then
                If .select Then
                    If 节点内容图片解析(富文本转义(.content), tempX, tempY, tempC) Then
                        If firstLock Then
                            .x = imageToNtx_StepX * (tempX - firstX) + Node(firstNid).x
                            .y = imageToNtx_StepY * (tempY - firstY) + Node(firstNid).y
                            .setColor = tempC
                        Else
                            firstX = tempX
                            firstY = tempY
                            firstNid = i
                            .setColor = tempC
                            firstLock = True
                        End If
                    End If
                End If
            End If
        End With
    Next
End Function

Public Function NodeWave(rootNid As Long, rootValue As Single, firstLock As Boolean, dX As Single, dY As Single)
    Dim tempValue As Single, i As Long
    Node(rootNid).toTreeTxtLock = True
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                If .source = rootNid And Node(.target).toTreeTxtLock = False Then
                    If NCF_NodeValueControl(富文本转义(Node(.target).content), tempValue) Then
                        If firstLock Then
                            Node(.target).x = dX + Node(rootNid).x
                            Node(.target).y = dY * (tempValue - rootValue) + Node(rootNid).y
                            NodeWave .target, tempValue, firstLock, dX, dY
                        Else
                            dX = Node(.target).x - Node(rootNid).x
                            dY = (Node(.target).y - Node(rootNid).y) / (tempValue - rootValue)
                            firstLock = True
                            NodeWave .target, tempValue, firstLock, dX, dY
                        End If
                    End If
                End If
            End If
        End With
    Next
End Function

Public Function NCF_NodeValueControl(content As String, getValue As Single) As Boolean
    Dim sT As String, tT As Long, vT() As String, r As Long, g As Long, b As Long
    On Error GoTo Er
    tT = InStr(1, content, "波值[")
    If tT > 0 Then
        sT = Mid(content, tT + 3, InStr(1, content, "]") - tT - 3)
        getValue = Val(sT)
        NCF_NodeValueControl = True
    End If
    Exit Function
Er:
End Function

Public Function NCF_NodeColorControl(content As String, color As Long) As Long
    Dim sT As String, tT As Long, vT() As String, r As Long, g As Long, b As Long
    On Error GoTo Er
    tT = InStr(1, content, "颜色[")
    If tT > 0 Then
        sT = Mid(content, tT + 3, InStr(1, content, "]") - tT - 3)
        If InStr(1, sT, ",") > 0 Then
            vT = Split(sT, ",")
            If UBound(vT) >= 2 Then
                NCF_NodeColorControl = RGB(Val(vT(0)), Val(vT(1)), Val(vT(2)))
            End If
        Else
            NCF_NodeColorControl = Val(sT)
        End If
    Else
        NCF_NodeColorControl = color
    End If
    Exit Function
Er:
    NCF_NodeColorControl = color
End Function
