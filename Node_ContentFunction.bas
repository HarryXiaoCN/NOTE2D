Attribute VB_Name = "Node_ContentFunction"
Public Function NodePixel()
    Dim tempX As Long, tempY As Long, tempC As Long, i As Long, firstLock As Boolean, secondLock As Boolean, dX As Single, dY As Single
    Dim firstX As Long, firstY As Long, firstNid As Long
    For i = 0 To nSum
        With node(i)
            If .b Then
                If .select Then
                    If 节点内容图片解析(.text, tempX, tempY, tempC) Then
                        If firstLock Then
                            .X = imageToNtx_StepX * (tempX - firstX) + node(firstNid).X
                            .Y = imageToNtx_StepY * (tempY - firstY) + node(firstNid).Y
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
    node(rootNid).toTreeTxtLock = True
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                If .Source = rootNid And node(.target).toTreeTxtLock = False Then
                    If NCF_NodeValueControl(富文本转义(node(.target).content), tempValue) Then
                        If firstLock Then
                            node(.target).X = dX + node(rootNid).X
                            node(.target).Y = dY * (tempValue - rootValue) + node(rootNid).Y
                            NodeWave .target, tempValue, firstLock, dX, dY
                        Else
                            dX = node(.target).X - node(rootNid).X
                            dY = (node(.target).Y - node(rootNid).Y) / (tempValue - rootValue)
                            firstLock = True
                            NodeWave .target, tempValue, firstLock, dX, dY
                        End If
                    End If
                End If
            End If
        End With
    Next
End Function

Public Function NCF_NodeGravitationalControl(content As String, getValue As String, feature As String) As Boolean
    Dim sT As String, tT As Long, vT() As String, r As Long, g As Long, b As Long
    On Error GoTo Er
        tT = InStr(1, content, feature)
        If tT > 0 Then
            sT = Mid(content, tT + Len(feature), InStr(1, content, ")") - tT - Len(feature))
            getValue = sT
            NCF_NodeGravitationalSourceControl = True
        End If
    Exit Function
Er:
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

Public Function NCF_NodeLikeControl(content As String, getValue As String, Source As Boolean) As Boolean
    Dim sT As String, tT As Long, vT() As String, r As Long, g As Long, b As Long
    On Error GoTo Er
        tT = InStr(1, content, "去(")
        If tT > 0 Then
            sT = Mid(content, tT + 2, InStr(1, content, ")") - tT - 2)
            getValue = sT
            NCF_NodeLikeControl = True
        Else
            tT = InStr(1, content, "源(")
            If tT > 0 Then
                sT = Mid(content, tT + 2, InStr(1, content, ")") - tT - 2)
                getValue = sT
                Source = True
                NCF_NodeLikeControl = True
            End If
        End If
    Exit Function
Er:
End Function
