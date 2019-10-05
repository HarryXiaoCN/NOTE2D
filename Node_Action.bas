Attribute VB_Name = "Node_Action"
Public Type ACTION_LIST
    be As Boolean
    name As String '动作名字
    nodeID() As Long '作用的节点ID
    interval As Long '多少时间周期执行一次动作
    intervalTemp As Long '当前计数的时间周期
    process() As 二维坐标 '记忆节点在动作中执行的中间总向量，方便回到初始位置
    route As String '直线还是圆周
    vector As 二维坐标 '直线移动的向量
    endTime As Long '直线/圆周移动的终止次数
    endTimeTemp As Long '终止次数计数
    angle As Single '圆周运动的角度
    relativePosition() As 二维坐标 '与目标节点的相对位置
    radius() As Single '与目标节点的半径
    initialAngle() As Single '初始夹角
    aimNode As Long '圆周目标运动的节点
    repeat As Boolean '到终止位置后节点是否回到初始位置重新执行动作
End Type
Public actionList() As ACTION_LIST

Public Sub ObjectAction()
    Dim i As Long, j As Long
    For i = 1 To UBound(actionList)
        If actionList(i).be Then
            If actionList(i).interval <= actionList(i).intervalTemp Then
                With actionList(i)
                    .intervalTemp = 0
                    Select Case .route
                        Case "直线"
                            For j = 0 To UBound(.nodeID)
                                node(.nodeID(j)).X = node(.nodeID(j)).X + .vector.X
                                node(.nodeID(j)).Y = node(.nodeID(j)).Y + .vector.Y
                                .process(j).X = .process(j).X - .vector.X
                                .process(j).Y = .process(j).Y - .vector.Y
                            Next
                        Case "圆周"
                            For j = 0 To UBound(.nodeID)
                                node(.nodeID(j)).X = node(.aimNode).X + .relativePosition(j).X
                                node(.nodeID(j)).Y = node(.aimNode).Y + .relativePosition(j).Y
                                获得角度后坐标 node(.nodeID(j)), node(.aimNode), .angle, .endTimeTemp, .radius(j), .relativePosition(j), .process(j), .initialAngle(j)
                            Next
                    End Select
                    .endTimeTemp = .endTimeTemp + 1
                    If .endTimeTemp = .endTime Then
                        If .repeat Then
                            For j = 0 To UBound(.nodeID)
                                node(.nodeID(j)).X = node(.nodeID(j)).X + .process(j).X
                                node(.nodeID(j)).Y = node(.nodeID(j)).Y + .process(j).Y
                                .process(j).X = 0
                                .process(j).Y = 0
                            Next
                        Else
                            .be = False
                        End If
                        .endTimeTemp = 0
                    End If
                End With
            Else
                actionList(i).intervalTemp = actionList(i).intervalTemp + 1
            End If
        End If
    Next
End Sub
Public Function 获得角度后坐标(arrayObj As 节点, aimNode As 节点, angle As Single, endTimeTemp As Long, radius As Single, relativePosition As 二维坐标, process As 二维坐标, initialAngle As Single)
    Dim arrAngle As Single, pT As 二维坐标
    arrAngle = 角度转弧度(angle) * endTimeTemp + initialAngle
    pT.X = arrayObj.X
    pT.Y = arrayObj.Y
    If arrAngle > PI / 2 Then
        arrAngle = PI - arrAngle
        arrayObj.X = radius * -Cos(arrAngle) + aimNode.X
    Else
        arrayObj.X = radius * Cos(arrAngle) + aimNode.X
    End If
    arrayObj.Y = radius * Sin(arrAngle) + aimNode.Y
    pT.X = arrayObj.X - pT.X
    pT.Y = arrayObj.Y - pT.Y
    process.X = -pT.X
    process.Y = -pT.Y
End Function
Private Function 角度转弧度(a) As Single
    角度转弧度 = a / 180 * PI
End Function
Public Function 关闭动作(动作名 As String) As String
    Dim aId As Long
    aId = 获得动作ID(动作名)
    If aId > 0 Then
        actionList(aId).be = False
        关闭动作 = "动作[" & 动作名 & "]已关闭！"
    Else
        关闭动作 = "未找到匹配动作，关闭无效。"
    End If
End Function
Public Function 重启动作(动作名 As String) As String
    Dim aId As Long
    aId = 获得动作ID(动作名)
    If aId > 0 Then
        With actionList(aId)
            node(.nodeID(j)).X = node(.nodeID(j)).X + .process(j).X
            node(.nodeID(j)).Y = node(.nodeID(j)).Y + .process(j).Y
            .process(j).X = 0
            .process(j).Y = 0
            .be = True
        End With
        重启动作 = "重启成功！"
    Else
        重启动作 = "未找到匹配动作，重启失败。"
    End If
End Function
Public Function 获得动作ID(动作名 As String) As Long
    Dim i As Long
    For i = 1 To UBound(actionList)
        If actionList(i).name = 动作名 Then
            获得动作ID = i
            Exit Function
        End If
    Next
End Function
Public Function 定义动作(语句 As String)
    Dim sT() As String, idT() As String, i As Long, aId As Long
    sT = Split(语句, ",")
    aId = 获得动作ID(sT(0))
    If aId = 0 Then
        ReDim Preserve actionList(UBound(actionList) + 1)
        aId = UBound(actionList)
    End If
    With actionList(aId)
        .name = sT(0)
        idT = Split(sT(1), "|")
        ReDim actionList(aId).nodeID(UBound(idT)), actionList(aId).process(UBound(idT))
        For i = 0 To UBound(idT)
            .nodeID(i) = Val(idT(i))
        Next
        .interval = Val(sT(2))
        Select Case UCase(sT(3))
            Case "L", "LINE", "直线", "直"
                .route = "直线"
                .vector.X = Val(sT(4))
                .vector.Y = Val(sT(5))
            Case "C", "CIRCLE", "圆周", "圆"
                .route = "圆周"
                .angle = Val(sT(4))
                .aimNode = Val(sT(5))
                ReDim actionList(aId).relativePosition(UBound(idT)), actionList(aId).radius(UBound(idT)), actionList(aId).initialAngle(UBound(idT))
                For i = 0 To UBound(.nodeID)
                    .relativePosition(i).X = node(.aimNode).X - node(.nodeID(i)).X
                    .relativePosition(i).Y = node(.aimNode).Y - node(.nodeID(i)).Y
                    .radius(i) = Sqr(.relativePosition(i).X ^ 2 + .relativePosition(i).Y ^ 2)
                    .initialAngle(i) = Atn(.relativePosition(i).Y / .relativePosition(i).X)
                Next
        End Select
        .endTime = Val(sT(6))
        .repeat = 字符串转布尔值(sT(7))
        .be = True
    End With
End Function
