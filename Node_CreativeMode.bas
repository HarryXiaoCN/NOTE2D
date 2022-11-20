Attribute VB_Name = "Node_CreativeMode"
Public Sub 进入创建模式()
    nodeCreativeMode = True
End Sub
Public Sub 新节点放置检查(p As 二维坐标)
    If (p.X - nodeCreativeStartPos.X) ^ 2 + (p.Y - nodeCreativeStartPos.Y) ^ 2 >= nodeAttributedToIntegers ^ 2 Then
        nodeCreativeStartPos = p
        顺序放置节点 p
    End If
End Sub
Public Sub 顺序放置节点(p As 二维坐标)
    Dim i As Long, tmpID As Long
    BehaviorIdSet
    For i = nodeCreativeListStart To UBound(NodeCreativeList)
        With NodeCreativeList(i)
            If .b Then
                tmpID = NodeEdit_NewNode(.t, .content, .setColor, .setSize, p.X, p.Y)
                LineAdd nodeCreativeSourceId, tmpID, "", lineDefaultSize, , , True
                .b = False
                nodeCreativeListStart = i + 1
                Exit For
            End If
        End With
    Next
    If nodeCreativeListStart > UBound(NodeCreativeList) Then
'        ReDim NodeCreativeList(0)
        nodeCreativeMode = False
    End If
End Sub
Public Sub 剩余节点阵列(起点 As 二维坐标)
    Dim i As Long, tmpID As Long
    BehaviorIdSet
    For i = nodeCreativeListStart To UBound(NodeCreativeList)
        With NodeCreativeList(i)
            If .b Then
                tmpID = NodeEdit_NewNode(.t, .content, .setColor, .setSize, 起点.X + (i - nodeCreativeListStart - 1) * nodeAttributedToIntegers, 起点.Y)
                LineAdd nodeCreativeSourceId, tmpID, "", lineDefaultSize, , , True
                .b = False
            End If
        End With
    Next
    ReDim NodeCreativeList(0)
    nodeCreativeMode = False
End Sub

Public Function 获取创建节点列表剩余节点个数() As Long
    获取创建节点列表剩余节点个数 = UBound(NodeCreativeList) - nodeCreativeListStart + 1
End Function
