Attribute VB_Name = "Node_CreativeMode"
Public Sub ���봴��ģʽ()
    nodeCreativeMode = True
End Sub
Public Sub �½ڵ���ü��(p As ��ά����)
    If (p.X - nodeCreativeStartPos.X) ^ 2 + (p.Y - nodeCreativeStartPos.Y) ^ 2 >= nodeAttributedToIntegers ^ 2 Then
        nodeCreativeStartPos = p
        ˳����ýڵ� p
    End If
End Sub
Public Sub ˳����ýڵ�(p As ��ά����)
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
Public Sub ʣ��ڵ�����(��� As ��ά����)
    Dim i As Long, tmpID As Long
    BehaviorIdSet
    For i = nodeCreativeListStart To UBound(NodeCreativeList)
        With NodeCreativeList(i)
            If .b Then
                tmpID = NodeEdit_NewNode(.t, .content, .setColor, .setSize, ���.X + (i - nodeCreativeListStart - 1) * nodeAttributedToIntegers, ���.Y)
                LineAdd nodeCreativeSourceId, tmpID, "", lineDefaultSize, , , True
                .b = False
            End If
        End With
    Next
    ReDim NodeCreativeList(0)
    nodeCreativeMode = False
End Sub

Public Function ��ȡ�����ڵ��б�ʣ��ڵ����() As Long
    ��ȡ�����ڵ��б�ʣ��ڵ���� = UBound(NodeCreativeList) - nodeCreativeListStart + 1
End Function
