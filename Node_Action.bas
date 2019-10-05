Attribute VB_Name = "Node_Action"
Public Type ACTION_LIST
    be As Boolean
    name As String '��������
    nodeID() As Long '���õĽڵ�ID
    interval As Long '����ʱ������ִ��һ�ζ���
    intervalTemp As Long '��ǰ������ʱ������
    process() As ��ά���� '����ڵ��ڶ�����ִ�е��м�������������ص���ʼλ��
    route As String 'ֱ�߻���Բ��
    vector As ��ά���� 'ֱ���ƶ�������
    endTime As Long 'ֱ��/Բ���ƶ�����ֹ����
    endTimeTemp As Long '��ֹ��������
    angle As Single 'Բ���˶��ĽǶ�
    relativePosition() As ��ά���� '��Ŀ��ڵ�����λ��
    radius() As Single '��Ŀ��ڵ�İ뾶
    initialAngle() As Single '��ʼ�н�
    aimNode As Long 'Բ��Ŀ���˶��Ľڵ�
    repeat As Boolean '����ֹλ�ú�ڵ��Ƿ�ص���ʼλ������ִ�ж���
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
                        Case "ֱ��"
                            For j = 0 To UBound(.nodeID)
                                node(.nodeID(j)).X = node(.nodeID(j)).X + .vector.X
                                node(.nodeID(j)).Y = node(.nodeID(j)).Y + .vector.Y
                                .process(j).X = .process(j).X - .vector.X
                                .process(j).Y = .process(j).Y - .vector.Y
                            Next
                        Case "Բ��"
                            For j = 0 To UBound(.nodeID)
                                node(.nodeID(j)).X = node(.aimNode).X + .relativePosition(j).X
                                node(.nodeID(j)).Y = node(.aimNode).Y + .relativePosition(j).Y
                                ��ýǶȺ����� node(.nodeID(j)), node(.aimNode), .angle, .endTimeTemp, .radius(j), .relativePosition(j), .process(j), .initialAngle(j)
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
Public Function ��ýǶȺ�����(arrayObj As �ڵ�, aimNode As �ڵ�, angle As Single, endTimeTemp As Long, radius As Single, relativePosition As ��ά����, process As ��ά����, initialAngle As Single)
    Dim arrAngle As Single, pT As ��ά����
    arrAngle = �Ƕ�ת����(angle) * endTimeTemp + initialAngle
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
Private Function �Ƕ�ת����(a) As Single
    �Ƕ�ת���� = a / 180 * PI
End Function
Public Function �رն���(������ As String) As String
    Dim aId As Long
    aId = ��ö���ID(������)
    If aId > 0 Then
        actionList(aId).be = False
        �رն��� = "����[" & ������ & "]�ѹرգ�"
    Else
        �رն��� = "δ�ҵ�ƥ�䶯�����ر���Ч��"
    End If
End Function
Public Function ��������(������ As String) As String
    Dim aId As Long
    aId = ��ö���ID(������)
    If aId > 0 Then
        With actionList(aId)
            node(.nodeID(j)).X = node(.nodeID(j)).X + .process(j).X
            node(.nodeID(j)).Y = node(.nodeID(j)).Y + .process(j).Y
            .process(j).X = 0
            .process(j).Y = 0
            .be = True
        End With
        �������� = "�����ɹ���"
    Else
        �������� = "δ�ҵ�ƥ�䶯��������ʧ�ܡ�"
    End If
End Function
Public Function ��ö���ID(������ As String) As Long
    Dim i As Long
    For i = 1 To UBound(actionList)
        If actionList(i).name = ������ Then
            ��ö���ID = i
            Exit Function
        End If
    Next
End Function
Public Function ���嶯��(��� As String)
    Dim sT() As String, idT() As String, i As Long, aId As Long
    sT = Split(���, ",")
    aId = ��ö���ID(sT(0))
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
            Case "L", "LINE", "ֱ��", "ֱ"
                .route = "ֱ��"
                .vector.X = Val(sT(4))
                .vector.Y = Val(sT(5))
            Case "C", "CIRCLE", "Բ��", "Բ"
                .route = "Բ��"
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
        .repeat = �ַ���ת����ֵ(sT(7))
        .be = True
    End With
End Function
