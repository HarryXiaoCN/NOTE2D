Attribute VB_Name = "Note_Map"
Public Function MapUpdata()
    On Error GoTo Er
        Note.GlobalView.Cls
        ���ƴ��������ӽ�
        MapUpdata_DetermineTheBoundary
        �ʼǶ������
Er:
End Function
Public Function MapUpdata_DetermineTheBoundary()
    Dim mapRange As �߽�
    mapRange = ȷ���߽�
    Note.GlobalView.Scale (mapRange.left, mapRange.up)-(mapRange.right, mapRange.down) '����map��ȫ������ϵ
    mapGetMousePosLock = False
End Function
Public Function ���ƴ��������ӽ�()
    Note.GlobalView.FillColor = RGB(255, 127, 80)
    Note.GlobalView.Line (-angleOfView.x, -angleOfView.y)-(Note.width * zoomFactor - angleOfView.x, Note.height * zoomFactor - angleOfView.y), RGB(255, 69, 0), B
    Note.GlobalView.FillColor = 10156544
End Function
Public Function �ʼǶ������()
    Dim i As Long, ��ɫ As Long
    For i = 0 To lSum
        With nodeLine(i)
            If .b = True Then
                If .select Then
                    ��ɫ = 255
                Else
                    ��ɫ = 10526880
                End If
                Note.GlobalView.Line (node(.Source).x, node(.Source).y)-(node(.target).x, node(.target).y), ��ɫ
            End If
        End With
    Next
    For i = 0 To nSum
        With node(i)
            If .b = True Then
                If .select Then
                    ��ɫ = 255
                Else
                    ��ɫ = 10156544
                End If
                Note.GlobalView.FillColor = ��ɫ
                If Note.�ص�.Checked = False Then
                    Note.GlobalView.Circle (.x, .y), 100, ��ɫ
                Else
                    Note.GlobalView.Line (.x - 100, .y - 100)-(.x + 100, .y + 100), ��ɫ, BF
                End If
            End If
        End With
    Next
End Function
Public Function ȷ���߽�() As �߽�
    Dim maxX As Single, maxY  As Single, minX As Single, minY As Single
    Dim i As Long
    For i = 0 To nSum
        With node(i)
            If .b = True Then
                If i = 0 Then
                    maxX = .x: minX = .x
                    maxY = .y: minY = .y
                Else
                    If .x > maxX Then
                        maxX = .x
                    ElseIf .x < minX Then
                        minX = .x
                    End If
                    If .y > maxY Then
                        maxY = .y
                    ElseIf .y < minY Then
                        minY = .y
                    End If
                End If
            End If
        End With
    Next
    ȷ���߽�.up = minY - 1000: ȷ���߽�.down = maxY + 1000: ȷ���߽�.left = minX - 1000: ȷ���߽�.right = maxX + 1000
End Function
