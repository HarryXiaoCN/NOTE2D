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
    Note.GlobalView.Line (-angleOfView.X, -angleOfView.Y)-(Note.width * zoomFactor - angleOfView.X, Note.height * zoomFactor - angleOfView.Y), RGB(255, 69, 0), B
    Note.GlobalView.FillColor = RGB(0, 250, 154)
End Function
Public Function �ʼǶ������()
    Dim i As Long
    For i = 0 To lSum
        With nodeLine(i)
            If .b = True Then
                Note.GlobalView.Line (node(.Source).X, node(.Source).Y)-(node(.target).X, node(.target).Y), RGB(160, 160, 160)
            End If
        End With
    Next
    For i = 0 To nSum
        With node(i)
            If .b = True Then
                If Note.�ص�.Checked = False Then
                    Note.GlobalView.Circle (.X, .Y), 100, RGB(0, 250, 154)
                Else
                    Note.GlobalView.Line (.X - 100, .Y - 100)-(.X + 100, .Y + 100), RGB(0, 250, 154), BF
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
                    maxX = .X: minX = .X
                    maxY = .Y: minY = .Y
                Else
                    If .X > maxX Then
                        maxX = .X
                    ElseIf .X < minX Then
                        minX = .X
                    End If
                    If .Y > maxY Then
                        maxY = .Y
                    ElseIf .Y < minY Then
                        minY = .Y
                    End If
                End If
            End If
        End With
    Next
    ȷ���߽�.up = maxY + 1000: ȷ���߽�.down = minY - 1000: ȷ���߽�.left = minX - 1000: ȷ���߽�.right = maxX + 1000
End Function
