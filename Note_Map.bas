Attribute VB_Name = "Note_Map"
Public Function MapUpdata()

Note.GlobalView.Cls
MapUpdata_AoVMove
���ƴ��������ӽ�
�ʼǶ������

End Function
Public Function MapUpdata_DetermineTheBoundary()
Dim mapRange As �߽�
mapRange = ȷ���߽�
Note.GlobalView.Scale (mapRange.left, mapRange.up)-(mapRange.right, mapRange.down) '����map��ȫ������ϵ
mapGetMousePosLock = False
End Function
Public Function MapUpdata_AoVMove()
Dim dx As Single: Dim dy As Single
Dim zFx, zFy As Single
If mapMoveLock = True Then
    zFx = Note.Width * zoomFactor / 2
    zFy = Note.Height * zoomFactor / 2
    dx = zFx - mouseMapPos.x
    dy = zFy - mouseMapPos.y
    MapUpdata_AoVMove_Moving dx, dy
    mouseMapPos.x = zFx
    mouseMapPos.y = zFy
    mapGetMousePosLock = True
End If
MapUpdata_DetermineTheBoundary
End Function
Public Function MapUpdata_AoVMove_Moving(ByRef dx As Single, ByRef dy As Single)
Dim i As Long
For i = 0 To nSum
    With node(i)
        If .b = True Then
                .x = .x + dx
                .y = .y + dy
        End If
    End With
Next
End Function
Public Function ���ƴ��������ӽ�()
'Note.GlobalView.FillStyle = 7
Note.GlobalView.FillColor = RGB(255, 127, 80)
Note.GlobalView.Line (0, 0)-(Note.Width * zoomFactor, Note.Height * zoomFactor), RGB(255, 69, 0), B
Note.GlobalView.FillColor = RGB(0, 250, 154)
'Note.GlobalView.FillStyle = 0
End Function
Public Function �ʼǶ������()
Dim i As Long
For i = 0 To lSum
    With nodeLine(i)
        If .b = True Then
            Note.GlobalView.Line (node(.source).x, node(.source).y)-(node(.target).x, node(.target).y), RGB(160, 160, 160)
        End If
    End With
Next
For i = 0 To nSum
    With node(i)
        If .b = True Then
            Note.GlobalView.Circle (.x, .y), 100, RGB(0, 250, 154)
        End If
    End With
Next
End Function
Public Function ȷ���߽�() As �߽�
Dim maxX, maxY, minX, minY As Single
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
ȷ���߽�.up = maxY + 1000: ȷ���߽�.down = minY - 1000: ȷ���߽�.left = minX - 1000: ȷ���߽�.right = maxX + 1000
End Function
