Attribute VB_Name = "Note_Map"
Public Function MapUpdata()

Note.GlobalView.Cls
MapUpdata_AoVMove
绘制窗口世界视角
笔记对象绘制

End Function
Public Function MapUpdata_DetermineTheBoundary()
Dim mapRange As 边界
mapRange = 确定边界
Note.GlobalView.Scale (mapRange.left, mapRange.up)-(mapRange.right, mapRange.down) '建立map的全局坐标系
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
Public Function 绘制窗口世界视角()
'Note.GlobalView.FillStyle = 7
Note.GlobalView.FillColor = RGB(255, 127, 80)
Note.GlobalView.Line (0, 0)-(Note.Width * zoomFactor, Note.Height * zoomFactor), RGB(255, 69, 0), B
Note.GlobalView.FillColor = RGB(0, 250, 154)
'Note.GlobalView.FillStyle = 0
End Function
Public Function 笔记对象绘制()
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
Public Function 确定边界() As 边界
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
确定边界.up = maxY + 1000: 确定边界.down = minY - 1000: 确定边界.left = minX - 1000: 确定边界.right = maxX + 1000
End Function
