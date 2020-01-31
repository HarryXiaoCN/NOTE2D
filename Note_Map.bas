Attribute VB_Name = "Note_Map"
Public Function MapUpdata()
    On Error GoTo Er
        Note.GlobalView.Cls
        绘制窗口世界视角
        MapUpdata_DetermineTheBoundary
        笔记对象绘制
Er:
End Function
Public Function MapUpdata_DetermineTheBoundary()
    Dim mapRange As 边界
    mapRange = 确定边界
    Note.GlobalView.Scale (mapRange.left, mapRange.up)-(mapRange.right, mapRange.down) '建立map的全局坐标系
    mapGetMousePosLock = False
End Function
Public Function 绘制窗口世界视角()
    Note.GlobalView.FillColor = RGB(255, 127, 80)
    Note.GlobalView.Line (-angleOfView.x, -angleOfView.y)-(Note.width * zoomFactor - angleOfView.x, Note.height * zoomFactor - angleOfView.y), RGB(255, 69, 0), B
    Note.GlobalView.FillColor = 10156544
End Function
Public Function 笔记对象绘制()
    Dim i As Long, 颜色 As Long
    For i = 0 To lSum
        With nodeLine(i)
            If .b = True Then
                If .select Then
                    颜色 = 255
                Else
                    颜色 = 10526880
                End If
                Note.GlobalView.Line (node(.Source).x, node(.Source).y)-(node(.target).x, node(.target).y), 颜色
            End If
        End With
    Next
    For i = 0 To nSum
        With node(i)
            If .b = True Then
                If .select Then
                    颜色 = 255
                Else
                    颜色 = 10156544
                End If
                Note.GlobalView.FillColor = 颜色
                If Note.矩点.Checked = False Then
                    Note.GlobalView.Circle (.x, .y), 100, 颜色
                Else
                    Note.GlobalView.Line (.x - 100, .y - 100)-(.x + 100, .y + 100), 颜色, BF
                End If
            End If
        End With
    Next
End Function
Public Function 确定边界() As 边界
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
    确定边界.up = minY - 1000: 确定边界.down = maxY + 1000: 确定边界.left = minX - 1000: 确定边界.right = maxX + 1000
End Function
