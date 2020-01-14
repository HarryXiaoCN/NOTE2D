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
    Note.GlobalView.Line (-angleOfView.X, -angleOfView.Y)-(Note.width * zoomFactor - angleOfView.X, Note.height * zoomFactor - angleOfView.Y), RGB(255, 69, 0), B
    Note.GlobalView.FillColor = RGB(0, 250, 154)
End Function
Public Function 笔记对象绘制()
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
                If Note.矩点.Checked = False Then
                    Note.GlobalView.Circle (.X, .Y), 100, RGB(0, 250, 154)
                Else
                    Note.GlobalView.Line (.X - 100, .Y - 100)-(.X + 100, .Y + 100), RGB(0, 250, 154), BF
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
    确定边界.up = maxY + 1000: 确定边界.down = minY - 1000: 确定边界.left = minX - 1000: 确定边界.right = maxX + 1000
End Function
