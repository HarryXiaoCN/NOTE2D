Attribute VB_Name = "Note_Print"
Public Function NotePrint(pB As PictureBox) As 四元数
    Dim i As Long, mX As Single, mY As Single, dX As Single, dY As Single, pX As Single, pY As Single, pX2 As Single, pY2 As Single
    pB.Cls
    获得最小值 NotePrint.xS, NotePrint.yS, NotePrint.xE, NotePrint.yE
    dX = 3000 - NotePrint.xS
    dY = 3000 - NotePrint.yS
    pB.width = NotePrint.xE + dX + 3000
    pB.height = NotePrint.yE + dY + 3000
    pB.Scale (0, 0)-(pB.width, pB.height)
    pB.Font = Note.Font
    pB.Font.size = Note.Font.size
    pB.ForeColor = Note.ForeColor
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                pX = node(.Source).X + dX
                pY = node(.Source).Y + dY
                pX2 = node(.target).X + dX
                pY2 = node(.target).Y + dY
'                mX = (pX + pX2) / 2
'                mY = (pY + pY2) / 2
                mX = (pX2 - pX) / 3 * 2 + pX
                mY = (pY2 - pY) / 3 * 2 + pY
                pB.DrawWidth = .size * 2
                pB.Line (pX, pY)-(mX, mY), node(.Source).setColor
                pB.DrawWidth = .size
                pB.Line (mX, mY)-(pX2, pY2), node(.target).setColor
                pB.CurrentX = mX
                pB.CurrentY = mY
                pB.Print .content
            End If
        End With
    Next
    
    For i = 0 To nSum
        With node(i)
            If .b Then
                pX = .X + dX
                pY = .Y + dY
                pB.FillColor = .setColor
                pB.Circle (pX, pY), .setSize, .setColor
                pB.CurrentX = pX
                pB.CurrentY = pY
                pB.Print .t
            End If
        End With
    Next
End Function
Private Function 获得最小值(minX, minY, maxX, maxY)
    Dim i As Long
    For i = 0 To nSum
        With node(i)
            If .b Then
                minX = .X
                minY = .Y
                maxX = .X
                maxY = .Y
                Exit For
            End If
        End With
    Next
    For i = 0 To nSum
        With node(i)
            If .b Then
                If minX > .X Then minX = .X
                If minY > .Y Then minY = .Y
                If maxX < .X Then maxX = .X
                If maxY < .Y Then maxY = .Y
            End If
        End With
    Next
End Function
