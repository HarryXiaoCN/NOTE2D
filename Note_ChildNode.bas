Attribute VB_Name = "Note_ChildNode"
Public Function �ӽڵ���(nC As String) As Boolean
    Dim sT As String, tT As Long, vT() As String
    On Error GoTo Er:
    tT = InStr(1, nC, "�ʼ�[")
    If tT > 0 Then
        sT = Mid(nC, tT + 3, InStr(1, nC, "]") - tT - 3)
        If Dir(sT) <> "" Then
            childNodeVisNtxPath = sT
            �ӽڵ��� = �ӽڵ���д(sT, Note.�ӽڵ���ͼ)
        ElseIf Dir(ntxPathNoName & "\" & sT) <> "" Then
            childNodeVisNtxPath = ntxPathNoName & "\" & sT
            �ӽڵ��� = �ӽڵ���д(sT, Note.�ӽڵ���ͼ)
        End If
    End If
Er:
End Function
Public Function ȥ��·���ļ���(s As String) As String
    Dim sT() As String
    On Error GoTo Er
        sT = Split(s, "\")
        ReDim Preserve sT(UBound(sT) - 1)
        ȥ��·���ļ��� = Join(sT, "\")
    Exit Function
Er:
    Debug.Print "ȥ��·���ļ���"; s
End Function
Public Function �ӽڵ���д(nFP As String, pic As PictureBox) As Boolean
    Dim ntx() As String
    On Error GoTo Er
    If ChildNodeFileRead(nFP, ntx) = True Then
        pic.Cls
        ChildNoteFileRead_202 ntx, pic
        �ӽڵ���д = True
    End If
    Exit Function
Er:
End Function

Public Function ChildNoteFileRead_202(ntx() As String, pic As PictureBox)
    Dim i As Long, nodeSum As Long, lineSum As Long, startNodeId As Long, lineTmp() As String, j As Long
    Dim picMinX As Single, picMinY As Single, picMaxX As Single, picMaxY As Single
    Dim dX As Single, dY As Single, midX As Single, midY As Single
    Dim childNode() As �ڵ�, childLine() As ����
    
    lineTmp = Split(ntx(0), LINEBREAK)
    nodeSum = Val(lineTmp(1))
    lineSum = Val(lineTmp(2))
    startNodeId = nSum
    If nodeSum > 0 Then
        ReDim childNode(nodeSum - 1)
        For i = 1 To nodeSum
            lineTmp = Split(ntx(i), LINEBREAK)
            childNode(i - 1).X = Val(lineTmp(0))
            childNode(i - 1).Y = Val(lineTmp(1))
            childNode(i - 1).t = lineTmp(2)
            childNode(i - 1).setColor = Val(lineTmp(4))
            childNode(i - 1).setSize = Val(lineTmp(5))
        Next
        If lineSum > 0 Then
            ReDim childLine(lineSum - 1)
            For i = nodeSum + 1 To nodeSum + lineSum
                lineTmp = Split(ntx(i), LINEBREAK)
                childLine(i - nodeSum - 1).Source = Val(lineTmp(0))
                childLine(i - nodeSum - 1).target = Val(lineTmp(1))
                If UBound(lineTmp) > 2 Then
                    childLine(i - nodeSum - 1).content = lineTmp(2)
                    childLine(i - nodeSum - 1).size = Val(lineTmp(3))
                End If
            Next
        End If
        picMinX = childNode(0).X
        picMinY = childNode(0).Y
        For i = 0 To UBound(childNode)
            With childNode(i)
                If .X < picMinX Then
                    picMinX = .X
                ElseIf .X > picMaxX Then
                    picMaxX = .X
                End If
                If .Y < picMinY Then
                    picMinY = .Y
                ElseIf .Y > picMaxY Then
                    picMaxY = .Y
                End If
            End With
        Next
        dX = 3000 - picMinX
        dY = 3000 - picMinY
        
        pic.width = picMaxX + 3000 + dX
        pic.height = picMaxY + 3000 + dY
        
        pic.Scale (0, pic.height)-(pic.width, 0)
        
        For i = 0 To UBound(childNode)
            With childNode(i)
                .X = .X + dX
                .Y = .Y + dY
            End With
        Next
        If lineSum > 0 Then
            For i = 0 To UBound(childLine)
                With childLine(i)
                    pic.DrawWidth = .size
                    midX = (childNode(.target).X - childNode(.Source).X) / 3 * 2 + childNode(.Source).X
                    midY = (childNode(.target).Y - childNode(.Source).Y) / 3 * 2 + childNode(.Source).Y
                    pic.Line (childNode(.Source).X, childNode(.Source).Y)-(midX, midY), childNode(.Source).setColor
                    pic.Line (midX, midY)-(childNode(.target).X, childNode(.target).Y), childNode(.target).setColor
                    pic.CurrentX = midX
                    pic.CurrentY = midY
                    pic.Print .content
                End With
            Next
        End If
        For i = 0 To UBound(childNode, 1)
            With childNode(i)
                pic.FillColor = .setColor
                pic.Circle (.X, .Y), .setSize, .setColor
                pic.CurrentX = .X
                pic.CurrentY = .Y
                pic.Print .t
            End With
        Next
    End If
End Function

Public Function ChildNodeFileRead(filePath As String, ntx() As String) As Boolean
    Dim fN As Integer, i As Long, version As Long, lT As String
    fN = FreeFile
    Open filePath For Input As #fN
            Do While Not EOF(fN)
                Line Input #fN, lT
                If lT = "" Then
                    Exit Do
                Else
                    ReDim Preserve ntx(i)
                    ntx(i) = lT
                    i = i + 1
                End If
            Loop
    Close #fN
    version = NoteFileRead_VersionCheck(ntx(0))
    Select Case version
        Case 202, 203
            ChildNodeFileRead = True
    End Select
End Function
