Attribute VB_Name = "Node_Fictitious"
Public Function FictitiousCheck()
    Dim i As Long, j As Long, m As Long
    fictitiousIndexLock = False
    On Error GoTo Er:
    For i = 0 To UBound(fictitiousNote)
        With fictitiousNote(i)
            If .be Then
                FictitiousIndex_LineNameErase .nodeLine
                For j = 0 To UBound(.node)
                    If .node(j).t = fictitiousIndexName Then
                        For m = 0 To UBound(.nodeLine)
                            If .nodeLine(m).Source = j Then
                                If FictitiousIndex_NodeNameCheck(.node(.nodeLine(m).target).t, .nodeLine(m).realityId) = False Then
                                    fictitiousIndexLock = True
                                    .nodeLine(m).direction = 1
                                    .node(.nodeLine(m).target).realityX = .node(.nodeLine(m).target).x - .node(j).x + node(fictitiousRootNodeId).x
                                    .node(.nodeLine(m).target).realityY = .node(.nodeLine(m).target).y - .node(j).y + node(fictitiousRootNodeId).y
                                ElseIf FictitiousIndex_LineNameCheck(node(fictitiousRootNodeId).t, .node(.nodeLine(m).target).t) = False Then
                                    fictitiousIndexLock = True
                                    .nodeLine(m).direction = 3
                                End If
                            ElseIf .nodeLine(m).target = j Then
                                If FictitiousIndex_NodeNameCheck(.node(.nodeLine(m).Source).t, .nodeLine(m).realityId) = False Then
                                    fictitiousIndexLock = True
                                    .nodeLine(m).direction = 2
                                    .node(.nodeLine(m).Source).realityX = .node(.nodeLine(m).Source).x - .node(j).x + node(fictitiousRootNodeId).x
                                    .node(.nodeLine(m).Source).realityY = .node(.nodeLine(m).Source).y - .node(j).y + node(fictitiousRootNodeId).y
                                ElseIf FictitiousIndex_LineNameCheck(.node(.nodeLine(m).Source).t, node(fictitiousRootNodeId).t) = False Then
                                    fictitiousIndexLock = True
                                    .nodeLine(m).direction = 4
                                End If
                            End If
                        Next
                    End If
                Next
            End If
        End With
    Next
Er:
End Function
Public Sub FictitiousIndex_LineNameErase(fL() As 虚拟连接)
    Dim i As Long
    For i = 0 To UBound(fL)
        fL(i).direction = 0
    Next
End Sub
Public Function FictitiousIndex_LineNameCheck(s As String, t As String) As Boolean
    Dim i As Long
    For i = 0 To lSum
        If nodeLine(i).b Then
            If (node(fictitiousRootNodeId).t = t And node(nodeLine(i).Source).t = s) Or (node(fictitiousRootNodeId).t = s And node(nodeLine(i).Source).t = t) Then
                FictitiousIndex_LineNameCheck = True
                Exit Function
            End If
        End If
    Next
End Function
Public Function FictitiousIndex_NodeNameCheck(t As String, realityId As Long) As Boolean
    Dim i As Long
    For i = 0 To nSum
        If node(i).b Then
            If node(i).t = t Then
                realityId = i
                FictitiousIndex_NodeNameCheck = True
                Exit Function
            End If
        End If
    Next
End Function
Public Function Fictitious_NoteFileRead(ficNote As 虚拟笔记, filePath As String)
    Dim ntx() As String, i As Long
    On Error GoTo Er
        Open filePath For Input As #1
                Do While Not EOF(1)
                    ReDim Preserve ntx(i)
                    Line Input #1, ntx(i)
                    If ntx(i) = "" Then Exit Do
                    i = i + 1
                Loop
        Close #1
        Select Case NoteFileRead_VersionCheck(ntx(0))
            Case 202, 203, 204
            Fictitious_NoteFileRead_202 ficNote, ntx
            ficNote.be = True
        End Select
    Exit Function
Er:
    Debug.Print "文件读取失败，原因：" & Err.Description
End Function
Private Function GetFileList(ByVal path As String, ByRef filename() As String, Optional fExp As String = "*.*") As Boolean
    Dim fName As String, i As Long
    fName = Dir(path & fExp)
    ReDim filename(0)
    Do While fName <> ""
        ReDim Preserve filename(UBound(filename) + 1)
        filename(UBound(filename) - 1) = fName
        fName = Dir
        GetFileList = True
    Loop
    If GetFileList Then
        ReDim Preserve filename(UBound(filename) - 1)
    End If
End Function

Public Function FictitiousNtxLoad()
    Dim i As Long, fN() As String
    Debug.Print fictitiousNtxPath
    'fictitiousNtxPath
    On Error GoTo Er
    If Dir(fictitiousNtxPath, vbDirectory) <> "" Then
        If GetFileList(fictitiousNtxPath, fN, "*.ntx") Then
            ReDim fictitiousNote(UBound(fN))
            For i = 0 To UBound(fN)
                Fictitious_NoteFileRead fictitiousNote(i), fictitiousNtxPath & fN(i)
            Next
        End If
    Else
        MkDir fictitiousNtxPath
    End If
Er:
End Function

Public Function Fictitious_NoteFileRead_202(ficNote As 虚拟笔记, ntx() As String)
    Dim i As Long, nodeSum As Long, lineSum As Long, startNodeId As Long: Dim lineTmp() As String
    lineTmp = Split(ntx(0), LINEBREAK)
    nodeSum = Val(lineTmp(1))
    lineSum = Val(lineTmp(2))
    ReDim ficNote.node(nodeSum - 1)
    ReDim ficNote.nodeLine(lineSum - 1)
    For i = 1 To nodeSum
        lineTmp = Split(ntx(i), LINEBREAK)
        With ficNote.node(i - 1)
            .t = lineTmp(2)
            .content = Replace(lineTmp(3), NODELINEBREAK, vbCrLf)
            .setColor = Val(lineTmp(4))
            .setSize = Val(lineTmp(5))
            .x = Val(lineTmp(0))
            .y = Val(lineTmp(1))
        End With
        DoEvents
    Next
    For i = nodeSum + 1 To nodeSum + lineSum
        lineTmp = Split(ntx(i), LINEBREAK)
        With ficNote.nodeLine(i - 1 - nodeSum)
            .Source = Val(lineTmp(0))
            .target = Val(lineTmp(1))
            If UBound(lineTmp) < 3 Then
                .content = ""
                .size = lineDefaultSize
            Else
                .content = lineTmp(2)
                .size = Val(lineTmp(3))
            End If
        End With
        DoEvents
    Next
End Function
