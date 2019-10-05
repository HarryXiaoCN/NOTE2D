Attribute VB_Name = "Note_To"
Public 导入TXT文件路径 As String
Public excelFile As Excel.Application
Private 最大深度 As Long

Public Sub NoteToTreeTXT(fP As String, nId As Long)
    Dim wB As Excel.Workbook, wS As Excel.Worksheet
    Set excelFile = New Excel.Application
    excelFile.Workbooks.Add
    Set wB = excelFile.Workbooks(1)
    Set wS = wB.Sheets(1)
    NodeDepthCls
    最大深度 = 0
    NoteToTreeTXT_Write wS, nId
    
    wB.SaveAs fP
    wB.Close
    excelFile.Quit
    Set excelFile = Nothing
End Sub
Private Sub NodeDepthCls()
    Dim i As Long
    For i = 0 To nSum
        With node(i)
            If .b Then
                .depth = 0
                .depthSum = 0
            End If
        End With
    Next
End Sub
Private Function NoteToTreeTXT_Write(wS As Excel.Worksheet, soucerNodeId As Long)
    Dim maxDepthSum As Long
    Dim i As Long
    node(soucerNodeId).depth = 1
    GetSourceNodeMaxDepth soucerNodeId, soucerNodeId, 1
    maxDepthSum = GetmaxDepthSum(最大深度)
    For i = 最大深度 - 1 To 1 Step -1
        GetmaxDepthSum i
    Next
    For i = 0 To nSum
        With node(i)
            If .b Then
                If .depth > 0 And .depthSum > 0 Then
                    wS.Cells(.depthSum, .depth * 2 - 1).value = .t
                    Note.RTBtemp.TextRTF = .content
                    wS.Cells(.depthSum, .depth * 2).value = Note.RTBtemp.Text
                End If
            End If
        End With
    Next
End Function
Private Function GetmaxDepthSum(mD As Long) As Long
    Dim i As Long
    For i = 0 To nSum
        With node(i)
            If .b Then
                If .depth = mD Then
                    GetmaxDepthSum = GetmaxDepthSum + 1
                    .depthSum = GetmaxDepthSum
                End If
            End If
        End With
    Next
End Function
Private Function GetSourceNodeMaxDepth(soucerNodeId As Long, ByVal tempNId As Long, ByVal d As Long) As Long
    Dim i As Long, dT As Long
    dT = d
    For i = 0 To lSum
        d = dT
        With nodeLine(i)
            If .b Then
                If .source = tempNId And .target <> soucerNodeId Then
                    d = d + 1
                    node(.target).depth = d
                    GetSourceNodeMaxDepth = GetSourceNodeMaxDepth(soucerNodeId, .target, d)
                    If 最大深度 < GetSourceNodeMaxDepth Then
                        最大深度 = GetSourceNodeMaxDepth
                    End If
                ElseIf .target = soucerNodeId Then
                    Exit For
                End If
            End If
        End With
    Next
    GetSourceNodeMaxDepth = d
End Function



Public Sub TreeTXTToNtx()
    Dim wB As Excel.Workbook, wS As Excel.Worksheet
    Set excelFile = New Excel.Application
    excelFile.Workbooks.Open 导入TXT文件路径
    Set wB = excelFile.Workbooks(1)
    Set wS = wB.Sheets(1)
    analysisTreeTXT wS
    excelFile.Workbooks.Close
    excelFile.Quit
    Set excelFile = Nothing
End Sub

Private Function analysisTreeTXT(wS As Excel.Worksheet)
    Dim maxColumn As Long, maxRow As Long
    Dim i As Long, j As Long
    Dim deepNode() As Long
    maxColumn = wS.UsedRange.Columns.Count
    maxRow = wS.UsedRange.Rows.Count
    maxColumn = (maxColumn \ 2) * 2
    maxRow = (maxRow \ 2) * 2
    ReDim deepNode(maxColumn)
    With wS
        For i = 1 To maxRow
            For j = 1 To maxColumn Step 2
                If j = 1 Then
                    If .Cells(i, j).value <> "" Then
                        deepNode(j) = NodeEdit_NewNode(.Cells(i, j).value, .Cells(i, j + 1).value, &HFFBF00, nodeDefaultSize, OUT_CONST_X + OUT_CONST_STEP * i, OUT_CONST_Y + OUT_CONST_STEP * j, True)
                    End If
                Else
                    If .Cells(i, j).value <> "" Then
                        deepNode(j) = NodeEdit_NewNode(.Cells(i, j).value, .Cells(i, j + 1).value, &HFFBF00, nodeDefaultSize, OUT_CONST_X + OUT_CONST_STEP * i, OUT_CONST_Y + OUT_CONST_STEP * j, True)
                        LineAdd deepNode(j - 2), deepNode(j), "", lineDefaultSize, True
                    End If
                End If
            Next
        Next
    End With
End Function
