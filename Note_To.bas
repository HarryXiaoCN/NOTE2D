Attribute VB_Name = "Note_To"
Public 导入TXT文件路径 As String, 有效像素个数 As Long
Public Function NoteToImage(fP As String, image As PictureBox) As Boolean
    Dim X As Long, Y As Long, c As Long, mX As Long, mY As Long, dX As Single, dY As Single
    image.Cls
    有效像素个数 = 0
    图片边界确定 mX, mY
    If 有效像素个数 > 0 Then
        dX = image.width / image.ScaleWidth
        dY = image.height / image.ScaleHeight
        image.width = (mX + 2) * dX
        image.height = (mY + 2) * dY
        图片递归 image
        SavePicture image.image, fP
        NoteToImage = True
    Else
        MsgBox "未选中像素节点，导出终止。", 16, "导出BMP位图"
    End If
End Function
Private Function 图片边界确定(w As Long, h As Long)
    Dim i As Long, X As Long, Y As Long, c As Long
    For i = 0 To nSum
        With node(i)
            If .b Then
                If .select Then
                    If 节点内容图片解析(.text, X, Y, c) Then
                        有效像素个数 = 有效像素个数 + 1
                        If X > w Then
                            w = X
                        End If
                        If Y > h Then
                            h = Y
                        End If
                    End If
                End If
            End If
        End With
    Next
End Function
Private Function 图片递归(image As PictureBox)
    Dim i As Long, X As Long, Y As Long, c As Long
    For i = 0 To nSum
        With node(i)
            If .b Then
                If .select Then
                    If 节点内容图片解析(.text, X, Y, c) Then
                        image.Line (X, Y)-(X + 1, Y + 1), c, BF
                    End If
                End If
            End If
        End With
    Next
End Function
Public Function 节点内容图片解析(nC As String, X As Long, Y As Long, c As Long) As Boolean
    Dim sT As String, tT As Long, vT() As String
    tT = InStr(1, nC, "像素[")
    If tT > 0 Then
        sT = Mid(nC, tT + 3, InStr(1, nC, "]") - tT - 3)
        vT = Split(sT, ",")
        If UBound(vT) >= 2 Then
            X = Val(vT(0))
            Y = Val(vT(1))
            c = Val(vT(2))
            节点内容图片解析 = True
        End If
    End If
End Function

Public Function 位图节点化(image As PictureBox)
    Dim X As Long, Y As Long
    Note.MainTime.Enabled = False
    Note.MapUpdataTimer.Enabled = False
    If image.ScaleWidth * image.ScaleHeight <= 10000 Then
        For X = 0 To image.ScaleWidth - 1
            For Y = 0 To image.ScaleHeight - 1
                NodeEdit_Save nSum, "[" & X & "," & Y & "]", "像素[" & X & "," & Y & "," & image.Point(X, Y) & "]", image.Point(X, Y), nodeDefaultSize, imageToNtx_StartX + X * imageToNtx_StepX, imageToNtx_StartY + Y * imageToNtx_StepY
                nSum = nSum + 1
            Next
        Next
    Else
        MsgBox "位图像素个数大于10000，加载终止。", 64, "位图节点化"
    End If
    Note.MainTime.Enabled = True
    Note.MapUpdataTimer.Enabled = True
End Function

Public Sub NoteToTreeTXT(fP As String, nid As Long)
    Dim ntx() As String
    ClearNode_ToTreeTxtLock
    树状递归 nid, ntx, 1, 1, "", True, , True
    SaveFile_All fP, Join(ntx, vbCrLf)
End Sub
Private Function 树状递归(母节点 As Long, ntx() As String, 行 As Long, ByVal 列 As Long, lC As String, ByVal 首行 As Boolean, Optional 终结 As Boolean, Optional 主节点 As Boolean)
    Dim nS As Long, nL() As Long, lL() As Long, i As Long, eNL() As Long, eNS As Long, eLL() As Long
    node(母节点).toTreeTxtLock = True
    ReDim Preserve ntx(1 To 行)
    If 主节点 Then
        ntx(行) = node(母节点).t & vbTab & 富文本转义(node(母节点).content) & vbTab
    ElseIf 首行 Then
        ntx(行) = ntx(行) & lC & vbTab & node(母节点).t & vbTab & 富文本转义(node(母节点).content) & vbTab
    Else
        ntx(行) = 字符累积(vbTab, 2 + (列 - 2) * 3) & lC & vbTab & node(母节点).t & vbTab & 富文本转义(node(母节点).content) & vbTab
    End If
    If 终结 = False Then
        nS = 获得子节点数量(母节点, nL, lL, eNL, eLL, eNS)
        If nS > 0 Or eNS > 0 Then
            If nS > 0 Then
                For i = 1 To nS
                    If i = 1 Then
                        树状递归 nL(i), ntx, 行, 列 + 1, nodeLine(lL(i)).content, True
                    Else
                        树状递归 nL(i), ntx, 行, 列 + 1, nodeLine(lL(i)).content, False
                    End If
                Next
            Else
                For i = 1 To eNS
                    If i = 1 Then
                        树状递归 eNL(i), ntx, 行, 列 + 1, nodeLine(eLL(i)).content, True, True
                    Else
                        树状递归 eNL(i), ntx, 行, 列 + 1, nodeLine(eLL(i)).content, False, True
                    End If
                Next
            End If
        Else
            行 = 行 + 1
        End If
    Else
        行 = 行 + 1
    End If
End Function
Public Sub ClearNode_ToTreeTxtLock()
    Dim i As Long
    For i = 0 To nSum
        With node(i)
            If .b Then
                .toTreeTxtLock = False
            End If
        End With
    Next
End Sub
Private Function 字符累积(d As String, j As Long) As String
    Dim i As Long
    For i = 1 To j
        字符累积 = 字符累积 & d
    Next
End Function
Private Function 获得子节点数量(母节点 As Long, l() As Long, lL() As Long, eL() As Long, eLL() As Long, eNS As Long) As Long
    Dim i As Long
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                If .Source = 母节点 Then
                    If node(.target).toTreeTxtLock = False Then
                        获得子节点数量 = 获得子节点数量 + 1
                        ReDim Preserve l(获得子节点数量), lL(获得子节点数量)
                        l(获得子节点数量) = .target
                        lL(获得子节点数量) = i
                    Else
                        eNS = eNS + 1
                        ReDim Preserve eL(eNS), eLL(eNS)
                        eL(eNS) = .target
                        eLL(eNS) = i
                    End If
                End If
            End If
        End With
    Next
End Function




Public Sub TreeTXTToNtx()
    Dim allTxt As String, lineTemp() As String, ntxTxt() As String, lT() As String, ntx() As String, i As Long, j As Long
    ReadFile_ALL_HV 导入TXT文件路径, allTxt
    lineTemp = Split(allTxt, vbCrLf)
    ReDim ntxTxt(0)
    For i = 0 To UBound(lineTemp)
        If lineTemp(i) <> "" Then
            ReDim Preserve ntxTxt(UBound(ntxTxt) + 1)
            ntxTxt(UBound(ntxTxt)) = lineTemp(i)
        End If
    Next
    ReDim ntx(UBound(ntxTxt), GetTxtMaxColunm(ntxTxt, vbTab) + 1)
    For i = 1 To UBound(ntxTxt)
        lT = Split(ntxTxt(i), vbTab)
        For j = 0 To UBound(lT)
            ntx(i, j + 1) = lT(j)
        Next
    Next
    AnalysisTreeTXT ntx, UBound(ntx, 1), UBound(ntx, 2)
End Sub

Private Function GetTxtMaxColunm(t() As String, cS As String) As Long
    Dim i As Long, tmp() As String
    For i = 0 To UBound(t)
        tmp = Split(t(i), cS)
        If GetTxtMaxColunm < UBound(tmp) Then
            GetTxtMaxColunm = UBound(tmp)
        End If
    Next
End Function

Private Function AnalysisTreeTXT(wS() As String, maxRow As Long, maxColumn As Long)
    Dim i As Long, j As Long
    Dim deepNode() As Long
    ReDim deepNode(maxColumn)
    For i = 1 To maxRow
        For j = 1 To maxColumn Step 3
            If j = 1 Then
                If wS(i, j) <> "" Then
                    deepNode(j) = NodeEdit_NewNode(wS(i, j), wS(i, j + 1), nodeDefaultColor, nodeDefaultSize, treeTxtToNtx_StartX + treeTxtToNtx_StepX * i, treeTxtToNtx_StartY + treeTxtToNtx_StepY * j, True)
                End If
            Else
                If wS(i, j) <> "" Then
                    deepNode(j) = NodeEdit_NewNode(wS(i, j), wS(i, j + 1), nodeDefaultColor, nodeDefaultSize, treeTxtToNtx_StartX + treeTxtToNtx_StepX * i, treeTxtToNtx_StartY + treeTxtToNtx_StepY * j, True)
                    LineAdd deepNode(j - 3), deepNode(j), wS(i, j - 1), lineDefaultSize, True
                End If
            End If
        Next
    Next
End Function
Public Function SaveFile_All(fPath As String, outString As String)
    Dim fN As Integer
    fN = FreeFile
    Open fPath For Output As #fN
        Print #fN, outString
    Close #fN
End Function
Public Function ReadFile_ALL_HV(fPath As String, sourceString As String)
    Dim fN As Integer
    fN = FreeFile
    Open fPath For Binary As #fN
        sourceString = Input(LOF(1), #fN)
    Close #fN
End Function

