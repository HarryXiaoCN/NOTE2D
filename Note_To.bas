Attribute VB_Name = "Note_To"
Public ����TXT�ļ�·�� As String, ��Ч���ظ��� As Long
Public Function NoteToImage(fP As String, image As PictureBox) As Boolean
    Dim X As Long, Y As Long, c As Long, mX As Long, mY As Long, dX As Single, dY As Single
    image.Cls
    ��Ч���ظ��� = 0
    ͼƬ�߽�ȷ�� mX, mY
    If ��Ч���ظ��� > 0 Then
        dX = image.width / image.ScaleWidth
        dY = image.height / image.ScaleHeight
        image.width = (mX + 2) * dX
        image.height = (mY + 2) * dY
        ͼƬ�ݹ� image
        SavePicture image.image, fP
        NoteToImage = True
    Else
        MsgBox "δѡ�����ؽڵ㣬������ֹ��", 16, "����BMPλͼ"
    End If
End Function
Private Function ͼƬ�߽�ȷ��(w As Long, h As Long)
    Dim i As Long, X As Long, Y As Long, c As Long
    For i = 0 To nSum
        With node(i)
            If .b Then
                If .select Then
                    If �ڵ�����ͼƬ����(.text, X, Y, c) Then
                        ��Ч���ظ��� = ��Ч���ظ��� + 1
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
Private Function ͼƬ�ݹ�(image As PictureBox)
    Dim i As Long, X As Long, Y As Long, c As Long
    For i = 0 To nSum
        With node(i)
            If .b Then
                If .select Then
                    If �ڵ�����ͼƬ����(.text, X, Y, c) Then
                        image.Line (X, Y)-(X + 1, Y + 1), c, BF
                    End If
                End If
            End If
        End With
    Next
End Function
Public Function �ڵ�����ͼƬ����(nC As String, X As Long, Y As Long, c As Long) As Boolean
    Dim sT As String, tT As Long, vT() As String
    tT = InStr(1, nC, "����[")
    If tT > 0 Then
        sT = Mid(nC, tT + 3, InStr(1, nC, "]") - tT - 3)
        vT = Split(sT, ",")
        If UBound(vT) >= 2 Then
            X = Val(vT(0))
            Y = Val(vT(1))
            c = Val(vT(2))
            �ڵ�����ͼƬ���� = True
        End If
    End If
End Function

Public Function λͼ�ڵ㻯(image As PictureBox)
    Dim X As Long, Y As Long
    Note.MainTime.Enabled = False
    Note.MapUpdataTimer.Enabled = False
    If image.ScaleWidth * image.ScaleHeight <= 10000 Then
        For X = 0 To image.ScaleWidth - 1
            For Y = 0 To image.ScaleHeight - 1
                NodeEdit_Save nSum, "[" & X & "," & Y & "]", "����[" & X & "," & Y & "," & image.Point(X, Y) & "]", image.Point(X, Y), nodeDefaultSize, imageToNtx_StartX + X * imageToNtx_StepX, imageToNtx_StartY + Y * imageToNtx_StepY
                nSum = nSum + 1
            Next
        Next
    Else
        MsgBox "λͼ���ظ�������10000��������ֹ��", 64, "λͼ�ڵ㻯"
    End If
    Note.MainTime.Enabled = True
    Note.MapUpdataTimer.Enabled = True
End Function

Public Sub NoteToTreeTXT(fP As String, nid As Long)
    Dim ntx() As String
    ClearNode_ToTreeTxtLock
    ��״�ݹ� nid, ntx, 1, 1, "", True, , True
    SaveFile_All fP, Join(ntx, vbCrLf)
End Sub
Private Function ��״�ݹ�(ĸ�ڵ� As Long, ntx() As String, �� As Long, ByVal �� As Long, lC As String, ByVal ���� As Boolean, Optional �ս� As Boolean, Optional ���ڵ� As Boolean)
    Dim nS As Long, nL() As Long, lL() As Long, i As Long, eNL() As Long, eNS As Long, eLL() As Long
    node(ĸ�ڵ�).toTreeTxtLock = True
    ReDim Preserve ntx(1 To ��)
    If ���ڵ� Then
        ntx(��) = node(ĸ�ڵ�).t & vbTab & ���ı�ת��(node(ĸ�ڵ�).content) & vbTab
    ElseIf ���� Then
        ntx(��) = ntx(��) & lC & vbTab & node(ĸ�ڵ�).t & vbTab & ���ı�ת��(node(ĸ�ڵ�).content) & vbTab
    Else
        ntx(��) = �ַ��ۻ�(vbTab, 2 + (�� - 2) * 3) & lC & vbTab & node(ĸ�ڵ�).t & vbTab & ���ı�ת��(node(ĸ�ڵ�).content) & vbTab
    End If
    If �ս� = False Then
        nS = ����ӽڵ�����(ĸ�ڵ�, nL, lL, eNL, eLL, eNS)
        If nS > 0 Or eNS > 0 Then
            If nS > 0 Then
                For i = 1 To nS
                    If i = 1 Then
                        ��״�ݹ� nL(i), ntx, ��, �� + 1, nodeLine(lL(i)).content, True
                    Else
                        ��״�ݹ� nL(i), ntx, ��, �� + 1, nodeLine(lL(i)).content, False
                    End If
                Next
            Else
                For i = 1 To eNS
                    If i = 1 Then
                        ��״�ݹ� eNL(i), ntx, ��, �� + 1, nodeLine(eLL(i)).content, True, True
                    Else
                        ��״�ݹ� eNL(i), ntx, ��, �� + 1, nodeLine(eLL(i)).content, False, True
                    End If
                Next
            End If
        Else
            �� = �� + 1
        End If
    Else
        �� = �� + 1
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
Private Function �ַ��ۻ�(d As String, j As Long) As String
    Dim i As Long
    For i = 1 To j
        �ַ��ۻ� = �ַ��ۻ� & d
    Next
End Function
Private Function ����ӽڵ�����(ĸ�ڵ� As Long, l() As Long, lL() As Long, eL() As Long, eLL() As Long, eNS As Long) As Long
    Dim i As Long
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                If .Source = ĸ�ڵ� Then
                    If node(.target).toTreeTxtLock = False Then
                        ����ӽڵ����� = ����ӽڵ����� + 1
                        ReDim Preserve l(����ӽڵ�����), lL(����ӽڵ�����)
                        l(����ӽڵ�����) = .target
                        lL(����ӽڵ�����) = i
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
    ReadFile_ALL_HV ����TXT�ļ�·��, allTxt
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

