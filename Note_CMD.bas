Attribute VB_Name = "Note_CMD"
Public Const ����̨���� = "Control Desk"
Public userDic As New Dictionary

Public Function CMD_In(cmd As String) As String
    Dim lineCMD() As String
    lineCMD = Split(cmd, vbCrLf)
    For i = 0 To UBound(lineCMD)
        If lineCMD(i) <> "" Then
            CMD_In = CMD_In & CMD_LineExecute(lineCMD(i)) & vbCrLf
        End If
    Next
End Function
Private Function CMD_LineExecute(cmd As String) As String
    Dim cT() As String
    cT = Split(cmd & " ", " ")
    Select Case UCase(cT(0))
        Case "����", "HELP"
            CMD_LineExecute = vbCrLf & "���������ڵ�[FORNODEADD] xStart(��ֵ) xStep(��ֵ) xCounts(��ֵ) yStart(��ֵ) yStep(��ֵ) yCounts(��ֵ) nodeTitle(�ַ���) nodeContent(�ַ���) pitchOn(0/1) size(��ֵ) color(��ֵ)" _
                            & vbCrLf & "��ʾ�������[VISMOUSEPOS] 1(��ʾ)/0(����ʾ)" _
                            & vbCrLf & "����ƫ����[SELFIM] iƫ�� xƫ�� yƫ��" _
                            & vbCrLf & "�ֵ�������[DICITEMADD] ��A:ֵA,��B:ֵB����" _
                            & vbCrLf & "�ֵ������[DICREMOVEALL]" _
                            & vbCrLf & "��ӡ�ֵ�[PRINTDIC]" _
                            & vbCrLf & "��ӡ�����б�[PRINTREVOKE]" _
                            & vbCrLf & "��ӡ�����б�[PRINTREDO]" _
                            & vbCrLf & "������״�ı�����λ�ÿ��Ƴ���[SETTREETXTINPOSCONTROLCONST/STTIPCC/STIPC] ���ڵ�X(��ֵ) ���ڵ�Y(��ֵ) �ڵ�X���(��ֵ) �ڵ�Y���(��ֵ)" _
                            & vbCrLf & "����λͼ����λ�ÿ��Ƴ���[SETIMAGEINPOSCONTROLCONST/SIIPCC/SIPC] ���ڵ�X(��ֵ) ���ڵ�Y(��ֵ) �ڵ�X���(��ֵ) �ڵ�Y���(��ֵ)" _
                            & vbCrLf & "������ɫ[RECTANGLECOLOR/RECCOLOR] VBColor(��ֵ)[RColor(��ֵ) GColor(��ֵ) BColor(��ֵ)]" _
                            & vbCrLf & "�ڵ����[NODEZOOM] ����ڵ���(�ַ���) X���������(��ֵ) Y���������(��ֵ)" _
                            & vbCrLf & "�����ڵ�[NEWBUILTNODE/NBN] Xλ��(��ֵ) Yλ��(��ֵ) ����(�ַ���) ����(�ַ���) VBColor(��ֵ) ��С(��ֵ) ѡ��(0/1)" _
                            & vbCrLf & "�༭�ڵ�[EDITNODE/EN] �ڵ����ID(��ֵ) ����(�ַ���) ����(�ַ���) VBColor(��ֵ) ��С(��ֵ)" _
                            & vbCrLf & "λ�ƽڵ�[MOVENODE/MN] �ڵ����ID(��ֵ) Xλ��(��ֵ) Yλ��(��ֵ)" _
                            & vbCrLf & "ɾ���ڵ�[DELETENODE/DN] �ڵ����ID1(��ֵ),�ڵ����ID2(��ֵ),�ڵ����ID3(��ֵ)..." _
                            & vbCrLf & "ѡ�нڵ�[SELECTNODE/SN] �ڵ����ID1(��ֵ),�ڵ����ID2(��ֵ),�ڵ����ID3(��ֵ)..." _
                            & vbCrLf & "��������[NEWBUILTNODE/NBL] ����Դ�ڵ����ID(��ֵ) ����ȥ�ڵ����ID(��ֵ) ��������(�ַ���) ���Ӵ�ϸ(��ֵ) ѡ��(0/1) *�����Ѵ��ڻᱻɾ��" _
                            & vbCrLf & "�༭��������[EDITLINE/EL] Դ�ڵ����ID(��ֵ) ȥ�ڵ����ID(��ֵ) ��������(�ַ���) ���Ӵ�ϸ(��ֵ)" _
                            & vbCrLf & "ѡ������[SELECTLINE/SL] ����1Դ�ڵ����ID(��ֵ):����1ȥ�ڵ����ID(��ֵ),����2Դ�ڵ����ID(��ֵ):����2ȥ�ڵ����ID(��ֵ),����3Դ�ڵ����ID(��ֵ):����3ȥ�ڵ����ID(��ֵ)..." _
                            & vbCrLf & "���ö��������ٶ�[SETACTIONUPDATASPEED/SAUS] ���¼��(��ֵ)" _
                            & vbCrLf & "��������ʱ��[STARTACTIONTIMER/SAT] 1(����)/0(�ر�)" _
                            & vbCrLf & "���嶯��[DEFINEACTION/DEFA/DA] ������(�ַ���),�����ڵ�ID1(��ֵ)[|�����ڵ�ID2(��ֵ)[|�����ڵ�ID3(��ֵ)[...]]],����ʱ��ִ�м��(��ֵ),��������(ֱ��/Բ��),ֱ��:����X(��ֵ),����Y(��ֵ)/[Բ��:�Ƕ�(��ֵ),���Ľڵ�ID(��ֵ)],��������(��ֵ),�Ƿ�ѭ��(0/1)" _
                            & vbCrLf & "��������[RESTARTACTION/RA] ������(�ַ���)"
            CMD_LineExecute = CMD_LineExecute _
                            & vbCrLf & "�رն���[OFFACTION/OA] ������(�ַ���)" _
                            & vbCrLf & "��ӡ�����б�[PRINTACTIONLIST/PAL]" _
                            & vbCrLf & "��ӡ��ִ�ж����б�[PRINTEXECUTABLEACTIONLIST/PEAL]" _
                            & vbCrLf & "���þ�ͷλ��[RESETLENSPOSITION/PLP] X(��ֵ),Y(��ֵ)" _
                            & vbCrLf & "ɫ����·���޸�[COLORLINKDICMOD/CLDM] �ֵ��ַ���(VBColor1:VBColor2,VBColor2:VBColor3...)" _
                            & vbCrLf & "ɫ����·������[COLORLINKDICRESET/CLDS]" _
                            & vbCrLf & "��ӡɫ����·��[PRINTCOLORLINKDIC/PCLD]"
            Exit Function
        Case "���������ڵ�", "FORNODEADD"
            ���������ڵ� Val(cT(1)), Val(cT(2)), Val(cT(3)), Val(cT(4)), Val(cT(5)), Val(cT(6)), cT(7), cT(8), cT(9), Val(cT(10)), Val(cT(11))
            GoTo Success
        Case "��ʾ�������", "VISMOUSEPOS"
            If cT(1) = "1" Or cT(1) = "TRUE" Then
                NoteControlDesk.CDMouseUpdataTimer.Enabled = True
            Else
                NoteControlDesk.CDMouseUpdataTimer.Enabled = False
                NoteControlDesk.Caption = ����̨����
            End If
            GoTo Success
        Case "vbCrLf", "SELFIM"
            oneselfAddI = Val(cT(1))
            oneselfAddX = Val(cT(2))
            oneselfAddY = Val(cT(3))
            GoTo Success
        Case "�ֵ�������", "DICITEMADD"
            �ֵ������� cT(1)
            CMD_LineExecute = vbCrLf & "��ǰ�ֵ��С��" & userDic.Count
            Exit Function
        Case "�ֵ������", "DICREMOVEALL"
            userDic.RemoveAll
            GoTo Success
        Case "��ӡ�ֵ�", "PRINTDIC"
            CMD_LineExecute = vbCrLf & �ֵ��ӡ(userDic)
            Exit Function
        Case "��ӡ�����б�", "PRINTREVOKE"
            CMD_LineExecute = vbCrLf & Join(behaviorList, vbCrLf)
            Exit Function
        Case "��ӡ�����б�", "PRINTREDO"
            CMD_LineExecute = vbCrLf & Join(redolist, vbCrLf)
            Exit Function
        Case "������״�ı�����λ�ÿ��Ƴ���", "SETTREETXTINPOSCONTROLCONST", "STTIPCC", "STIPC"
            treeTxtToNtx_StartX = Val(cT(1))
            treeTxtToNtx_StartY = Val(cT(2))
            treeTxtToNtx_StepX = Val(cT(3))
            treeTxtToNtx_StepY = Val(cT(4))
            GoTo Success
        Case "����λͼ����λ�ÿ��Ƴ���", "SETIMAGEINPOSCONTROLCONST", "SIIPCC", "SIPC"
            imageToNtx_StartX = Val(cT(1))
            imageToNtx_StartY = Val(cT(2))
            imageToNtx_StepX = Val(cT(3))
            imageToNtx_StepY = Val(cT(4))
            GoTo Success
        Case "������ɫ", "RECTANGLECOLOR", "RECCOLOR"
            If UBound(cT) > 2 Then
                rectangleLineColor = RGB(Val(cT(1)), Val(cT(2)), Val(cT(3)))
            Else
                rectangleLineColor = Val(cT(1))
            End If
            GoTo Success
        Case "�ڵ����", "NODEZOOM"
            �ڵ���� cT(1), Val(cT(2)), Val(cT(3))
            GoTo Success
        Case "�����ڵ�", "NEWBUILTNODE", "NBN"
            �ڵ㴴�� Val(cT(1)), Val(cT(2)), cT(3), cT(4), Val(cT(5)), Val(cT(6)), cT(7)
            GoTo Success
        Case "�༭�ڵ�", "EDITNODE", "EN"
            �༭�ڵ� Val(cT(1)), cT(2), cT(3), Val(cT(4)), Val(cT(5))
            GoTo Success
        Case "λ�ƽڵ�", "MOVENODE", "MN"
            λ�ƽڵ� Val(cT(1)), Val(cT(2)), Val(cT(3))
            GoTo Success
        Case "ɾ���ڵ�", "DELETENODE", "DN"
            ɾ���ڵ� cT(1)
            GoTo Success
        Case "ѡ�нڵ�", "SELECTNODE", "SN"
            ѡ�нڵ� cT(1)
            GoTo Success
        Case "��������", "NEWBUILTNODE", "NBL"
            ���Ӵ��� Val(cT(1)), Val(cT(2)), cT(3), Val(cT(4)), cT(5)
            GoTo Success
        Case "�༭��������", "EDITLINE", "EL"
            �༭���� Val(cT(1)), Val(cT(2)), cT(3), Val(cT(4))
            GoTo Success
        Case "ѡ������", "SELECTLINE", "SL"
            ѡ������ cT(1)
            GoTo Success
        Case "���ö��������ٶ�", "SETACTIONUPDATASPEED", "SAUS"
            ���ö��������ٶ� Val(cT(1))
            GoTo Success
        Case "��������ʱ��", "STARTACTIONTIMER", "SAT"
            ��������ʱ�� cT(1)
            GoTo Success
        Case "���嶯��", "DEFINEACTION", "DEFA", "DA"
            ���嶯�� cT(1)
            GoTo Success
        Case "��������", "RESTARTACTION", "RA"
            CMD_LineExecute = ��������(cT(1))
            Exit Function
        Case "�رն���", "OFFACTION", "OA"
            CMD_LineExecute = �رն���(cT(1))
            Exit Function
        Case "��ӡ�����б�", "PRINTACTIONLIST", "PAL"
            CMD_LineExecute = ��ӡ�����б�
            Exit Function
        Case "��ӡ��ִ�ж����б�", "PRINTEXECUTABLEACTIONLIST", "PEAL"
            CMD_LineExecute = ��ӡ��ִ�ж����б�
            Exit Function
        Case "���þ�ͷλ��", "RESETLENSPOSITION", "PLP"
            angleOfView.X = Val(cT(1))
            angleOfView.Y = Val(cT(2))
            MainCoordinateSystemDefinition
            GoTo Success
        Case "ɫ����·���޸�", "COLORLINKDICMOD", "CLDM"
            ɫ����·�ֵ��޸� cT(1)
            GoTo Success
        Case "ɫ����·������", "COLORLINKDICRESET", "CLDS"
            ɫ����·��ʼ��
            GoTo Success
        Case "��ӡɫ����·��", "PRINTCOLORLINKDIC", "PCLD"
            CMD_LineExecute = ɫ����·�ֵ䵼��
            Exit Function
    End Select
    CMD_LineExecute = "δ֪���"
Exit Function
Success:
    CMD_LineExecute = "����ִ�гɹ���"
End Function
Public Sub ɫ����·�ֵ��޸�(s As String)
    Dim i As Long, sT() As String, sT2() As String
    colorLinkDic.RemoveAll
    If InStr(1, s, ",") > 0 Then
        sT = Split(s, ",")
        For i = 0 To UBound(sT)
            If InStr(1, sT(i), ":") > 0 Then
                sT2 = Split(sT(i), ":")
                colorLinkDic.Add Val(sT2(0)), Val(sT2(1))
            End If
        Next
    End If
End Sub
Public Function ɫ����·��ʼ��()
    colorLinkDic.RemoveAll
    colorLinkDic.Add &HFFBF00, &HC000C0
    colorLinkDic.Add &HC000C0, &HFF&
    colorLinkDic.Add &HFF&, &H80FF&
    colorLinkDic.Add &H80FF&, &HFFFF&
    colorLinkDic.Add &HFFFF&, &HFF00&
    colorLinkDic.Add &HFF00&, &HFFFF00
    colorLinkDic.Add &HFFFF00, &HFFBF00
End Function
Public Function ɫ����·�ֵ䵼��() As String
    Dim i As Long
    For i = 0 To colorLinkDic.Count - 1
        ɫ����·�ֵ䵼�� = ɫ����·�ֵ䵼�� & colorLinkDic.Keys(i) & ":" & colorLinkDic.Items(i) & ","
    Next
End Function
Private Function ��ӡ��ִ�ж����б�() As String
    Dim i As Long, j As Long
    ��ӡ��ִ�ж����б� = "���嶯�� "
    For i = 1 To UBound(actionList)
        With actionList(i)
            ��ӡ��ִ�ж����б� = ��ӡ��ִ�ж����б� & .name & ","
            For j = 0 To UBound(.nodeID)
                ��ӡ��ִ�ж����б� = ��ӡ��ִ�ж����б� & .nodeID(j) & "|"
            Next
            ��ӡ��ִ�ж����б� = Mid(��ӡ��ִ�ж����б�, 1, Len(��ӡ��ִ�ж����б�) - 1) & "," & .interval & "," & .route
            Select Case .route
                Case "ֱ��"
                    ��ӡ��ִ�ж����б� = ��ӡ��ִ�ж����б� & "," & .vector.X & "," & .vector.Y
                Case "Բ��"
                    ��ӡ��ִ�ж����б� = ��ӡ��ִ�ж����б� & "," & .angle & "," & .aimNode
            End Select
            ��ӡ��ִ�ж����б� = ��ӡ��ִ�ж����б� & "," & .endTime & "," & .repeat & vbCrLf
        End With
    Next
End Function
Private Function ��ӡ�����б�() As String
    Dim i As Long, j As Long
    For i = 1 To UBound(actionList)
        With actionList(i)
            ��ӡ�����б� = ��ӡ�����б� & "������(" & .name & "),"
            For j = 0 To UBound(.nodeID)
                ��ӡ�����б� = ��ӡ�����б� & "�ڵ�(" & .nodeID(j) & ")|"
            Next
            ��ӡ�����б� = Mid(��ӡ�����б�, 1, Len(��ӡ�����б�) - 1) & ",����ʱ��ִ�м��(" & .interval & "),��������(" & .route & ")"
            Select Case .route
                Case "ֱ��"
                    ��ӡ�����б� = ��ӡ�����б� & ",����X(" & .vector.X & "),����Y(" & .vector.Y & ")"
                Case "Բ��"
                    ��ӡ�����б� = ��ӡ�����б� & ",�Ƕ�(" & .angle & "),���Ľڵ�ID(" & .aimNode & ")"
            End Select
            ��ӡ�����б� = ��ӡ�����б� & ",��������(" & .endTime & "),�Ƿ�ѭ��(" & .repeat & ")" & vbCrLf
        End With
    Next
End Function
Private Function ��������ʱ��(c As String)
    Note.ActionTimer.Enabled = �ַ���ת����ֵ(c)
End Function
Private Function ���ö��������ٶ�(interval As Long)
    Note.ActionTimer.interval = interval
End Function
Private Function �༭����(nS As Long, nT As Long, content As String, size As Single)
    Dim i As Long
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                If .Source = nS And .target = nT Then
                    .content = �ո�ת��(content)
                    .size = size
                    Exit Function
                End If
            End If
        End With
    Next
End Function
Private Function ѡ������(allNid As String)
    Dim allNidTemp() As String, i As Long, temp() As String
    allNidTemp = Split(allNid, ",")
    For i = 0 To UBound(allNidTemp)
        If allNidTemp(i) <> "" Then
            temp = Split(allNidTemp(i), ":")
            nodeLine(LineAdd_RepeatedChecking(Val(temp(0)), Val(temp(1)))).select = True
        End If
    Next
End Function
Private Function ���Ӵ���(nidA As Long, nidB As Long, content As String, size As Single, pichOn As String)
    LineAdd nidA, nidB, �ո�ת��(content), size, �ַ���ת����ֵ(pichOn), False
End Function
Private Function ѡ�нڵ�(allNid As String)
    Dim allNidTemp() As String, i As Long
    allNidTemp = Split(allNid, ",")
    For i = 0 To UBound(allNidTemp)
        If allNidTemp(i) <> "" Then
            node(Val(allNidTemp(i))).select = True
        End If
    Next
End Function
Private Function ɾ���ڵ�(nid As String)
    Dim allNidTemp() As String, i As Long
    allNidTemp = Split(nid, ",")
    For i = 0 To UBound(allNidTemp)
        If allNidTemp(i) <> "" Then
            NodeDelete Val(allNidTemp(i))
        End If
    Next
End Function
Private Function λ�ƽڵ�(nid As Long, X As Single, Y As Single)
    With node(nid)
        .X = X
        .Y = Y
    End With
End Function
Private Function �༭�ڵ�(nid As Long, t As String, content As String, color As Long, size As Single)
    NodeEdit_ReviseNode nid, �ո�ת��(t), �ո�ת��(content), color, size
End Function
Public Function �ַ���ת����ֵ(s As String) As Boolean
    If s = "1" Then
        �ַ���ת����ֵ = True
    End If
End Function
Private Function �ո�ת��(s As String) As String
    �ո�ת�� = Replace(s, "\_", " ")
End Function
Private Function �ڵ㴴��(X As Single, Y As Single, t As String, content As String, color As Long, size As Single, pichOn As String)
    NodeEdit_NewNode �ո�ת��(t), �ո�ת��(content), color, size, X, Y, �ַ���ת����ֵ(pichOn)
End Function
Private Function �ڵ����������(nodeName As String) As Long
    Dim i As Long
    �ڵ���������� = -1
    For i = 0 To nSum
        With node(i)
            If .b Then
                If nodeName = .t And .select = True Then
                    �ڵ���������� = i
                    Exit Function
                End If
            End If
        End With
    Next
End Function
Private Sub �ڵ����(centre As String, zoomX As Single, zoomY As Single)
    Dim i As Long, iLock As Boolean, iX As Single, iY As Single, rootNid As Long
    rootNid = �ڵ����������(centre)
    If rootNid <> -1 Then
        With node(rootNid)
                iX = .X * zoomX - .X
                iY = .Y * zoomY - .Y
        End With
        For i = 0 To nSum
            With node(i)
                If .b Then
                    If .select Then
                        .X = .X * zoomX - iX
                        .Y = .Y * zoomY - iY
                    End If
                End If
            End With
        Next
    End If
End Sub
Private Function �ֵ��ӡ(dic As Dictionary) As String
    Dim i As Long
    For i = 0 To dic.Count - 1
        �ֵ��ӡ = �ֵ��ӡ & dic.Keys(i) & ":" & dic.Items(i) & vbCrLf
    Next
End Function

Private Sub �ֵ�������(dS As String)
    Dim dSTmp() As String, dTmp() As String
    dSTmp = Split(dS, ",")
    For i = 0 To UBound(dSTmp)
        If dSTmp(i) <> "" Then
            dTmp = Split(dSTmp(i), ":")
            If Not userDic.Exists(dTmp(0)) Then userDic.Add dTmp(0), dTmp(1)
        End If
    Next
End Sub

Private Sub ���������ڵ�(xS As Single, xStep As Single, xE As Single, yS As Single, yStep As Single, yE As Single, t As String, c As String, p As String, size As Single, color As Long)
    Dim X As Long, Y As Long, m As Single, n As Single, pL As Boolean, dN As Long
    If p = "1" Or UCase(p) = "TRUE" Then
        pL = True
    End If
    For X = 1 To xE
        For Y = 1 To yE
            dN = dN + 1
            NodeEdit_NewNode �ֵ��滻(t, dN, X, Y), �ֵ��滻(c, dN, X, Y), color, size, xStep * (X - 1) + xS, yStep * (Y - 1) + yS, pL
        Next
    Next
End Sub

Private Function �ֵ��滻(ByVal s As String, dN As Long, dX As Long, dY As Long) As String
    Dim dNT As String
    dNT = dN + oneselfAddI
    If userDic.Exists(dNT) Then
        s = Replace(s, "[i]", userDic(dNT))
    Else
        s = Replace(s, "[i]", dNT)
    End If
    dNT = oneselfAddX + dX
    If userDic.Exists(dNT) Then
        s = Replace(s, "[x]", userDic(dNT))
    Else
        s = Replace(s, "[x]", dNT)
    End If
    dNT = oneselfAddY + dY
    If userDic.Exists(dNT) Then
        s = Replace(s, "[y]", userDic(dNT))
    Else
        s = Replace(s, "[y]", dNT)
    End If
    �ֵ��滻 = s
End Function
