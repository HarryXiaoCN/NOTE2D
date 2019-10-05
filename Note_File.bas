Attribute VB_Name = "Note_File"
Public Function LoadProfile()
Dim lineTmp As String
On Error GoTo Er:
MainFormFontSize = note.Font.size
NodeInputBackColor = &H8000000F
If Dir(ProfilePath & PROFILENAME) <> "" Then
    Open ProfilePath & PROFILENAME For Input As #1
        Do While Not EOF(1)
            Line Input #1, lineTmp
            LoadProfile_ReadLine lineTmp
        Loop
    Close #1
End If
Er:
End Function
Public Function LoadProfile_InitializationBool()
note.��ʾȫ���ڵ���.Checked = False
note.��ʾȫ������.Checked = False
End Function
Public Function LoadProfile_ReadLine(ByRef lineTmp As String)
Dim ESRStr As String: Dim ESRBool As Boolean
ESRStr = LoadProfile_ReadLine_GetEqualSignRight(lineTmp)
ESRBool = StrToBool(ESRStr)
If InStr(1, lineTmp, "��ʾȫ���ڵ���=") Then
    note.��ʾȫ���ڵ���.Checked = ESRBool
ElseIf InStr(1, lineTmp, "��ʾ˳��ڵ���=") Then
    note.��ʾ˳��ڵ���.Checked = ESRBool
ElseIf InStr(1, lineTmp, "��ʾ����ڵ���=") Then
    note.��ʾ����ڵ���.Checked = ESRBool
ElseIf InStr(1, lineTmp, "ʼ����ʾѡ����=") Then
    note.ʼ����ʾѡ����.Checked = ESRBool
ElseIf InStr(1, lineTmp, "��ʾ�ڵ����ID=") Then
    note.��ʾ�ڵ����ID.Checked = ESRBool
ElseIf InStr(1, lineTmp, "��ʾȫ������=") Then
    note.��ʾȫ������.Checked = ESRBool
ElseIf InStr(1, lineTmp, "��ʾ˳������=") Then
    note.��ʾ˳������.Checked = ESRBool
ElseIf InStr(1, lineTmp, "��ʾ��������=") Then
    note.��ʾ��������.Checked = ESRBool
ElseIf InStr(1, lineTmp, "ʼ����ʾѡ��=") Then
    note.ʼ����ʾѡ��.Checked = ESRBool
ElseIf InStr(1, lineTmp, "ȫ����ͼ=") Then
    NoteGlobalViewSet ESRBool
ElseIf InStr(1, lineTmp, "����=") Then
    note.Font.name = ESRStr
ElseIf InStr(1, lineTmp, "�ֺ�=") Then
    MainFormFontSize = Val(ESRStr): note.Font.size = MainFormFontSize
ElseIf InStr(1, lineTmp, "�Ӵ�=") Then
    note.Font.Bold = ESRBool
ElseIf InStr(1, lineTmp, "��б=") Then
    note.Font.Italic = ESRBool
ElseIf InStr(1, lineTmp, "������ȫ��͸��=") Then
    note.ȫ��͸��.Checked = ESRBool: FormTransparent note, 50
ElseIf InStr(1, lineTmp, "������ȫ��͸��=") Then
    note.ȫ��͸��.Checked = ESRBool: FormTransparent note, 125
ElseIf InStr(1, lineTmp, "������ȫ��͸��=") Then
    note.ȫ��͸��.Checked = ESRBool: FormTransparent note, 200
ElseIf InStr(1, lineTmp, "�����汳��ɫ=") Then
    note.BackColor = Val(ESRStr)
ElseIf InStr(1, lineTmp, "������������ɫ=") Then
    note.ForeColor = Val(ESRStr)
ElseIf InStr(1, lineTmp, "�ص�=") Then
    note.�ص�.Checked = ESRBool
ElseIf InStr(1, lineTmp, "����=") Then
    note.����.Checked = ESRBool
ElseIf InStr(1, lineTmp, "����ͼ·��=") Then
    If ESRStr <> "" Then
        note.homeBackPicPath = ESRStr
        note.���ر���ͼ note.homeBackPicPath
    End If
ElseIf InStr(1, lineTmp, "�ʺ�Ȧ=") Then
    note.�ʺ�Ȧ.Checked = ESRBool
ElseIf InStr(1, lineTmp, "�ʺ���=") Then
    note.�ʺ���.Checked = ESRBool
ElseIf InStr(1, lineTmp, "�������=") Then
    note.�������.Checked = ESRBool
ElseIf InStr(1, lineTmp, "�������ȫ��͸��=") Then
    note.ȫ��͸��2.Checked = ESRBool: FormTransparent NodeInput, 50
ElseIf InStr(1, lineTmp, "�������ȫ��͸��=") Then
    note.ȫ��͸��2.Checked = ESRBool: FormTransparent NodeInput, 125
ElseIf InStr(1, lineTmp, "�������ȫ��͸��=") Then
    note.ȫ��͸��2.Checked = ESRBool: FormTransparent NodeInput, 200
ElseIf InStr(1, lineTmp, "������汳��ɫ=") Then
    NodeInputBackColor = Val(ESRStr)
ElseIf InStr(1, lineTmp, "�������ȫ��͸��=") Then
    note.ȫ��͸��3.Checked = ESRBool: FormTransparent NodePrint, 50
ElseIf InStr(1, lineTmp, "�������ȫ��͸��=") Then
    note.ȫ��͸��3.Checked = ESRBool: FormTransparent NodePrint, 125
ElseIf InStr(1, lineTmp, "�������ȫ��͸��=") Then
    note.ȫ��͸��3.Checked = ESRBool: FormTransparent NodePrint, 200
ElseIf InStr(1, lineTmp, "��������ö�=") Then
    note.�ö�.Checked = ESRBool: FormStick NodePrint, ESRBool
ElseIf InStr(1, lineTmp, "��ǩ��=") Then
    note.��ǩ��.Checked = ESRBool
ElseIf InStr(1, lineTmp, "��������=") Then
    nodeInputFormHeight = Val(ESRStr)
ElseIf InStr(1, lineTmp, "��������=") Then
    nodeInputFormWidth = Val(ESRStr)
ElseIf InStr(1, lineTmp, "�Զ�����ʱ����=") Then
    saveNtxTime = Val(ESRStr)
ElseIf InStr(1, lineTmp, "�ڵ�Ĭ�ϴ�С=") Then
    nodeDefaultSize = Val(ESRStr)
    nodeDefaultSize = ������ֵ(nodeDefaultSize, 50, 500)
ElseIf InStr(1, lineTmp, "����Ĭ�Ͽ��=") Then
    lineDefaultSize = Val(ESRStr)
    lineDefaultSize = ������ֵ(lineDefaultSize, 1, 10)
ElseIf InStr(1, lineTmp, "��ͼ���=") Then
    note.updataSpeed = Val(ESRStr)
    If note.updataSpeed < 10 Then note.updataSpeed = 10
    If note.updataSpeed > 100 Then note.updataSpeed = 100
    note.MainTime.interval = note.updataSpeed
ElseIf InStr(1, lineTmp, "������ɫ=") Then
    rectangleLineColor = Val(ESRStr)
ElseIf InStr(1, lineTmp, "���߲���=") Then
    rectangleStep = Val(ESRStr)
ElseIf InStr(1, lineTmp, "�ڵ��������=") Then
    nodeAttributedToIntegers = Val(ESRStr)
ElseIf InStr(1, lineTmp, "�ڵ����=") Then
    note.�ڵ����.Checked = ESRBool
End If
End Function
Public Function LoadProfile_ReadLine_GetEqualSignRight(ByRef str As String)
    Dim strTmps
    On Error GoTo Er
    strTmp = Split(str, "=")
    LoadProfile_ReadLine_GetEqualSignRight = strTmp(1)
Er:
End Function
Public Function SaveProfile()
Dim profileStr As String: Dim i As Long
On Error GoTo Er
profileStr = "-��ͼ-" & vbCrLf _
& "[�ڵ�����ʾ]" & vbCrLf & SaveProfile_GetTrueNodeViewName _
& vbCrLf & "ʼ����ʾѡ����=" & note.ʼ����ʾѡ����.Checked _
& vbCrLf & "��ʾ�ڵ����ID=" & note.��ʾ�ڵ����ID.Checked _
& vbCrLf & "[�ڵ�������ʾ]" & vbCrLf & SaveProfile_GetTrueLineViewName _
& vbCrLf & "ʼ����ʾѡ��=" & note.ʼ����ʾѡ��.Checked _
& vbCrLf & "ȫ����ͼ=" & note.ȫ����ͼ.Checked _
& vbCrLf & "-����-" & vbCrLf & "[������]" _
& vbCrLf & "����=" & note.Font.name _
& vbCrLf & "�ֺ�=" & MainFormFontSize _
& vbCrLf & "�Ӵ�=" & note.Font.Bold _
& vbCrLf & "��б=" & note.Font.Italic _
& vbCrLf & SaveProfile_GetTrueTransparentViewName _
& vbCrLf & "�����汳��ɫ=" & note.BackColor & vbCrLf & "����ͼ·��=" & note.homeBackPicPath _
& vbCrLf & "������������ɫ=" & note.ForeColor _
& vbCrLf & "�ص�=" & note.�ص�.Checked & vbCrLf & "����=" & note.����.Checked _
& vbCrLf & "�ʺ�Ȧ=" & note.�ʺ�Ȧ.Checked _
& vbCrLf & "�ʺ���=" & note.�ʺ���.Checked _
& vbCrLf & "�������=" & note.�������.Checked _
& vbCrLf & "[�������]" _
& vbCrLf & SaveProfile_GetTrueTransparent2ViewName _
& vbCrLf & "������汳��ɫ=" & NodeInputBackColor & vbCrLf & "��ǩ��=" & note.��ǩ��.Checked _
& vbCrLf & "[�������]" & vbCrLf & SaveProfile_GetTrueTransparent3ViewName _
& vbCrLf & "��������ö�=" & note.�ö�.Checked & vbCrLf & "-ϵͳ-" & vbCrLf & "��������=" & nodeInputFormHeight & vbCrLf & "��������=" & nodeInputFormWidth _
& vbCrLf & "�Զ�����ʱ����=" & saveNtxTime & vbCrLf & "��ͼ���=" & note.updataSpeed & vbCrLf & "�ڵ�Ĭ�ϴ�С=" & nodeDefaultSize & vbCrLf & "����Ĭ�Ͽ��=" & lineDefaultSize & vbCrLf
profileStr = profileStr & "-ϵͳ-" & vbCrLf _
& vbCrLf & "������ɫ=" & rectangleLineColor _
& vbCrLf & "���߲���=" & rectangleStep _
& vbCrLf & "�ڵ��������=" & nodeAttributedToIntegers _
& vbCrLf & "�ڵ����=" & note.�ڵ����.Checked
'& vbCrLf & "-ϵͳ-" & vbCrLf & "������=" & magnification
If Dir(ProfilePath, vbDirectory) = "" Then
    Shell "cmd /c md """ & ProfilePath & """", vbHide
    Do While Dir(ProfilePath, vbDirectory) = ""
        Sleep 30
        DoEvents
    Loop
Else
    If Dir(ProfilePath & PROFILENAME) <> "" Then
        SetAttr ProfilePath & PROFILENAME, vbNormal
    End If
End If

Open ProfilePath & PROFILENAME For Output As #3
    Print #3, profileStr
Close #3
SetAttr ProfilePath & PROFILENAME, vbReadOnly
Exit Function
Er:
MsgBox "���������ļ�����ʧ�ܣ�����·��" & ProfilePath & "�Ƿ���ڡ�", , "���棡"
End Function
Public Function SaveProfile_GetTrueTransparent3ViewName() As String
If note.ȫ��͸��3.Checked = True Then SaveProfile_GetTrueTransparent3ViewName = "�������ȫ��͸��=True": Exit Function
If note.ȫ��͸��3.Checked = True Then SaveProfile_GetTrueTransparent3ViewName = "�������ȫ��͸��=True": Exit Function
If note.ȫ��͸��3.Checked = True Then SaveProfile_GetTrueTransparent3ViewName = "�������ȫ��͸��=True": Exit Function
End Function
Public Function SaveProfile_GetTrueTransparent2ViewName() As String
If note.ȫ��͸��2.Checked = True Then SaveProfile_GetTrueTransparent2ViewName = "�������ȫ��͸��=True": Exit Function
If note.ȫ��͸��2.Checked = True Then SaveProfile_GetTrueTransparent2ViewName = "�������ȫ��͸��=True": Exit Function
If note.ȫ��͸��2.Checked = True Then SaveProfile_GetTrueTransparent2ViewName = "�������ȫ��͸��=True": Exit Function
End Function
Public Function SaveProfile_GetTrueTransparentViewName() As String
If note.ȫ��͸��.Checked = True Then SaveProfile_GetTrueTransparentViewName = "������ȫ��͸��=True": Exit Function
If note.ȫ��͸��.Checked = True Then SaveProfile_GetTrueTransparentViewName = "������ȫ��͸��=True": Exit Function
If note.ȫ��͸��.Checked = True Then SaveProfile_GetTrueTransparentViewName = "������ȫ��͸��=True": Exit Function
End Function
Public Function SaveProfile_GetTrueNodeViewName() As String
If note.��ʾȫ���ڵ���.Checked = True Then SaveProfile_GetTrueNodeViewName = "��ʾȫ���ڵ���=True": Exit Function
If note.��ʾ˳��ڵ���.Checked = True Then SaveProfile_GetTrueNodeViewName = "��ʾ˳��ڵ���=True": Exit Function
If note.��ʾ����ڵ���.Checked = True Then SaveProfile_GetTrueNodeViewName = "��ʾ����ڵ���=True": Exit Function
End Function
Public Function SaveProfile_GetTrueLineViewName() As String
If note.��ʾȫ������.Checked = True Then SaveProfile_GetTrueLineViewName = "��ʾȫ������=True": Exit Function
If note.��ʾ˳������.Checked = True Then SaveProfile_GetTrueLineViewName = "��ʾ˳������=True": Exit Function
If note.��ʾ��������.Checked = True Then SaveProfile_GetTrueLineViewName = "��ʾ��������=True": Exit Function
End Function
Public Function NoteFileRead(ByRef filePath As String)
Dim ntx() As String: Dim i As Long, version As Long
newAddNote
On Error GoTo Er
ntxPath = filePath
note.Caption = NOTEFORMNAME & ntxPath
Open filePath For Input As #1
        Do While Not EOF(1)
            ReDim Preserve ntx(i)
            Line Input #1, ntx(i)
            If ntx(i) = "" Then Exit Do
            i = i + 1
        Loop
Close #1
version = NoteFileRead_VersionCheck(ntx(0))
Select Case version
    Case -1
        MsgBox "�ļ��޷�ʶ��", , "���棡"
        newAddNote
    Case 201
        NoteFileRead_201 ntx, False
    Case 202
        NoteFileRead_202 ntx, False
        lastNtx = ntx
    Case 301
        NoteFileRead_301 ntx
End Select
note.MainTime.Enabled = True
Exit Function
Er:
    MsgBox "�ļ���ȡʧ�ܣ�ԭ��" & Err.Description, 16, "����"
End Function

Public Function NoteFileRead_201(ByRef ntx() As String, ByRef fromCopy As Boolean)
Dim i, nodeSum, lineSum, startNodeId As Long: Dim lineTmp
lineTmp = Split(ntx(0), "^|`")
nodeSum = Val(lineTmp(1))
lineSum = Val(lineTmp(2))
If UBound(lineTmp) > 2 Then
    magnification = Val(lineTmp(3)): zoomFactor = MToZF(magnification)
    MainCoordinateSystemDefinition
End If
startNodeId = nSum
For i = 1 To nodeSum
    lineTmp = Split(ntx(i), "^|`")
    If fromCopy = True Then
        NodeEdit_NewNode lineTmp(2), Replace(lineTmp(3), "^||`", vbCrLf), &HFFBF00, nodeDefaultSize, Val(lineTmp(0)) + mousePos.X, Val(lineTmp(1)) + mousePos.Y, True
    Else
        NodeEdit_NewNode lineTmp(2), Replace(lineTmp(3), "^||`", vbCrLf), &HFFBF00, nodeDefaultSize, Val(lineTmp(0)), Val(lineTmp(1))
    End If
Next
For i = nodeSum + 1 To nodeSum + lineSum
    lineTmp = Split(ntx(i), "^|`")
    If fromCopy = True Then
        LineAdd_Save Val(lineTmp(0)) + startNodeId, Val(lineTmp(1)) + startNodeId, "", lineDefaultSize, True
    Else
        LineAdd_Save Val(lineTmp(0)), Val(lineTmp(1)), "", lineDefaultSize
    End If
Next
End Function
Public Function NoteFileRead_202(ByRef ntx() As String, ByRef fromCopy As Boolean)
    Dim i As Long, nodeSum As Long, lineSum As Long, startNodeId As Long: Dim lineTmp() As String
    lineTmp = Split(ntx(0), LINEBREAK)
    nodeSum = Val(lineTmp(1))
    lineSum = Val(lineTmp(2))
    If UBound(lineTmp) > 2 Then
        magnification = Val(lineTmp(3)): zoomFactor = MToZF(magnification)
        MainCoordinateSystemDefinition
        If UBound(lineTmp) > 3 And lineTmp(4) <> "" Then
            NoteFileRead_202_GetDic lineTmp(4)
        End If
    End If
    startNodeId = nSum
    For i = 1 To nodeSum
        lineTmp = Split(ntx(i), LINEBREAK)
        If fromCopy = True Then
            NodeEdit_NewNode lineTmp(2), Replace(lineTmp(3), NODELINEBREAK, vbCrLf), Val(lineTmp(4)), Val(lineTmp(5)), Val(lineTmp(0)) + mousePos.X, Val(lineTmp(1)) + mousePos.Y, True
        Else
            NodeEdit_NewNode lineTmp(2), Replace(lineTmp(3), NODELINEBREAK, vbCrLf), Val(lineTmp(4)), Val(lineTmp(5)), Val(lineTmp(0)), Val(lineTmp(1))
        End If
    Next
    For i = nodeSum + 1 To nodeSum + lineSum
        lineTmp = Split(ntx(i), LINEBREAK)
        If fromCopy = True Then
            LineAdd_Save Val(lineTmp(0)) + startNodeId, Val(lineTmp(1)) + startNodeId, lineTmp(2), Val(lineTmp(3)), True
        Else
            If UBound(lineTmp) < 3 Then
                LineAdd_Save Val(lineTmp(0)), Val(lineTmp(1)), "", lineDefaultSize
            Else
                LineAdd_Save Val(lineTmp(0)), Val(lineTmp(1)), lineTmp(2), Val(lineTmp(3))
            End If
        End If
    Next
End Function
Public Function NoteFileRead_202_GetDic(lT As String)
    Dim dic() As String, key() As String, value() As String
    dic = Split(lT, DICBREAK)
    key = Split(dic(0), KEYBREAK)
    For i = 0 To UBound(key)
        If key(i) <> "" Then
            value = Split(key(i), VALUEBREAK)
            nodeSelectKeyDic.Add value(0), value(1)
        End If
    Next
    key = Split(dic(1), KEYBREAK)
    For i = 0 To UBound(key)
        If key(i) <> "" Then
            value = Split(key(i), VALUEBREAK)
            lineSelectKeyDic.Add value(0), value(1)
        End If
    Next
End Function
Public Function noteSaveCheck() As Long
If Dir(ntxPath) = "" Then
    noteSaveCheck = 0
Else
    If noteSaveCheck_ContentCheck(ntxPath) = True Then
        noteSaveCheck = 1
    Else
        noteSaveCheck = 2
    End If
End If
End Function
Public Function noteSaveCheck_ContentCheck(filePath As String) As Boolean
    Dim ntx() As String
    Dim i As Long
    ntx = NoteFileWrite_202_Coding(node, nSum, nodeLine, lSum)
    On Error GoTo Er
        If UBound(ntx) = UBound(lastNtx) Then
            For i = 0 To UBound(ntx)
                If ntx(i) <> lastNtx(i) Then
                    noteSaveCheck_ContentCheck = True
                    Exit Function
                End If
            Next
        Else
            noteSaveCheck_ContentCheck = True
        End If
    Exit Function
Er:
    noteSaveCheck_ContentCheck = True
End Function
Public Function NoteFileWrite_202_Coding(ByRef nodeObj() As �ڵ�, ByRef nodeObjSum As Long, _
ByRef lineObj() As ����, ByRef lineObjSum As Long)
Dim ntx() As String: Dim i As Long, j As Long, nodeSum As Long, lineSum As Long, nodeIdList() As Long
ReDim nodeIdList(nodeObjSum)
ReDim ntx(nodeObjSum + lineObjSum + 1)
j = 1
For i = 0 To nodeObjSum - 1
    With nodeObj(i)
        If .b = True Then
            ntx(j) = .X & LINEBREAK & .Y & LINEBREAK & .t & LINEBREAK & Replace(.content, vbCrLf, NODELINEBREAK) & LINEBREAK & .setColor & LINEBREAK & .setSize
            nodeIdList(i) = j - 1
            j = j + 1
        End If
    End With
Next
nodeSum = j - 1
For i = 0 To lineObjSum - 1
    With lineObj(i)
        If .b = True Then
            ntx(j) = nodeIdList(.Source) & LINEBREAK & nodeIdList(.target) & LINEBREAK & .content & LINEBREAK & .size
            j = j + 1
        End If
    End With
Next
lineSum = j - nodeSum - 1
ntx(0) = VERSIONID & LINEBREAK & nodeSum & LINEBREAK & lineSum & LINEBREAK & magnification & LINEBREAK & NoteFileWrite_202_Coding_DIC
NoteFileWrite_202_Coding = ntx
End Function
Public Function NoteFileWrite_202_Coding_DIC() As String
    Dim i As Long
    
    For i = 0 To nodeSelectKeyDic.Count - 1
        NoteFileWrite_202_Coding_DIC = NoteFileWrite_202_Coding_DIC & nodeSelectKeyDic.Keys(i) & VALUEBREAK & nodeSelectKeyDic.Items(i) & KEYBREAK
    Next
        
    If Len(NoteFileWrite_202_Coding_DIC) > 0 Then
        NoteFileWrite_202_Coding_DIC = Mid(NoteFileWrite_202_Coding_DIC, 1, Len(NoteFileWrite_202_Coding_DIC) - 1)
    End If
    NoteFileWrite_202_Coding_DIC = NoteFileWrite_202_Coding_DIC & DICBREAK
    
    For i = 0 To lineSelectKeyDic.Count - 1
        NoteFileWrite_202_Coding_DIC = NoteFileWrite_202_Coding_DIC & lineSelectKeyDic.Keys(i) & VALUEBREAK & lineSelectKeyDic.Items(i) & KEYBREAK
    Next
    If Len(NoteFileWrite_202_Coding_DIC) > 0 Then
        NoteFileWrite_202_Coding_DIC = Mid(NoteFileWrite_202_Coding_DIC, 1, Len(NoteFileWrite_202_Coding_DIC) - 1)
    End If
End Function
Public Function NoteFileWrite_202(ByRef filePath As String)
    Dim ntx() As String: Dim i As Long, fN As Integer
    ntxPath = filePath
    note.Caption = NOTEFORMNAME & ntxPath
    ntx = NoteFileWrite_202_Coding(node, nSum, nodeLine, lSum)
    lastNtx = ntx
    fN = FreeFile
    Open filePath For Output As #fN
        For i = 0 To UBound(ntx)
            Print #fN, ntx(i)
        Next
    Close #fN
End Function
Public Function NoteFileRead_301(ByRef ntx() As String)
Dim lineStr, lineTmp
Dim i, nodeSum As Long
Dim v3 As ��ά����
Dim v2 As ��ά����
lineStr = Split(ntx(0), LINEBREAK)
nodeSum = Val(lineStr(1))
For i = 1 To nodeSum
    lineTmp = Split(ntx(i), LINEBREAK)
    v3 = StrToV3(lineTmp(0))
    v2 = V3ToV2Pos(v3.X, v3.Y, v3.z)
    NodeEdit_NewNode lineTmp(1), Replace(lineTmp(2), NODELINEBREAK, vbCrLf), &HFFBF00, nodeDefaultSize, v2.X, v2.Y
Next
For i = nodeSum + 1 To UBound(ntx)
    lineTmp = Split(ntx(i), LINEBREAK)
    LineAdd_Save Val(lineTmp(0)), Val(lineTmp(1)), "", lineDefaultSize
Next
End Function

Public Function NoteFileRead_VersionCheck(ByRef firstLine As String) As Long
If InStr(1, firstLine, "Note3D_1") Then
    NoteFileRead_VersionCheck = 301
ElseIf InStr(1, firstLine, "Note2D_1") Then
    NoteFileRead_VersionCheck = 201
ElseIf InStr(1, firstLine, "Note2D_2") Then
    NoteFileRead_VersionCheck = 202
Else
    NoteFileRead_VersionCheck = -1
End If
End Function
Public Function noteArrInitialization()
ReDim node(1000): ReDim nodeLine(1000): ReDim behaviorList(1000): ReDim redolist(1000): ReDim actionList(0)
End Function
Public Function newAddNote()
MeExeIdSet
note.MainTime.Enabled = False
noteArrInitialization
nSum = 0: lSum = 0: bHLSum = 0: magnification = 0: zoomFactor = 1
note.GlobalView.Cls
ntxPath = App.path & "\�½��ڵ�ʼ�.ntx"
note.Caption = NOTEFORMNAME & ntxPath
MainCoordinateSystemDefinition
note.MainTime.Enabled = True
End Function
