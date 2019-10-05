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
note.显示全部节点名.Checked = False
note.显示全部连接.Checked = False
End Function
Public Function LoadProfile_ReadLine(ByRef lineTmp As String)
Dim ESRStr As String: Dim ESRBool As Boolean
ESRStr = LoadProfile_ReadLine_GetEqualSignRight(lineTmp)
ESRBool = StrToBool(ESRStr)
If InStr(1, lineTmp, "显示全部节点名=") Then
    note.显示全部节点名.Checked = ESRBool
ElseIf InStr(1, lineTmp, "显示顺向节点名=") Then
    note.显示顺向节点名.Checked = ESRBool
ElseIf InStr(1, lineTmp, "显示逆向节点名=") Then
    note.显示逆向节点名.Checked = ESRBool
ElseIf InStr(1, lineTmp, "始终显示选点名=") Then
    note.始终显示选点名.Checked = ESRBool
ElseIf InStr(1, lineTmp, "显示节点遍历ID=") Then
    note.显示节点遍历ID.Checked = ESRBool
ElseIf InStr(1, lineTmp, "显示全部连接=") Then
    note.显示全部连接.Checked = ESRBool
ElseIf InStr(1, lineTmp, "显示顺向连接=") Then
    note.显示顺向连接.Checked = ESRBool
ElseIf InStr(1, lineTmp, "显示逆向连接=") Then
    note.显示逆向连接.Checked = ESRBool
ElseIf InStr(1, lineTmp, "始终显示选接=") Then
    note.始终显示选接.Checked = ESRBool
ElseIf InStr(1, lineTmp, "全局视图=") Then
    NoteGlobalViewSet ESRBool
ElseIf InStr(1, lineTmp, "字体=") Then
    note.Font.name = ESRStr
ElseIf InStr(1, lineTmp, "字号=") Then
    MainFormFontSize = Val(ESRStr): note.Font.size = MainFormFontSize
ElseIf InStr(1, lineTmp, "加粗=") Then
    note.Font.Bold = ESRBool
ElseIf InStr(1, lineTmp, "倾斜=") Then
    note.Font.Italic = ESRBool
ElseIf InStr(1, lineTmp, "主界面全高透明=") Then
    note.全高透明.Checked = ESRBool: FormTransparent note, 50
ElseIf InStr(1, lineTmp, "主界面全半透明=") Then
    note.全半透明.Checked = ESRBool: FormTransparent note, 125
ElseIf InStr(1, lineTmp, "主界面全低透明=") Then
    note.全低透明.Checked = ESRBool: FormTransparent note, 200
ElseIf InStr(1, lineTmp, "主界面背景色=") Then
    note.BackColor = Val(ESRStr)
ElseIf InStr(1, lineTmp, "主界面字体颜色=") Then
    note.ForeColor = Val(ESRStr)
ElseIf InStr(1, lineTmp, "矩点=") Then
    note.矩点.Checked = ESRBool
ElseIf InStr(1, lineTmp, "矩线=") Then
    note.矩线.Checked = ESRBool
ElseIf InStr(1, lineTmp, "背景图路径=") Then
    If ESRStr <> "" Then
        note.homeBackPicPath = ESRStr
        note.加载背景图 note.homeBackPicPath
    End If
ElseIf InStr(1, lineTmp, "彩虹圈=") Then
    note.彩虹圈.Checked = ESRBool
ElseIf InStr(1, lineTmp, "彩虹线=") Then
    note.彩虹线.Checked = ESRBool
ElseIf InStr(1, lineTmp, "流光溢彩=") Then
    note.流光溢彩.Checked = ESRBool
ElseIf InStr(1, lineTmp, "输入界面全高透明=") Then
    note.全高透明2.Checked = ESRBool: FormTransparent NodeInput, 50
ElseIf InStr(1, lineTmp, "输入界面全半透明=") Then
    note.全半透明2.Checked = ESRBool: FormTransparent NodeInput, 125
ElseIf InStr(1, lineTmp, "输入界面全低透明=") Then
    note.全低透明2.Checked = ESRBool: FormTransparent NodeInput, 200
ElseIf InStr(1, lineTmp, "输入界面背景色=") Then
    NodeInputBackColor = Val(ESRStr)
ElseIf InStr(1, lineTmp, "输出界面全高透明=") Then
    note.全高透明3.Checked = ESRBool: FormTransparent NodePrint, 50
ElseIf InStr(1, lineTmp, "输出界面全半透明=") Then
    note.全半透明3.Checked = ESRBool: FormTransparent NodePrint, 125
ElseIf InStr(1, lineTmp, "输出界面全低透明=") Then
    note.全低透明3.Checked = ESRBool: FormTransparent NodePrint, 200
ElseIf InStr(1, lineTmp, "输出界面置顶=") Then
    note.置顶.Checked = ESRBool: FormStick NodePrint, ESRBool
ElseIf InStr(1, lineTmp, "标签化=") Then
    note.标签化.Checked = ESRBool
ElseIf InStr(1, lineTmp, "输入界面高=") Then
    nodeInputFormHeight = Val(ESRStr)
ElseIf InStr(1, lineTmp, "输入界面宽=") Then
    nodeInputFormWidth = Val(ESRStr)
ElseIf InStr(1, lineTmp, "自动保存时间间隔=") Then
    saveNtxTime = Val(ESRStr)
ElseIf InStr(1, lineTmp, "节点默认大小=") Then
    nodeDefaultSize = Val(ESRStr)
    nodeDefaultSize = 限制数值(nodeDefaultSize, 50, 500)
ElseIf InStr(1, lineTmp, "连接默认宽度=") Then
    lineDefaultSize = Val(ESRStr)
    lineDefaultSize = 限制数值(lineDefaultSize, 1, 10)
ElseIf InStr(1, lineTmp, "绘图间隔=") Then
    note.updataSpeed = Val(ESRStr)
    If note.updataSpeed < 10 Then note.updataSpeed = 10
    If note.updataSpeed > 100 Then note.updataSpeed = 100
    note.MainTime.interval = note.updataSpeed
ElseIf InStr(1, lineTmp, "矩线颜色=") Then
    rectangleLineColor = Val(ESRStr)
ElseIf InStr(1, lineTmp, "矩线步长=") Then
    rectangleStep = Val(ESRStr)
ElseIf InStr(1, lineTmp, "节点归整长度=") Then
    nodeAttributedToIntegers = Val(ESRStr)
ElseIf InStr(1, lineTmp, "节点归整=") Then
    note.节点归整.Checked = ESRBool
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
profileStr = "-视图-" & vbCrLf _
& "[节点名显示]" & vbCrLf & SaveProfile_GetTrueNodeViewName _
& vbCrLf & "始终显示选点名=" & note.始终显示选点名.Checked _
& vbCrLf & "显示节点遍历ID=" & note.显示节点遍历ID.Checked _
& vbCrLf & "[节点连接显示]" & vbCrLf & SaveProfile_GetTrueLineViewName _
& vbCrLf & "始终显示选接=" & note.始终显示选接.Checked _
& vbCrLf & "全局视图=" & note.全局视图.Checked _
& vbCrLf & "-界面-" & vbCrLf & "[主界面]" _
& vbCrLf & "字体=" & note.Font.name _
& vbCrLf & "字号=" & MainFormFontSize _
& vbCrLf & "加粗=" & note.Font.Bold _
& vbCrLf & "倾斜=" & note.Font.Italic _
& vbCrLf & SaveProfile_GetTrueTransparentViewName _
& vbCrLf & "主界面背景色=" & note.BackColor & vbCrLf & "背景图路径=" & note.homeBackPicPath _
& vbCrLf & "主界面字体颜色=" & note.ForeColor _
& vbCrLf & "矩点=" & note.矩点.Checked & vbCrLf & "矩线=" & note.矩线.Checked _
& vbCrLf & "彩虹圈=" & note.彩虹圈.Checked _
& vbCrLf & "彩虹线=" & note.彩虹线.Checked _
& vbCrLf & "流光溢彩=" & note.流光溢彩.Checked _
& vbCrLf & "[输入界面]" _
& vbCrLf & SaveProfile_GetTrueTransparent2ViewName _
& vbCrLf & "输入界面背景色=" & NodeInputBackColor & vbCrLf & "标签化=" & note.标签化.Checked _
& vbCrLf & "[输出界面]" & vbCrLf & SaveProfile_GetTrueTransparent3ViewName _
& vbCrLf & "输出界面置顶=" & note.置顶.Checked & vbCrLf & "-系统-" & vbCrLf & "输入界面高=" & nodeInputFormHeight & vbCrLf & "输入界面宽=" & nodeInputFormWidth _
& vbCrLf & "自动保存时间间隔=" & saveNtxTime & vbCrLf & "绘图间隔=" & note.updataSpeed & vbCrLf & "节点默认大小=" & nodeDefaultSize & vbCrLf & "连接默认宽度=" & lineDefaultSize & vbCrLf
profileStr = profileStr & "-系统-" & vbCrLf _
& vbCrLf & "矩线颜色=" & rectangleLineColor _
& vbCrLf & "矩线步长=" & rectangleStep _
& vbCrLf & "节点归整长度=" & nodeAttributedToIntegers _
& vbCrLf & "节点归整=" & note.节点归整.Checked
'& vbCrLf & "-系统-" & vbCrLf & "放缩率=" & magnification
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
MsgBox "程序配置文件保存失败！请检查路径" & ProfilePath & "是否存在。", , "警告！"
End Function
Public Function SaveProfile_GetTrueTransparent3ViewName() As String
If note.全高透明3.Checked = True Then SaveProfile_GetTrueTransparent3ViewName = "输出界面全高透明=True": Exit Function
If note.全半透明3.Checked = True Then SaveProfile_GetTrueTransparent3ViewName = "输出界面全半透明=True": Exit Function
If note.全低透明3.Checked = True Then SaveProfile_GetTrueTransparent3ViewName = "输出界面全低透明=True": Exit Function
End Function
Public Function SaveProfile_GetTrueTransparent2ViewName() As String
If note.全高透明2.Checked = True Then SaveProfile_GetTrueTransparent2ViewName = "输入界面全高透明=True": Exit Function
If note.全半透明2.Checked = True Then SaveProfile_GetTrueTransparent2ViewName = "输入界面全半透明=True": Exit Function
If note.全低透明2.Checked = True Then SaveProfile_GetTrueTransparent2ViewName = "输入界面全低透明=True": Exit Function
End Function
Public Function SaveProfile_GetTrueTransparentViewName() As String
If note.全高透明.Checked = True Then SaveProfile_GetTrueTransparentViewName = "主界面全高透明=True": Exit Function
If note.全半透明.Checked = True Then SaveProfile_GetTrueTransparentViewName = "主界面全半透明=True": Exit Function
If note.全低透明.Checked = True Then SaveProfile_GetTrueTransparentViewName = "主界面全低透明=True": Exit Function
End Function
Public Function SaveProfile_GetTrueNodeViewName() As String
If note.显示全部节点名.Checked = True Then SaveProfile_GetTrueNodeViewName = "显示全部节点名=True": Exit Function
If note.显示顺向节点名.Checked = True Then SaveProfile_GetTrueNodeViewName = "显示顺向节点名=True": Exit Function
If note.显示逆向节点名.Checked = True Then SaveProfile_GetTrueNodeViewName = "显示逆向节点名=True": Exit Function
End Function
Public Function SaveProfile_GetTrueLineViewName() As String
If note.显示全部连接.Checked = True Then SaveProfile_GetTrueLineViewName = "显示全部连接=True": Exit Function
If note.显示顺向连接.Checked = True Then SaveProfile_GetTrueLineViewName = "显示顺向连接=True": Exit Function
If note.显示逆向连接.Checked = True Then SaveProfile_GetTrueLineViewName = "显示逆向连接=True": Exit Function
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
        MsgBox "文件无法识别！", , "警告！"
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
    MsgBox "文件读取失败，原因：" & Err.Description, 16, "错误"
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
Public Function NoteFileWrite_202_Coding(ByRef nodeObj() As 节点, ByRef nodeObjSum As Long, _
ByRef lineObj() As 连接, ByRef lineObjSum As Long)
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
Dim v3 As 三维坐标
Dim v2 As 二维坐标
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
ntxPath = App.path & "\新建节点笔记.ntx"
note.Caption = NOTEFORMNAME & ntxPath
MainCoordinateSystemDefinition
note.MainTime.Enabled = True
End Function
