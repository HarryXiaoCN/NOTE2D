Attribute VB_Name = "Note_File"
Public Function LoadProfile()
    Dim lineTmp As String
    On Error GoTo Er:
    MainFormFontSize = Note.Font.size
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
Note.显示全部节点名.Checked = False
Note.显示全部连接.Checked = False
End Function
Public Function LoadProfile_ReadLine(ByRef lineTmp As String)
    Dim ESRStr As String: Dim ESRBool As Boolean
    ESRStr = LoadProfile_ReadLine_GetEqualSignRight(lineTmp)
    ESRBool = StrToBool(ESRStr)
    If InStr(1, lineTmp, "显示全部节点名=") Then
        Note.显示全部节点名.Checked = ESRBool
    ElseIf InStr(1, lineTmp, "显示顺向节点名=") Then
        Note.显示顺向节点名.Checked = ESRBool
    ElseIf InStr(1, lineTmp, "显示逆向节点名=") Then
        Note.显示逆向节点名.Checked = ESRBool
    ElseIf InStr(1, lineTmp, "始终显示选点名=") Then
        Note.始终显示选点名.Checked = ESRBool
    ElseIf InStr(1, lineTmp, "显示节点遍历ID=") Then
        Note.显示节点遍历ID.Checked = ESRBool
    ElseIf InStr(1, lineTmp, "显示全部连接=") Then
        Note.显示全部连接.Checked = ESRBool
    ElseIf InStr(1, lineTmp, "显示顺向连接=") Then
        Note.显示顺向连接.Checked = ESRBool
    ElseIf InStr(1, lineTmp, "显示逆向连接=") Then
        Note.显示逆向连接.Checked = ESRBool
    ElseIf InStr(1, lineTmp, "始终显示选接=") Then
        Note.始终显示选接.Checked = ESRBool
    ElseIf InStr(1, lineTmp, "全局视图=") Then
        NoteGlobalViewSet ESRBool
    ElseIf InStr(1, lineTmp, "字体=") Then
        Note.Font.name = ESRStr
    ElseIf InStr(1, lineTmp, "字号=") Then
        MainFormFontSize = Val(ESRStr): Note.Font.size = MainFormFontSize
    ElseIf InStr(1, lineTmp, "加粗=") Then
        Note.Font.Bold = ESRBool
    ElseIf InStr(1, lineTmp, "倾斜=") Then
        Note.Font.Italic = ESRBool
    ElseIf InStr(1, lineTmp, "主界面全高透明=") Then
        Note.全高透明.Checked = ESRBool: FormTransparent Note, 50
    ElseIf InStr(1, lineTmp, "主界面全半透明=") Then
        Note.全半透明.Checked = ESRBool: FormTransparent Note, 125
    ElseIf InStr(1, lineTmp, "主界面全低透明=") Then
        Note.全低透明.Checked = ESRBool: FormTransparent Note, 200
    ElseIf InStr(1, lineTmp, "主界面背景色=") Then
        Note.BackColor = Val(ESRStr)
    ElseIf InStr(1, lineTmp, "主界面字体颜色=") Then
        Note.ForeColor = Val(ESRStr)
    ElseIf InStr(1, lineTmp, "矩点=") Then
        Note.矩点.Checked = ESRBool
    ElseIf InStr(1, lineTmp, "矩线=") Then
        Note.矩线.Checked = ESRBool
    ElseIf InStr(1, lineTmp, "背景图路径=") Then
        If ESRStr <> "" Then
            Note.homeBackPicPath = ESRStr
            Note.加载背景图 Note.homeBackPicPath
        End If
    ElseIf InStr(1, lineTmp, "彩虹圈=") Then
        Note.彩虹圈.Checked = ESRBool
    ElseIf InStr(1, lineTmp, "彩虹线=") Then
        Note.彩虹线.Checked = ESRBool
    ElseIf InStr(1, lineTmp, "流光溢彩=") Then
        Note.流光溢彩.Checked = ESRBool
    ElseIf InStr(1, lineTmp, "输入界面全高透明=") Then
        Note.全高透明2.Checked = ESRBool: FormTransparent NodeInput, 50
    ElseIf InStr(1, lineTmp, "输入界面全半透明=") Then
        Note.全半透明2.Checked = ESRBool: FormTransparent NodeInput, 125
    ElseIf InStr(1, lineTmp, "输入界面全低透明=") Then
        Note.全低透明2.Checked = ESRBool: FormTransparent NodeInput, 200
    ElseIf InStr(1, lineTmp, "输入界面背景色=") Then
        NodeInputBackColor = Val(ESRStr)
    ElseIf InStr(1, lineTmp, "输出界面全高透明=") Then
        Note.全高透明3.Checked = ESRBool: FormTransparent NodePrint, 50
    ElseIf InStr(1, lineTmp, "输出界面全半透明=") Then
        Note.全半透明3.Checked = ESRBool: FormTransparent NodePrint, 125
    ElseIf InStr(1, lineTmp, "输出界面全低透明=") Then
        Note.全低透明3.Checked = ESRBool: FormTransparent NodePrint, 200
    ElseIf InStr(1, lineTmp, "输出界面置顶=") Then
        Note.置顶.Checked = ESRBool: FormStick NodePrint, ESRBool
    ElseIf InStr(1, lineTmp, "标签化=") Then
        Note.标签化.Checked = ESRBool
    ElseIf InStr(1, lineTmp, "创建模式=") Then
        Note.创建模式.Checked = ESRBool
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
        Note.updataSpeed = Val(ESRStr)
        If Note.updataSpeed < 10 Then Note.updataSpeed = 10
        If Note.updataSpeed > 100 Then Note.updataSpeed = 100
        Note.MainTime.interval = Note.updataSpeed
    ElseIf InStr(1, lineTmp, "矩线颜色=") Then
        rectangleLineColor = Val(ESRStr)
    ElseIf InStr(1, lineTmp, "节点归整长度=") Then
        nodeAttributedToIntegers = Val(ESRStr)
    ElseIf InStr(1, lineTmp, "节点归整=") Then
        Note.节点归整.Checked = ESRBool
    ElseIf InStr(1, lineTmp, "精简内容=") Then
        Note.精简内容.Checked = ESRBool
    ElseIf InStr(1, lineTmp, "色彩链路显示=") Then
        Note.色彩链路.Checked = ESRBool
    ElseIf InStr(1, lineTmp, "长=") Then
        Note.width = Val(ESRStr)
    ElseIf InStr(1, lineTmp, "宽=") Then
        Note.height = Val(ESRStr)
    ElseIf InStr(1, lineTmp, "主界面X位置=") Then
        Note.left = Val(ESRStr)
    ElseIf InStr(1, lineTmp, "主界面Y位置=") Then
        Note.Top = Val(ESRStr)
    ElseIf InStr(1, lineTmp, "输入X位置=") Then
        nodeInputFormLeft = Val(ESRStr)
    ElseIf InStr(1, lineTmp, "输入Y位置=") Then
        nodeInputFormTop = Val(ESRStr)
    ElseIf InStr(1, lineTmp, "色彩链路=") Then
        色彩链路字典修改 ESRStr
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
        & vbCrLf & "始终显示选点名=" & Note.始终显示选点名.Checked _
        & vbCrLf & "显示节点遍历ID=" & Note.显示节点遍历ID.Checked _
        & vbCrLf & "[节点连接显示]" & vbCrLf & SaveProfile_GetTrueLineViewName _
        & vbCrLf & "始终显示选接=" & Note.始终显示选接.Checked _
        & vbCrLf & "全局视图=" & Note.全局视图.Checked _
        & vbCrLf & "-界面-" & vbCrLf & "[主界面]" _
        & vbCrLf & "字体=" & Note.Font.name _
        & vbCrLf & "字号=" & MainFormFontSize _
        & vbCrLf & "加粗=" & Note.Font.Bold _
        & vbCrLf & "倾斜=" & Note.Font.Italic _
        & vbCrLf & SaveProfile_GetTrueTransparentViewName _
        & vbCrLf & "主界面背景色=" & Note.BackColor & vbCrLf & "背景图路径=" & Note.homeBackPicPath _
        & vbCrLf & "主界面字体颜色=" & Note.ForeColor & vbCrLf & "精简内容=" & Note.精简内容.Checked _
        & vbCrLf & "矩点=" & Note.矩点.Checked & vbCrLf & "矩线=" & Note.矩线.Checked _
        & vbCrLf & "彩虹圈=" & Note.彩虹圈.Checked _
        & vbCrLf & "彩虹线=" & Note.彩虹线.Checked _
        & vbCrLf & "流光溢彩=" & Note.流光溢彩.Checked _
        & vbCrLf & "[输入界面]" _
        & vbCrLf & SaveProfile_GetTrueTransparent2ViewName _
        & vbCrLf & "输入界面背景色=" & NodeInputBackColor & vbCrLf & "标签化=" & Note.标签化.Checked _
        & vbCrLf & "[输出界面]" & vbCrLf & SaveProfile_GetTrueTransparent3ViewName _
        & vbCrLf & "输出界面置顶=" & Note.置顶.Checked & vbCrLf & "-系统-" & vbCrLf & "输入界面高=" & nodeInputFormHeight & vbCrLf & "输入界面宽=" & nodeInputFormWidth _
        & vbCrLf & "自动保存时间间隔=" & saveNtxTime & vbCrLf & "绘图间隔=" & Note.updataSpeed & vbCrLf & "节点默认大小=" & nodeDefaultSize & vbCrLf & "连接默认宽度=" & lineDefaultSize & vbCrLf
        profileStr = profileStr & "-系统-" & vbCrLf _
        & vbCrLf & "矩线颜色=" & rectangleLineColor _
        & vbCrLf & "节点归整长度=" & nodeAttributedToIntegers _
        & vbCrLf & "节点归整=" & Note.节点归整.Checked _
        & vbCrLf & "长=" & Note.width _
        & vbCrLf & "宽=" & Note.height _
        & vbCrLf & "主界面X位置=" & Note.left _
        & vbCrLf & "主界面Y位置=" & Note.Top _
        & vbCrLf & "输入X位置=" & NodeInput.left _
        & vbCrLf & "输入Y位置=" & NodeInput.Top _
        & vbCrLf & "创建模式=" & Note.创建模式.Checked _
        & vbCrLf & "色彩链路=" & 色彩链路字典导出 _
        & vbCrLf & "色彩链路显示=" & Note.色彩链路.Checked
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
If Note.全高透明3.Checked = True Then SaveProfile_GetTrueTransparent3ViewName = "输出界面全高透明=True": Exit Function
If Note.全半透明3.Checked = True Then SaveProfile_GetTrueTransparent3ViewName = "输出界面全半透明=True": Exit Function
If Note.全低透明3.Checked = True Then SaveProfile_GetTrueTransparent3ViewName = "输出界面全低透明=True": Exit Function
End Function
Public Function SaveProfile_GetTrueTransparent2ViewName() As String
If Note.全高透明2.Checked = True Then SaveProfile_GetTrueTransparent2ViewName = "输入界面全高透明=True": Exit Function
If Note.全半透明2.Checked = True Then SaveProfile_GetTrueTransparent2ViewName = "输入界面全半透明=True": Exit Function
If Note.全低透明2.Checked = True Then SaveProfile_GetTrueTransparent2ViewName = "输入界面全低透明=True": Exit Function
End Function
Public Function SaveProfile_GetTrueTransparentViewName() As String
If Note.全高透明.Checked = True Then SaveProfile_GetTrueTransparentViewName = "主界面全高透明=True": Exit Function
If Note.全半透明.Checked = True Then SaveProfile_GetTrueTransparentViewName = "主界面全半透明=True": Exit Function
If Note.全低透明.Checked = True Then SaveProfile_GetTrueTransparentViewName = "主界面全低透明=True": Exit Function
End Function
Public Function SaveProfile_GetTrueNodeViewName() As String
If Note.显示全部节点名.Checked = True Then SaveProfile_GetTrueNodeViewName = "显示全部节点名=True": Exit Function
If Note.显示顺向节点名.Checked = True Then SaveProfile_GetTrueNodeViewName = "显示顺向节点名=True": Exit Function
If Note.显示逆向节点名.Checked = True Then SaveProfile_GetTrueNodeViewName = "显示逆向节点名=True": Exit Function
End Function
Public Function SaveProfile_GetTrueLineViewName() As String
If Note.显示全部连接.Checked = True Then SaveProfile_GetTrueLineViewName = "显示全部连接=True": Exit Function
If Note.显示顺向连接.Checked = True Then SaveProfile_GetTrueLineViewName = "显示顺向连接=True": Exit Function
If Note.显示逆向连接.Checked = True Then SaveProfile_GetTrueLineViewName = "显示逆向连接=True": Exit Function
End Function
Public Function NoteFileRead(ByRef filePath As String)
Dim ntx() As String: Dim i As Long, version As Long
newAddNote
On Error GoTo Er
    ntxPath = filePath
    ntxPathNoName = 去除路径文件名(ntxPath)
    Note.Caption = NOTEFORMNAME & ntxPath
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
        Case 203
            NoteFileRead_203 ntx, False
            lastNtx = ntx
        Case 204
            NoteFileRead_204 ntx, False
            lastNtx = ntx
        Case 301
            NoteFileRead_301 ntx
    End Select
    Note.MainTime.Enabled = True
Exit Function
Er:
    MsgBox "文件读取失败，原因：" & Err.Description, 16, "错误"
End Function
Public Function NoteFileRead_204(ByRef ntx() As String, ByRef fromCopy As Boolean)
    Dim i As Long, nodeSum As Long, lineSum As Long, startNodeId As Long: Dim lineTmp() As String
    lineTmp = Split(ntx(0), LINEBREAK)
    nodeSum = Val(lineTmp(1))
    lineSum = Val(lineTmp(2))
    If UBound(lineTmp) > 2 Then
        magnification = Val(lineTmp(3)): zoomFactor = MToZF(magnification)
        MainCoordinateSystemDefinition
        If lineTmp(4) = "True" Then Note.节点归整.Checked = True Else Note.节点归整.Checked = False
        If UBound(lineTmp) > 4 And lineTmp(5) <> "" Then
            NoteFileRead_202_GetDic lineTmp(5)
        End If
        If UBound(lineTmp) > 6 And lineTmp(6) <> "" And lineTmp(7) <> "" Then
            angleOfView.X = Val(lineTmp(6))
            angleOfView.Y = Val(lineTmp(7))
            MainCoordinateSystemDefinition
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
        NodeEdit_NewNode lineTmp(2), Replace(lineTmp(3), "^||`", vbCrLf), nodeDefaultColor, nodeDefaultSize, Val(lineTmp(0)) + mousePos.X, Val(lineTmp(1)) + mousePos.Y, True
    Else
        NodeEdit_NewNode lineTmp(2), Replace(lineTmp(3), "^||`", vbCrLf), nodeDefaultColor, nodeDefaultSize, Val(lineTmp(0)), Val(lineTmp(1))
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
Public Function NoteFileRead_203(ByRef ntx() As String, ByRef fromCopy As Boolean)
    Dim i As Long, nodeSum As Long, lineSum As Long, startNodeId As Long: Dim lineTmp() As String
    lineTmp = Split(ntx(0), LINEBREAK)
    nodeSum = Val(lineTmp(1))
    lineSum = Val(lineTmp(2))
    If UBound(lineTmp) > 2 Then
        magnification = Val(lineTmp(3)): zoomFactor = MToZF(magnification)
        MainCoordinateSystemDefinition
        If lineTmp(4) = "True" Then Note.节点归整.Checked = True Else Note.节点归整.Checked = False
        If UBound(lineTmp) > 4 And lineTmp(5) <> "" Then
            NoteFileRead_202_GetDic lineTmp(5)
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
    ntx = NoteFileWrite_204_Coding(node, nSum, nodeLine, lSum)
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
Public Function NoteFileWrite_204_Coding(ByRef nodeObj() As 节点, ByRef nodeObjSum As Long, _
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
    ntx(0) = VERSIONID & LINEBREAK & nodeSum & LINEBREAK & lineSum & LINEBREAK & magnification & LINEBREAK & Note.节点归整.Checked & LINEBREAK & NoteFileWrite_204_Coding_DIC & LINEBREAK & angleOfView.X & LINEBREAK & angleOfView.Y
    NoteFileWrite_204_Coding = ntx
End Function
Public Function NoteFileWrite_204_Coding_DIC() As String
    Dim i As Long
    For i = 0 To nodeSelectKeyDic.Count - 1
        NoteFileWrite_204_Coding_DIC = NoteFileWrite_204_Coding_DIC & nodeSelectKeyDic.Keys(i) & VALUEBREAK & nodeSelectKeyDic.Items(i) & KEYBREAK
    Next
        
    If Len(NoteFileWrite_204_Coding_DIC) > 0 Then
        NoteFileWrite_204_Coding_DIC = Mid(NoteFileWrite_204_Coding_DIC, 1, Len(NoteFileWrite_204_Coding_DIC) - 1)
    End If
    NoteFileWrite_204_Coding_DIC = NoteFileWrite_204_Coding_DIC & DICBREAK
    
    For i = 0 To lineSelectKeyDic.Count - 1
        NoteFileWrite_204_Coding_DIC = NoteFileWrite_204_Coding_DIC & lineSelectKeyDic.Keys(i) & VALUEBREAK & lineSelectKeyDic.Items(i) & KEYBREAK
    Next
    If Len(NoteFileWrite_204_Coding_DIC) > 0 Then
        NoteFileWrite_204_Coding_DIC = Mid(NoteFileWrite_204_Coding_DIC, 1, Len(NoteFileWrite_204_Coding_DIC) - 1)
    End If
End Function
Public Function NoteFileWrite_204(ByRef filePath As String)
    Dim ntx() As String: Dim i As Long, fN As Integer
    ntxPath = filePath
    ntxPathNoName = 去除路径文件名(ntxPath)
    Note.Caption = NOTEFORMNAME & ntxPath
    ntx = NoteFileWrite_204_Coding(node, nSum, nodeLine, lSum)
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
    NodeEdit_NewNode lineTmp(1), Replace(lineTmp(2), NODELINEBREAK, vbCrLf), nodeDefaultColor, nodeDefaultSize, v2.X, v2.Y
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
    ElseIf InStr(1, firstLine, "Note2D_3") Then
        NoteFileRead_VersionCheck = 203
    ElseIf InStr(1, firstLine, "Note2D_4") Then
        NoteFileRead_VersionCheck = 204
    Else
        NoteFileRead_VersionCheck = -1
    End If
End Function
Public Function noteArrInitialization()
    ReDim node(1000): ReDim nodeLine(1000): ReDim behaviorList(1000): ReDim redolist(1000): ReDim actionList(0)
End Function
Public Function newAddNote()
    MeExeIdSet
    Note.MainTime.Enabled = False
    noteArrInitialization
    nSum = 0: lSum = 0: bHLSum = 0: magnification = 0: zoomFactor = 1
    Note.GlobalView.Cls
    ntxPath = App.path & "\新建节点笔记.ntx"
    ntxPathNoName = 去除路径文件名(ntxPath)
    Note.Caption = NOTEFORMNAME & ntxPath
    MainCoordinateSystemDefinition
    Note.MainTime.Enabled = True
End Function
