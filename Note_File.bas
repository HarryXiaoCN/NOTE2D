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
If InStr(1, lineTmp, "显示全部节点名=") Then Note.显示全部节点名.Checked = ESRBool: Exit Function
If InStr(1, lineTmp, "显示顺向节点名=") Then Note.显示顺向节点名.Checked = ESRBool: Exit Function
If InStr(1, lineTmp, "显示逆向节点名=") Then Note.显示逆向节点名.Checked = ESRBool: Exit Function
If InStr(1, lineTmp, "始终显示选点名=") Then Note.始终显示选点名.Checked = ESRBool: Exit Function
If InStr(1, lineTmp, "显示节点遍历ID=") Then Note.显示节点遍历ID.Checked = ESRBool: Exit Function
If InStr(1, lineTmp, "显示全部连接=") Then Note.显示全部连接.Checked = ESRBool: Exit Function
If InStr(1, lineTmp, "显示顺向连接=") Then Note.显示顺向连接.Checked = ESRBool: Exit Function
If InStr(1, lineTmp, "显示逆向连接=") Then Note.显示逆向连接.Checked = ESRBool: Exit Function
If InStr(1, lineTmp, "始终显示选接=") Then Note.始终显示选接.Checked = ESRBool: Exit Function
If InStr(1, lineTmp, "全局视图=") Then NoteGlobalViewSet ESRBool: Exit Function
If InStr(1, lineTmp, "字体=") Then Note.Font.Name = ESRStr: Exit Function
If InStr(1, lineTmp, "字号=") Then MainFormFontSize = Val(ESRStr): Note.Font.size = MainFormFontSize: Exit Function
If InStr(1, lineTmp, "加粗=") Then Note.Font.Bold = ESRBool: Exit Function
If InStr(1, lineTmp, "倾斜=") Then Note.Font.Italic = ESRBool: Exit Function
If InStr(1, lineTmp, "主界面全高透明=") Then Note.全高透明.Checked = ESRBool: FormTransparent Note, 50: Exit Function
If InStr(1, lineTmp, "主界面全半透明=") Then Note.全半透明.Checked = ESRBool: FormTransparent Note, 125: Exit Function
If InStr(1, lineTmp, "主界面全低透明=") Then Note.全低透明.Checked = ESRBool: FormTransparent Note, 200: Exit Function
If InStr(1, lineTmp, "主界面背景色=") Then Note.BackColor = Val(ESRStr): Exit Function
If InStr(1, lineTmp, "主界面字体颜色=") Then Note.ForeColor = Val(ESRStr): Exit Function
If InStr(1, lineTmp, "彩虹圈=") Then Note.彩虹圈.Checked = ESRBool: Exit Function
If InStr(1, lineTmp, "彩虹线=") Then Note.彩虹线.Checked = ESRBool: Exit Function
If InStr(1, lineTmp, "流光溢彩=") Then If Note.彩虹线.Checked = True And ESRBool = True Then Note.流光溢彩.Checked = ESRBool: Exit Function
If InStr(1, lineTmp, "输入界面全高透明=") Then Note.全高透明2.Checked = ESRBool: FormTransparent NodeInput, 50: Exit Function
If InStr(1, lineTmp, "输入界面全半透明=") Then Note.全半透明2.Checked = ESRBool: FormTransparent NodeInput, 125: Exit Function
If InStr(1, lineTmp, "输入界面全低透明=") Then Note.全低透明2.Checked = ESRBool: FormTransparent NodeInput, 200: Exit Function
If InStr(1, lineTmp, "输入界面背景色=") Then NodeInputBackColor = Val(ESRStr): Exit Function
If InStr(1, lineTmp, "输出界面全高透明=") Then Note.全高透明3.Checked = ESRBool: FormTransparent NodePrint, 50: Exit Function
If InStr(1, lineTmp, "输出界面全半透明=") Then Note.全半透明3.Checked = ESRBool: FormTransparent NodePrint, 125: Exit Function
If InStr(1, lineTmp, "输出界面全低透明=") Then Note.全低透明3.Checked = ESRBool: FormTransparent NodePrint, 200: Exit Function
If InStr(1, lineTmp, "输出界面置顶=") Then Note.置顶.Checked = ESRBool: FormStick NodePrint, ESRBool: Exit Function
If InStr(1, lineTmp, "标签化=") Then Note.标签化.Checked = ESRBool:  Exit Function
If InStr(1, lineTmp, "自动保存时间间隔=") Then saveNtxTime = Val(ESRStr):  Exit Function
'If InStr(1, lineTmp, "放缩率=") Then magnification = Val(ESRStr): zoomFactor = MToZF(magnification): Exit Function
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
& vbCrLf & "字体=" & Note.Font.Name _
& vbCrLf & "字号=" & MainFormFontSize _
& vbCrLf & "加粗=" & Note.Font.Bold _
& vbCrLf & "倾斜=" & Note.Font.Italic _
& vbCrLf & SaveProfile_GetTrueTransparentViewName _
& vbCrLf & "主界面背景色=" & Note.BackColor _
& vbCrLf & "主界面字体颜色=" & Note.ForeColor _
& vbCrLf & "彩虹圈=" & Note.彩虹圈.Checked _
& vbCrLf & "彩虹线=" & Note.彩虹线.Checked _
& vbCrLf & "流光溢彩=" & Note.流光溢彩.Checked _
& vbCrLf & "[输入界面]" _
& vbCrLf & SaveProfile_GetTrueTransparent2ViewName _
& vbCrLf & "输入界面背景色=" & NodeInputBackColor & vbCrLf & "标签化=" & Note.标签化.Checked _
& vbCrLf & "[输出界面]" _
& vbCrLf & SaveProfile_GetTrueTransparent3ViewName _
& vbCrLf & "输出界面置顶=" & Note.置顶.Checked _
& vbCrLf & "自动保存时间间隔=" & saveNtxTime
'& vbCrLf & "-系统-" & vbCrLf & "放缩率=" & magnification
If Dir(ProfilePath, vbDirectory) = "" Then
    Shell "cmd /c md " & ProfilePath, vbHide
    Do While Dir(ProfilePath, vbDirectory) = ""
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
Dim ntx() As String: Dim i, version As Long
newAddNote
On Error GoTo Er
ntxPath = filePath
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
    Case 200
        NoteFileRead_200 filePath
    Case 201
        NoteFileRead_201 ntx, False
    Case 301
        NoteFileRead_301 ntx
End Select
Note.MainTime.Enabled = True
Er:
End Function

Public Function NoteFileRead_201(ByRef ntx() As String, ByRef fromCopy As Boolean)
Dim i, nodeSum, lineSum, startNodeId As Long: Dim lineTmp
lineTmp = Split(ntx(0), LINEBREAK)
nodeSum = Val(lineTmp(1))
lineSum = Val(lineTmp(2))
If UBound(lineTmp) > 2 Then
    magnification = Val(lineTmp(3)): zoomFactor = MToZF(magnification)
    MainCoordinateSystemDefinition
End If
startNodeId = nSum
For i = 1 To nodeSum
    lineTmp = Split(ntx(i), LINEBREAK)
    If fromCopy = True Then
        NodeEdit_NewNode lineTmp(2), Replace(lineTmp(3), NODELINEBREAK, vbCrLf), Val(lineTmp(0)) + mousePos.x, Val(lineTmp(1)) + mousePos.y, True
    Else
        NodeEdit_NewNode lineTmp(2), Replace(lineTmp(3), NODELINEBREAK, vbCrLf), Val(lineTmp(0)), Val(lineTmp(1))
    End If
Next
For i = nodeSum + 1 To nodeSum + lineSum
    lineTmp = Split(ntx(i), LINEBREAK)
    If fromCopy = True Then
        LineAdd_Save Val(lineTmp(0)) + startNodeId, Val(lineTmp(1)) + startNodeId, True
    Else
        LineAdd_Save Val(lineTmp(0)), Val(lineTmp(1))
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
Public Function noteSaveCheck_ContentCheck(ByRef filePath As String) As Boolean
Dim ntx() As String
Dim lineTmp As String
Dim i As Long
ntx = NoteFileWrite_201_Coding(node, nSum, nodeLine, lSum)
On Error GoTo Er
Open filePath For Input As #2
    Do While Not EOF(2)
        Line Input #2, lineTmp
        If lineTmp = "" Then Exit Do
        If lineTmp <> ntx(i) Then
            noteSaveCheck_ContentCheck = True
            Exit Do
        End If
        i = i + 1
        DoEvents
    Loop
Close #2
If i = 0 Then GoTo Er
Exit Function
Er:
noteSaveCheck_ContentCheck = True
End Function
Public Function NoteFileWrite_201_Coding(ByRef nodeObj() As 节点, ByRef nodeObjSum As Long, _
ByRef lineObj() As 连接, ByRef lineObjSum As Long)
Dim ntx() As String: Dim i As Long, j As Long, nodeSum As Long, lineSum As Long, nodeIdList() As Long
ReDim nodeIdList(nodeObjSum)
ReDim ntx(nodeObjSum + lineObjSum + 1)
j = 1
For i = 0 To nodeObjSum - 1
    With nodeObj(i)
        If .b = True Then
            ntx(j) = .x & LINEBREAK & .y & LINEBREAK & .t & LINEBREAK & Replace(.content, vbCrLf, NODELINEBREAK)
            nodeIdList(i) = j - 1
            j = j + 1
        End If
    End With
Next
nodeSum = j - 1
For i = 0 To lineObjSum - 1
    With lineObj(i)
        If .b = True Then
            ntx(j) = nodeIdList(.source) & LINEBREAK & nodeIdList(.target)
            j = j + 1
        End If
    End With
Next
lineSum = j - nodeSum - 1
ntx(0) = VERSIONID & LINEBREAK & nodeSum & LINEBREAK & lineSum & LINEBREAK & magnification
NoteFileWrite_201_Coding = ntx
End Function
Public Function NoteFileWrite_201(ByRef filePath As String)
Dim ntx() As String: Dim i As Long
ntxPath = filePath
Note.Caption = NOTEFORMNAME & ntxPath
ntx = NoteFileWrite_201_Coding(node, nSum, nodeLine, lSum)
Open filePath For Output As #1
    For i = 0 To UBound(ntx)
        Print #1, ntx(i)
        DoEvents
    Next
Close #1
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
    v2 = V3ToV2Pos(v3.x, v3.y, v3.z)
    NodeEdit_NewNode lineTmp(1), Replace(lineTmp(2), NODELINEBREAK, vbCrLf), v2.x, v2.y
Next
For i = nodeSum + 1 To UBound(ntx)
    lineTmp = Split(ntx(i), LINEBREAK)
    LineAdd_Save Val(lineTmp(0)), Val(lineTmp(1))
Next
End Function

Public Function NoteFileRead_VersionCheck(ByRef firstLine As String) As Long
If InStr(1, firstLine, "Note3D_1") Then
    NoteFileRead_VersionCheck = 301
ElseIf InStr(1, firstLine, "Note2D_1") Then
    NoteFileRead_VersionCheck = 201
ElseIf Val(firstLine) = 0 Then
    NoteFileRead_VersionCheck = -1
Else
    NoteFileRead_VersionCheck = 200
End If
End Function
Public Function noteArrInitialization()
ReDim node(1000): ReDim nodeLine(1000): ReDim behaviorList(1000): ReDim redolist(1000)
End Function
Public Function newAddNote()
MeExeIdSet
Note.MainTime.Enabled = False
noteArrInitialization
nSum = 0: lSum = 0: bHLSum = 0: magnification = 0: zoomFactor = 1
ntxPath = App.Path & "\新建节点笔记.ntx"
Note.Caption = NOTEFORMNAME & ntxPath
MainCoordinateSystemDefinition
Note.MainTime.Enabled = True
End Function
Public Function NoteFileRead_200(ByRef filePath As String)
Dim i, c As Long: Dim lineStr As String
Dim 版本 As Boolean
Open filePath For Input As #1
    Do While Not EOF(1)
        Line Input #1, lineStr
        Line Input #1, lineStr
        Line Input #1, lineStr
        Line Input #1, lineStr
        Line Input #1, lineStr
        Line Input #1, lineStr
        Line Input #1, lineStr
        Line Input #1, lineStr
        If lineStr <> "True" And lineStr <> "False" Then
            版本 = True
        End If
        Exit Do
    Loop
Close #1
If 版本 = True Then
    Open filePath For Input As #1
        Do While Not EOF(1)
            Line Input #1, lineStr
            nSum = Val(lineStr)
            Line Input #1, lineStr
            lSum = Val(lineStr)
            NodeUboundAdd
            LineUboundAdd
            For i = 0 To nSum - 1
                Line Input #1, lineStr
                node(i).b = CBool(lineStr)
                Line Input #1, node(i).t
                Line Input #1, lineStr
                node(i).x = Val(lineStr)
                Line Input #1, lineStr
                node(i).y = Val(lineStr)
                Line Input #1, lineStr
                node(i).content = Replace(lineStr, "_/_", vbCrLf)
                Line Input #1, lineStr
            Next
            For i = 0 To lSum - 1
                Line Input #1, lineStr
                nodeLine(i).b = CBool(lineStr)
                Line Input #1, lineStr
                nodeLine(i).source = Val(lineStr)
                Line Input #1, lineStr
                nodeLine(i).target = Val(lineStr)
            Next
        Loop
   Close #1
Else
    Open filePath For Input As #1
        Do While Not EOF(1)
            Line Input #1, lineStr
            nSum = Val(lineStr)
            Line Input #1, lineStr
            lSum = Val(lineStr)
            NodeUboundAdd
            LineUboundAdd
            For i = 0 To nSum - 1
                Line Input #1, lineStr
                node(i).b = CBool(lineStr)
                Line Input #1, node(i).t
                Line Input #1, lineStr
                node(i).x = Val(lineStr)
                Line Input #1, lineStr
                node(i).y = Val(lineStr)
                Line Input #1, lineStr
                node(i).content = Replace(lineStr, "_/_", vbCrLf)
            Next
            For i = 0 To lSum - 1
                Line Input #1, lineStr
                nodeLine(i).b = CBool(lineStr)
                Line Input #1, lineStr
                nodeLine(i).source = Val(lineStr)
                Line Input #1, lineStr
                nodeLine(i).target = Val(lineStr)
            Next
        Loop
    Close #1
End If
End Function
