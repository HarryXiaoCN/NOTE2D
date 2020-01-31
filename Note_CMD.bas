Attribute VB_Name = "Note_CMD"
Public Const 控制台名字 = "Control Desk"
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
        Case "帮助", "HELP"
            CMD_LineExecute = vbCrLf & "阵列新增节点[FORNODEADD] xStart(数值) xStep(数值) xCounts(数值) yStart(数值) yStep(数值) yCounts(数值) nodeTitle(字符串) nodeContent(字符串) pitchOn(0/1) size(数值) color(数值)" _
                            & vbCrLf & "显示鼠标坐标[VISMOUSEPOS] 1(显示)/0(不显示)" _
                            & vbCrLf & "自增偏移量[SELFIM] i偏移 x偏移 y偏移" _
                            & vbCrLf & "字典项增加[DICITEMADD] 键A:值A,键B:值B……" _
                            & vbCrLf & "字典项清空[DICREMOVEALL]" _
                            & vbCrLf & "打印字典[PRINTDIC]" _
                            & vbCrLf & "打印撤销列表[PRINTREVOKE]" _
                            & vbCrLf & "打印重做列表[PRINTREDO]" _
                            & vbCrLf & "设置树状文本导入位置控制常数[SETTREETXTINPOSCONTROLCONST/STTIPCC/STIPC] 根节点X(数值) 根节点Y(数值) 节点X间隔(数值) 节点Y间隔(数值)" _
                            & vbCrLf & "设置位图导入位置控制常数[SETIMAGEINPOSCONTROLCONST/SIIPCC/SIPC] 根节点X(数值) 根节点Y(数值) 节点X间隔(数值) 节点Y间隔(数值)" _
                            & vbCrLf & "矩线颜色[RECTANGLECOLOR/RECCOLOR] VBColor(数值)[RColor(数值) GColor(数值) BColor(数值)]" _
                            & vbCrLf & "节点放缩[NODEZOOM] 基点节点名(字符串) X轴放缩倍数(数值) Y轴放缩倍数(数值)" _
                            & vbCrLf & "创建节点[NEWBUILTNODE/NBN] X位置(数值) Y位置(数值) 标题(字符串) 内容(字符串) VBColor(数值) 大小(数值) 选中(0/1)" _
                            & vbCrLf & "编辑节点[EDITNODE/EN] 节点遍历ID(数值) 标题(字符串) 内容(字符串) VBColor(数值) 大小(数值)" _
                            & vbCrLf & "位移节点[MOVENODE/MN] 节点遍历ID(数值) X位置(数值) Y位置(数值)" _
                            & vbCrLf & "删除节点[DELETENODE/DN] 节点遍历ID1(数值),节点遍历ID2(数值),节点遍历ID3(数值)..." _
                            & vbCrLf & "选中节点[SELECTNODE/SN] 节点遍历ID1(数值),节点遍历ID2(数值),节点遍历ID3(数值)..." _
                            & vbCrLf & "创建连接[NEWBUILTNODE/NBL] 连接源节点遍历ID(数值) 连接去节点遍历ID(数值) 连接内容(字符串) 连接粗细(数值) 选中(0/1) *连接已存在会被删除" _
                            & vbCrLf & "编辑连接内容[EDITLINE/EL] 源节点遍历ID(数值) 去节点遍历ID(数值) 连接内容(字符串) 连接粗细(数值)" _
                            & vbCrLf & "选中连接[SELECTLINE/SL] 连接1源节点遍历ID(数值):连接1去节点遍历ID(数值),连接2源节点遍历ID(数值):连接2去节点遍历ID(数值),连接3源节点遍历ID(数值):连接3去节点遍历ID(数值)..." _
                            & vbCrLf & "设置动作更新速度[SETACTIONUPDATASPEED/SAUS] 更新间隔(数值)" _
                            & vbCrLf & "启动动作时钟[STARTACTIONTIMER/SAT] 1(启动)/0(关闭)" _
                            & vbCrLf & "定义动作[DEFINEACTION/DEFA/DA] 动作名(字符串),动作节点ID1(数值)[|动作节点ID2(数值)[|动作节点ID3(数值)[...]]],动作时间执行间隔(数值),动作类型(直线/圆周),直线:向量X(数值),向量Y(数值)/[圆周:角度(数值),中心节点ID(数值)],动作次数(数值),是否循环(0/1)" _
                            & vbCrLf & "重启动作[RESTARTACTION/RA] 动作名(字符串)"
            CMD_LineExecute = CMD_LineExecute _
                            & vbCrLf & "关闭动作[OFFACTION/OA] 动作名(字符串)" _
                            & vbCrLf & "打印动作列表[PRINTACTIONLIST/PAL]" _
                            & vbCrLf & "打印可执行动作列表[PRINTEXECUTABLEACTIONLIST/PEAL]" _
                            & vbCrLf & "重置镜头位置[RESETLENSPOSITION/PLP] X(数值),Y(数值)" _
                            & vbCrLf & "色彩链路表修改[COLORLINKDICMOD/CLDM] 字典字符串(VBColor1:VBColor2,VBColor2:VBColor3...)" _
                            & vbCrLf & "色彩链路表重置[COLORLINKDICRESET/CLDS]" _
                            & vbCrLf & "打印色彩链路表[PRINTCOLORLINKDIC/PCLD]"
            Exit Function
        Case "阵列新增节点", "FORNODEADD"
            阵列新增节点 Val(cT(1)), Val(cT(2)), Val(cT(3)), Val(cT(4)), Val(cT(5)), Val(cT(6)), cT(7), cT(8), cT(9), Val(cT(10)), Val(cT(11))
            GoTo Success
        Case "显示鼠标坐标", "VISMOUSEPOS"
            If cT(1) = "1" Or cT(1) = "TRUE" Then
                NoteControlDesk.CDMouseUpdataTimer.Enabled = True
            Else
                NoteControlDesk.CDMouseUpdataTimer.Enabled = False
                NoteControlDesk.Caption = 控制台名字
            End If
            GoTo Success
        Case "vbCrLf", "SELFIM"
            oneselfAddI = Val(cT(1))
            oneselfAddX = Val(cT(2))
            oneselfAddY = Val(cT(3))
            GoTo Success
        Case "字典项增加", "DICITEMADD"
            字典项增加 cT(1)
            CMD_LineExecute = vbCrLf & "当前字典大小：" & userDic.Count
            Exit Function
        Case "字典项清空", "DICREMOVEALL"
            userDic.RemoveAll
            GoTo Success
        Case "打印字典", "PRINTDIC"
            CMD_LineExecute = vbCrLf & 字典打印(userDic)
            Exit Function
        Case "打印撤销列表", "PRINTREVOKE"
            CMD_LineExecute = vbCrLf & Join(behaviorList, vbCrLf)
            Exit Function
        Case "打印重做列表", "PRINTREDO"
            CMD_LineExecute = vbCrLf & Join(redolist, vbCrLf)
            Exit Function
        Case "设置树状文本导入位置控制常数", "SETTREETXTINPOSCONTROLCONST", "STTIPCC", "STIPC"
            treeTxtToNtx_StartX = Val(cT(1))
            treeTxtToNtx_StartY = Val(cT(2))
            treeTxtToNtx_StepX = Val(cT(3))
            treeTxtToNtx_StepY = Val(cT(4))
            GoTo Success
        Case "设置位图导入位置控制常数", "SETIMAGEINPOSCONTROLCONST", "SIIPCC", "SIPC"
            imageToNtx_StartX = Val(cT(1))
            imageToNtx_StartY = Val(cT(2))
            imageToNtx_StepX = Val(cT(3))
            imageToNtx_StepY = Val(cT(4))
            GoTo Success
        Case "矩线颜色", "RECTANGLECOLOR", "RECCOLOR"
            If UBound(cT) > 2 Then
                rectangleLineColor = RGB(Val(cT(1)), Val(cT(2)), Val(cT(3)))
            Else
                rectangleLineColor = Val(cT(1))
            End If
            GoTo Success
        Case "节点放缩", "NODEZOOM"
            节点放缩 cT(1), Val(cT(2)), Val(cT(3))
            GoTo Success
        Case "创建节点", "NEWBUILTNODE", "NBN"
            节点创建 Val(cT(1)), Val(cT(2)), cT(3), cT(4), Val(cT(5)), Val(cT(6)), cT(7)
            GoTo Success
        Case "编辑节点", "EDITNODE", "EN"
            编辑节点 Val(cT(1)), cT(2), cT(3), Val(cT(4)), Val(cT(5))
            GoTo Success
        Case "位移节点", "MOVENODE", "MN"
            位移节点 Val(cT(1)), Val(cT(2)), Val(cT(3))
            GoTo Success
        Case "删除节点", "DELETENODE", "DN"
            删除节点 cT(1)
            GoTo Success
        Case "选中节点", "SELECTNODE", "SN"
            选中节点 cT(1)
            GoTo Success
        Case "创建连接", "NEWBUILTNODE", "NBL"
            连接创建 Val(cT(1)), Val(cT(2)), cT(3), Val(cT(4)), cT(5)
            GoTo Success
        Case "编辑连接内容", "EDITLINE", "EL"
            编辑连接 Val(cT(1)), Val(cT(2)), cT(3), Val(cT(4))
            GoTo Success
        Case "选中连接", "SELECTLINE", "SL"
            选中连接 cT(1)
            GoTo Success
        Case "设置动作更新速度", "SETACTIONUPDATASPEED", "SAUS"
            设置动作更新速度 Val(cT(1))
            GoTo Success
        Case "启动动作时钟", "STARTACTIONTIMER", "SAT"
            启动动作时钟 cT(1)
            GoTo Success
        Case "定义动作", "DEFINEACTION", "DEFA", "DA"
            定义动作 cT(1)
            GoTo Success
        Case "重启动作", "RESTARTACTION", "RA"
            CMD_LineExecute = 重启动作(cT(1))
            Exit Function
        Case "关闭动作", "OFFACTION", "OA"
            CMD_LineExecute = 关闭动作(cT(1))
            Exit Function
        Case "打印动作列表", "PRINTACTIONLIST", "PAL"
            CMD_LineExecute = 打印动作列表
            Exit Function
        Case "打印可执行动作列表", "PRINTEXECUTABLEACTIONLIST", "PEAL"
            CMD_LineExecute = 打印可执行动作列表
            Exit Function
        Case "重置镜头位置", "RESETLENSPOSITION", "PLP"
            angleOfView.X = Val(cT(1))
            angleOfView.Y = Val(cT(2))
            MainCoordinateSystemDefinition
            GoTo Success
        Case "色彩链路表修改", "COLORLINKDICMOD", "CLDM"
            色彩链路字典修改 cT(1)
            GoTo Success
        Case "色彩链路表重置", "COLORLINKDICRESET", "CLDS"
            色彩链路初始化
            GoTo Success
        Case "打印色彩链路表", "PRINTCOLORLINKDIC", "PCLD"
            CMD_LineExecute = 色彩链路字典导出
            Exit Function
    End Select
    CMD_LineExecute = "未知命令！"
Exit Function
Success:
    CMD_LineExecute = "命令执行成功！"
End Function
Public Sub 色彩链路字典修改(s As String)
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
Public Function 色彩链路初始化()
    colorLinkDic.RemoveAll
    colorLinkDic.Add &HFFBF00, &HC000C0
    colorLinkDic.Add &HC000C0, &HFF&
    colorLinkDic.Add &HFF&, &H80FF&
    colorLinkDic.Add &H80FF&, &HFFFF&
    colorLinkDic.Add &HFFFF&, &HFF00&
    colorLinkDic.Add &HFF00&, &HFFFF00
    colorLinkDic.Add &HFFFF00, &HFFBF00
End Function
Public Function 色彩链路字典导出() As String
    Dim i As Long
    For i = 0 To colorLinkDic.Count - 1
        色彩链路字典导出 = 色彩链路字典导出 & colorLinkDic.Keys(i) & ":" & colorLinkDic.Items(i) & ","
    Next
End Function
Private Function 打印可执行动作列表() As String
    Dim i As Long, j As Long
    打印可执行动作列表 = "定义动作 "
    For i = 1 To UBound(actionList)
        With actionList(i)
            打印可执行动作列表 = 打印可执行动作列表 & .name & ","
            For j = 0 To UBound(.nodeID)
                打印可执行动作列表 = 打印可执行动作列表 & .nodeID(j) & "|"
            Next
            打印可执行动作列表 = Mid(打印可执行动作列表, 1, Len(打印可执行动作列表) - 1) & "," & .interval & "," & .route
            Select Case .route
                Case "直线"
                    打印可执行动作列表 = 打印可执行动作列表 & "," & .vector.X & "," & .vector.Y
                Case "圆周"
                    打印可执行动作列表 = 打印可执行动作列表 & "," & .angle & "," & .aimNode
            End Select
            打印可执行动作列表 = 打印可执行动作列表 & "," & .endTime & "," & .repeat & vbCrLf
        End With
    Next
End Function
Private Function 打印动作列表() As String
    Dim i As Long, j As Long
    For i = 1 To UBound(actionList)
        With actionList(i)
            打印动作列表 = 打印动作列表 & "动作名(" & .name & "),"
            For j = 0 To UBound(.nodeID)
                打印动作列表 = 打印动作列表 & "节点(" & .nodeID(j) & ")|"
            Next
            打印动作列表 = Mid(打印动作列表, 1, Len(打印动作列表) - 1) & ",动作时间执行间隔(" & .interval & "),动作类型(" & .route & ")"
            Select Case .route
                Case "直线"
                    打印动作列表 = 打印动作列表 & ",向量X(" & .vector.X & "),向量Y(" & .vector.Y & ")"
                Case "圆周"
                    打印动作列表 = 打印动作列表 & ",角度(" & .angle & "),中心节点ID(" & .aimNode & ")"
            End Select
            打印动作列表 = 打印动作列表 & ",动作次数(" & .endTime & "),是否循环(" & .repeat & ")" & vbCrLf
        End With
    Next
End Function
Private Function 启动动作时钟(c As String)
    Note.ActionTimer.Enabled = 字符串转布尔值(c)
End Function
Private Function 设置动作更新速度(interval As Long)
    Note.ActionTimer.interval = interval
End Function
Private Function 编辑连接(nS As Long, nT As Long, content As String, size As Single)
    Dim i As Long
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                If .Source = nS And .target = nT Then
                    .content = 空格转义(content)
                    .size = size
                    Exit Function
                End If
            End If
        End With
    Next
End Function
Private Function 选中连接(allNid As String)
    Dim allNidTemp() As String, i As Long, temp() As String
    allNidTemp = Split(allNid, ",")
    For i = 0 To UBound(allNidTemp)
        If allNidTemp(i) <> "" Then
            temp = Split(allNidTemp(i), ":")
            nodeLine(LineAdd_RepeatedChecking(Val(temp(0)), Val(temp(1)))).select = True
        End If
    Next
End Function
Private Function 连接创建(nidA As Long, nidB As Long, content As String, size As Single, pichOn As String)
    LineAdd nidA, nidB, 空格转义(content), size, 字符串转布尔值(pichOn), False
End Function
Private Function 选中节点(allNid As String)
    Dim allNidTemp() As String, i As Long
    allNidTemp = Split(allNid, ",")
    For i = 0 To UBound(allNidTemp)
        If allNidTemp(i) <> "" Then
            node(Val(allNidTemp(i))).select = True
        End If
    Next
End Function
Private Function 删除节点(nid As String)
    Dim allNidTemp() As String, i As Long
    allNidTemp = Split(nid, ",")
    For i = 0 To UBound(allNidTemp)
        If allNidTemp(i) <> "" Then
            NodeDelete Val(allNidTemp(i))
        End If
    Next
End Function
Private Function 位移节点(nid As Long, X As Single, Y As Single)
    With node(nid)
        .X = X
        .Y = Y
    End With
End Function
Private Function 编辑节点(nid As Long, t As String, content As String, color As Long, size As Single)
    NodeEdit_ReviseNode nid, 空格转义(t), 空格转义(content), color, size
End Function
Public Function 字符串转布尔值(s As String) As Boolean
    If s = "1" Then
        字符串转布尔值 = True
    End If
End Function
Private Function 空格转义(s As String) As String
    空格转义 = Replace(s, "\_", " ")
End Function
Private Function 节点创建(X As Single, Y As Single, t As String, content As String, color As Long, size As Single, pichOn As String)
    NodeEdit_NewNode 空格转义(t), 空格转义(content), color, size, X, Y, 字符串转布尔值(pichOn)
End Function
Private Function 节点名序号索引(nodeName As String) As Long
    Dim i As Long
    节点名序号索引 = -1
    For i = 0 To nSum
        With node(i)
            If .b Then
                If nodeName = .t And .select = True Then
                    节点名序号索引 = i
                    Exit Function
                End If
            End If
        End With
    Next
End Function
Private Sub 节点放缩(centre As String, zoomX As Single, zoomY As Single)
    Dim i As Long, iLock As Boolean, iX As Single, iY As Single, rootNid As Long
    rootNid = 节点名序号索引(centre)
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
Private Function 字典打印(dic As Dictionary) As String
    Dim i As Long
    For i = 0 To dic.Count - 1
        字典打印 = 字典打印 & dic.Keys(i) & ":" & dic.Items(i) & vbCrLf
    Next
End Function

Private Sub 字典项增加(dS As String)
    Dim dSTmp() As String, dTmp() As String
    dSTmp = Split(dS, ",")
    For i = 0 To UBound(dSTmp)
        If dSTmp(i) <> "" Then
            dTmp = Split(dSTmp(i), ":")
            If Not userDic.Exists(dTmp(0)) Then userDic.Add dTmp(0), dTmp(1)
        End If
    Next
End Sub

Private Sub 阵列新增节点(xS As Single, xStep As Single, xE As Single, yS As Single, yStep As Single, yE As Single, t As String, c As String, p As String, size As Single, color As Long)
    Dim X As Long, Y As Long, m As Single, n As Single, pL As Boolean, dN As Long
    If p = "1" Or UCase(p) = "TRUE" Then
        pL = True
    End If
    For X = 1 To xE
        For Y = 1 To yE
            dN = dN + 1
            NodeEdit_NewNode 字典替换(t, dN, X, Y), 字典替换(c, dN, X, Y), color, size, xStep * (X - 1) + xS, yStep * (Y - 1) + yS, pL
        Next
    Next
End Sub

Private Function 字典替换(ByVal s As String, dN As Long, dX As Long, dY As Long) As String
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
    字典替换 = s
End Function
