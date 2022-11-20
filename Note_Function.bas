Attribute VB_Name = "Note_Function"
Public 放缩率需要提示提示倒计时 As Long

Public Function NodePositionVague(vT As Single, Optional allNode As Boolean)
    Dim i As Long
    For i = 0 To nSum
        With node(i)
            If .b Then
                If .select = True Or allNode = True Then
                    .X = (.X \ vT) * vT
                    .Y = (.Y \ vT) * vT
                End If
            End If
        End With
    Next
End Function

Public Function ConnectionReversal() As Boolean
    Dim i As Long
    BehaviorIdSet
    For i = 0 To lSum
        With nodeLine(i)
            If .b = True And (.select = True Or lineTargetAim = i) Then
                LineDelete i
                LineAdd .target, .Source, .content, .size
            End If
        End With
    Next
End Function
Public Function 限制数值(ByVal num As Double, ByVal min As Double, ByVal max As Double) As Double
    If num > max Then num = max
    If num < min Then num = min
    限制数值 = num
End Function
Public Function MultipointConnection() As Boolean
Dim i As Long
If MultipointConnection_Check = True And nodeClickAim <> -1 Then
    MultipointConnection = True
    BehaviorIdSet
    For i = 0 To nSum
        With node(i)
            If .b = True And .select = True And i <> nodeClickAim Then
                LineAdd i, nodeClickAim, "", lineDefaultSize
            End If
        End With
    Next
End If
End Function
Public Function MultipointConnection_Check() As Boolean
Dim i As Long
For i = 0 To nSum
    With node(i)
        If .b = True And .select = True Then
            MultipointConnection_Check = True
            Exit Function
        End If
    End With
Next
End Function
Public Function NoteGlobalViewSet(ByRef bool As Boolean)
Note.全局视图.Checked = bool
Note.GlobalView.Visible = bool
End Function
Public Function IconSet(ByRef formObj As Form)
On Error GoTo Er:
formObj.Icon = LoadPicture(App.path & "\note.ico")
formObj.Show
Exit Function
Er:
MsgBox "出错了！" & vbCrLf & "错误编号：" & Err.Number & " 错误描述：" & Err.Description, , "警告"
End Function
Public Function SelectDisplayObjcet()
If Note.显示全部节点名.Checked = True Then
    AllSelection
ElseIf Note.显示顺向节点名.Checked = True Then
    SelectDisplayObjcet_Forward True
Else
    SelectDisplayObjcet_Forward False
End If
End Function
Public Function SelectDisplayObjcet_Forward(ByRef forward As Boolean)
Dim i As Long
For i = 0 To nSum
    With node(i)
        If .b = True Then
            If forward = True Then
                .select = .forward
            Else
                .select = .backward
            End If
        End If
    End With
Next
For i = 0 To lSum
    With nodeLine(i)
        If .b = True Then
            If forward = True Then
                .select = .forward
            Else
                .select = .backward
            End If
        End If
    End With
Next
End Function
Public Function FindNode(ByRef findStr As String, ByRef capsLook As Boolean, ByRef selectMode As Boolean, ByRef newNoteOutput As Boolean)
Dim i, fNSum As Long: Dim tStr As String, contentText As String, findStrCase As String
For i = 0 To nSum
    With node(i)
        If .b = True And ((selectMode = True And .select = True) Or selectMode = False) Then
            If capsLook = True Then
                NodeFind.nodeTmp.TextRTF = .content
                contentText = NodeFind.nodeTmp.text
                tStr = .t
                findStrCase = findStr
            Else
                NodeFind.nodeTmp.TextRTF = .content
                contentText = UCase(NodeFind.nodeTmp.text)
                tStr = UCase(.t)
                findStrCase = UCase(findStr)
            End If
            If InStr(1, tStr, findStrCase) Or InStr(1, contentText, findStrCase) Then
                .select = True: fNSum = fNSum + 1
            Else
                .select = False
            End If
        End If
    End With
Next
If newNoteOutput = True And fNSum > 0 Then
    FindNode_NewNoteOutput fNSum, findStr
End If
End Function
Public Function FindNode_NewNoteOutput(ByRef fNSum As Long, ByRef findStr As String)
Dim tempFilePath As String: Dim i, j As Long: Dim angle As Single, filename As Integer
Dim fN() As 节点: Dim fL() As 连接: Dim fLSum As Long: Dim ntx
'tempFilePath = App.Path & "\" & App.EXEName & "_FindTemp.ntx"
tempFilePath = ntxPath & "~FT.ntx"
fLSum = fNSum
angle = 2 * PI / fNSum
ReDim fL(fLSum)
fNSum = fNSum + 1
ReDim fN(fNSum)
j = 1
For i = 0 To nSum
    With node(i)
        If .b = True And .select = True Then
            fN(j).b = True
            fN(j).t = .t
            fN(j).content = .content
            fN(j).setSize = .setSize
            fN(j).setColor = .setColor
            FindNode_NewNoteOutput_CircularArray fN(j), angle * (j - 1), fNSum
            fL(j - 1).b = True
            fL(j - 1).Source = 0
            fL(j - 1).target = j
            fL(j - 1).size = lineDefaultSize
            j = j + 1
        End If
    End With
Next
With fN(0)
    .b = True
    .t = findStr
    .setSize = nodeDefaultSize
    .setColor = nodeDefaultColor
    .X = Note.width / 2
    .Y = Note.height / 2
End With
ntx = NoteFileWrite_204_Coding(fN, fNSum, fL, fLSum)
filename = FreeFile
Open tempFilePath For Output As #filename
    For i = 0 To UBound(ntx)
        Print #filename, ntx(i)
    Next
Close #filename
'Shell "C:\ProgramData\Note\Note2D.exe " & tempFilePath, vbNormalFocus
Shell """" & App.path & "\" & App.EXEName & ".exe"" " & tempFilePath, vbNormalFocus
Kill tempFilePath
End Function
Public Sub NodeListVisUpdata()
    Dim i As Long
    NodeListVis.节点列表.Clear
    For i = 0 To nSum
        If node(i).b Then
            NodeListVis.节点列表.AddItem i & " - [" & node(i).t & "]：" & 富文本转义(node(i).content)
        End If
    Next
End Sub
Public Sub LineListVisUpdata()
    Dim i As Long
    LineListVis.连接列表.Clear
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                LineListVis.连接列表.AddItem i & " - [" & node(.Source).t & "](" & .Source & ")-[" & node(.target).t & "](" & .target & ")：" & .content
            End If
        End With
    Next
End Sub
Public Function 富文本转义(s As String) As String
    NodeListVis.转义文本.TextRTF = s
    富文本转义 = NodeListVis.转义文本.text
End Function
Public Function 转为富文本(s As String) As String
    NodeListVis.转义文本.text = s
    转为富文本 = NodeListVis.转义文本.TextRTF
End Function
Public Function RollerEventHandling(ByRef narrow As Boolean)
    Dim oldZF As Single, mousePrimaryPos As 三维坐标
    mousePrimaryPos = mouseV3Pos
    oldZF = zoomFactor
    If narrow = True Then
        If magnification < 4 Then magnification = magnification + 0.5
    Else
        If magnification > -4 Then magnification = magnification - 0.5
    End If
    zoomFactor = MToZF(magnification)
    MainCoordinateSystemDefinition
'    MainCoordinateSystemReduction mousePrimaryPos, oldZF
    放缩率需要提示提示倒计时 = 2
End Function
Public Function MainCoordinateSystemDefinition()
'    Note.Scale (-angleOfView.X, Note.height * zoomFactor - angleOfView.Y)-(Note.width * zoomFactor - angleOfView.X, -angleOfView.Y)
    Note.Scale (-angleOfView.X, -angleOfView.Y)-(Note.width * zoomFactor - angleOfView.X, Note.height * zoomFactor - angleOfView.Y)
End Function
Public Function MainCoordinateSystemReduction(mousePrimaryPos As 三维坐标, oldZF As Single)
'    angleOfView.X = mousePrimaryPos.X / mousePrimaryPos.z * (zoomFactor - oldZF)
'    angleOfView.Y = mousePrimaryPos.Y / mousePrimaryPos.z * (zoomFactor - oldZF)
    angleOfView.X = mousePrimaryPos.X * (zoomFactor - 1)
    angleOfView.Y = mousePrimaryPos.Y * (zoomFactor - 1)
    MainCoordinateSystemDefinition
End Function
Public Function FindNode_NewNoteOutput_CircularArray(ByRef arrayObj As 节点, ByRef arrAngle As Single, ByRef fNSum As Long)
If arrAngle > PI / 2 Then
    arrAngle = PI - arrAngle
    arrayObj.X = 300 * fNSum * -Cos(arrAngle) + Note.width / 2
Else
    arrayObj.X = 300 * fNSum * Cos(arrAngle) + Note.width / 2
End If
If NodeFind.圆形阵列.Checked = True Then
    arrayObj.Y = 300 * fNSum * Sin(arrAngle) + Note.height / 2
Else
    arrayObj.Y = 600 * fNSum * Sin(arrAngle) + Note.height / 2
End If
End Function

Public Function BehaviorListAdd(ByRef functionName As String, ParamArray fArr())
    Dim i As Variant
    If behaviorId = "" Then Exit Function '行为无集ID，退出行为记录
    behaviorList(bHLSum) = behaviorId & "," & functionName
    For Each i In fArr
        behaviorList(bHLSum) = behaviorList(bHLSum) & "," & i
    Next
    bHLSum = bHLSum + 1
    BehaviorListUboundAdd
End Function
Public Function RedoListAdd(ByRef functionName As String, ParamArray fArr())
    Dim i As Variant
    If redoId = "" Then Exit Function '行为无集ID，退出行为记录
    redolist(redoSum) = redoId & "," & functionName
    For Each i In fArr
        redolist(redoSum) = redolist(redoSum) & "," & i
    Next
    redoSum = redoSum + 1
    RedoListUboundAdd
End Function
Public Function BehaviorListUboundAdd()
    If UBound(behaviorList) < bHLSum + 100 Then
        ReDim Preserve behaviorList(bHLSum + 1000)
    End If
End Function
Public Function RedoListUboundAdd()
    If UBound(redolist) < redoSum + 100 Then
        ReDim Preserve redolist(redoSum + 1000)
    End If
End Function
Public Function CopyObject(ByRef delSoure As Boolean)
    If CopyObject_BeCheck = True Then
        CopyObject_Node delSoure
        CopyObject_Line delSoure
        CopyObject_Coding
        DeselectObjcet
    End If
End Function
Public Function CopyObject_BeCheck() As Boolean
    Dim i As Long
    For i = 0 To nSum
        With node(i)
            If .b = True Then
                If .select = True Or i = nodeTargetAim Then
                    CopyObject_BeCheck = True: Exit Function
                End If
            End If
        End With
    Next
End Function

Public Function PasteObject()
    Dim pasteStr As String: Dim listStr
    Dim ntx() As String
    Dim startNodeId, version As Long
    On Error GoTo Er
    pasteStr = Clipboard.GetText
    listStr = Split(pasteStr, COPYLINEBREAK)
    version = PasteObject_NtxFileCheck(listStr(1))
    If listStr(0) = meExeId Then
        startNodeId = nSum
        PasteObject_Local_Node
        PasteObject_Local_Line startNodeId
        Exit Function
    End If
    Select Case version
        Case 201
            ntx = PasteObject_GetNtx(listStr)
            NoteFileRead_201 ntx, True
        Case 202
            ntx = PasteObject_GetNtx(listStr)
            NoteFileRead_202 ntx, True
        Case 203
            ntx = PasteObject_GetNtx(listStr)
            NoteFileRead_203 ntx, True
        Case Else
            GoTo Er
    End Select
    Exit Function
Er:
    NodeEdit_NewNode "", pasteStr, nodeDefaultColor, nodeDefaultSize, mousePos.X, mousePos.Y
End Function
Public Function PasteObject_GetNtx(ByRef listStr)
    Dim i As Long, ntx() As String
    ReDim ntx(UBound(listStr) - 1)
    For i = 1 To UBound(listStr)
        ntx(i - 1) = listStr(i)
    Next
    PasteObject_GetNtx = ntx
End Function
Public Function PasteObject_NtxFileCheck(ByVal linStr As String) As Long
    If InStr(1, linStr, VERSIONID) Then
        PasteObject_NtxFileCheck = 203
    ElseIf InStr(1, linStr, "Note2D_2") Then
        PasteObject_NtxFileCheck = 202
    Else
        PasteObject_NtxFileCheck = -1
    End If
End Function
Public Function PasteObject_Local_Node()
    Dim i As Long: Dim firstPos As 二维坐标
    firstPos.X = node(copyNIdList(0)).X: firstPos.Y = node(copyNIdList(0)).Y
    For i = 0 To copyNSum - 1
        With node(copyNIdList(i))
            NodeEdit_NewNode .t, .content, .setColor, .setSize, .X - firstPos.X + mousePos.X, .Y - firstPos.Y + mousePos.Y, True
        End With
    Next
End Function
Public Function PasteObject_Local_Line(ByVal startNodeId As Long)
    Dim i As Long
    For i = 0 To copyLSum - 1
        With copyLineList(i)
            LineAdd .Source + startNodeId, .target + startNodeId, .content, .size, True
        End With
    Next
End Function
Public Function CopyObject_Coding()
    Dim ntx: Dim copyStr As String: Dim i As Long
    'ReDim ntx(copyNSum + copyLSum + 1)
    ntx = NoteFileWrite_204_Coding(copyNodeList, copyNSum, copyLineList, copyLSum)
    copyStr = meExeId & COPYLINEBREAK
    For i = 0 To copyNSum + copyLSum
        copyStr = copyStr & ntx(i) & COPYLINEBREAK
    Next
    Clipboard.Clear
    Clipboard.SetText copyStr
End Function
Public Function CopyObject_Line(ByRef delSoure As Boolean)
    Dim i As Long
    ReDim copyLineList(lSum): ReDim copyLIdList(lSum)
    copyLSum = 0
    For i = 0 To lSum
        With nodeLine(i)
            If .b = True And .select = True Then
                copyLIdList(copyLSum) = i
                copyLineList(copyLSum).b = True
                copyLineList(copyLSum).Source = CopyObject_Line_GetNodeRelativityId(.Source)
                copyLineList(copyLSum).target = CopyObject_Line_GetNodeRelativityId(.target)
                copyLineList(copyLSum).content = .content
                copyLineList(copyLSum).size = .size
                copyLSum = copyLSum + 1
                If delSoure = True Then .b = False
            End If
        End With
    Next
End Function
Public Function CopyObject_Node(ByRef delSoure As Boolean)
    Dim firstPos As 二维坐标, i As Long
    ReDim copyNodeList(nSum): ReDim copyNIdList(nSum)
    copyNSum = 0
    For i = 0 To nSum
        With node(i)
            If .b = True Then
                If .select = True Or i = nodeTargetAim Then
                    copyNIdList(copyNSum) = i
                    If copyNSum = 0 Then
                        firstPos.X = .X: firstPos.Y = .Y
                        copyNodeList(copyNSum).X = 0
                        copyNodeList(copyNSum).Y = 0
                    Else
                        copyNodeList(copyNSum).X = .X - firstPos.X
                        copyNodeList(copyNSum).Y = .Y - firstPos.Y
                    End If
                    copyNodeList(copyNSum).b = True
                    copyNodeList(copyNSum).t = .t
                    copyNodeList(copyNSum).content = .content
                    copyNodeList(copyNSum).setColor = .setColor
                    copyNodeList(copyNSum).setSize = .setSize
                    copyNSum = copyNSum + 1
                    If delSoure = True Then NodeDelete i
                End If
            End If
        End With
    Next
End Function
Public Function CopyObject_Line_GetNodeRelativityId(ByRef nId As Long) As Long
    Dim i As Long
    For i = 0 To copyNSum
        If copyNIdList(i) = nId Then
            CopyObject_Line_GetNodeRelativityId = i: Exit Function
        End If
    Next
End Function
Public Function MeExeIdSet()
    Dim i As Long
    Randomize
    meExeId = ""
    For i = 0 To 9
        meExeId = meExeId & Int(Rnd * 10)
    Next
End Function
Public Function BehaviorIdSet()
    Dim i As Long
    Randomize
    behaviorId = ""
    For i = 0 To 9
        behaviorId = behaviorId & Int(Rnd * 10)
    Next
End Function
Public Function RedoSet()
    Dim i As Long
    Randomize
    For i = 0 To 9
        redoId = redoId & Int(Rnd * 10)
    Next
End Function
Public Function DeleteSelectObjcet()
    Dim i As Long
    If nodeTargetAim <> -1 Then
        NodeDelete nodeTargetAim
    End If
    For i = 0 To nSum
        With node(i)
            If .b = True And .select = True Then
                NodeDelete i
            End If
        End With
    Next
    For i = 0 To lSum
        With nodeLine(i)
            If .b = True And .select = True Then
                LineDelete i
            End If
        End With
    Next
End Function
Public Function DeselectObjcet()
    Dim i As Long
    For i = 0 To nSum
        With node(i)
            If .b = True Then .select = False
        End With
    Next
    For i = 0 To lSum
        With nodeLine(i)
            If .b = True Then .select = False
        End With
    Next
End Function
Public Function ChainSelection(ByRef nId As Long, ByRef selectMode As Long)
    Select Case selectMode
        Case 0
            ChainSelection_All nId
    End Select
End Function
Public Function DirectSelect()
    Dim aim As Long, i As Long
    aim = NodeCheck(mousePos.X, mousePos.Y)
    If aim = -1 Then Exit Function
    node(aim).select = True
    For i = 0 To lSum
        With nodeLine(i)
            If .b = True Then
                If .Source = aim Then
                    .select = True
                    node(.target).select = True
                End If
                If .target = aim Then
                    .select = True
                    node(.Source).select = True
                End If
            End If
        End With
    Next
End Function
Public Function AllSelection()
    Dim i As Long
    For i = 0 To nSum
        With node(i)
            If .b = True Then .select = True
        End With
    Next
    For i = 0 To lSum
        With nodeLine(i)
            If .b = True Then .select = True
        End With
    Next
End Function

Public Function ChainSelection_All(ByRef nId As Long) '关联选区
    Dim i As Long
    node(nId).select = True
    For i = 0 To lSum
        With nodeLine(i)
            If .b = True And .search = False Then
                .search = True
                If .Source = nId Then ChainSelection_All .target: .select = True
                If .target = nId Then ChainSelection_All .Source: .select = True
                .search = False
            End If
        End With
    Next
End Function
