Attribute VB_Name = "Note_Function"
Public Function ConnectionReversal() As Boolean
Dim i As Long
BehaviorIdSet
For i = 0 To lSum
    With nodeLine(i)
        If .b = True And .select = True Then
            LineDelete i
            LineAdd .target, .source
        End If
    End With
Next
End Function
Public Function MultipointConnection() As Boolean
Dim i As Long
If MultipointConnection_Check = True And nodeClickAim <> -1 Then
    MultipointConnection = True
    BehaviorIdSet
    For i = 0 To nSum
        With node(i)
            If .b = True And .select = True And i <> nodeClickAim Then
                LineAdd i, nodeClickAim
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
Public Function NodesToTxt(ByRef outPath As String)
Dim i  As Long, startNid As Long, outTxt As String
startNid = NodesToTxt_GetStartNid
If startNid <> -1 Then
    Updata_GetNodeTargetAim_Select startNid
    outTxt = NodesToTxt_Coding
    Open outPath For Output As #8
        Print #8, outTxt
    Close #8
End If
End Function
Public Function NodesToTxt_Coding() As String
Dim i As Long
'dim j As Long, c As Long, dMax As Long, nMax As Long, tOrder As String
'dMax = NodesToTxt_Coding_GetDepthMax
'nMax = NodesToTxt_Coding_GetDepthNodeMax(dMax)
'ReDim depthList(dMax, nMax)
For i = 0 To nSum
    With node(i)
        If .b = True And .forward = True Then
'            c = .depth
'            j = NodesToTxt_Coding_GetNodeOrder(c, nMax)
'            tOrder = c + 1 & "." & j + 1 & " "
            NodePrint.NodePrintBox.TextRTF = .content
'            NodesToTxt_Coding = NodesToTxt_Coding & tOrder & .t & vbCrLf & NodePrint.NodePrintBox.Text & vbCrLf
            NodesToTxt_Coding = NodesToTxt_Coding & .t & vbCrLf & NodePrint.NodePrintBox.Text & vbCrLf
        End If
    End With
Next
End Function
Public Function NodesToTxt_Coding_GetNodeOrder(ByRef depth As Long, ByRef nMax As Long) As Long
Dim i As Long
For i = 0 To nMax
    If depthList(depth, i) = False Then
        depthList(depth, i) = True: NodesToTxt_Coding_GetNodeOrder = i: Exit Function
    End If
Next
End Function
Public Function NodesToTxt_Coding_GetDepthNodeMax(ByRef dMax As Long) As Long
Dim i As Long, sum() As Long, nMax As Long
ReDim sum(dMax) As Long
For i = 0 To nSum
    With node(i)
        If .b = True And .forward = True Then
            sum(.depth) = sum(.depth) + 1
        End If
    End With
Next
For i = 0 To dMax
    If sum(i) > nMax Then nMax = sum(i): NodesToTxt_Coding_GetDepthNodeMax = i
Next
End Function
Public Function NodesToTxt_Coding_GetDepthMax() As Long
Dim i As Long
For i = 0 To nSum
    With node(i)
        If .b = True And .forward = True Then
            If .depth > NodesToTxt_Coding_GetDepthMax Then NodesToTxt_Coding_GetDepthMax = .depth
        End If
    End With
Next
End Function
Public Function NodesToTxt_GetStartNid() As Long
NodesToTxt_GetStartNid = -1
For i = 0 To nSum
    With node(i)
        If .b = True And .select Then NodesToTxt_GetStartNid = i: Exit Function
    End With
Next
End Function
Public Function NoteGlobalViewSet(ByRef bool As Boolean)
Note.全局视图.Checked = bool
Note.GlobalView.Visible = bool
End Function
Public Function IconSet(ByRef formObj As Form)
On Error GoTo Er:
formObj.Icon = LoadPicture(App.Path & "\note.ico")
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
Dim i, fNSum As Long: Dim tStr, contentText, findStrCase As String
For i = 0 To nSum
    With node(i)
        If .b = True And ((selectMode = True And .select = True) Or selectMode = False) Then
            If capsLook = True Then
                NodeFind.nodeTmp.TextRTF = .content
                contentText = NodeFind.nodeTmp.Text
                tStr = .t
                findStrCase = findStr
            Else
                NodeFind.nodeTmp.TextRTF = .content
                contentText = UCase(NodeFind.nodeTmp.Text)
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
Dim tempFilePath As String: Dim i, j As Long: Dim angle As Single
Dim fN() As 节点: Dim fL() As 连接: Dim fLSum As Long: Dim ntx
tempFilePath = App.Path & "\" & App.EXEName & "_FindTemp.ntx"
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
            FindNode_NewNoteOutput_CircularArray fN(j), angle * (j - 1), fNSum
            fL(j - 1).b = True
            fL(j - 1).source = 0
            fL(j - 1).target = j
            j = j + 1
        End If
    End With
Next
With fN(0)
    .b = True
    .t = findStr
    .x = Note.Width / 2
    .y = Note.Height / 2
End With
ntx = NoteFileWrite_201_Coding(fN, fNSum, fL, fLSum)
Open tempFilePath For Output As #4
    For i = 0 To UBound(ntx)
        Print #4, ntx(i)
    Next
Close #4
Shell "C:\ProgramData\Note\Note2D.exe " & tempFilePath, vbNormalFocus
Kill tempFilePath
End Function
Public Function RollerEventHandling(ByRef narrow As Boolean)
Dim oldZF As Single: Dim mousePrimaryPos As 三维坐标
mousePrimaryPos = mouseV3Pos
oldZF = zoomFactor
'MainCoordinateSystemZero mousePrimaryPos
If narrow = True Then
'    If zoomFactor < 100 Then zoomFactor = zoomFactor * 1.1
    If magnification < 4 Then magnification = magnification + 0.5
Else
'    If zoomFactor > 0.01 Then zoomFactor = zoomFactor * 0.9
    If magnification > -4 Then magnification = magnification - 0.5
End If
zoomFactor = MToZF(magnification)
MainCoordinateSystemDefinition
MainCoordinateSystemReduction mousePrimaryPos, oldZF

End Function
Public Function MainCoordinateSystemZero(ByRef mousePrimaryPos As 三维坐标)
Dim dx As Single: Dim dy As Single
MapUpdata_AoVMove_Moving -mousePrimaryPos.x, -mousePrimaryPos.y
End Function
Public Function MainCoordinateSystemDefinition()
Note.Scale (0, Note.Height * zoomFactor)-(Note.Width * zoomFactor, 0)
End Function
Public Function MainCoordinateSystemReduction(ByRef mousePrimaryPos As 三维坐标, ByRef oldZF As Single)
Dim dx As Single: Dim dy As Single
dx = mousePrimaryPos.x / mousePrimaryPos.z * (zoomFactor - oldZF)
dy = mousePrimaryPos.y / mousePrimaryPos.z * (zoomFactor - oldZF)
'MapUpdata_AoVMove_Moving mousePrimaryPos.x + dx, mousePrimaryPos.y + dy
MapUpdata_AoVMove_Moving dx, dy
End Function
Public Function FindNode_NewNoteOutput_CircularArray(ByRef arrayObj As 节点, ByRef arrAngle As Single, ByRef fNSum As Long)
If arrAngle > PI / 2 Then
    arrAngle = PI - arrAngle
    arrayObj.x = 300 * fNSum * -Cos(arrAngle) + Note.Width / 2
Else
    arrayObj.x = 300 * fNSum * Cos(arrAngle) + Note.Width / 2
End If
If NodeFind.圆形阵列.Checked = True Then
    arrayObj.y = 300 * fNSum * Sin(arrAngle) + Note.Height / 2
Else
    arrayObj.y = 600 * fNSum * Sin(arrAngle) + Note.Height / 2
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
    Case Else
        GoTo Er
End Select
Exit Function
Er:
NodeEdit_NewNode "", pasteStr, mousePos.x, mousePos.y
End Function
Public Function PasteObject_GetNtx(ByRef listStr)
Dim i As Long: Dim ntx() As String
ReDim ntx(UBound(listStr) - 1)
For i = 1 To UBound(listStr)
    ntx(i - 1) = listStr(i)
Next
PasteObject_GetNtx = ntx
End Function
Public Function PasteObject_NtxFileCheck(ByVal linStr As String) As Long
If InStr(1, linStr, VERSIONID) Then
    PasteObject_NtxFileCheck = 201
Else
    PasteObject_NtxFileCheck = -1
End If
End Function
Public Function PasteObject_Local_Node()
Dim i As Long: Dim firstPos As 二维坐标
firstPos.x = node(copyNIdList(0)).x: firstPos.y = node(copyNIdList(0)).y
For i = 0 To copyNSum - 1
    With node(copyNIdList(i))
        NodeEdit_NewNode .t, .content, .x - firstPos.x + mousePos.x, .y - firstPos.y + mousePos.y, True
    End With
Next
End Function
Public Function PasteObject_Local_Line(ByVal startNodeId As Long)
Dim i As Long
For i = 0 To copyLSum
    With copyLineList(i)
        LineAdd .source + startNodeId, .target + startNodeId, True
    End With
Next
End Function
Public Function CopyObject_Coding()
Dim ntx: Dim copyStr As String: Dim i As Long
'ReDim ntx(copyNSum + copyLSum + 1)
ntx = NoteFileWrite_201_Coding(copyNodeList, copyNSum, copyLineList, copyLSum)
copyStr = meExeId & COPYLINEBREAK
For i = 0 To copyNSum + copyLSum
    copyStr = copyStr & ntx(i) & COPYLINEBREAK
Next
Clipboard.Clear
Clipboard.SetText copyStr
End Function
Public Function CopyObject_Line(ByRef delSoure As Boolean)
Dim i As Long
ReDim copyLineList(lSum): ReDim copyLIdList(nSum)
copyLSum = 0
For i = 0 To lSum
    With nodeLine(i)
        If .b = True And .select = True Then
            copyLIdList(copyLSum) = i
            copyLineList(copyLSum).b = True
            copyLineList(copyLSum).source = CopyObject_Line_GetNodeRelativityId(.source)
            copyLineList(copyLSum).target = CopyObject_Line_GetNodeRelativityId(.target)
            copyLSum = copyLSum + 1
            If delSoure = True Then .b = False
        End If
    End With
Next
End Function
Public Function CopyObject_Node(ByRef delSoure As Boolean)
Dim firstPos As 二维坐标: Dim i As Long
ReDim copyNodeList(nSum): ReDim copyNIdList(nSum)
copyNSum = 0
For i = 0 To nSum
    With node(i)
        If .b = True Then
            If .select = True Or i = nodeTargetAim Then
                copyNIdList(copyNSum) = i
                If copyNSum = 0 Then
                    firstPos.x = .x: firstPos.y = .y
                    copyNodeList(copyNSum).x = 0
                    copyNodeList(copyNSum).y = 0
                Else
                    copyNodeList(copyNSum).x = .x - firstPos.x
                    copyNodeList(copyNSum).y = .y - firstPos.y
                End If
                copyNodeList(copyNSum).b = True
                copyNodeList(copyNSum).t = .t
                copyNodeList(copyNSum).content = .content
                copyNSum = copyNSum + 1
                If delSoure = True Then .b = False
            End If
        End If
    End With
Next
End Function
Public Function CopyObject_Line_GetNodeRelativityId(ByRef nid As Long) As Long
Dim i As Long
For i = 0 To copyNSum
    If copyNIdList(i) = nid Then
        CopyObject_Line_GetNodeRelativityId = i: Exit Function
    End If
Next
End Function
Public Function MeExeIdSet()
Dim i As Long
Randomize
For i = 0 To 9
    meExeId = meExeId & Int(Rnd * 10)
Next
End Function
Public Function BehaviorIdSet()
Dim i As Long
Randomize
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
Public Function ChainSelection(ByRef nid As Long, ByRef selectMode As Long)
Select Case selectMode
    Case 0
        ChainSelection_All nid
End Select
End Function
Public Function DirectSelect()
Dim aim, i As Long
aim = NodeCheck(mousePos.x, mousePos.y)
If aim = -1 Then Exit Function
node(aim).select = True
For i = 0 To lSum
    With nodeLine(i)
        If .b = True Then
            If .source = aim Then
                .select = True
                node(.target).select = True
            End If
            If .target = aim Then
                .select = True
                node(.source).select = True
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

Public Function ChainSelection_All(ByRef nid As Long) '关联选区
Dim i As Long
node(nid).select = True
For i = 0 To lSum
    With nodeLine(i)
        If .b = True And .search = False Then
            .search = True
            If .source = nid Then ChainSelection_All .target: .select = True
            If .target = nid Then ChainSelection_All .source: .select = True
            .search = False
        End If
    End With
Next
End Function
