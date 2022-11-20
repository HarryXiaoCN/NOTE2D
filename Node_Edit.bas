Attribute VB_Name = "Node_Edit"
Public Function NodeEditeStart(ByRef X As Single, ByRef Y As Single)
    nodePreviousEditAim = nodeEditAim
    nodeEditAim = NodeCheck(X, Y)
    If nodeEditAim = -1 Then '新建节点
        nodeEditLock = False
        If NodeInput.保持内容.Checked = False Then
            NodeInput.NodeTitle.text = ""
            NodeInput.NodeInputBox.text = ""
        End If
        nodeEditPos.X = X
        nodeEditPos.Y = Y
    Else
        NodeInput.NodeTitle.text = node(nodeEditAim).t
        NodeInput.NodeInputBox.TextRTF = node(nodeEditAim).content
        NodeInput.节点颜色预览_Click NodeInput.色号匹配(node(nodeEditAim).setColor)
        nodeEditLock = True
    End If
    NodeInput.Show
End Function
Public Function NodeUboundAdd()
If UBound(node) < nSum + 100 Then
    ReDim Preserve node(nSum + 1000)
End If
End Function
Public Function LineUboundAdd()
If UBound(nodeLine) < lSum + 100 Then
    ReDim Preserve nodeLine(lSum + 1000)
End If
End Function
Public Function NodeCheck(ByRef X As Single, ByRef Y As Single) As Long
    Dim i As Long
    NodeCheck = -1
    For i = 0 To nSum
        If node(i).b = True Then
            If OverlappingJudgment(node(i).size, X, Y, node(i).X, node(i).Y) = True Then
                NodeCheck = i: Exit Function
            End If
        End If
    Next
End Function
Public Function LineCheck(X As Single, Y As Single) As Long
    Dim i As Long
    LineCheck = -1
    For i = 0 To lSum
        With nodeLine(i)
            If .b = True Then
                '点在线上
                If DistanceBetweenLinePoint(node(.Source).X, node(.Source).Y, node(.target).X, node(.target).Y, X, Y, .size) Then
                    LineCheck = i: Exit Function
                End If
            End If
        End With
    Next
End Function
Public Function NodeEdit_NewNode(ByVal title As String, ByVal content As String, setC As Long, setS As Single, ByRef X As Single, ByRef Y As Single, Optional pitchOn As Boolean) As Long
    BehaviorListAdd "NodeEdit_NewNode", nSum
    NodeEdit_Save nSum, title, content, setC, setS, X, Y, pitchOn
    nodeEditLock = True
    nodeEditAim = nSum
    nSum = nSum + 1
    NodeEdit_NewNode = nodeEditAim
End Function
Public Function NodeEdit_TitleFilter(ByRef nid As Long, ByRef title As String) As String
If title = "" Then
    NodeEdit_TitleFilter = NodeEdit_TitleFilter_StrCombination
Else
    NodeEdit_TitleFilter = title
End If
End Function
Public Function NodeEdit_ContentFilter(ByRef content As String) As Boolean
If content = "请输入节点内容..." Then
    NodeEdit_ContentFilter = True
End If
End Function
Public Function NodeEdit_TitleFilter_StrCombination() As String
NodeEdit_TitleFilter_StrCombination = "node[" & NodeEdit_TitleFilter_StrCombination_GetSureNId & "]"
End Function
Public Function NodeEdit_TitleFilter_StrCombination_GetSureNId() As Long
Dim i As Long
For i = 0 To nSum
    If node(i).b = True Then NodeEdit_TitleFilter_StrCombination_GetSureNId = NodeEdit_TitleFilter_StrCombination_GetSureNId + 1
Next
End Function
Public Function NodeEdit_ReviseNode(ByRef nid As Long, ByRef title As String, ByRef content As String, setC As Long, setS As Single, Optional noDo As Boolean)
    With node(nid)
        BehaviorListAdd "NodeEdit_ReviseNode", nid, .t, .content, .setColor, .setSize
        NodeEdit_Save nid, title, content, setC, setS, .X, .Y, , noDo
    End With
End Function
Public Function NodeEdit_Save(nid As Long, title As String, content As String, setC As Long, setS As Single, X As Single, Y As Single, Optional pitchOn As Boolean, Optional noDo As Boolean)
    With node(nid)
        .b = True
        .X = X
        .Y = Y
        .size = setS
    '    .t = NodeEdit_TitleFilter(nid, title)
        .t = title
        .content = content
        .text = 富文本转义(.content)
        If noDo = False Then
            .gravitational_s_name = GetNodeGravitational(.text, "引源名(")
            .gravitational_s_text = GetNodeGravitational(.text, "引源实(")
            .gravitational_t_name = GetNodeGravitational(.text, "引去名(")
            .gravitational_t_text = GetNodeGravitational(.text, "引去实(")
            
            If manuallyEstablishedLock Then
                If .gravitational_s_name <> "" Then
                    NodeEdit_Save_NameGravityJudgment nid, .gravitational_s_name, True
                End If
                If .gravitational_s_text <> "" Then
                    NodeEdit_Save_TextGravityJudgment nid, .gravitational_s_text, True
                End If
                If .gravitational_t_name <> "" Then
                    NodeEdit_Save_NameGravityJudgment nid, .gravitational_t_name, False
                End If
                If .gravitational_t_text <> "" Then
                    NodeEdit_Save_TextGravityJudgment nid, .gravitational_t_text, False
                End If
                NodeEdit_Save_ActiveGravityDecision nid, .t, .text
            End If
        End If
        manuallyEstablishedLock = False
        .setColor = setC
        .setSize = setS
        .select = pitchOn
        If nodeEditFormLock Then
            NodeInput.NodeTitle.text = .t
            NodeInput.NodeInputBox.TextRTF = .content
        End If
    End With
    NodeUboundAdd
End Function
Public Sub GetLineDic(lineDic As Dictionary)
    Dim i As Long
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                lineDic.Add .Source & "," & .target, i
                lineDic.Add .target & "," & .Source, i
            End If
        End With
    Next
End Sub
Public Sub NodeEdit_Save_RangeJoin(nid As Long)
    Dim j As Long, lineDic As New Dictionary, sizeTmp As Single
    GetLineDic lineDic
    For j = 0 To nSum
        With node(j)
            If .b = True And j <> nid Then
                If lineDic.Exists(j & "," & nid) = False Then
                    sizeTmp = .setSize * 25
                    If .gSource Then
                        If node(nid).X >= .X - sizeTmp And node(nid).X <= .X + sizeTmp _
                        And node(nid).Y >= .Y - sizeTmp And node(nid).Y <= .Y + sizeTmp Then
                            LineAdd j, nid, "", lineDefaultSize, , True, True
                        End If
                    End If
                    If .gTarget Then
                        If node(nid).X >= .X - sizeTmp And node(nid).X <= .X + sizeTmp _
                        And node(nid).Y >= .Y - sizeTmp And node(nid).Y <= .Y + sizeTmp Then
                            LineAdd nid, j, "", lineDefaultSize, , True, True
                        End If
                    End If
                End If
            End If
        End With
    Next
End Sub
Public Sub NodeEdit_Save_ActiveGravityDecision(nid As Long, beLikeName As String, beLikeText As String)
    Dim j As Long
    For j = 0 To nSum
        With node(j)
            If .b = True And j <> nid Then
                If LineAdd_RepeatedChecking(nid, j) = -1 Then
                    If .gravitational_s_name <> "" And beLikeName Like .gravitational_s_name Then
                        LineAdd nid, j, "", lineDefaultSize, , True, True
                    ElseIf .gravitational_s_text <> "" And beLikeText Like .gravitational_s_text Then
                        LineAdd nid, j, "", lineDefaultSize, , True, True
                    ElseIf .gravitational_t_name <> "" And beLikeName Like .gravitational_t_name Then
                        LineAdd j, nid, "", lineDefaultSize, , True, True
                    ElseIf .gravitational_t_text <> "" And beLikeText Like .gravitational_t_text Then
                        LineAdd j, nid, "", lineDefaultSize, , True, True
                    End If
                End If
            End If
        End With
    Next
End Sub
Public Sub NodeEdit_Save_TextGravityJudgment(nid As Long, likeStr As String, Source As Boolean)
    Dim j As Long
    For j = 0 To nSum
        With node(j)
            If .b = True And j <> nid Then
                If .text Like likeStr Then
                    If LineAdd_RepeatedChecking(nid, j) = -1 Then
                        If Source Then
                            LineAdd j, nid, "", lineDefaultSize, , True, True
                        Else
                            LineAdd nid, j, "", lineDefaultSize, , True, True
                        End If
                    End If
                End If
            End If
        End With
    Next
End Sub
Public Sub NodeEdit_Save_NameGravityJudgment(nid As Long, likeStr As String, Source As Boolean)
    Dim j As Long
    For j = 0 To nSum
        With node(j)
            If .b = True And j <> nid Then
                If .t Like likeStr Then
                    If LineAdd_RepeatedChecking(nid, j) = -1 Then
                        If Source Then
                            LineAdd j, nid, "", lineDefaultSize, , True, True
                        Else
                            LineAdd nid, j, "", lineDefaultSize, , True, True
                        End If
                    End If
                End If
            End If
        End With
    Next
End Sub
Public Function GetNodeGravitational(c As String, feature As String) As String
    Dim sT() As String, i As Long
    sT = Split(c, vbCrLf)
    For i = 0 To UBound(sT)
        If sT(i) <> "" Then
            If NCF_NodeGravitationalControl(sT(i), GetNodeGravitational, feature) Then
                Exit Function
            End If
        End If
    Next
End Function
Public Function LineAdd(ByRef Source As Long, ByRef target As Long, content As String, size As Single, Optional pitchOn As Boolean, Optional safe As Boolean, Optional artificial As Boolean)
    Dim LineRepeatedCheck As Long
    If safe = False Then
        LineRepeatedCheck = LineAdd_RepeatedChecking(Source, target)
    Else
        LineRepeatedCheck = -1
    End If
    If LineRepeatedCheck = -1 Then
        LineAdd_Save Source, target, content, size, pitchOn, artificial
    Else
        LineDelete LineRepeatedCheck
    End If
End Function
Public Function LineDelete(ByRef lid As Long)
    BehaviorListAdd "LineDelete", lid
    nodeLine(lid).b = False
End Function
Public Function LineReplace(lid As Long, f As Long, rN As Long, oldN As Long)
    BehaviorListAdd "LineReplace", lid, f, rN, oldN
    If f = 0 Then
        nodeLine(lid).Source = rN
    Else
        nodeLine(lid).target = rN
    End If
End Function
Public Function NodeDelete(ByRef nid As Long)
    BehaviorListAdd "NodeDelete", nid
    node(nid).b = False
    NodeDelete_RelevantLine nid
End Function
Public Function NodeDelete_RelevantLine(ByRef nid As Long)
Dim i As Long
For i = 0 To lSum
    With nodeLine(i)
        If .b = True And (.Source = nid Or .target = nid) Then
            LineDelete i
        End If
    End With
Next
End Function
Public Function LineAdd_Save(ByRef Source As Long, ByRef target As Long, content As String, size As Single, Optional pitchOn As Boolean, Optional artificial As Boolean)
    BehaviorListAdd "LineAdd_Save", lSum '行为记录函数
    With nodeLine(lSum)
        .b = True
        .Source = Source
        .target = target
        .select = pitchOn
        .content = content
        .size = size
    End With
    lSum = lSum + 1
    LineUboundAdd
    If Note.色彩链路.Checked = True And colorLinkDic.Exists(node(Source).setColor) = True And artificial = True Then
        NodeEdit_ReviseNode target, node(target).t, node(target).content, colorLinkDic(node(Source).setColor), node(target).setSize, True
    End If
End Function
Public Function LineAdd_RepeatedChecking(Source As Long, target As Long) As Long
    Dim i As Long
    LineAdd_RepeatedChecking = -1
    For i = 0 To lSum
        With nodeLine(i)
            If .b = True Then
                If (.Source = Source And .target = target) _
                    Or (.Source = target And .target = Source) Then
                    LineAdd_RepeatedChecking = i: Exit Function
                End If
            End If
        End With
    Next
End Function
