Attribute VB_Name = "Node_Edit"
Public Function NodeEditeStart(ByRef x As Single, ByRef y As Single)
nodeEditAim = NodeCheck(x, y)
If nodeEditAim = -1 Then '�½��ڵ�
    nodeEditLock = False
    NodeInput.NodeTitle.Text = "������ڵ����..."
    NodeInput.NodeInputBox.Text = "������ڵ�����..."
    nodeEditPos.x = x
    nodeEditPos.y = y
Else
    NodeInput.NodeTitle.Text = node(nodeEditAim).t
    NodeInput.NodeInputBox.TextRTF = node(nodeEditAim).content
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
Public Function NodeCheck(ByRef x As Single, ByRef y As Single) As Long
Dim i As Long
NodeCheck = -1
For i = 0 To nSum
    If node(i).b = True Then
        If OverlappingJudgment(120, x, y, node(i).x, node(i).y) = True Then
            NodeCheck = i: Exit Function
        End If
    End If
Next
End Function
Public Function NodeEdit_NewNode(ByVal title As String, ByVal content As String, ByRef x As Single, ByRef y As Single, Optional pitchOn As Boolean)
BehaviorListAdd "NodeEdit_NewNode", nSum
NodeEdit_Save nSum, title, content, x, y, pitchOn
nodeEditLock = True
nodeEditAim = nSum
nSum = nSum + 1
End Function
Public Function NodeEdit_TitleFilter(ByRef nid As Long, ByRef title As String) As String
If title = "" Or title = "������ڵ����..." Then
    NodeEdit_TitleFilter = NodeEdit_TitleFilter_StrCombination
Else
    NodeEdit_TitleFilter = title
End If
End Function
Public Function NodeEdit_ContentFilter(ByRef content As String) As Boolean
If content = "������ڵ�����..." Then
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
Public Function NodeEdit_ReviseNode(ByRef nid As Long, ByRef title As String, ByRef content As String)
With node(nid)
    BehaviorListAdd "NodeEdit_ReviseNode", nid, .t, .content
    NodeEdit_Save nid, title, content, .x, .y
End With
End Function
Public Function NodeEdit_Save(ByRef nid As Long, ByRef title As String, ByRef content As String, ByRef x As Single, ByRef y As Single, Optional pitchOn As Boolean)
With node(nid)
    .b = True
    .x = x
    .y = y
    .size = 100
    .t = NodeEdit_TitleFilter(nid, title)
    .content = content
    .select = pitchOn
    NodeInput.NodeTitle.Text = .t
    NodeInput.NodeInputBox.TextRTF = .content
End With
NodeUboundAdd
End Function
Public Function LineAdd(ByRef source As Long, ByRef target As Long, Optional pitchOn As Boolean)
Dim LineRepeatedCheck As Long
LineRepeatedCheck = LineAdd_RepeatedChecking(source, target)
If LineRepeatedCheck = -1 Then
    LineAdd_Save source, target, pitchOn
Else
    LineDelete LineRepeatedCheck
End If
End Function
Public Function LineDelete(ByRef lid As Long)
BehaviorListAdd "LineAdd_Save", lid
nodeLine(lid).b = False
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
        If .b = True And (.source = nid Or .target = nid) Then .b = False
    End With
Next
End Function
Public Function LineAdd_Save(ByRef source As Long, ByRef target As Long, Optional pitchOn As Boolean)
BehaviorListAdd "LineAdd_Save", lSum '��Ϊ��¼����
With nodeLine(lSum)
    .b = True
    .source = source
    .target = target
    .select = pitchOn
End With
lSum = lSum + 1
LineUboundAdd
End Function
Public Function LineAdd_RepeatedChecking(ByRef source As Long, ByRef target As Long) As Long
Dim i As Long
LineAdd_RepeatedChecking = -1
For i = 0 To lSum
    With nodeLine(i)
        If .b = True Then
            If (.source = source And .target = target) _
                Or (.source = target And .target = source) Then
                LineAdd_RepeatedChecking = i: Exit Function
            End If
        End If
    End With
Next
End Function
