Attribute VB_Name = "Node_Edit"
Public Function NodeEditeStart(ByRef X As Single, ByRef Y As Single)
nodeEditAim = NodeCheck(X, Y)
If nodeEditAim = -1 Then '新建节点
    nodeEditLock = False
    If NodeInput.保持内容.Checked = False Then
        NodeInput.NodeTitle.Text = "请输入节点标题..."
        NodeInput.NodeInputBox.Text = "请输入节点内容..."
    End If
    nodeEditPos.X = X
    nodeEditPos.Y = Y
Else
    NodeInput.NodeTitle.Text = node(nodeEditAim).t
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
Public Function NodeEdit_NewNode(ByVal title As String, ByVal content As String, setC As Long, setS As Single, ByRef X As Single, ByRef Y As Single, Optional pitchOn As Boolean) As Long
    BehaviorListAdd "NodeEdit_NewNode", nSum
    NodeEdit_Save nSum, title, content, setC, setS, X, Y, pitchOn
    nodeEditLock = True
    nodeEditAim = nSum
    nSum = nSum + 1
    NodeEdit_NewNode = nodeEditAim
End Function
Public Function NodeEdit_TitleFilter(ByRef nid As Long, ByRef title As String) As String
If title = "" Or title = "请输入节点标题..." Then
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
Public Function NodeEdit_ReviseNode(ByRef nid As Long, ByRef title As String, ByRef content As String, setC As Long, setS As Single)
With node(nid)
    BehaviorListAdd "NodeEdit_ReviseNode", nid, .t, .content, .setColor, .setSize
    NodeEdit_Save nid, title, content, setC, setS, .X, .Y
End With
End Function
Public Function NodeEdit_Save(ByRef nid As Long, ByRef title As String, ByRef content As String, setC As Long, setS As Single, ByRef X As Single, ByRef Y As Single, Optional pitchOn As Boolean)
With node(nid)
    .b = True
    .X = X
    .Y = Y
    .size = setS
    .t = NodeEdit_TitleFilter(nid, title)
    .content = content
    .setColor = setC
    .setSize = setS
    .select = pitchOn
    If nodeEditFormLock Then
        NodeInput.NodeTitle.Text = .t
        NodeInput.NodeInputBox.TextRTF = .content
    End If
End With
NodeUboundAdd
End Function
Public Function LineAdd(ByRef Source As Long, ByRef target As Long, content As String, size As Single, Optional pitchOn As Boolean, Optional safe As Boolean)
    Dim LineRepeatedCheck As Long
    If safe = False Then
        LineRepeatedCheck = LineAdd_RepeatedChecking(Source, target)
    Else
        LineRepeatedCheck = -1
    End If
    If LineRepeatedCheck = -1 Then
        LineAdd_Save Source, target, content, size, pitchOn
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
        If .b = True And (.Source = nid Or .target = nid) Then .b = False
    End With
Next
End Function
Public Function LineAdd_Save(ByRef Source As Long, ByRef target As Long, content As String, size As Single, Optional pitchOn As Boolean)
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
End Function
Public Function LineAdd_RepeatedChecking(ByRef Source As Long, ByRef target As Long) As Long
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
