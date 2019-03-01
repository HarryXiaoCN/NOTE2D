Attribute VB_Name = "Note_Engine"
Sub Main()
LoadPublicVar
On Error GoTo Er
Note.Show
Exit Sub
Er:
环境文件拷贝
Note.Show
End Sub
Public Function Updata()
Note.Cls
Updata_GetNodeTargetAim
Updata_NodeMove
Updata_SelectMove
Updata_AllNodeMove
Updata_NodeLine
Updata_Node
Updata_RegionalSelect
Updata_Colourful
End Function
Public Function Updata_Colourful()
If nodeTargetAim <> -1 And Note.彩虹圈.Checked = True Then
    With node(nodeTargetAim)
        If nodeMoveLock = False Then
            Rainbow_Crcle Note, .size * 2, .x, .y
        Else
            Rainbow_Crcle Note, .size * 3, .x, .y
        End If
    End With
End If
End Function
Public Function Updata_GetNodeTargetAim()
nodeTargetAim = NodeCheck(mousePos.x, mousePos.y)
Updata_GetNodeTargetAim_TagVis
Updata_GetNodeTargetAim_Deselect
If (Note.显示全部节点名.Checked = False Or Note.显示全部连接.Checked = False) And nodeTargetAim <> -1 Then
    Updata_GetNodeTargetAim_Select nodeTargetAim
End If
End Function
Public Function Updata_GetNodeTargetAim_TagVis()
If Note.标签化.Checked = True Then
    If nodeTargetAim <> -1 Then
        Note.NodePrintBox.left = node(nodeTargetAim).x + 50
        Note.NodePrintBox.Top = node(nodeTargetAim).y - 50
        Note.NodePrintBox.Visible = True
    Else
        Note.NodePrintBox.Visible = False
    End If
End If
End Function
Public Function Updata_GetNodeTargetAim_Deselect()
Dim i As Long
For i = 0 To nSum
    With node(i)
        If .b = True Then
            .forward = False: .backward = False: depth = 0
        End If
    End With
Next
For i = 0 To lSum
    With nodeLine(i)
        If .b = True Then
            .forward = False: .backward = False
        End If
    End With
Next
End Function

Public Function Updata_GetNodeTargetAim_Select(ByRef nid As Long)
Updata_GetNodeTargetAim_Select_Forward nid
Updata_GetNodeTargetAim_Select_Backward nid
node(nid).forward = True
node(nid).backward = True
End Function
Public Function Updata_GetNodeTargetAim_Select_Backward(ByRef nid As Long)
Dim i As Long
For i = 0 To lSum
    With nodeLine(i)
        If .b = True And .search = False Then
            If .target = nid Then
                .search = True
                .backward = True
                node(.source).backward = True
                Updata_GetNodeTargetAim_Select_Backward .source
                .search = False
            End If
        End If
    End With
Next
End Function
Public Function Updata_GetNodeTargetAim_Select_Forward(ByRef nid As Long, Optional depth As Long)
Dim i As Long
For i = 0 To lSum
    With nodeLine(i)
        If .b = True And .search = False Then
            If .source = nid Then
                .search = True
                .forward = True
                node(.target).forward = True
                node(.target).depth = depth
                Updata_GetNodeTargetAim_Select_Forward .target, depth + 1
                .search = False
            End If
        End If
    End With
Next
End Function
Public Function Updata_SelectMove()
If selectMoveLock = True Then Updata_AllNodeMove True
End Function
Public Function Updata_RegionalSelect()
If regionalSelectLock = True Then
    Note.FillStyle = 1
    Note.Line (regionalSelectStart.x, regionalSelectStart.y)-(mousePos.x, mousePos.y), RGB(126, 126, 126), B
    Note.FillStyle = 0
    Updata_RegionalSelect_Node
    Updata_RegionalSelect_Line
End If
End Function
Public Function Updata_RegionalSelect_Line()
Dim i As Long: Dim widerRange As Boolean
widerRange = RectangleRightStartCheck(regionalSelectStart, mousePos)
For i = 0 To lSum
    With nodeLine(i)
        If .b = True Then
            If node(.source).select = True And node(.target).select = True Then
                .select = True
            ElseIf widerRange = True And LineIntersectionArea(node(.source).x, node(.source).y, node(.target).x, node(.target).y, regionalSelectStart, mousePos) = True Then
                .select = True
            Else
                .select = False
            End If
        End If
    End With
Next
End Function
Public Function Updata_RegionalSelect_Node()
Dim i As Long: Dim checkPos As 二维坐标
For i = 0 To nSum
    With node(i)
        If .b = True Then
            checkPos.x = .x: checkPos.y = .y
            .select = RectangleOverlapCheck(regionalSelectStart, mousePos, checkPos)
        End If
    End With
Next
End Function
Public Function Updata_NodeMove()
Dim dx, dy As Single
If nodeMoveLock = True Then
    If OverlappingJudgment(50, lineAddStrat.x, lineAddStrat.y, mousePos.x, mousePos.y) = False Then lineAddLock = False
    dx = mousePos.x - nodeMoveStart.x
    nodeMoveStart.x = mousePos.x
    dy = mousePos.y - nodeMoveStart.y
    nodeMoveStart.y = mousePos.y
    With node(nodeClickAim)
        .x = .x + dx
        .y = .y + dy
    End With
End If
End Function
Public Function Updata_AllNodeMove(Optional selectMove As Boolean)
Dim dx As Single: Dim dy As Single
If allNodeMoveLock = True Or selectMove = True Then
    dx = mousePos.x - allNodeMoveStart.x
    allNodeMoveStart.x = mousePos.x
    dy = mousePos.y - allNodeMoveStart.y
    allNodeMoveStart.y = mousePos.y
    Updata_AllNodeMove_Moving dx, dy, selectMove
    nodeEditPos.x = nodeEditPos.x + dx: nodeEditPos.y = nodeEditPos.y + dy
End If
End Function
Public Function Updata_AllNodeMove_Moving(ByRef dx As Single, ByRef dy As Single, ByRef selectMove As Boolean)
Dim i As Long
For i = 0 To nSum
    With node(i)
        If .b = True Then
           If selectMove = False Then
                .x = .x + dx
                .y = .y + dy
            ElseIf .select = True Then
                .x = .x + dx
                .y = .y + dy
            End If
        End If
    End With
Next
End Function
Public Function Updata_Node()
Dim i As Long
Note.Font.size = MainFormFontSize / zoomFactor
For i = 0 To nSum
    If node(i).b = True Then
        With node(i)
            Updata_Node_SetColor i
            If Note.显示全部节点名.Checked = True Or nodeTargetAim = i Then
                Updata_Node_FormPrint .x, .y, .t
                If .select = True Then
                    Updata_Node_FontBold .x, .y, .t
                End If
            End If
            If Note.显示顺向节点名.Checked = True And .forward = True Then
                Updata_Node_FormPrint .x, .y, .t
                If .select = True Then
                    Updata_Node_FontBold .x, .y, .t
                End If
            End If
            If Note.显示逆向节点名.Checked = True And .backward = True Then
                Updata_Node_FormPrint .x, .y, .t
                If .select = True Then
                    Updata_Node_FontBold .x, .y, .t
                End If
            End If
            If Note.始终显示选点名.Checked = True And .select = True Then
                Updata_Node_FormPrint .x, .y, .t
                Updata_Node_FontBold .x, .y, .t
            End If
            If Note.显示节点遍历ID.Checked = True Then
                Updata_Node_FormPrint .x - 220, .y + 200, str(i)
            End If
            If nodeTargetAim = i And (notePrintNodeId <> i Or nodePrintBeLock = False) Then
                If nodePrintBeLock = False And Note.标签化.Checked = False Then NodePrint.Show
                notePrintNodeId = i
                NodePrint.Caption = .t
                If Note.标签化.Checked = True Then
                    Note.NodePrintBox.TextRTF = .content
                Else
                    Updata_Node_ContentPrint .content
                End If
            End If
        End With
    End If
Next
Updata_Node_addNew
End Function
Public Function Updata_Node_FormPrint(ByRef x As Single, ByRef y As Single, ByRef t As String)
Note.CurrentX = x
Note.CurrentY = y
Note.Print t
End Function
Public Function Updata_Node_FontBold(ByRef x As Single, ByRef y As Single, ByRef t As String)
Dim fCTmp As Long
fCTmp = Note.ForeColor
Note.ForeColor = RGB(220, 20, 60)
Note.CurrentX = x
Note.CurrentY = y
Note.Print t
Note.ForeColor = fCTmp
End Function
Public Function Updata_Node_addNew()
If nodeEditLock = False And nodeEditAim = -1 And nodeEditFormLock = True Then
    Note.FillColor = Updata_Node_GetColor(-1)
    Note.Circle (nodeEditPos.x, nodeEditPos.y), 100, Updata_Node_GetColor(-1)
    Note.FillColor = Updata_Node_GetColor(0)
End If
End Function
Public Function Updata_Node_SetColor(ByRef nid As Long)
With node(nid)
    If nodeTargetAim = nid Then .color = 1: .size = 120 Else .color = 0: .size = 100
    If .select = True Then .color = -2
    If nodeMoveLock = True And nodeClickAim = nid Then .color = 2: .size = 120
    Note.FillColor = Updata_Node_GetColor(.color)
    Note.Circle (.x, .y), .size, Updata_Node_GetColor(.color)
    Note.FillColor = Updata_Node_GetColor(0)
End With
End Function
Public Function Updata_Node_GetColor(ByRef color As Long)
Select Case color
    Case -2
        Updata_Node_GetColor = RGB(255, 215, 0)
    Case -1
        Updata_Node_GetColor = RGB(127, 255, 170)
    Case 0
        Updata_Node_GetColor = RGB(0, 191, 255)
    Case 1
        Updata_Node_GetColor = RGB(148, 0, 211)
    Case 2
        Updata_Node_GetColor = RGB(255, 20, 147)
End Select
End Function
Public Function Updata_Node_ContentPrint(ByRef content As String)
NodePrint.NodePrintBox.TextRTF = content
End Function
Public Function Updata_NodeLine()
Dim i As Long, c As Long: Dim midX, midY As Single
Dim lineStartPos() As 二维坐标
Dim lineEndPos() As 二维坐标
ReDim lineStartPos(lSum)
ReDim lineEndPos(lSum)
For i = 0 To lSum
    With nodeLine(i)
        If .b = True Then
            If (Note.显示全部连接.Checked = True _
            Or (Note.显示顺向连接.Checked = True And .forward = True) _
            Or (Note.显示逆向连接.Checked = True And .backward = True)) _
            Or (Note.始终显示选接.Checked = True And .select = True) Then
                midX = (node(.source).x + node(.target).x) / 2
                midY = (node(.source).y + node(.target).y) / 2
                If .select = False Then
                    If Note.彩虹线.Checked = False Then
                        Note.Line (node(.source).x, node(.source).y)-(midX, midY), RGB(255, 0, 0)
                        Note.Line (midX, midY)-(node(.target).x, node(.target).y), RGB(0, 0, 255)
                    Else
                        If Note.流光溢彩.Checked = False Then
                            Rainbow_Line Note, node(.source).x, node(.source).y, node(.target).x, node(.target).y
                        Else
'                            lineStartPos(c).x = node(.source).x
'                            lineStartPos(c).y = node(.source).y
'                            lineEndPos(c).x = node(.target).x
'                            lineEndPos(c).y = node(.target).y
                            DynamicRainbow_Line Note, node(.source).x, node(.source).y, node(.target).x, node(.target).y
                            c = c + 1
                        End If
                    End If
                Else
                    Note.DrawWidth = 3
                    If Note.彩虹线.Checked = False Then
                        Note.Line (node(.source).x, node(.source).y)-(midX, midY), RGB(255, 215, 0)
                        Note.Line (midX, midY)-(node(.target).x, node(.target).y), RGB(255, 255, 0)
                    Else
                        If Note.流光溢彩.Checked = False Then
                            Rainbow_Line Note, node(.source).x, node(.source).y, node(.target).x, node(.target).y
                        Else
                            DynamicRainbow_Line Note, node(.source).x, node(.source).y, node(.target).x, node(.target).y
                        End If
                    End If
                    Note.DrawWidth = 1
                End If
            End If
            
        End If
    End With
Next
'If Note.流光溢彩.Checked = True Then
'    For i = 0 To c - 1
'        DynamicRainbow_Line Note, lineStartPos(i).x, lineStartPos(i).y, lineEndPos(i).x, lineEndPos(i).y
'    Next
'End If
Updata_nodeLine_addNewLine
End Function
Public Function Updata_nodeLine_addNewLine()
If lineAddLock = True And nodeClickAim <> -1 Then
    If Note.彩虹线.Checked = False Then
        midX = (node(nodeClickAim).x + mousePos.x) / 2
        midY = (node(nodeClickAim).y + mousePos.y) / 2
        Note.Line (node(nodeClickAim).x, node(nodeClickAim).y)-(midX, midY), RGB(255, 0, 0)
        Note.Line (midX, midY)-(mousePos.x, mousePos.y), RGB(0, 0, 255)
    Else
        If Note.流光溢彩.Checked = False Then
            Rainbow_Line Note, node(nodeClickAim).x, node(nodeClickAim).y, mousePos.x, mousePos.y
        Else
            DynamicRainbow_Line Note, node(nodeClickAim).x, node(nodeClickAim).y, mousePos.x, mousePos.y
        End If
    End If
End If
End Function
