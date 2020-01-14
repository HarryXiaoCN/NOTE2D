Attribute VB_Name = "Note_Engine"
Private midNodeAT As Single, contentTemp As String
Sub Main()
    PublicVarLoad
    On Error GoTo Er
    Note.Show
    Exit Sub
Er:
    MsgBox "启动程序时出现错误，原因：" & Err.Description, 64, "启动节点笔记"
    环境文件拷贝
    Note.Show
End Sub
Public Function Updata()
        Note.Cls
        Updata_RectangleLine
        Updata_ColorUpdata
        Updata_GetNodeTargetAim
        Updata_NodeMove
        Updata_SelectMove
'        Updata_AllNodeMove
        Updata_NodeLine
        Updata_Node
        Updata_FictitiousIndex
        Updata_RegionalSelect
        Updata_Colourful
        Updata_pilotLightColor
        Updata_ChildNodeVis
End Function
Public Sub Updata_FictitiousIndex()
    Dim i As Long, j As Long
    If fictitiousIndexLock Then
        For i = 0 To UBound(fictitiousNote)
            With fictitiousNote(i)
                If .be Then
                    For j = 0 To UBound(.nodeLine)
                        If .nodeLine(j).direction = 1 Then
                            Note.FillColor = 11206527
                            Note.DrawWidth = .nodeLine(j).size
                            Note.Line (node(fictitiousRootNodeId).X, node(fictitiousRootNodeId).Y)-(.node(.nodeLine(j).target).realityX, .node(.nodeLine(j).target).realityY), 11206527
                            Note.Circle (.node(.nodeLine(j).target).realityX, .node(.nodeLine(j).target).realityY), .node(.nodeLine(j).target).setSize, 11206527
                            Note.CurrentX = .node(.nodeLine(j).target).realityX
                            Note.CurrentY = .node(.nodeLine(j).target).realityY
                            Note.Print .node(.nodeLine(j).target).t
                        ElseIf .nodeLine(j).direction = 2 Then
                            Note.FillColor = 9890667
                            Note.DrawWidth = .nodeLine(j).size
                            Note.Line (.node(.nodeLine(j).Source).realityX, .node(.nodeLine(j).Source).realityY)-(node(fictitiousRootNodeId).X, node(fictitiousRootNodeId).Y), 9890667
                            Note.Circle (.node(.nodeLine(j).Source).realityX, .node(.nodeLine(j).Source).realityY), .node(.nodeLine(j).Source).setSize, 9890667
                            Note.CurrentX = .node(.nodeLine(j).Source).realityX
                            Note.CurrentY = .node(.nodeLine(j).Source).realityY
                            Note.Print .node(.nodeLine(j).Source).t
                        ElseIf .nodeLine(j).direction = 3 Then
                            Note.Line (node(fictitiousRootNodeId).X, node(fictitiousRootNodeId).Y)-(node(.nodeLine(j).realityId).X, node(.nodeLine(j).realityId).Y), 11206527
                        ElseIf .nodeLine(j).direction = 4 Then
                            Note.Line (node(.nodeLine(j).realityId).X, node(.nodeLine(j).realityId).Y)-(node(fictitiousRootNodeId).X, node(fictitiousRootNodeId).Y), 9890667
                        End If
                    Next
                End If
            End With
        Next
    End If
End Sub

Public Sub Updata_RectangleLine()
    If Note.矩线.Checked Then
        Dim i As Long, height As Single, width As Single, lowW As Long, lowH As Long
        height = Note.height * zoomFactor
        width = Note.width * zoomFactor
        
        lowW = (-angleOfView.X \ nodeAttributedToIntegers - 1) * nodeAttributedToIntegers
        lowH = (-angleOfView.Y \ nodeAttributedToIntegers - 1) * nodeAttributedToIntegers
        
        For i = lowH To height - angleOfView.Y Step nodeAttributedToIntegers
            Note.Line (-angleOfView.X, i)-(width - angleOfView.X, i), rectangleLineColor
        Next
        For i = lowW To width - angleOfView.X Step nodeAttributedToIntegers
            Note.Line (i, -angleOfView.Y)-(i, height - angleOfView.Y), rectangleLineColor
        Next
    End If
End Sub
Public Sub Updata_ChildNodeVis()
    If childNodeVisLock Then Note.子节点视图容器.Visible = True
End Sub
Public Function Updata_ColorUpdata()
    If Note.流光溢彩.Checked Then
        If dColor > 70 Then
            dColor = 0
        End If
        dColor = dColor + 1
    End If
End Function
Public Function Updata_pilotLightColor()
    Note.DrawWidth = 1
    Note.Line (Note.pilotLightX0 * zoomFactor - angleOfView.X, Note.pilotLightY0 * zoomFactor - angleOfView.Y)-(Note.pilotLightX1 * zoomFactor - angleOfView.X, Note.pilotLightY1 * zoomFactor - angleOfView.Y), Note.pilotLightColor, BF
End Function
Public Function Updata_Colourful()
If nodeTargetAim <> -1 And Note.彩虹圈.Checked = True Then
    With node(nodeTargetAim)
        If nodeMoveLock = False Then
            Rainbow_Crcle Note, .size * 2, .X, .Y
        Else
            Rainbow_Crcle Note, .size * 3, .X, .Y
        End If
    End With
End If
End Function
Public Function Updata_GetNodeTargetAim()
    nodeTargetAim = NodeCheck(mousePos.X, mousePos.Y)
    Updata_GetNodeTargetAim_TagVis
    Updata_GetNodeTargetAim_Deselect
    If (Note.显示全部节点名.Checked = False Or Note.显示全部连接.Checked = False) And nodeTargetAim <> -1 Then
        Updata_GetNodeTargetAim_Select nodeTargetAim
    End If
End Function
Public Function Updata_GetNodeTargetAim_TagVis()
If Note.标签化.Checked = True Then
    If nodeTargetAim <> -1 Then
        Note.NodePrintBox.left = node(nodeTargetAim).X - Note.NodePrintBox.width / 2
        Note.NodePrintBox.Top = node(nodeTargetAim).Y + Note.NodePrintBox.height + node(nodeTargetAim).setSize * 0.8
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
                node(.Source).backward = True
                Updata_GetNodeTargetAim_Select_Backward .Source
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
            If .Source = nid Then
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
        Note.Line (regionalSelectStart.X, regionalSelectStart.Y)-(mousePos.X, mousePos.Y), RGB(126, 126, 126), B
        Note.FillStyle = 0
        Updata_RegionalSelect_Node
        Updata_RegionalSelect_Line
    End If
End Function
Public Function Updata_RegionalSelect_Line()
    Dim i As Long, widerRange As Boolean
    widerRange = RectangleRightStartCheck(regionalSelectStart, mousePos)
    For i = 0 To lSum
        With nodeLine(i)
            If .b = True Then
                If node(.Source).select = True And node(.target).select = True Then
                    .select = True
                ElseIf widerRange = True And LineIntersectionArea(node(.Source).X, node(.Source).Y, node(.target).X, node(.target).Y, regionalSelectStart, mousePos) = True Then
                    .select = True
                Else
                    .select = False
                End If
            End If
        End With
    Next
End Function
Public Function Updata_RegionalSelect_Node()
    Dim i As Long, checkPos As 二维坐标
    For i = 0 To nSum
        With node(i)
            If .b = True Then
                checkPos.X = .X: checkPos.Y = .Y
                .select = RectangleOverlapCheck(regionalSelectStart, mousePos, checkPos)
            End If
        End With
    Next
End Function
Public Function Updata_NodeMove()
    Dim dX As Single, dY As Single
    If nodeMoveLock = True Then
        If OverlappingJudgment(50, lineAddStrat.X, lineAddStrat.Y, mousePos.X, mousePos.Y) = False Then lineAddLock = False
        dX = mousePos.X - nodeMoveStart.X
        nodeMoveStart.X = mousePos.X
        dY = mousePos.Y - nodeMoveStart.Y
        nodeMoveStart.Y = mousePos.Y
        With node(nodeClickAim)
            .X = .X + dX
            .Y = .Y + dY
        End With
    End If
End Function
Public Function Updata_AllNodeMove(Optional selectMove As Boolean)
    Dim dX As Single: Dim dY As Single
    If selectMove = True Then
        dX = mousePos.X - allNodeMoveStart.X
        allNodeMoveStart.X = mousePos.X
        dY = mousePos.Y - allNodeMoveStart.Y
        allNodeMoveStart.Y = mousePos.Y
        Updata_AllNodeMove_Moving dX, dY, selectMove
        nodeEditPos.X = nodeEditPos.X + dX: nodeEditPos.Y = nodeEditPos.Y + dY
    End If 'MainCoordinateSystemDefinition
End Function
Public Function Updata_AllNodeMove_Moving(ByRef dX As Single, ByRef dY As Single, ByRef selectMove As Boolean)
Dim i As Long
For i = 0 To nSum
    With node(i)
        If .b = True Then
           If selectMove = False Then
                .X = .X + dX
                .Y = .Y + dY
            ElseIf .select = True Then
                .X = .X + dX
                .Y = .Y + dY
            End If
        End If
    End With
Next
End Function
Public Function Updata_Node()
    Dim i As Long
    Note.Font.size = MainFormFontSize / zoomFactor
    Note.DrawWidth = lineDefaultSize
    midNodeAT = nodeAttributedToIntegers * 0.5
    For i = 0 To nSum
        If node(i).b = True Then
            With node(i)
                If Note.节点归整.Checked = True And mainFormMouseState = False Then
                    .X = ((.X + midNodeAT) \ nodeAttributedToIntegers) * nodeAttributedToIntegers
                    .Y = ((.Y + midNodeAT) \ nodeAttributedToIntegers) * nodeAttributedToIntegers
                ElseIf Note.节点归整.Checked = True And mainFormMouseState = True And Note.矩线.Checked = True Then
                    Note.Line (.X, .Y)-(((.X + midNodeAT) \ nodeAttributedToIntegers) * nodeAttributedToIntegers, ((.Y + midNodeAT) \ nodeAttributedToIntegers) * nodeAttributedToIntegers), 11206527
                End If
                Updata_Node_SetColor i
                If Note.显示全部节点名.Checked = True Or nodeTargetAim = i Then
                    Updata_Node_FormPrint .X, .Y, .t, .content
                    If .select = True Then
                        Updata_Node_FontBold .X, .Y, .t
                    End If
                End If
                If Note.显示顺向节点名.Checked = True And .forward = True Then
                    Updata_Node_FormPrint .X, .Y, .t, .content
                    If .select = True Then
                        Updata_Node_FontBold .X, .Y, .t
                    End If
                End If
                If Note.显示逆向节点名.Checked = True And .backward = True Then
                    Updata_Node_FormPrint .X, .Y, .t, .content
                    If .select = True Then
                        Updata_Node_FontBold .X, .Y, .t
                    End If
                End If
                If Note.始终显示选点名.Checked = True And .select = True Then
                    Updata_Node_FormPrint .X, .Y, .t, .content
                    Updata_Node_FontBold .X, .Y, .t
                End If
                If Note.显示节点遍历ID.Checked = True Then
                    Updata_Node_FormPrint .X - 220, .Y + 200, str(i)
                End If
                If nodeTargetAim = i And (notePrintNodeId <> i Or nodePrintBeLock = False Or needUpdataNodePrint = True) Then
                    If nodePrintBeLock = False And Note.标签化.Checked = False Then NodePrint.Show
                    notePrintNodeId = i
                    NodePrint.Caption = .t
                    childNodeVisLock = 子节点检查(富文本转义(.content))
                    If Note.标签化.Checked = True Then
                        Note.NodePrintBox.TextRTF = .content
                    Else
                        Updata_Node_ContentPrint .content
                    End If
                    needUpdataNodePrint = False
                End If
            End With
        End If
    Next
    Updata_Node_addNew
End Function
Public Function Updata_Node_FormPrint(X As Single, Y As Single, t As String, Optional c As String)
    'form replace 2
    Note.CurrentX = X
    Note.CurrentY = Y
    If Note.精简内容.Checked Then
        contentTemp = 富文本转义(c)
        If Len(contentTemp) > 10 Then
            Note.Print t & ":" & Mid(contentTemp, 1, 10) & "...."
        Else
            Note.Print t & ":" & contentTemp
        End If
    Else
        Note.Print t
    End If
End Function
Public Function Updata_Node_FontBold(ByRef X As Single, ByRef Y As Single, ByRef t As String)
    Dim fCTmp As Long
    'form replace 3
    fCTmp = Note.ForeColor
    Note.ForeColor = RGB(220, 20, 60)
    Note.CurrentX = X
    Note.CurrentY = Y
    Note.Print t
    Note.ForeColor = fCTmp
End Function
Public Function Updata_Node_addNew()
If nodeEditLock = False And nodeEditAim = -1 And nodeEditFormLock = True Then
    If Note.矩点.Checked = False Then
        Note.FillColor = Updata_Node_GetColor(-1)
        Note.Circle (nodeEditPos.X, nodeEditPos.Y), nodeDefaultSize, Updata_Node_GetColor(-1)
        Note.FillColor = Updata_Node_GetColor(0)
    Else
        Note.Line (nodeEditPos.X - nodeDefaultSize, nodeEditPos.Y - nodeDefaultSize)-(nodeEditPos.X + nodeDefaultSize, nodeEditPos.Y + nodeDefaultSize), Updata_Node_GetColor(-1), BF
    End If
End If
End Function
Public Function Updata_Node_SetColor(ByRef nid As Long)
    'form replace 4
    With node(nid)
        If nodeTargetAim = nid Then .color = 1: .size = .setSize * 1.2 Else .color = 0: .size = .setSize
        If .select = True Then .color = -2
        If nodeMoveLock = True And nodeClickAim = nid Then .color = 2: .size = .setSize * 1.2
        If Note.矩点.Checked = False Then
            Note.FillColor = Updata_Node_GetColor(.color, .setColor)
            Note.Circle (.X, .Y), .size, Updata_Node_GetColor(.color, .setColor)
            Note.FillColor = Updata_Node_GetColor(0, .setColor)
        Else
            Note.Line (.X - .size, .Y - .size)-(.X + .size, .Y + .size), Updata_Node_GetColor(.color, .setColor), BF
        End If
    End With
End Function
Public Function Updata_Node_GetColor(ByRef color As Long, Optional oneselfColor As Long)
Select Case color
    Case -2
        If Note.流光溢彩.Checked = False Then
            Updata_Node_GetColor = 55295 '金
        Else
            Updata_Node_GetColor = Rainbow_RedEnd(dColor)
        End If
    Case -1
        Updata_Node_GetColor = 11206527 '9890667 绿
    Case 0
        Updata_Node_GetColor = oneselfColor
    Case 1
        Updata_Node_GetColor = 13828244 '紫
    Case 2
        Updata_Node_GetColor = 9639167 '品红
End Select
End Function
Public Function Updata_Node_ContentPrint(ByRef content As String)
    NodePrint.NodePrintBox.TextRTF = content
End Function
Public Function Updata_NodeLine()
    Dim i As Long, c As Long: Dim midX As Single, midY As Single
    Dim lineStartPos() As 二维坐标
    Dim lineEndPos() As 二维坐标
        'form replace 5
    ReDim lineStartPos(lSum)
    ReDim lineEndPos(lSum)
    For i = 0 To lSum
        With nodeLine(i)
            If .b = True Then
                If (Note.显示全部连接.Checked = True _
                Or (Note.显示顺向连接.Checked = True And .forward = True) _
                Or (Note.显示逆向连接.Checked = True And .backward = True)) _
                Or (Note.始终显示选接.Checked = True And .select = True) Then
                    midX = (node(.target).X - node(.Source).X) / 3 * 2 + node(.Source).X
                    midY = (node(.target).Y - node(.Source).Y) / 3 * 2 + node(.Source).Y
                    If .select = False Then
                        Note.DrawWidth = .size
                        If Note.彩虹线.Checked = False Then
                            If Note.流光溢彩.Checked = False Then
                                Note.Line (node(.Source).X, node(.Source).Y)-(midX, midY), RGB(255, 0, 0)
                                Note.Line (midX, midY)-(node(.target).X, node(.target).Y), RGB(0, 0, 255)
                            Else
                                DynamicRainbow_Line Note, node(.Source).X, node(.Source).Y, node(.target).X, node(.target).Y
                                c = c + 1
                            End If
                        Else
                            DoubleColorLine Note, node(.Source), node(.target), midX, midY, .size
                        End If
                    Else
                        Note.DrawWidth = .size * 2
                        If Note.流光溢彩.Checked = False Then
                            Rainbow_Line Note, node(.Source).X, node(.Source).Y, node(.target).X, node(.target).Y
                        Else
                            DynamicRainbow_Line Note, node(.Source).X, node(.Source).Y, node(.target).X, node(.target).Y
                        End If
                    End If
                    If .content <> "" Then
                        Note.CurrentX = midX
                        Note.CurrentY = midY
                        Note.Print .content
                    End If
                End If
            End If
        End With
    Next
    Updata_nodeLine_addNewLine
End Function
Public Function Updata_nodeLine_addNewLine()
If lineAddLock = True And nodeClickAim <> -1 Then
    Note.DrawWidth = lineDefaultSize
    If Note.流光溢彩.Checked = False Then
        If Note.彩虹线.Checked = False Then
            midX = (mousePos.X - node(nodeClickAim).X) / 3 * 2 + node(nodeClickAim).X
            midY = (mousePos.Y - node(nodeClickAim).Y) / 3 * 2 + node(nodeClickAim).Y
            Note.Line (node(nodeClickAim).X, node(nodeClickAim).Y)-(midX, midY), RGB(255, 0, 0)
            Note.Line (midX, midY)-(mousePos.X, mousePos.Y), RGB(0, 0, 255)
        Else
            Rainbow_Line Note, node(nodeClickAim).X, node(nodeClickAim).Y, mousePos.X, mousePos.Y
        End If
    Else
            DynamicRainbow_Line Note, node(nodeClickAim).X, node(nodeClickAim).Y, mousePos.X, mousePos.Y
    End If
End If
End Function
