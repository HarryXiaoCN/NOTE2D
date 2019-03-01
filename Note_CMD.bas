Attribute VB_Name = "Note_CMD"
Public Function CMD_In(ByRef cmd As String)
Dim cmdList, fOut
Dim localVarNameList() As String
Dim i As Long
cmdList = Split(cmd, vbCrLf)
For i = 0 To UBound(cmdList)
    If cmdList(i) <> "" Then
        fOut = CMD_In_Line(cmdList(i))
    End If
Next
End Function
Public Function CMD_In_Line(ByVal cmdLine As String)
Dim temp, fPar
Dim fName As String
temp = Split(cmdLine, ":")
fName = temp(0)
fPar = Split(temp(1), ",")
CMD_In_Line = CMD_In_Line_GetFunction(fName, fPar)
End Function
Public Function CMD_In_Line_GetFunction(ByRef fName As String, ByRef fPar)
Dim fOut
If InStr(1, fName, "NodeEditeStart", 1) Then
    NodeEditeStart Val(fPar(0)), Val(fPar(1))
ElseIf InStr(1, fName, "NodeUboundAdd", 1) Then
    NodeUboundAdd fPar(0)
ElseIf InStr(1, fName, "LineUboundAdd", 1) Then
    LineUboundAdd fPar(0)
ElseIf InStr(1, fName, "NodeCheck", 1) Then
    fOut = NodeCheck(Val(fPar(0)), Val(fPar(1)))
ElseIf InStr(1, fName, "NodeEdit_NewNode", 1) Then
    NodeEdit_NewNode fPar(0), fPar(1), Val(fPar(2)), Val(fPar(3)), StrToBool(fPar(4))
ElseIf InStr(1, fName, "NodeEdit_ContentFilter", 1) Then
    fOut = NodeEdit_ContentFilter(StrToBool(fPar(0)))
ElseIf InStr(1, fName, "LineAdd", 1) Then
    LineAdd Val(fPar(0)), Val(fPar(1)), StrToBool(fPar(2))
ElseIf InStr(1, fName, "LineDelete", 1) Then
    LineDelete Val(fPar(0))
ElseIf InStr(1, fName, "NodeDelete", 1) Then
    NodeDelete Val(fPar(0))
ElseIf InStr(1, fName, "NodeDelete_RelevantLine", 1) Then
    NodeDelete_RelevantLine Val(fPar(0))
ElseIf InStr(1, fName, "LineAdd_Save", 1) Then
    LineAdd_Save Val(fPar(0)), Val(fPar(1)), StrToBool(fPar(2))
ElseIf InStr(1, fName, "LineAdd_RepeatedChecking", 1) Then
    fOut = LineAdd_RepeatedChecking(Val(fPar(0)), Val(fPar(1)))
ElseIf InStr(1, fName, "Updata", 1) Then
    Updata
ElseIf InStr(1, fName, "Updata_Colourful", 1) Then
    Updata_Colourful fPar(0)
ElseIf InStr(1, fName, "Updata_GetNodeTargetAim", 1) Then
    Updata_GetNodeTargetAim fPar(0)
ElseIf InStr(1, fName, "Updata_GetNodeTargetAim_Deselect", 1) Then
    Updata_GetNodeTargetAim_Deselect fPar(0)
ElseIf InStr(1, fName, "Updata_GetNodeTargetAim_Select", 1) Then
    Updata_GetNodeTargetAim_Select Val(fPar(0))
ElseIf InStr(1, fName, "Updata_GetNodeTargetAim_Select_Backward", 1) Then
    Updata_GetNodeTargetAim_Select_Backward Val(fPar(0))
ElseIf InStr(1, fName, "Updata_GetNodeTargetAim_Select_Forward", 1) Then
    Updata_GetNodeTargetAim_Select_Forward Val(fPar(0))
ElseIf InStr(1, fName, "Updata_SelectMove", 1) Then
    Updata_SelectMove fPar(0)
ElseIf InStr(1, fName, "Updata_RegionalSelect", 1) Then
    Updata_RegionalSelect fPar(0)
ElseIf InStr(1, fName, "Updata_RegionalSelect_Line", 1) Then
    Updata_RegionalSelect_Line fPar(0)
ElseIf InStr(1, fName, "Updata_RegionalSelect_Node", 1) Then
    Updata_RegionalSelect_Node fPar(0)
ElseIf InStr(1, fName, "Updata_NodeMove", 1) Then
    Updata_NodeMove fPar(0)
ElseIf InStr(1, fName, "Updata_AllNodeMove", 1) Then
    Updata_AllNodeMove StrToBool(fPar(0))
ElseIf InStr(1, fName, "Updata_AllNodeMove_Moving", 1) Then
    Updata_AllNodeMove_Moving Val(fPar(0)), Val(fPar(1)), StrToBool(fPar(2))
ElseIf InStr(1, fName, "Updata_Node", 1) Then
    Updata_Node fPar(0)
ElseIf InStr(1, fName, "Updata_Node_addNew", 1) Then
    Updata_Node_addNew fPar(0)
ElseIf InStr(1, fName, "Updata_Node_SetColor", 1) Then
    Updata_Node_SetColor Val(fPar(0))
ElseIf InStr(1, fName, "Updata_Node_GetColor", 1) Then
    fOut = Updata_Node_GetColor(Val(fPar(0)))
ElseIf InStr(1, fName, "Updata_NodeLine", 1) Then
    Updata_NodeLine fPar(0)
ElseIf InStr(1, fName, "Updata_nodeLine_addNewLine", 1) Then
    Updata_nodeLine_addNewLine fPar(0)
ElseIf InStr(1, fName, "LoadProfile", 1) Then
    LoadProfile fPar(0)
ElseIf InStr(1, fName, "LoadProfile_InitializationBool", 1) Then
    LoadProfile_InitializationBool fPar(0)
ElseIf InStr(1, fName, "SaveProfile", 1) Then
    SaveProfile fPar(0)
ElseIf InStr(1, fName, "noteSaveCheck_ContentCheck", 1) Then
    fOut = noteSaveCheck_ContentCheck(StrToBool(fPar(0)))
ElseIf InStr(1, fName, "NoteFileRead_VersionCheck", 1) Then
    fOut = NoteFileRead_VersionCheck(Val(fPar(0)))
ElseIf InStr(1, fName, "noteArrInitialization", 1) Then
    noteArrInitialization fPar(0)
ElseIf InStr(1, fName, "newAddNote", 1) Then
    newAddNote fPar(0)
ElseIf InStr(1, fName, "NoteGlobalViewSet", 1) Then
    NoteGlobalViewSet StrToBool(fPar(0))
ElseIf InStr(1, fName, "SelectDisplayObjcet", 1) Then
    SelectDisplayObjcet fPar(0)
ElseIf InStr(1, fName, "SelectDisplayObjcet_Forward", 1) Then
    SelectDisplayObjcet_Forward StrToBool(fPar(0))
ElseIf InStr(1, fName, "RollerEventHandling", 1) Then
    RollerEventHandling StrToBool(fPar(0))
ElseIf InStr(1, fName, "MainCoordinateSystemDefinition", 1) Then
    MainCoordinateSystemDefinition fPar(0)
ElseIf InStr(1, fName, "BehaviorListUboundAdd", 1) Then
    BehaviorListUboundAdd fPar(0)
ElseIf InStr(1, fName, "RedoListUboundAdd", 1) Then
    RedoListUboundAdd fPar(0)
ElseIf InStr(1, fName, "CopyObject", 1) Then
    CopyObject StrToBool(fPar(0))
ElseIf InStr(1, fName, "PasteObject", 1) Then
    PasteObject fPar(0)
ElseIf InStr(1, fName, "PasteObject_GetNtx", 1) Then
    fOut = PasteObject_GetNtx(fPar(0))
ElseIf InStr(1, fName, "PasteObject_NtxFileCheck", 1) Then
    fOut = PasteObject_NtxFileCheck(Val(fPar(0)))
ElseIf InStr(1, fName, "PasteObject_Local_Node", 1) Then
    PasteObject_Local_Node fPar(0)
ElseIf InStr(1, fName, "PasteObject_Local_Line", 1) Then
    PasteObject_Local_Line Val(fPar(0))
ElseIf InStr(1, fName, "CopyObject_Coding", 1) Then
    CopyObject_Coding fPar(0)
ElseIf InStr(1, fName, "CopyObject_Line", 1) Then
    CopyObject_Line StrToBool(fPar(0))
ElseIf InStr(1, fName, "CopyObject_Node", 1) Then
    CopyObject_Node StrToBool(fPar(0))
ElseIf InStr(1, fName, "CopyObject_Line_GetNodeRelativityId", 1) Then
    fOut = CopyObject_Line_GetNodeRelativityId(Val(fPar(0)))
ElseIf InStr(1, fName, "MeExeIdSet", 1) Then
    MeExeIdSet fPar(0)
ElseIf InStr(1, fName, "BehaviorIdSet", 1) Then
    BehaviorIdSet fPar(0)
ElseIf InStr(1, fName, "RedoSet", 1) Then
    RedoSet fPar(0)
ElseIf InStr(1, fName, "DeleteSelectObjcet", 1) Then
    DeleteSelectObjcet fPar(0)
ElseIf InStr(1, fName, "DeselectObjcet", 1) Then
    DeselectObjcet fPar(0)
ElseIf InStr(1, fName, "ChainSelection", 1) Then
    ChainSelection Val(fPar(0)), Val(fPar(1))
ElseIf InStr(1, fName, "DirectSelect", 1) Then
    DirectSelect fPar(0)
ElseIf InStr(1, fName, "AllSelection", 1) Then
    AllSelection fPar(0)
ElseIf InStr(1, fName, "ChainSelection_All", 1) Then
    ChainSelection_All Val(fPar(0))
ElseIf InStr(1, fName, "RedoBehavior", 1) Then
    RedoBehavior fPar(0)
ElseIf InStr(1, fName, "Redo_LineAdd_Save", 1) Then
    Redo_LineAdd_Save Val(fPar(0))
ElseIf InStr(1, fName, "Redo_LineDelete", 1) Then
    Redo_LineDelete Val(fPar(0))
ElseIf InStr(1, fName, "Redo_NodeEdit_NewNode", 1) Then
    Redo_NodeEdit_NewNode Val(fPar(0))
ElseIf InStr(1, fName, "Redo_NodeEdit_ReviseNode", 1) Then
    Redo_NodeEdit_ReviseNode Val(fPar(0)), fPar(1), fPar(2)
ElseIf InStr(1, fName, "Redo_NodeDelete", 1) Then
    Redo_NodeDelete Val(fPar(0))
ElseIf InStr(1, fName, "RevokeBehavior", 1) Then
    RevokeBehavior fPar(0)
ElseIf InStr(1, fName, "Revoke_LineAdd_Save", 1) Then
    Revoke_LineAdd_Save Val(fPar(0))
ElseIf InStr(1, fName, "Revoke_LineDelete", 1) Then
    Revoke_LineDelete Val(fPar(0))
ElseIf InStr(1, fName, "Revoke_NodeEdit_NewNode", 1) Then
    Revoke_NodeEdit_NewNode Val(fPar(0))
ElseIf InStr(1, fName, "Revoke_NodeEdit_ReviseNode", 1) Then
    Revoke_NodeEdit_ReviseNode Val(fPar(0)), fPar(1), fPar(2)
ElseIf InStr(1, fName, "Revoke_NodeDelete", 1) Then
    Revoke_NodeDelete Val(fPar(0))
ElseIf InStr(1, fName, "MapUpdata", 1) Then
    MapUpdata fPar(0)
ElseIf InStr(1, fName, "MapUpdata_DetermineTheBoundary", 1) Then
    MapUpdata_DetermineTheBoundary fPar(0)
ElseIf InStr(1, fName, "MapUpdata_AoVMove", 1) Then
    MapUpdata_AoVMove fPar(0)
ElseIf InStr(1, fName, "MapUpdata_AoVMove_Moving", 1) Then
    MapUpdata_AoVMove_Moving Val(fPar(0)), Val(fPar(1))
ElseIf InStr(1, fName, "绘制窗口世界视角", 1) Then
    绘制窗口世界视角 fPar(0)
ElseIf InStr(1, fName, "笔记对象绘制", 1) Then
    笔记对象绘制 fPar(0)
ElseIf InStr(1, fName, "MToZF", 1) Then
    fOut = MToZF(Val(fPar(0)))
ElseIf InStr(1, fName, "StrToBool", 1) Then
    fOut = StrToBool(StrToBool(fPar(0)))
ElseIf InStr(1, fName, "TwoPointDistance", 1) Then
    fOut = TwoPointDistance(Val(fPar(0)), Val(fPar(1)), Val(fPar(2)), Val(fPar(3)))
ElseIf InStr(1, fName, "OverlappingJudgment", 1) Then
    fOut = OverlappingJudgment(Val(fPar(0)), Val(fPar(1)), Val(fPar(2)), Val(fPar(3)), Val(fPar(4)))
ElseIf InStr(1, fName, "Rainbow_BackEnd", 1) Then
    fOut = Rainbow_BackEnd(Val(fPar(0)))
ElseIf InStr(1, fName, "Rainbow_RedEnd", 1) Then
    fOut = Rainbow_RedEnd(Val(fPar(0)))
ElseIf InStr(1, fName, "注册表注册", 1) Then
    fOut = 注册表注册(fPar(0), Val(fPar(1)))
ElseIf InStr(1, fName, "环境文件拷贝", 1) Then
    环境文件拷贝 fPar(0)
End If
CMD_In_Line_GetFunction = fOut
End Function
Public Function Judge_constants_and_variables(ByRef localVarNameList() As String, ByRef varName As String) As Long
Dim i As Long: Dim ucaseVN As String
ucaseVN = UCase(varName)
For i = 0 To UBound(publicVarName)
    If UCase(publicVarName(i)) = ucaseVN Then
        Judge_constants_and_variables = 10 + i: Exit Function
    End If
Next
For i = 0 To UBound(publicArrVarName)
    If UCase(publicArrVarName(i)) = ucaseVN Then
        Judge_constants_and_variables = 100 + i: Exit Function
    End If
Next
For i = 0 To UBound(publicFormName)
    If UCase(publicFormName(i)) = ucaseVN Then
        Judge_constants_and_variables = 1000 + i: Exit Function
    End If
Next
For i = 0 To UBound(localVarNameList)
    If UCase(localVarNameList(i)) = ucaseVN Then
        Judge_constants_and_variables = 10000 + i: Exit Function
    End If
Next
End Function
Public Function LoadPublicVar()
ReDim publicVarName(40)
publicVarName(0) = "nSum"
publicVarName(1) = "lSum"
publicVarName(2) = "copyNSum"
publicVarName(3) = "copyLSum"
publicVarName(4) = "NodeInputBackColor"
publicVarName(5) = "nodeEditLock"
publicVarName(6) = "nodeEditFormLock"
publicVarName(7) = "allNodeMoveLock"
publicVarName(8) = "nodeMoveLock"
publicVarName(9) = "regionalSelectLock"
publicVarName(10) = "selectMoveLock"
publicVarName(11) = "lineAddLock"
publicVarName(12) = "mapMoveLock"
publicVarName(13) = "mapGetMousePosLock"
publicVarName(14) = "nodePrintBeLock"
publicVarName(15) = "iconCompatible"
publicVarName(16) = "lineAddSource"
publicVarName(17) = "nodeEditAim"
publicVarName(18) = "nodeMoveAim"
publicVarName(19) = "nodeClickAim"
publicVarName(20) = "nodeTargetAim"
publicVarName(21) = "notePrintNodeId"
publicVarName(22) = "bHLSum"
publicVarName(23) = "redoSum"
publicVarName(24) = "ntxPath"
publicVarName(25) = "meExeId"
publicVarName(26) = "behaviorId"
publicVarName(27) = "redoId"
publicVarName(28) = "ProfilePath"
publicVarName(29) = "InstallPath"
publicVarName(30) = "nodeEditPos"
publicVarName(31) = "regionalSelectStart"
publicVarName(32) = "allNodeMoveStart"
publicVarName(33) = "nodeMoveStart"
publicVarName(34) = "mousePos"
publicVarName(35) = "mouseMapPos"
publicVarName(36) = "lineAddStrat"
publicVarName(37) = "mouseV3Pos"
publicVarName(38) = "zoomFactor"
publicVarName(39) = "magnification"
publicVarName(40) = "MainFormFontSize"
ReDim publicArrVarName(7)
publicArrVarName(0) = "node"
publicArrVarName(1) = "nodeLine"
publicArrVarName(2) = "copyNodeList"
publicArrVarName(3) = "copyLineList"
publicArrVarName(4) = "copyNIdList"
publicArrVarName(5) = "copyLIdList"
publicArrVarName(6) = "behaviorList"
publicArrVarName(7) = "redolist"
ReDim publicFormName(5)
publicFormName(0) = "Note"
publicFormName(1) = "NoteControlDesk"
publicFormName(2) = "NodePrint"
publicFormName(3) = "NodeInput"
publicFormName(4) = "NodeFind"
publicFormName(5) = "AboutNote"
ReDim ControlCharacter(5)
ControlCharacter(0) = "="
ControlCharacter(1) = "+"
ControlCharacter(2) = "-"
ControlCharacter(3) = "/"
ControlCharacter(4) = "*"
ControlCharacter(5) = "^"
ReDim publicFunctionName(0)
publicFunctionName(0) = "Msgbox"
End Function

