Attribute VB_Name = "Note_Var"
Public nodeSelectKeyDic As New Dictionary
Public lineSelectKeyDic As New Dictionary

Public lastNtx() As String
Public oneselfAddX As Double, oneselfAddY As Double, oneselfAddI As Double


Public nodeDefaultSize As Single
Public lineDefaultSize As Single
Public nodeInputFormHeight As Single
Public nodeInputFormWidth As Single
Public nodeInputFormTop As Single
Public nodeInputFormLeft As Single

Public 节点列表更新锁 As Boolean
Public 连接列表更新锁 As Boolean
Public needUpdataNodePrint As Boolean
Public fictitiousIndexLock As Boolean

Public mainFormMouseState As Boolean

Public node() As 节点
Public nodeLine() As 连接
Public copyNodeList() As 节点
Public copyLineList() As 连接
Public fictitiousNote() As 虚拟笔记

Public nSum  As Long
Public lSum   As Long
Public copyNSum As Long
Public copyLSum As Long
Public NodeInputBackColor As Long
Public copyNIdList() As Long
Public copyLIdList() As Long
Public rectangleLineColor As Long

Public nodeEditLock As Boolean '节点编辑锁
Public nodeEditFormLock As Boolean '节点编辑窗体状态锁
Public allNodeMoveLock As Boolean '全部节点移动锁
Public nodeMoveLock As Boolean
Public regionalSelectLock As Boolean '滚轮点击触发的区域选择锁
Public selectMoveLock As Boolean '选中的点移动锁
Public lineAddLock As Boolean
Public mapMoveLock As Boolean
Public mapGetMousePosLock As Boolean
Public nodePrintBeLock As Boolean
Public iconCompatible As Boolean
Public depthList() As Boolean
Public childNodeVisLock As Boolean

Public lineAddSource As Long
Public nodeEditAim As Long
Public nodeMoveAim As Long
Public nodeClickAim As Long
Public nodeTargetAim As Long
Public notePrintNodeId As Long
Public bHLSum As Long
Public redoSum As Long
Public fictitiousRootNodeId As Long

Public ntxPath As String
Public ntxPathNoName As String
Public meExeId As String
Public behaviorId As String
Public redoId As String
Public behaviorList() As String 'bHLSum 本数组上限索引
Public redolist() As String
Public ProfilePath As String
Public InstallPath As String
Public publicVarName() As String
Public publicArrVarName() As String
Public publicFormName() As String
Public ControlCharacter() As String
Public publicFunctionName() As String
Public childNodeVisNtxPath As String
Public fictitiousNtxPath As String
Public fictitiousIndexName As String

Public nodeEditPos As 二维坐标
Public regionalSelectStart As 二维坐标
Public allNodeMoveStart As 二维坐标
Public nodeMoveStart  As 二维坐标
Public mousePos As 二维坐标
Public mouseMapPos As 二维坐标
Public lineAddStrat As 二维坐标
Public angleOfView As 二维坐标

Public mouseV3Pos As 三维坐标

Public zoomFactor As Single
Public magnification As Single
Public MainFormFontSize As Single
Public saveNtxTime As Single
Public nodeAttributedToIntegers As Single
Public treeTxtToNtx_StartX As Single, treeTxtToNtx_StartY As Single, treeTxtToNtx_StepX As Single, treeTxtToNtx_StepY As Single
Public imageToNtx_StartX As Single, imageToNtx_StartY As Single, imageToNtx_StepX As Single, imageToNtx_StepY As Single

Public Const PI As Double = 3.14159265358979
Public Const VERSIONID As String = "Note2D_3"
Public Const NOTEFORMNAME As String = "节点笔记 - "
Public LINEBREAK As String
Public COPYLINEBREAK As String
Public NODELINEBREAK As String
Public DICBREAK As String
Public KEYBREAK As String
Public VALUEBREAK As String
Public Const PROFILENAME As String = "NoteConfig.ini"
Public Const TEXTINDENT As String = "    "

Public inputRecord() As String, cursorPos As Long

Public Sub PublicVarLoad()
    LINEBREAK = Chr(1)
    LINEBREAK = Chr(2)
    COPYLINEBREAK = Chr(3)
    DICBREAK = Chr(4)
    KEYBREAK = Chr(5)
    VALUEBREAK = Chr(6)
    nodeDefaultSize = 100
    lineDefaultSize = 2
    nodeInputFormHeight = 9750
    nodeInputFormWidth = 6480
    treeTxtToNtx_StartX = 1000
    treeTxtToNtx_StartY = 1000
    treeTxtToNtx_StepX = 1500
    treeTxtToNtx_StepY = 1000
    
    rectangleLineColor = 16443110
    nodeAttributedToIntegers = 3000
    
    fictitiousNtxPath = Environ("USERPROFILE") & "\Documents\Note\Fictitious\"
    ReDim inputRecord(0)
End Sub

Public Sub PublicVarLoad2()
    imageToNtx_StartX = 100
    imageToNtx_StartY = Note.height
    imageToNtx_StepX = 600
    imageToNtx_StepY = -600
End Sub
