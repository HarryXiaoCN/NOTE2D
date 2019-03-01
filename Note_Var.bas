Attribute VB_Name = "Note_Var"
Public node() As 节点
Public nodeLine() As 连接
Public copyNodeList() As 节点
Public copyLineList() As 连接

Public nSum  As Long
Public lSum   As Long
Public copyNSum As Long
Public copyLSum As Long
Public NodeInputBackColor As Long
Public copyNIdList() As Long
Public copyLIdList() As Long


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

Public lineAddSource As Long
Public nodeEditAim As Long
Public nodeMoveAim As Long
Public nodeClickAim As Long
Public nodeTargetAim As Long
Public notePrintNodeId As Long
Public bHLSum As Long
Public redoSum As Long

Public ntxPath As String
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

Public nodeEditPos As 二维坐标
Public regionalSelectStart As 二维坐标
Public allNodeMoveStart As 二维坐标
Public nodeMoveStart  As 二维坐标
Public mousePos As 二维坐标
Public mouseMapPos As 二维坐标
Public lineAddStrat As 二维坐标

Public mouseV3Pos As 三维坐标

Public zoomFactor As Single
Public magnification As Single
Public MainFormFontSize As Single
Public saveNtxTime As Single

Public Const PI As Double = 3.14159265358979
Public Const VERSIONID As String = "Note2D_1"
Public Const NOTEFORMNAME As String = "Note - "
Public Const LINEBREAK As String = "^|`"
Public Const COPYLINEBREAK As String = "^CoPy`"
Public Const NODELINEBREAK As String = "^||`"
Public Const PROFILENAME As String = "NoteConfig.ini"
Public Const TEXTINDENT As String = "    "
