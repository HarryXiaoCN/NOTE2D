Attribute VB_Name = "Note_Var"
Public node() As �ڵ�
Public nodeLine() As ����
Public copyNodeList() As �ڵ�
Public copyLineList() As ����

Public nSum  As Long
Public lSum   As Long
Public copyNSum As Long
Public copyLSum As Long
Public NodeInputBackColor As Long
Public copyNIdList() As Long
Public copyLIdList() As Long


Public nodeEditLock As Boolean '�ڵ�༭��
Public nodeEditFormLock As Boolean '�ڵ�༭����״̬��
Public allNodeMoveLock As Boolean 'ȫ���ڵ��ƶ���
Public nodeMoveLock As Boolean
Public regionalSelectLock As Boolean '���ֵ������������ѡ����
Public selectMoveLock As Boolean 'ѡ�еĵ��ƶ���
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
Public behaviorList() As String 'bHLSum ��������������
Public redolist() As String
Public ProfilePath As String
Public InstallPath As String
Public publicVarName() As String
Public publicArrVarName() As String
Public publicFormName() As String
Public ControlCharacter() As String
Public publicFunctionName() As String

Public nodeEditPos As ��ά����
Public regionalSelectStart As ��ά����
Public allNodeMoveStart As ��ά����
Public nodeMoveStart  As ��ά����
Public mousePos As ��ά����
Public mouseMapPos As ��ά����
Public lineAddStrat As ��ά����

Public mouseV3Pos As ��ά����

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
