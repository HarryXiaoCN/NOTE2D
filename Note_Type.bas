Attribute VB_Name = "Note_Type"
Public Type �ڵ�
    b As Boolean
    select As Boolean
    forward As Boolean
    backward As Boolean
    depth As Long
    toTreeTxtLock As Boolean
    X As Single
    Y As Single
    t As String
    size As Single
    setSize As Single
    content As String
    color As Long
    setColor As Long
End Type
Public Type ����
    b As Boolean
    select As Boolean
    search As Boolean
    forward As Boolean
    backward As Boolean
    content As String
    size As Single
    Source As Long
    target As Long
End Type
Public Type ��ά����
    X As Single
    Y As Single
End Type
Public Type ��ά����
    X As Single
    Y As Single
    z As Single
End Type
Public Type ��Ԫ��
    xS As Single
    yS As Single
    xE As Single
    yE As Single
End Type
Public Type �߽�
    up As Single
    down As Single
    left As Single
    right As Single
End Type
Public Type ����ڵ�
    realityX As Single
    realityY As Single
    X As Single
    Y As Single
    t As String
    setSize As Single
    content As String
    setColor As Long
End Type
Public Type ��������
    direction As Byte
    realityId As Long
    content As String
    size As Single
    Source As Long
    target As Long
End Type
Public Type ����ʼ�
    be As Boolean
    node() As ����ڵ�
    nodeLine() As ��������
End Type
