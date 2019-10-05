Attribute VB_Name = "Note_Type"
Public Type 节点
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
Public Type 连接
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
Public Type 二维坐标
    X As Single
    Y As Single
End Type
Public Type 三维坐标
    X As Single
    Y As Single
    z As Single
End Type
Public Type 四元数
    xS As Single
    yS As Single
    xE As Single
    yE As Single
End Type
Public Type 边界
    up As Single
    down As Single
    left As Single
    right As Single
End Type
Public Type 虚拟节点
    realityX As Single
    realityY As Single
    X As Single
    Y As Single
    t As String
    setSize As Single
    content As String
    setColor As Long
End Type
Public Type 虚拟连接
    direction As Byte
    realityId As Long
    content As String
    size As Single
    Source As Long
    target As Long
End Type
Public Type 虚拟笔记
    be As Boolean
    node() As 虚拟节点
    nodeLine() As 虚拟连接
End Type
