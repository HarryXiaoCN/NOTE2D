Attribute VB_Name = "Note_Type"
Public Type 节点
    b As Boolean
    select As Boolean
    forward As Boolean
    backward As Boolean
    depth As Long
    x As Single
    y As Single
    t As String
    size As Single
    content As String
    color As Long
End Type
Public Type 连接
    b As Boolean
    select As Boolean
    search As Boolean
    forward As Boolean
    backward As Boolean
    source As Long
    target As Long
End Type
Public Type 二维坐标
    x As Single
    y As Single
End Type
Public Type 三维坐标
    x As Single
    y As Single
    z As Single
End Type
Public Type 边界
    up As Single
    down As Single
    left As Single
    right As Single
End Type
