VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Note 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Note"
   ClientHeight    =   9120
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   14760
   FillColor       =   &H00FFFF00&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000013&
   Icon            =   "Note_Main.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   9120
   ScaleWidth      =   14760
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer MapUpdataTimer 
      Interval        =   100
      Left            =   2040
      Top             =   8520
   End
   Begin VB.PictureBox GlobalView 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FAE6E6&
      FillColor       =   &H009AFA00&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2500
      Left            =   11000
      ScaleHeight     =   2475
      ScaleWidth      =   3585
      TabIndex        =   0
      Top             =   6500
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Timer PLC 
      Interval        =   500
      Left            =   120
      Top             =   8520
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer MainTime 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1440
      Top             =   8520
   End
   Begin RichTextLib.RichTextBox NodePrintBox 
      Height          =   555
      Left            =   2640
      TabIndex        =   1
      Top             =   8400
      Visible         =   0   'False
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   979
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Note_Main.frx":1084A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label PilotLight 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "  "
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   14480
      TabIndex        =   2
      ToolTipText     =   "绿色：修改已保存；黄色：笔记已存在但修改未保存；红色：新建笔记还未保存"
      Top             =   120
      Width           =   120
   End
   Begin VB.Menu 文件 
      Caption         =   "文件"
      Begin VB.Menu 新建笔记 
         Caption         =   "新建"
         Shortcut        =   ^N
      End
      Begin VB.Menu 打开笔记 
         Caption         =   "打开"
         Shortcut        =   ^O
      End
      Begin VB.Menu 保存笔记 
         Caption         =   "保存"
         Shortcut        =   ^S
      End
      Begin VB.Menu 另存为 
         Caption         =   "另存为"
      End
      Begin VB.Menu Cut3 
         Caption         =   "-"
      End
      Begin VB.Menu 退出 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu 编辑 
      Caption         =   "编辑"
      Begin VB.Menu 撤销 
         Caption         =   "撤销"
         Shortcut        =   ^Z
      End
      Begin VB.Menu 重做 
         Caption         =   "重做"
         Shortcut        =   ^Y
      End
      Begin VB.Menu Cut1 
         Caption         =   "-"
      End
      Begin VB.Menu 复制 
         Caption         =   "复制"
         Shortcut        =   ^C
      End
      Begin VB.Menu 剪切 
         Caption         =   "剪切"
         Shortcut        =   ^X
      End
      Begin VB.Menu 粘贴 
         Caption         =   "粘贴"
         Shortcut        =   ^V
      End
      Begin VB.Menu Cut4 
         Caption         =   "-"
      End
      Begin VB.Menu 查找 
         Caption         =   "查找"
         Shortcut        =   ^F
      End
      Begin VB.Menu Cut2 
         Caption         =   "-"
      End
      Begin VB.Menu 选显 
         Caption         =   "选显"
         Shortcut        =   ^R
      End
      Begin VB.Menu 全选 
         Caption         =   "全选"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu 视图 
      Caption         =   "视图"
      Begin VB.Menu 节点名显示 
         Caption         =   "节点名显示"
         Begin VB.Menu 显示全部节点名 
            Caption         =   "显示全部节点名"
            Checked         =   -1  'True
         End
         Begin VB.Menu 显示顺向节点名 
            Caption         =   "显示顺向节点名"
         End
         Begin VB.Menu 显示逆向节点名 
            Caption         =   "显示逆向节点名"
         End
         Begin VB.Menu Cut6 
            Caption         =   "-"
         End
         Begin VB.Menu 始终显示选点名 
            Caption         =   "始终显示选点名"
         End
         Begin VB.Menu Cut5 
            Caption         =   "-"
         End
         Begin VB.Menu 显示节点遍历ID 
            Caption         =   "显示节点遍历ID"
         End
      End
      Begin VB.Menu 节点连接显示 
         Caption         =   "节点连接显示"
         Begin VB.Menu 显示全部连接 
            Caption         =   "显示全部连接"
            Checked         =   -1  'True
         End
         Begin VB.Menu 显示顺向连接 
            Caption         =   "显示顺向连接"
         End
         Begin VB.Menu 显示逆向连接 
            Caption         =   "显示逆向连接"
         End
         Begin VB.Menu Cut7 
            Caption         =   "-"
         End
         Begin VB.Menu 始终显示选接 
            Caption         =   "始终显示选接"
         End
      End
      Begin VB.Menu 全局视图 
         Caption         =   "全局视图"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu 界面 
      Caption         =   "界面"
      Begin VB.Menu 主界面 
         Caption         =   "主界面"
         Begin VB.Menu 字体 
            Caption         =   "字体"
         End
         Begin VB.Menu 透明2 
            Caption         =   "透明"
            Begin VB.Menu 全高透明 
               Caption         =   "全高透明"
            End
            Begin VB.Menu 全半透明 
               Caption         =   "全半透明"
            End
            Begin VB.Menu 全低透明 
               Caption         =   "全低透明"
            End
         End
         Begin VB.Menu 背景色 
            Caption         =   "背景色"
         End
         Begin VB.Menu 文字颜色 
            Caption         =   "文字颜色"
         End
         Begin VB.Menu Cut9 
            Caption         =   "-"
         End
         Begin VB.Menu 彩虹圈 
            Caption         =   "彩虹圈"
            Checked         =   -1  'True
         End
         Begin VB.Menu 彩虹线 
            Caption         =   "彩虹线"
            Checked         =   -1  'True
         End
         Begin VB.Menu 流光溢彩 
            Caption         =   "流光溢彩"
         End
      End
      Begin VB.Menu 输入界面 
         Caption         =   "输入界面"
         Begin VB.Menu 透明 
            Caption         =   "透明"
            Begin VB.Menu 全高透明2 
               Caption         =   "全高透明"
            End
            Begin VB.Menu 全半透明2 
               Caption         =   "全半透明"
            End
            Begin VB.Menu 全低透明2 
               Caption         =   "全低透明"
            End
         End
         Begin VB.Menu 背景色2 
            Caption         =   "背景色"
         End
      End
      Begin VB.Menu 输出界面 
         Caption         =   "输出界面"
         Begin VB.Menu 置顶 
            Caption         =   "置顶"
            Shortcut        =   ^T
         End
         Begin VB.Menu 标签化 
            Caption         =   "标签化"
            Shortcut        =   ^L
         End
         Begin VB.Menu 透明3 
            Caption         =   "透明"
            Begin VB.Menu 全高透明3 
               Caption         =   "全高透明"
            End
            Begin VB.Menu 全半透明3 
               Caption         =   "全半透明"
            End
            Begin VB.Menu 全低透明3 
               Caption         =   "全低透明"
            End
         End
      End
   End
   Begin VB.Menu 功能 
      Caption         =   "功能"
      Begin VB.Menu 导出 
         Caption         =   "导出"
         Begin VB.Menu 节点树到TXT文本 
            Caption         =   "节点树到TXT文本"
         End
      End
      Begin VB.Menu 连接反转 
         Caption         =   "连接反转(R)"
      End
      Begin VB.Menu Cut11 
         Caption         =   "-"
      End
      Begin VB.Menu 自动保存间隔 
         Caption         =   "自动保存间隔"
      End
   End
   Begin VB.Menu 帮助 
      Caption         =   "帮助"
      Begin VB.Menu 控制台 
         Caption         =   "控制台"
         Shortcut        =   {F12}
         Visible         =   0   'False
      End
      Begin VB.Menu 关于节点笔记 
         Caption         =   "关于节点笔记"
      End
   End
End
Attribute VB_Name = "Note"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private saveNtxTimeNow As Single
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
Select Case KeyCode
    Case 27
        DeselectObjcet
    Case 46
        BehaviorIdSet
        DeleteSelectObjcet
    Case 107 '+
        RollerEventHandling False
    Case 109 '-
        RollerEventHandling True
    Case vbKey0
        MainCoordinateSystemZero mouseV3Pos
    Case 37
        MapUpdata_AoVMove_Moving -10, 0
    Case 38
        MapUpdata_AoVMove_Moving 0, 10
    Case 39
        MapUpdata_AoVMove_Moving 10, 0
    Case 40
        MapUpdata_AoVMove_Moving 0, -10
    Case vbKeyR
        连接反转_Click
End Select
End Sub

Private Sub Form_Load()
Dim dirPath As String
zoomFactor = 1
If App.LogMode <> 0 Then
    HookMouse Me.hWnd
End If
notePrintNodeId = -1
MeExeIdSet
ProfilePath = Environ("USERPROFILE") & "\Documents\Note\"
InstallPath = Environ("systemdrive") & "\ProgramData\Note\"
LoadProfile
Select Case 注册表注册("NodeNote", ".ntx")
    Case 0
        MsgBox "软件注册失败！请用管理员身份运行！"
    Case 1
        MsgBox "软件注册成功！"
End Select
MainCoordinateSystemDefinition
If 标签化.Checked = False Then NodePrint.Show
If Command <> "" Then
    dirPath = Replace(Command, """", "")
    NoteFileRead dirPath
Else
    newAddNote
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'C:2;S:1;A:4
Select Case Button
    Case 1
        nodeClickAim = NodeCheck(x, y)
        If nodeClickAim = -1 Then  '移动坐标系
            Select Case Shift
                Case 0
                    allNodeMoveLock = True
                    allNodeMoveStart.x = x: allNodeMoveStart.y = y
                    If lineAddLock = True Then lineAddLock = False
                Case 2
                    allNodeMoveStart.x = x: allNodeMoveStart.y = y
                    selectMoveLock = True
                Case 4
                    regionalSelectStart.x = x: regionalSelectStart.y = y
                    regionalSelectLock = True
            End Select
        Else
            Select Case Shift
                Case 0
                    If MultipointConnection = False Then
                        If lineAddLock = False Then
                            lineAddLock = True
                            lineAddStrat.x = x: lineAddStrat.y = y
                            lineAddSource = nodeClickAim
                        Else
                            BehaviorIdSet
                            If lineAddSource <> nodeClickAim Then
                                LineAdd lineAddSource, nodeClickAim
                            End If
                            lineAddLock = False
                        End If
                        nodeMoveLock = True
                        nodeMoveStart.x = x: nodeMoveStart.y = y
                    End If
                Case 4
                    ChainSelection nodeClickAim, 0
            End Select
        End If
    Case 2
        NodeEditeStart x, y
    Case 4
        Select Case Shift
            Case 4
                DirectSelect
        End Select
End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
mousePos.x = x: mousePos.y = y
mouseV3Pos.x = x: mouseV3Pos.y = y: mouseV3Pos.z = zoomFactor
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
allNodeMoveLock = False
nodeMoveLock = False
regionalSelectLock = False
selectMoveLock = False
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
NoteFileRead Data.Files(1)
End Sub

Private Sub Form_Resize()
If WindowState = 1 Then Exit Sub
MainCoordinateSystemDefinition
If Me.Height < 10000 Then Me.Enabled = False: Me.Height = 10000: Me.Enabled = True
If Me.Width < 15000 Then Me.Enabled = False: Me.Width = 15000: Me.Enabled = True
PilotLight.left = Me.Width * zoomFactor - 240 * zoomFactor
GlobalView.left = Me.Width * zoomFactor - GlobalView.Width - 120 * zoomFactor
GlobalView.Top = GlobalView.Height + 120 * zoomFactor
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveProfile
'UnHookMouse Me.hWnd
If App.LogMode <> 0 Then
    UnHookMouse Me.hWnd
End If
End
End Sub

Private Sub GlobalView_KeyDown(KeyCode As Integer, Shift As Integer)
Form_KeyDown KeyCode, Shift
End Sub

Private Sub GlobalView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim dx As Single: Dim dy As Single
If Button = 1 Or Button = 2 Then
    dx = Note.Width / 2 - x
    dy = Note.Height / 2 - y
    MapUpdata_AoVMove_Moving dx, dy
    mouseMapPos.x = Note.Width / 2
    mouseMapPos.y = Note.Height / 2
    mapMoveLock = True
End If
End Sub

Private Sub GlobalView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If mapMoveLock = True And mapGetMousePosLock = False Then mouseMapPos.x = x: mouseMapPos.y = y
End Sub

Private Sub GlobalView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
mapMoveLock = False
End Sub

Private Sub MainTime_Timer()
Updata

End Sub

Private Sub MapUpdataTimer_Timer()
If nSum > 0 And 全局视图.Checked = True Then MapUpdata
End Sub

Private Sub PLC_Timer()
Select Case noteSaveCheck
    Case 0
        PilotLight.BackColor = RGB(255, 0, 0)
    Case 1
        PilotLight.BackColor = RGB(255, 165, 0)
    Case 2
        PilotLight.BackColor = RGB(0, 255, 0)
End Select
On Error GoTo Er
If Clipboard.GetText = "" Then 粘贴.Enabled = False Else 粘贴.Enabled = True
If bHLSum < 1 Then 撤销.Enabled = False Else 撤销.Enabled = True
If redoSum < 1 Then 重做.Enabled = False Else 重做.Enabled = True
菜单单项控制
If saveNtxTime <> 0 Then
    saveNtxTimeNow = saveNtxTimeNow + 0.5
    If saveNtxTimeNow > saveNtxTime And PilotLight.BackColor = RGB(255, 165, 0) Then
        保存笔记_Click
        saveNtxTimeNow = 0
    End If
End If
Er:
End Sub
Private Function 菜单单项控制()
If 显示逆向节点名.Checked = True Then 显示顺向节点名.Checked = False: 显示全部节点名.Checked = False
If 显示顺向节点名.Checked = True Then 显示逆向节点名.Checked = False: 显示全部节点名.Checked = False
If 显示全部节点名.Checked = True Then 显示逆向节点名.Checked = False: 显示顺向节点名.Checked = False

If 显示逆向连接.Checked = True Then 显示顺向连接.Checked = False: 显示全部连接.Checked = False
If 显示顺向连接.Checked = True Then 显示逆向连接.Checked = False: 显示全部连接.Checked = False
If 显示全部连接.Checked = True Then 显示逆向连接.Checked = False: 显示顺向连接.Checked = False

End Function


Private Sub 保存笔记_Click()
Dim filePath As String
' 设置“CancelError”为 True
If Dir(ntxPath) = "" Then
    Note.CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' 设置标志
'    Note.CommonDialog1.Flags = cdlOFNHideReadOnly
    Note.CommonDialog1.Flags = cdlOFNOverwritePrompt
    ' 设置过滤器
    Note.CommonDialog1.Filter = "节点笔记 (*.ntx)|*.ntx|All Files (*.*)|*.*"
    ' 指定缺省的过滤器
    Note.CommonDialog1.FilterIndex = 1
    ' 显示“打开”对话框
    Note.CommonDialog1.ShowSave
    ' 显示选定文件的名字
    filePath = Note.CommonDialog1.FileName
    NoteFileWrite_201 filePath
Else
    NoteFileWrite_201 ntxPath
End If
Exit Sub
ErrHandler:
' 用户按了“取消”按钮
End Sub

Private Sub 背景色_Click()
With CommonDialog1
    .Flags = 1
    .color = Me.BackColor
    .ShowColor
    Me.BackColor = .color
End With
End Sub

Private Sub 背景色2_Click()
With CommonDialog1
    .Flags = 1
    .color = NodeInput.BackColor
    .ShowColor
    NodeInputBackColor = .color
End With
NodeInput.BackColor = NodeInputBackColor
End Sub

Private Sub 标签化_Click()
If 标签化.Checked = False Then
    标签化.Checked = True
    Unload NodePrint
Else
    标签化.Checked = False
    nodePrintBeLock = False
End If
End Sub

Private Sub 彩虹圈_Click()
If 彩虹圈.Checked = False Then 彩虹圈.Checked = True Else 彩虹圈.Checked = False
End Sub

Private Sub 彩虹线_Click()
If 彩虹线.Checked = False Then 彩虹线.Checked = True Else 彩虹线.Checked = False: 流光溢彩.Checked = False
End Sub

Private Sub 查找_Click()
NodeFind.Show
End Sub

Private Sub 撤销_Click()
RedoSet
RevokeBehavior
End Sub

Private Sub 打开笔记_Click()
Dim filePath As String
' 设置“CancelError”为 True
newAddNote
Note.CommonDialog1.CancelError = True
On Error GoTo ErrHandler
' 设置标志
Note.CommonDialog1.Flags = cdlOFNHideReadOnly
' 设置过滤器
Note.CommonDialog1.Filter = "节点笔记 (*.ntx)|*.ntx|All Files (*.*)|*.*"
' 指定缺省的过滤器
Note.CommonDialog1.FilterIndex = 1
' 显示“打开”对话框
Note.CommonDialog1.ShowOpen
' 显示选定文件的名字
filePath = Note.CommonDialog1.FileName
NoteFileRead filePath
Exit Sub
ErrHandler:
' 用户按了“取消”按钮
End Sub

Private Sub 复制_Click()
CopyObject False
End Sub

Private Sub 关于节点笔记_Click()
AboutNote.Show
End Sub

Private Sub 剪切_Click()
BehaviorIdSet
CopyObject True
End Sub

Private Sub 节点树到TXT文本_Click()
Dim filePath As String
Note.CommonDialog1.CancelError = True
On Error GoTo ErrHandler
' 设置标志
'    Note.CommonDialog1.Flags = cdlOFNHideReadOnly
Note.CommonDialog1.Flags = cdlOFNOverwritePrompt
' 设置过滤器
Note.CommonDialog1.Filter = "文本文档 (*.txt)|*.txt"
' 指定缺省的过滤器
Note.CommonDialog1.FilterIndex = 1
' 显示“打开”对话框
Note.CommonDialog1.ShowSave
' 显示选定文件的名字
filePath = Note.CommonDialog1.FileName
NodesToTxt filePath
Exit Sub
ErrHandler:
' 用户按了“取消”按钮
End Sub

Private Sub 控制台_Click()
NoteControlDesk.Show
End Sub

Private Sub 连接反转_Click()
ConnectionReversal
End Sub

Private Sub 另存为_Click()
Dim filePath As String
Note.CommonDialog1.CancelError = True
On Error GoTo ErrHandler
' 设置标志
'    Note.CommonDialog1.Flags = cdlOFNHideReadOnly
Note.CommonDialog1.Flags = cdlOFNOverwritePrompt
' 设置过滤器
Note.CommonDialog1.Filter = "节点笔记 (*.ntx)|*.ntx|All Files (*.*)|*.*"
' 指定缺省的过滤器
Note.CommonDialog1.FilterIndex = 1
' 显示“打开”对话框
Note.CommonDialog1.ShowSave
' 显示选定文件的名字
filePath = Note.CommonDialog1.FileName
NoteFileWrite_201 filePath
Exit Sub
ErrHandler:
' 用户按了“取消”按钮
End Sub

Private Sub 流光溢彩_Click()
If 流光溢彩.Checked = True Then
    流光溢彩.Checked = False
ElseIf 彩虹线.Checked = True Then
    流光溢彩.Checked = True
End If
End Sub

Private Sub 全高透明_Click()
If 全高透明.Checked = True Then
    全高透明.Checked = False: FormTransparent Me, 255
Else
    全高透明.Checked = True: FormTransparent Me, 50
    全半透明.Checked = False: 全低透明.Checked = False
End If
End Sub

Private Sub 全半透明_Click()
If 全半透明.Checked = True Then
    全半透明.Checked = False: FormTransparent Me, 255
Else
    全半透明.Checked = True: FormTransparent Me, 125
    全高透明.Checked = False: 全低透明.Checked = False
End If
End Sub

Private Sub 全低透明_Click()
If 全低透明.Checked = True Then
    全低透明.Checked = False: FormTransparent Me, 255
Else
    全低透明.Checked = True: FormTransparent Me, 200
    全半透明.Checked = False: 全高透明.Checked = False
End If
End Sub

Private Sub 全高透明2_Click()
If 全高透明2.Checked = True Then
    全高透明2.Checked = False: FormTransparent NodeInput, 255
Else
    全高透明2.Checked = True: FormTransparent NodeInput, 50
    全半透明2.Checked = False: 全低透明2.Checked = False
End If
End Sub

Private Sub 全半透明2_Click()
If 全半透明2.Checked = True Then
    全半透明2.Checked = False: FormTransparent NodeInput, 255
Else
    全半透明2.Checked = True: FormTransparent NodeInput, 125
    全高透明2.Checked = False: 全低透明2.Checked = False
End If
End Sub

Private Sub 全低透明2_Click()
If 全低透明2.Checked = True Then
    全低透明2.Checked = False: FormTransparent NodeInput, 255
Else
    全低透明2.Checked = True: FormTransparent NodeInput, 200
    全半透明2.Checked = False: 全高透明2.Checked = False
End If
End Sub
Private Sub 全高透明3_Click()
If 全高透明3.Checked = True Then
    全高透明3.Checked = False: FormTransparent NodePrint, 255
Else
    全高透明3.Checked = True: FormTransparent NodePrint, 50
    全半透明3.Checked = False: 全低透明3.Checked = False
End If
End Sub

Private Sub 全半透明3_Click()
If 全半透明3.Checked = True Then
    全半透明3.Checked = False: FormTransparent NodePrint, 255
Else
    全半透明3.Checked = True: FormTransparent NodePrint, 125
    全高透明3.Checked = False: 全低透明3.Checked = False
End If
End Sub

Private Sub 全低透明3_Click()
If 全低透明3.Checked = True Then
    全低透明3.Checked = False: FormTransparent NodePrint, 255
Else
    全低透明3.Checked = True: FormTransparent NodePrint, 200
    全半透明3.Checked = False: 全高透明3.Checked = False
End If
End Sub
Private Sub 全局视图_Click()
If 全局视图.Checked = True Then
    全局视图.Checked = False: GlobalView.Visible = False
Else
    全局视图.Checked = True: GlobalView.Visible = True
End If
End Sub

Private Sub 全选_Click()
AllSelection
End Sub

Private Sub 始终显示选点名_Click()
If 始终显示选点名.Checked = True Then 始终显示选点名.Checked = False Else 始终显示选点名.Checked = True
End Sub

Private Sub 始终显示选接_Click()
If 始终显示选接.Checked = True Then 始终显示选接.Checked = False Else 始终显示选接.Checked = True
End Sub

Private Sub 退出_Click()
End
End Sub

Private Sub 文字颜色_Click()
With CommonDialog1
    .Flags = 1
    .color = Me.ForeColor
    .ShowColor
    Me.ForeColor = .color
End With
End Sub

Private Sub 显示节点遍历ID_Click()
If 显示节点遍历ID.Checked = True Then 显示节点遍历ID.Checked = False Else 显示节点遍历ID.Checked = True
End Sub

Private Sub 显示逆向节点名_Click()
If 显示逆向节点名.Checked = False Then
    显示逆向节点名.Checked = True: 显示全部节点名.Checked = False: 显示顺向节点名.Checked = False
Else
    显示逆向节点名.Checked = False
End If

End Sub

Private Sub 显示逆向连接_Click()
If 显示逆向连接.Checked = False Then
    显示逆向连接.Checked = True: 显示全部连接.Checked = False: 显示顺向连接.Checked = False
Else
    显示逆向连接.Checked = False
End If

End Sub

Private Sub 显示全部连接_Click()
If 显示全部连接.Checked = False Then
    显示全部连接.Checked = True: 显示顺向连接.Checked = False: 显示逆向连接.Checked = False
Else
    显示全部连接.Checked = False
End If

End Sub

Private Sub 显示顺向节点名_Click()
If 显示顺向节点名.Checked = False Then
    显示顺向节点名.Checked = True: 显示全部节点名.Checked = False: 显示逆向节点名.Checked = False
Else
    显示顺向节点名.Checked = False
End If

End Sub

Private Sub 显示顺向连接_Click()
If 显示顺向连接.Checked = False Then
    显示顺向连接.Checked = True: 显示全部连接.Checked = False: 显示逆向连接.Checked = False
Else
    显示顺向连接.Checked = False
End If

End Sub

Private Sub 显示全部节点名_Click()
If 显示全部节点名.Checked = False Then
    显示全部节点名.Checked = True: 显示顺向节点名.Checked = False: 显示逆向节点名.Checked = False
Else
    显示全部节点名.Checked = False
End If

End Sub

Private Sub 新建笔记_Click()
newAddNote
End Sub

Private Sub 选显_Click()
SelectDisplayObjcet
End Sub

Private Sub 粘贴_Click()
BehaviorIdSet
PasteObject
End Sub

Private Sub 置顶_Click()
If 置顶.Checked = True Then
    置顶.Checked = False: FormStick NodePrint, False
Else
    置顶.Checked = True: FormStick NodePrint, True
End If
Me.SetFocus
End Sub

Private Sub 重做_Click()
RedoBehavior
End Sub

Private Sub 自动保存间隔_Click()
On Error GoTo Er
saveNtxTime = Val(InputBox("请输入自动保存时间间隔(单位:秒,输入0代表不自动保存！)", "节点笔记自动保存时间间隔设置"))
Exit Sub
Er:
MsgBox "设置失败！", , "警告！"
saveNtxTime = 0
End Sub

Private Sub 字体_Click()
On Error GoTo Er
With CommonDialog1
    .Flags = 1
    .FontName = Me.Font.Name
    .FontBold = Me.Font.Bold
    .FontSize = MainFormFontSize
    .FontItalic = Me.Font.Italic
    .FontUnderline = Me.Font.Underline
    .FontStrikethru = Me.Font.Strikethrough
    .ShowFont
    Me.Font.Name = .FontName '字体名称
    Me.Font.Bold = .FontBold  '加粗？
    MainFormFontSize = .FontSize '字体大小
    Me.Font.Italic = .FontItalic '倾斜？
    Me.Font.Underline = .FontUnderline '下划线？
    Me.Font.Strikethrough = .FontStrikethru '删除线
End With
Er:
End Sub
