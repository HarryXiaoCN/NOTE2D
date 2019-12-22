VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Note 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Note"
   ClientHeight    =   9120
   ClientLeft      =   6375
   ClientTop       =   3465
   ClientWidth     =   14760
   DrawWidth       =   2
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   9120
   ScaleWidth      =   14760
   Begin VB.Timer ActionTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   7920
   End
   Begin VB.PictureBox 子节点视图容器 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   1200
      ScaleHeight     =   7785
      ScaleWidth      =   12825
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   12855
      Begin VB.PictureBox 子节点视图标题栏 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   12825
         TabIndex        =   9
         Top             =   0
         Width           =   12855
         Begin VB.Label 子节点视图按钮 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "○"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   315
            Index           =   2
            Left            =   0
            TabIndex        =   12
            ToolTipText     =   "打开预览笔记"
            Top             =   0
            Width           =   375
         End
         Begin VB.Label 子节点视图按钮 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "□"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   315
            Index           =   1
            Left            =   12000
            TabIndex        =   11
            ToolTipText     =   "最大/最小化子节点视图"
            Top             =   0
            Width           =   375
         End
         Begin VB.Label 子节点视图按钮 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "×"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   0
            Left            =   12360
            TabIndex        =   10
            ToolTipText     =   "关闭子节点视图"
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.PictureBox 子节点视图 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   10000
         Left            =   -2000
         ScaleHeight     =   9975
         ScaleWidth      =   9975
         TabIndex        =   8
         Top             =   -2000
         Width           =   10000
      End
   End
   Begin VB.PictureBox 位图输出器 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   2520
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox 位图读取缓存器 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   1320
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox 打印缓存器 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      ScaleHeight     =   945
      ScaleWidth      =   1065
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox RTBtemp 
      Height          =   495
      Left            =   6600
      TabIndex        =   3
      Top             =   8400
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"Note_Main.frx":700A
   End
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
      TabStop         =   0   'False
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
      Height          =   1755
      Left            =   2640
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   3096
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Note_Main.frx":70B7
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
   Begin VB.Label 放缩倍率显示标题 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   75
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
      Begin VB.Menu wencut2 
         Caption         =   "-"
      End
      Begin VB.Menu 打印 
         Caption         =   "打印"
         Begin VB.Menu 打印成PNG图片 
            Caption         =   "打印成PNG图片"
         End
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
      Begin VB.Menu bjcut4 
         Caption         =   "-"
      End
      Begin VB.Menu 印窃 
         Caption         =   "印窃（Ctrl+Shift+C）"
      End
      Begin VB.Menu Cut4 
         Caption         =   "-"
      End
      Begin VB.Menu 查找 
         Caption         =   "查找"
         Shortcut        =   ^F
      End
      Begin VB.Menu 有损替换 
         Caption         =   "有损替换"
         Shortcut        =   ^H
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
      Begin VB.Menu jmcut1 
         Caption         =   "-"
      End
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
         Begin VB.Menu 文字颜色 
            Caption         =   "文字颜色"
         End
         Begin VB.Menu bjcut1 
            Caption         =   "-"
         End
         Begin VB.Menu 背景色 
            Caption         =   "背景色"
         End
         Begin VB.Menu 背景图 
            Caption         =   "背景图"
         End
         Begin VB.Menu 删除背景图 
            Caption         =   "删除背景图"
         End
         Begin VB.Menu Cut9 
            Caption         =   "-"
         End
         Begin VB.Menu 矩点 
            Caption         =   "矩点"
         End
         Begin VB.Menu 矩线 
            Caption         =   "矩线"
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
      Begin VB.Menu jmcut2 
         Caption         =   "-"
      End
      Begin VB.Menu 全局视图 
         Caption         =   "全局视图"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu 操作 
      Caption         =   "操作"
      Begin VB.Menu 连接反转 
         Caption         =   "连接反转(R)"
      End
      Begin VB.Menu 圆阵节点 
         Caption         =   "圆阵节点(A)"
      End
      Begin VB.Menu 连接内容 
         Caption         =   "连接内容(C)"
      End
      Begin VB.Menu 深度上色 
         Caption         =   "深度上色(N)"
      End
      Begin VB.Menu 波化节点 
         Caption         =   "波化节点(W)"
      End
      Begin VB.Menu 像素节点 
         Caption         =   "像素节点(P)"
      End
      Begin VB.Menu 坐标化整 
         Caption         =   "坐标化整(V)"
      End
   End
   Begin VB.Menu 功能 
      Caption         =   "功能"
      Begin VB.Menu 导入 
         Caption         =   "导入"
         Begin VB.Menu 导入文本文件 
            Caption         =   "树状TAB分隔格式文本"
         End
         Begin VB.Menu 导入位图 
            Caption         =   "BMP位图"
         End
         Begin VB.Menu 导入TXT文章 
            Caption         =   "TXT文章"
         End
      End
      Begin VB.Menu 导出 
         Caption         =   "导出"
         Begin VB.Menu 导出文本文件 
            Caption         =   "树状TAB分割式文本"
            Shortcut        =   ^E
         End
         Begin VB.Menu 导出位图 
            Caption         =   "BMP位图"
            Shortcut        =   ^I
         End
         Begin VB.Menu 导出TXT文章 
            Caption         =   "TXT文章"
         End
      End
      Begin VB.Menu gncut6 
         Caption         =   "-"
      End
      Begin VB.Menu 选域消点 
         Caption         =   "选域消点(Shift+N)"
      End
      Begin VB.Menu 选域消线 
         Caption         =   "选域消线(Shift+L)"
      End
      Begin VB.Menu gncut7 
         Caption         =   "-"
      End
      Begin VB.Menu 节点归一 
         Caption         =   "节点归一(Ctrl+Shift+O)"
      End
      Begin VB.Menu 节点归整 
         Caption         =   "节点归整"
      End
      Begin VB.Menu Cut11 
         Caption         =   "-"
      End
      Begin VB.Menu 节点清单 
         Caption         =   "节点清单"
      End
      Begin VB.Menu 连接清单 
         Caption         =   "连接清单"
      End
      Begin VB.Menu gncut3 
         Caption         =   "-"
      End
      Begin VB.Menu 设置默认节点大小 
         Caption         =   "设置默认节点大小"
      End
      Begin VB.Menu 设置默认连接宽度 
         Caption         =   "设置默认连接宽度"
      End
      Begin VB.Menu gncut4 
         Caption         =   "-"
      End
      Begin VB.Menu 绘图刷新间隔 
         Caption         =   "绘图刷新间隔"
      End
      Begin VB.Menu 自动保存间隔 
         Caption         =   "自动保存间隔"
      End
      Begin VB.Menu gncut5 
         Caption         =   "-"
      End
      Begin VB.Menu RGB色与VBColor互转工具 
         Caption         =   "RGB色与VBColor互转工具"
      End
      Begin VB.Menu 打开联想节点文件目录 
         Caption         =   "打开联想节点文件目录"
      End
      Begin VB.Menu 打开网络接口 
         Caption         =   "打开网络接口"
      End
   End
   Begin VB.Menu 帮助 
      Caption         =   "帮助"
      Begin VB.Menu 控制台 
         Caption         =   "控制台"
         Shortcut        =   {F12}
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
Public pilotLightColor As Long, pilotLightX0 As Single, pilotLightY0 As Single, pilotLightX1 As Single, pilotLightY1 As Single
Public homeBackPicPath As String, updataSpeed As Long, 圆阵半径记忆 As Single
Private childNodeVisOld As 二维坐标, childNodeVisPos As 四元数, childNodeFormOld As 二维坐标

Private Sub ActionTimer_Timer()
    ObjectAction
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print KeyCode; ; Shift
    Select Case KeyCode
        Case vbKeyA
            If Shift = 0 Then
                圆阵节点_Click
            End If
        Case 13 '回车
            确认创建虚拟节点
        Case 27 'ESC
            DeselectObjcet
            If 子节点视图容器.Visible Then
                子节点视图按钮_Click 0
            End If
            fictitiousIndexLock = False
        Case vbKeyW
            If Shift = 0 Then 波化节点_Click
        Case vbKeyP
            If Shift = 0 Then 像素节点_Click
        Case vbKeyV
            If Shift = 0 Then 坐标化整_Click
        Case vbKeyO
            If Shift = 3 Then
                归一化选中节点
            End If
        Case 46
            BehaviorIdSet
            DeleteSelectObjcet
        Case 107, 187 '+
            If Shift = 2 Then
                RollerEventHandling False
            ElseIf Shift = 0 Then
                节点连接大小变化 10, 1
            End If
        Case vbKeyC
            If Shift = 3 Then
                印窃_Click
            ElseIf Shift = 0 Then
                连接内容_Click
            End If
        Case 109, 189 '-
            If Shift = 2 Then
                RollerEventHandling True
            ElseIf Shift = 0 Then
                节点连接大小变化 -10, -1
            End If
'        Case vbKey0
'            MainCoordinateSystemZero mouseV3Pos
        Case 37
            MapUpdata_AoVMove_Moving -10, 0
        Case 38
            MapUpdata_AoVMove_Moving 0, 10
        Case 39
            MapUpdata_AoVMove_Moving 10, 0
        Case 40
            MapUpdata_AoVMove_Moving 0, -10
        Case 192
            NoteControlDesk.Show
        Case vbKeyR
            连接反转_Click
        Case vbKeyL
            If Shift = 1 Then
                取消选区内的所有连接
            End If
        Case vbKeyN
            If Shift = 1 Then
                取消选区内的所有节点
            ElseIf Shift = 0 Then
                设置子节点颜色
            End If
        Case 48 To 57
            If Shift = 2 Then
                节点连接选域设置 KeyCode - 48
            ElseIf Shift = 0 Then
                节点连接选域使用 KeyCode - 48
            ElseIf Shift = 1 Then
                节点连接选域删除 KeyCode - 48
            End If
    End Select
End Sub
Private Sub 确认创建虚拟节点()
    Dim i As Long, j As Long
    BehaviorIdSet
    If fictitiousIndexLock Then
        For i = 0 To UBound(fictitiousNote)
            With fictitiousNote(i)
                If .be Then
                    For j = 0 To UBound(.nodeLine)
                        If .nodeLine(j).direction = 1 Then
                            LineAdd fictitiousRootNodeId, NodeEdit_NewNode(.node(.nodeLine(j).target).t, .node(.nodeLine(j).target).content, .node(.nodeLine(j).target).setColor, .node(.nodeLine(j).target).setSize, .node(.nodeLine(j).target).realityX, .node(.nodeLine(j).target).realityY), .nodeLine(j).content, .nodeLine(j).size
                        ElseIf .nodeLine(j).direction = 2 Then
                            LineAdd NodeEdit_NewNode(.node(.nodeLine(j).Source).t, .node(.nodeLine(j).Source).content, .node(.nodeLine(j).Source).setColor, .node(.nodeLine(j).Source).setSize, .node(.nodeLine(j).Source).realityX, .node(.nodeLine(j).Source).realityY), fictitiousRootNodeId, .nodeLine(j).content, .nodeLine(j).size
                        ElseIf .nodeLine(j).direction = 3 Then
                            LineAdd fictitiousRootNodeId, .nodeLine(j).realityId, .nodeLine(j).content, .nodeLine(j).size
                        ElseIf .nodeLine(j).direction = 4 Then
                            LineAdd .nodeLine(j).realityId, fictitiousRootNodeId, .nodeLine(j).content, .nodeLine(j).size
                        End If
                    Next
                End If
            End With
        Next
        fictitiousIndexLock = False
    End If
End Sub
Private Sub 归一化选中节点()
    Dim i As Long
    BehaviorIdSet
    For i = 0 To nSum
        With node(i)
            If .b Then
                If .select = True Or nodeTargetAim = i Then
                    节点去重 i
                End If
            End If
        End With
    Next
End Sub
Private Function 节点去重(sN As Long)
    Dim i As Long
    For i = sN + 1 To nSum
        With node(i)
            If .b Then
                If .t = node(sN).t And .content = node(sN).content And .setColor = node(sN).setColor And .setSize = node(sN).setSize Then
                    连接转换 i, sN
                    NodeDelete i
                End If
            End If
        End With
    Next
End Function
Private Function 连接转换(aN As Long, rN As Long)
    Dim i As Long
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                If .Source = aN Then
                    If .target = rN Then
                        LineDelete i
                    Else
                        LineReplace i, 0, rN, aN
                    End If
                    
                ElseIf .target = aN Then
                    If .Source = rN Then
                        LineDelete i
                    Else
                        LineReplace i, 1, rN, aN
                    End If
                End If
            End If
        End With
    Next
End Function
Private Sub 设置子节点颜色()
    If nodeTargetAim <> -1 Then
        NodeColorSelectForm.锁定母节点序号 = nodeTargetAim
        NodeColorSelectForm.Show 1
    End If
End Sub
Private Sub 圆阵子节点()
    Dim 缓存 As String, nidT As Long
    nidT = 未选中代替(nodeTargetAim)
    If nidT <> -1 Then
        缓存 = InBox("请输入圆阵半径[1000,100000]：", 圆阵半径记忆)
        If promptBoxSelect = 0 Then
            圆阵半径记忆 = 限制数值(Val(缓存), 1000, 100000)
            NodeArray nidT, 圆阵半径记忆
        End If
    End If
End Sub
Private Sub 取消选区内的所有节点()
    Dim i As Long
    For i = 0 To nSum
        With node(i)
            If .b Then
                .select = False
            End If
        End With
    Next
End Sub
Private Sub 取消选区内的所有连接()
    Dim i As Long
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                .select = False
            End If
        End With
    Next
End Sub
Private Sub 节点连接选域删除(ByVal key As String)
    If nodeSelectKeyDic.Exists(key) Then
        nodeSelectKeyDic.Remove key
        lineSelectKeyDic.Remove key
    End If
End Sub
Private Sub 节点连接选域使用(ByVal key As String)
    Dim i As Long
    If nodeSelectKeyDic.Exists(key) Then
        For i = 0 To nSum
            With node(i)
                If .b Then
                    If InStr(1, nodeSelectKeyDic(key), "," & i & ",") > 0 Then
                        .select = True
                    Else
                        .select = False
                    End If
                End If
            End With
        Next
        For i = 0 To lSum
            With nodeLine(i)
                If .b Then
                    If InStr(1, lineSelectKeyDic(key), "," & i & ",") > 0 Then
                        .select = True
                    Else
                        .select = False
                    End If
                End If
            End With
        Next
    End If
End Sub
Private Sub 节点连接选域设置(ByVal key As String)
    Dim i As Long, n As String, l As String
    n = ","
    For i = 0 To nSum
        With node(i)
            If .b Then
                If .select Then
                    n = n & i & ","
                End If
            End If
        End With
    Next
    nodeSelectKeyDic.Add key, n
    l = ","
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                If .select Then
                    l = l & i & ","
                End If
            End If
        End With
    Next
    lineSelectKeyDic.Add key, l
End Sub
Private Sub 节点连接大小变化(点增量 As Single, 线增量 As Single)
    Dim i As Long
    For i = 0 To nSum
        With node(i)
            If .b Then
                If nodeTargetAim = i Or .select = True Then
                    .setSize = .setSize + 点增量
                    If .setSize > 500 Then
                        .setSize = 500
                    ElseIf .setSize < 50 Then
                        .setSize = 50
                    End If
                End If
            End If
        End With
    Next
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                If .select Then
                    .size = .size + 线增量
                    If .size > 10 Then
                        .size = 10
                    ElseIf .size < 1 Then
                        .size = 1
                    End If
                End If
            End If
        End With
    Next
End Sub
Private Sub 连接内容赋予(内容 As String)
    Dim i As Long
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                If .select Then
                    .content = 内容
                End If
            End If
        End With
    Next
End Sub
Private Sub 复制节点为纯文本()
    Clipboard.Clear
    Clipboard.SetText nodeToTxt
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim dirPath As String
    PublicVarLoad2
    zoomFactor = 1: 圆阵半径记忆 = 1000
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
    子节点视图容器.Scale (0, 子节点视图容器.height)-(子节点视图容器.width, 0)
    子节点视图标题栏.Scale (0, 子节点视图标题栏.height)-(子节点视图标题栏.width, 0)
    If 标签化.Checked = False Then NodePrint.Show
    If Command <> "" Then
        pilotLightColor = RGB(0, 255, 0)
        dirPath = Replace(Command, """", "")
'        MsgBox dirPath
        NoteFileRead dirPath
    Else
        pilotLightColor = RGB(255, 0, 0)
        newAddNote
    End If
    FictitiousNtxLoad
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'C:2;S:1;A:4
mainFormMouseState = True
Select Case Button
    Case 1
        nodeClickAim = NodeCheck(X, Y)
        If nodeClickAim = -1 Then  '移动坐标系
            Select Case Shift
                Case 0
                    allNodeMoveLock = True
                    allNodeMoveStart.X = X: allNodeMoveStart.Y = Y
                    If lineAddLock = True Then lineAddLock = False
                Case 2
                    allNodeMoveStart.X = X: allNodeMoveStart.Y = Y
                    selectMoveLock = True
                Case 4
                    regionalSelectStart.X = X: regionalSelectStart.Y = Y
                    regionalSelectLock = True
            End Select
        Else
            Select Case Shift
                Case 0
                    If MultipointConnection = False Then
                        If lineAddLock = False Then
                            lineAddLock = True
                            lineAddStrat.X = X: lineAddStrat.Y = Y
                            lineAddSource = nodeClickAim
                        Else
                            BehaviorIdSet
                            If lineAddSource <> nodeClickAim Then
                                LineAdd lineAddSource, nodeClickAim, "", lineDefaultSize
                            End If
                            lineAddLock = False
                        End If
                        nodeMoveLock = True
                        nodeMoveStart.X = X: nodeMoveStart.Y = Y
                    End If
                Case 4
                    ChainSelection nodeClickAim, 0
            End Select
        End If
    Case 2
        NodeEditeStart X, Y
    Case 4
        Select Case Shift
            Case 4
                DirectSelect
        End Select
End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mousePos.X = X: mousePos.Y = Y
    mouseV3Pos.X = X: mouseV3Pos.Y = Y: mouseV3Pos.z = zoomFactor
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mainFormMouseState = False
allNodeMoveLock = False
nodeMoveLock = False
regionalSelectLock = False
selectMoveLock = False
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
NoteFileRead Data.Files(1)
End Sub

Private Sub Form_Resize()
If WindowState = 1 Then Exit Sub
MainCoordinateSystemDefinition
If Me.height < 4000 Then Me.Enabled = False: Me.height = 4000: Me.Enabled = True
If Me.width < 4000 Then Me.Enabled = False: Me.width = 4000: Me.Enabled = True
'PilotLight.left = Me.Width * zoomFactor - 240 * zoomFactor
Note.pilotLightX0 = Me.width - 300
Note.pilotLightX1 = Me.width - 230
Note.pilotLightY0 = Me.height - 300
Note.pilotLightY1 = Me.height - 200

GlobalView.left = Me.width * zoomFactor - GlobalView.width - 120 * zoomFactor
GlobalView.Top = GlobalView.height + 120 * zoomFactor
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveProfile
'UnHookMouse Me.hWnd
If App.LogMode <> 0 Then
    UnHookMouse Me.hWnd
End If
End
End Sub

Private Sub GlobalView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim dX As Single: Dim dY As Single
    If Button = 1 Or Button = 2 Then
        dX = Note.width / 2 - X
        dY = Note.height / 2 - Y
        MapUpdata_AoVMove_Moving dX, dY
        mouseMapPos.X = Note.width / 2
        mouseMapPos.Y = Note.height / 2
        mapMoveLock = True
    End If
End Sub

Private Sub GlobalView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If mapMoveLock = True And mapGetMousePosLock = False Then mouseMapPos.X = X: mouseMapPos.Y = Y
End Sub

Private Sub GlobalView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mapMoveLock = False
End Sub

Private Sub MainTime_Timer()
    If 子节点视图容器.Visible = False Then
        Updata
    End If
End Sub

Private Sub MapUpdataTimer_Timer()
    If nSum > 0 And 全局视图.Checked = True And 子节点视图容器.Visible = False Then
        MapUpdata
    End If
End Sub

Private Sub NodePrintBox_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub NodePrintBox_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub PLC_Timer()
    Select Case noteSaveCheck
        Case 0
            pilotLightColor = RGB(255, 0, 0)
        Case 1
            pilotLightColor = RGB(255, 165, 0)
        Case 2
            pilotLightColor = RGB(0, 255, 0)
    End Select
    On Error GoTo Er
        If Clipboard.GetText = "" Then 粘贴.Enabled = False Else 粘贴.Enabled = True
        If bHLSum < 1 Then 撤销.Enabled = False Else 撤销.Enabled = True
        If redoSum < 1 Then 重做.Enabled = False Else 重做.Enabled = True
        If 放缩率需要提示提示倒计时 > 0 Then
            放缩率需要提示提示倒计时 = 放缩率需要提示提示倒计时 - 1
            放缩倍率显示标题.Caption = Format(zoomFactor, "0.0") & "X"
            If 放缩倍率显示标题.Visible = False Then 放缩倍率显示标题.Visible = True
        ElseIf 放缩倍率显示标题.Visible Then
            放缩倍率显示标题.Visible = False
        End If
        菜单单项控制
        If saveNtxTime <> 0 Then
            saveNtxTimeNow = saveNtxTimeNow + 0.5
            If saveNtxTimeNow > saveNtxTime And pilotLightColor = RGB(255, 165, 0) Then
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


Private Sub RGB色与VBColor互转工具_Click()
    RGBTOVBColorForm.Show
End Sub

Private Sub 文章节点化(txt As String)
    Dim i As Long, 回车数量 As Long, 列标 As Long, sT As String, idT As Long
    回车数量 = 1
    For i = 1 To Len(txt)
        sT = Mid(txt, i, 1)
        列标 = 列标 + 1
        idT = NodeEdit_NewNode(回车转义(sT), "", &HFFBF00, nodeDefaultSize, 列标 * imageToNtx_StepX, Me.height + 回车数量 * imageToNtx_StepY)
        If i > 1 Then LineAdd idT - 1, idT, "", lineDefaultSize
        If sT = vbLf Then 回车数量 = 回车数量 + 1: 列标 = 0
    Next
End Sub
Private Function 回车转义(s As String) As String
    Select Case s
        Case vbCr
            回车转义 = "\n"
        Case vbLf
            回车转义 = "\r"
        Case "\n"
            回车转义 = vbCr
        Case "\r"
            回车转义 = vbLf
        Case Else
            回车转义 = s
    End Select
End Function

Private Sub 保存笔记_Click()
    Dim filePath As String
    On Error GoTo Er
    If Dir(ntxPath) = "" Then
        filePath = 对话框选取保存文件路径("节点笔记 (*.ntx)|*.ntx|所有文件 (*.*)|*.*")
        If filePath <> "" Then NoteFileWrite_203 filePath
    Else
        NoteFileWrite_203 ntxPath
    End If

Exit Sub
Er:
    MsgBox "保存笔记失败，原因：" & Err.Description, 16, "保存笔记"
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

Private Sub 背景图_Click()
    On Error GoTo Er
        With CommonDialog1
            .Flags = cdlOFNHideReadOnly
            .Filter = "图片文件 (*.bmp;*.png;*jpg)|*.bmp;*.png;*jpg|所有文件 (*.*)|*.*"
            .FilterIndex = 1
            .ShowOpen
            homeBackPicPath = .filename
            加载背景图 homeBackPicPath
        End With
Er:
End Sub
Public Sub 加载背景图(fP As String)
    On Error GoTo Er
        Me.Picture = LoadPicture(fP)
    Exit Sub
Er:
    MsgBox "加载背景图失败！原因：" & Err.Description
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

Private Sub 波化节点_Click()
    Dim rootNid As Long, tempValue As Single
    rootNid = 未选中代替(nodeTargetAim)
    If rootNid <> -1 Then
        If NCF_NodeValueControl(富文本转义(node(rootNid).content), tempValue) Then
            ClearNode_ToTreeTxtLock
            NodeWave rootNid, tempValue, False, 0, 0
        End If
    End If
End Sub

Private Sub 彩虹圈_Click()
If 彩虹圈.Checked = False Then 彩虹圈.Checked = True Else 彩虹圈.Checked = False
End Sub

Private Sub 彩虹线_Click()
    彩虹线.Checked = Not 彩虹线.Checked
    If 彩虹线.Checked Then
        流光溢彩.Checked = False
    End If
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
    filePath = 对话框选取打开文件路径("节点笔记 (*.ntx)|*.ntx|所有文件 (*.*)|*.*")
    If filePath <> "" Then
        newAddNote
        NoteFileRead filePath
    End If
End Sub
Private Function 对话框选取打开文件路径(打开类型 As String) As String
    Dim filePath As String
    With CommonDialog1
        .CancelError = True
        On Error GoTo ErrHandler
            .Flags = cdlOFNHideReadOnly
            .Filter = 打开类型
            .FilterIndex = 1
            .ShowOpen
            对话框选取打开文件路径 = .filename
    End With
    Exit Function
ErrHandler:
End Function
Private Function 对话框选取保存文件路径(保存类型 As String) As String
    Dim filePath As String
    With CommonDialog1
        .CancelError = True
        On Error GoTo ErrHandler
            .Flags = cdlOFNOverwritePrompt
            .Filter = 保存类型
            .FilterIndex = 1
            .ShowSave
            对话框选取保存文件路径 = .filename
    End With
    Exit Function
ErrHandler:
End Function

Private Sub 打开联想节点文件目录_Click()
    On Error GoTo Er
        Shell "cmd.exe /c start """" """ & fictitiousNtxPath & """"
    Exit Sub
Er:
    MsgBox "打开联想节点文件目录失败，原因：" & Err.Description, 16, "错误"
End Sub

Private Sub 打开网络接口_Click()
    NodeNetwork.Show
End Sub

Private Sub 打印成PNG图片_Click()
    Dim fP As String
    On Error GoTo Er:
        fP = 对话框选取保存文件路径("图片 (*.png)|*.png")
        If fP <> "" Then
            NotePrint 打印缓存器
            SavePicture 打印缓存器.image, fP
        End If
    Exit Sub
Er:
    MsgBox "打印失败，失败原因：" & Err.Description, 16, "打印错误"
End Sub

Private Sub 导出TXT文章_Click()
    Dim filePath As String, nIdTemp As Long, outS As String
    On Error GoTo Er
        nIdTemp = 未选中代替(nodeTargetAim)
        If nIdTemp <> -1 Then
            filePath = 对话框选取保存文件路径("文本文件 (*.txt)|*.txt")
            If filePath <> "" Then
                ClearNode_ToTreeTxtLock
                节点转文章 outS, nIdTemp
                SaveFile_All filePath, outS
                MsgBox "导出成功！", 64, "导出TXT文章"
            End If
        End If
    Exit Sub
Er:
    MsgBox "导出错误，原因：" & Err.Description, 16, "导出TXT文章"
End Sub

Private Function 节点转文章(outS As String, sNid As Long)
    Dim i As Long
    node(sNid).toTreeTxtLock = True
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                If .Source = sNid And node(.target).toTreeTxtLock = False Then
                    outS = outS & 回车转义(node(.target).t)
                    节点转文章 outS, .target
                    Exit Function
                End If
            End If
        End With
    Next
End Function

Private Sub 导出位图_Click()
    Dim filePath As String, nIdTemp As Long
    On Error GoTo Er
        nIdTemp = 未选中代替(nodeTargetAim)
        If nIdTemp <> -1 Then
            filePath = 对话框选取保存文件路径("图片文件 (*.bmp)|*.bmp")
            If filePath <> "" Then
                If NoteToImage(filePath, 位图输出器) Then
                    MsgBox "导出成功！", 64, "导出笔记到位图文件"
                End If
            End If
        End If
    Exit Sub
Er:
    MsgBox "导出错误，原因：" & Err.Description, 16, "导出笔记到位图文件"
End Sub

Private Sub 导出文本文件_Click()
    Dim filePath As String, nIdTemp As Long
On Error GoTo Er
    nIdTemp = 未选中代替(nodeTargetAim)
    If nIdTemp <> -1 Then
        filePath = 对话框选取保存文件路径("文本文件 (*.txt)|*.txt")
        If filePath <> "" Then
            NoteToTreeTXT filePath, nIdTemp
            MsgBox "导出成功！", 64, "导出笔记到文本文件"
        End If
    Else
        MsgBox "未选中节点，导出操作无效。", 64, "导出笔记到文本文件"
    End If
    Exit Sub
Er:
    MsgBox "导出错误，原因：" & Err.Description, 16, "导出笔记到文本文件"
End Sub
Public Function 未选中代替(n As Long) As Long
    Dim i As Long
    If n = -1 Then
        未选中代替 = -1
        For i = 0 To nSum
            If node(i).b Then
                If node(i).select = True Then
                    未选中代替 = i
                End If
            End If
        Next
    Else
        未选中代替 = n
    End If
End Function

Private Sub 导入TXT文章_Click()
    Dim fP As String, txt As String
    On Error GoTo Er
        fP = 对话框选取打开文件路径("文本文件 (*.txt)|*.txt")
        If fP <> "" Then
            ReadFile_ALL_HV fP, txt
            BehaviorIdSet
            文章节点化 txt
        End If
    Exit Sub
Er:
    MsgBox "导入失败，原因：" & Err.Description, 16, "导入TXT文章"
End Sub

Private Sub 导入位图_Click()
    Dim fP As String
    On Error GoTo Er
    fP = 对话框选取打开文件路径("图片文件 (*.bmp;*.png;*.jpg)|*.bmp;*.png;*.jpg")
    If fP <> "" Then
        位图读取缓存器.Picture = LoadPicture(fP)
        BehaviorIdSet
        位图节点化 位图读取缓存器
    End If
    Exit Sub
Er:
    MsgBox "导入失败，原因：" & Err.Description, 16, "导入位图"
End Sub

Private Sub 导入文本文件_Click()
    On Error GoTo Er:
        导入TXT文件路径 = 对话框选取打开文件路径("文本文档 (*.txt)|*.txt")
        If 导入TXT文件路径 <> "" Then
            BehaviorIdSet
            TreeTXTToNtx
        End If
    Exit Sub
Er:
    MsgBox "导入失败，原因：" & Err.Description, 16, "导入文本文件"
End Sub

Private Sub 复制_Click()
    CopyObject False
End Sub

Private Function nodeToTxt() As String
    Dim i As Long
    For i = 0 To nSum
        With node(i)
            If .b = True Then
                If .select = True Or i = nodeTargetAim Then
                    RTBtemp.TextRTF = .content
                    nodeToTxt = nodeToTxt & Chr(34) & .t & """ : """ & RTBtemp.Text & """ , "
                End If
            End If
        End With
    Next
    nodeToTxt = Mid(nodeToTxt, 1, Len(nodeToTxt) - 2)
End Function

Private Sub 关于节点笔记_Click()
    AboutNote.Show
End Sub

Private Sub 绘图刷新间隔_Click()
    Dim sT As String
    sT = InBox("请输入绘图刷新间隔，间隔越小性能要求越高，程序更灵敏[10~100]：", updataSpeed)
    If promptBoxSelect = 0 Then
        updataSpeed = 限制数值(Val(sT), 10, 100)
        MainTime.interval = updataSpeed
    End If
End Sub

Private Sub 剪切_Click()
    BehaviorIdSet
    CopyObject True
End Sub

Private Sub 节点归一_Click()
    归一化选中节点
End Sub

Private Sub 节点归整_Click()
    节点归整.Checked = Not 节点归整.Checked
End Sub

Private Sub 节点清单_Click()
    NodeListVis.Show
End Sub


Private Sub 矩点_Click()
    矩点.Checked = Not 矩点.Checked
End Sub

Private Sub 矩线_Click()
    矩线.Checked = Not 矩线.Checked
End Sub

Private Sub 控制台_Click()
NoteControlDesk.Show
End Sub

Private Sub 连接反转_Click()
    ConnectionReversal
End Sub

Private Sub 连接内容_Click()
    Dim sT As String
    sT = InBox("请输入选中连接的显示内容：")
    If promptBoxSelect = 0 Then
        连接内容赋予 sT
    End If
End Sub

Private Sub 连接清单_Click()
    LineListVis.Show
End Sub

Private Sub 另存为_Click()
    Dim filePath As String
    On Error GoTo Er
    filePath = 对话框选取保存文件路径("节点笔记 (*.ntx)|*.ntx|所有文件 (*.*)|*.*")
    If filePath <> "" Then NoteFileWrite_203 filePath
    Exit Sub
Er:
    MsgBox "另存为笔记失败，原因：" & Err.Description, 16, "另存为笔记"
End Sub

Private Sub 流光溢彩_Click()
    流光溢彩.Checked = Not 流光溢彩.Checked
    If 流光溢彩.Checked Then
        彩虹线.Checked = False
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

Private Sub 删除背景图_Click()
    Me.Picture = Nothing
    homeBackPicPath = ""
End Sub

Private Sub 设置默认节点大小_Click()
    Dim sT As String
    sT = InBox("输入每次新建节点的大小[50,500]：", nodeDefaultSize)
    If promptBoxSelect = 0 Then
        nodeDefaultSize = 限制数值(Val(sT), 50, 500)
    End If
End Sub

Private Sub 设置默认连接宽度_Click()
    Dim sT As String
    sT = InBox("请输入新建连接宽度[1,10]：", lineDefaultSize)
    If promptBoxSelect = 0 Then
        lineDefaultSize = 限制数值(Val(sT), 1, 10)
    End If
End Sub

Private Sub 深度上色_Click()
    设置子节点颜色
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

Private Sub 位图输出器_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print X; ; Y
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

Private Sub 像素节点_Click()
    NodePixel
End Sub

Private Sub 新建笔记_Click()
    newAddNote
End Sub

Private Sub 选显_Click()
    SelectDisplayObjcet
End Sub

Private Sub 选域消点_Click()
    取消选区内的所有节点
End Sub

Private Sub 选域消线_Click()
    取消选区内的所有连接
End Sub

Private Sub 印窃_Click()
    On Error GoTo Er
    复制节点为纯文本
Er:
End Sub

Private Sub 有损替换_Click()
    NodeReplace.Show
End Sub

Private Sub 圆阵节点_Click()
    圆阵子节点
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

Private Sub 子节点视图_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    childNodeVisOld.X = X
    childNodeVisOld.Y = Y
End Sub

Private Sub 子节点视图_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And childNodeVisOld.X <> 0 And childNodeVisOld.Y <> 0 Then
        子节点视图.left = 子节点视图.left + X - childNodeVisOld.X
        子节点视图.Top = 子节点视图.Top + Y - childNodeVisOld.Y
    End If
End Sub

Private Sub 子节点视图_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    子节点视图_MouseMove Button, Shift, X, Y
End Sub

Private Sub 子节点视图标题栏_DblClick()
    子节点视图按钮_Click 1
End Sub

Private Sub 子节点视图标题栏_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    childNodeFormOld.X = X
    childNodeFormOld.Y = Y
End Sub

Private Sub 子节点视图标题栏_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And childNodeFormOld.X <> 0 And childNodeFormOld.Y <> 0 Then
        If 子节点视图按钮(1).Caption = "-" Then
            子节点视图按钮_Click 1
        End If
        子节点视图容器.left = 子节点视图容器.left + X - childNodeFormOld.X
        子节点视图容器.Top = 子节点视图容器.Top + Y - childNodeFormOld.Y
    End If
End Sub
Private Sub 子节点视图标题栏_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    子节点视图标题栏_MouseMove Button, Shift, X, Y
End Sub
Private Sub 子节点视图按钮_Click(Index As Integer)
    Select Case Index
        Case 0
            nodeTargetAim = -1
            childNodeVisLock = False
            子节点视图容器.Visible = False
        Case 1
            If 子节点视图按钮(1).Caption = "□" Then
                With childNodeVisPos
                    .xE = 子节点视图容器.width
                    .xS = 子节点视图容器.left
                    .yE = 子节点视图容器.height
                    .yS = 子节点视图容器.Top
                End With
                子节点视图容器.Top = Me.height
                子节点视图容器.left = 0
                子节点视图容器.height = Me.height
                子节点视图容器.width = Me.width
                子节点视图按钮(1).Caption = "-"
                子节点视图标题栏.left = (子节点视图容器.width - 子节点视图标题栏.width) / 2
            Else
                With childNodeVisPos
                    子节点视图容器.Top = .yS
                    子节点视图容器.left = .xS
                    子节点视图容器.height = .yE
                    子节点视图容器.width = .xE
                    子节点视图按钮(1).Caption = "□"
                    子节点视图标题栏.left = 0
                End With
            End If
        Case 2
            On Error Resume Next
                Shell App.EXEName & ".exe " & childNodeVisNtxPath, vbNormalFocus
    End Select
End Sub


Private Sub 自动保存间隔_Click()
    Dim sT As String
    On Error GoTo Er
    sT = InBox("请输入自动保存时间间隔(单位:秒,输入0代表不自动保存！)", saveNtxTime)
    If promptBoxSelect = 0 Then
        saveNtxTime = Val(sT)
    End If
    Exit Sub
Er:
    MsgBox "设置失败，原因：" & Err.Description, 16, "设置自动保存间隔"
    saveNtxTime = 0
End Sub

Private Sub 字体_Click()
On Error GoTo Er
With CommonDialog1
    .Flags = 1
    .FontName = Me.Font.name
    .FontBold = Me.Font.Bold
    .FontSize = MainFormFontSize
    .FontItalic = Me.Font.Italic
    .FontUnderline = Me.Font.Underline
    .FontStrikethru = Me.Font.Strikethrough
    .ShowFont
    Me.Font.name = .FontName '字体名称
    Me.Font.Bold = .FontBold  '加粗？
    MainFormFontSize = .FontSize '字体大小
    Me.Font.Italic = .FontItalic '倾斜？
    Me.Font.Underline = .FontUnderline '下划线？
    Me.Font.Strikethrough = .FontStrikethru '删除线
End With
Er:
End Sub

Private Sub 坐标化整_Click()
    On Error GoTo Er
    nodeAttributedToIntegers = Val(InBox("请输入化整数量级：", nodeAttributedToIntegers))
    If promptBoxSelect = 0 Then
        NodePositionVague nodeAttributedToIntegers
    End If
    Exit Sub
Er:
    MsgBox "坐标化整失败，原因：" & Err.Description, 16, "坐标化整"
End Sub
