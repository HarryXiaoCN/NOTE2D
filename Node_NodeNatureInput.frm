VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form NodeInput 
   AutoRedraw      =   -1  'True
   Caption         =   "Node Edit"
   ClientHeight    =   8865
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6240
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
   Icon            =   "Node_NodeNatureInput.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   6240
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer TxtCheck 
      Interval        =   100
      Left            =   5760
      Top             =   8400
   End
   Begin RichTextLib.RichTextBox NodeInputBox 
      Height          =   8000
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Ctrl+S保存节点编辑内容，ESC键关闭节点编辑窗口"
      Top             =   740
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   14129
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Node_NodeNatureInput.frx":1084A
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
   Begin VB.TextBox NodeTitle 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "请输入节点标题..."
      Top             =   120
      Width           =   6000
   End
   Begin VB.Label PilotLight 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "  "
      Height          =   315
      Left            =   5830
      TabIndex        =   2
      ToolTipText     =   "绿色：修改已保存；黄色：节点已存在但修改未保存；红色：新建节点还未保存"
      Top             =   520
      Width           =   150
   End
   Begin VB.Menu 节点 
      Caption         =   "节点"
      Begin VB.Menu 保存 
         Caption         =   "保存"
         Shortcut        =   ^S
      End
      Begin VB.Menu Cut1 
         Caption         =   "-"
      End
      Begin VB.Menu 退出 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu 格式 
      Caption         =   "格式"
      Begin VB.Menu 加粗 
         Caption         =   "加粗"
         Shortcut        =   ^B
      End
      Begin VB.Menu 倾斜 
         Caption         =   "倾斜"
         Shortcut        =   ^Q
      End
      Begin VB.Menu 下划线 
         Caption         =   "下划线"
         Shortcut        =   ^U
      End
      Begin VB.Menu 删除线 
         Caption         =   "删除线"
         Shortcut        =   ^D
      End
      Begin VB.Menu Cut4 
         Caption         =   "-"
      End
      Begin VB.Menu 增大字号 
         Caption         =   "增大字号 [Ctrl+Shift+>]"
      End
      Begin VB.Menu 减小字号 
         Caption         =   "减小字号 [Ctrl+Shift+<]"
      End
      Begin VB.Menu Cut5 
         Caption         =   "-"
      End
      Begin VB.Menu 左对齐 
         Caption         =   "左对齐 [Ctrl+L]"
      End
      Begin VB.Menu 右对齐 
         Caption         =   "右对齐 [Ctrl+R]"
      End
      Begin VB.Menu 居中对齐 
         Caption         =   "居中对齐  [Ctrl+E]"
      End
      Begin VB.Menu Cut7 
         Caption         =   "-"
      End
      Begin VB.Menu 换行缩进 
         Caption         =   "换行缩进"
         Checked         =   -1  'True
         Shortcut        =   ^T
      End
      Begin VB.Menu Cut2 
         Caption         =   "-"
      End
      Begin VB.Menu 字体 
         Caption         =   "字体"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu 颜色 
      Caption         =   "颜色"
      Begin VB.Menu 红色 
         Caption         =   "红色"
         Shortcut        =   {F1}
      End
      Begin VB.Menu 橙色 
         Caption         =   "橙色"
         Shortcut        =   {F2}
      End
      Begin VB.Menu 黄色 
         Caption         =   "黄色"
         Shortcut        =   {F3}
      End
      Begin VB.Menu 绿色 
         Caption         =   "绿色"
         Shortcut        =   {F4}
      End
      Begin VB.Menu 青色 
         Caption         =   "青色"
         Shortcut        =   {F5}
      End
      Begin VB.Menu 蓝色 
         Caption         =   "蓝色"
         Shortcut        =   {F6}
      End
      Begin VB.Menu 紫色 
         Caption         =   "紫色"
         Shortcut        =   {F7}
      End
      Begin VB.Menu Cut6 
         Caption         =   "-"
      End
      Begin VB.Menu 自定义 
         Caption         =   "万色"
         Shortcut        =   {F8}
      End
   End
End
Attribute VB_Name = "NodeInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private inputBoxContent As String
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.FontBold = True
End Sub

Private Sub Form_Load()
nodeEditFormLock = True
Me.BackColor = NodeInputBackColor
If Note.全高透明2.Checked = True Then
    FormTransparent Me, 50
ElseIf Note.全半透明2.Checked = True Then
    FormTransparent Me, 125
ElseIf Note.全低透明2.Checked = True Then
    FormTransparent Me, 200
End If
End Sub

Private Sub Form_Resize()
If WindowState = 1 Then Exit Sub
If Me.Height < 9450 Then Me.Enabled = False: Me.Height = 9450: Me.Enabled = True
If Me.Width < 6480 Then Me.Enabled = False: Me.Width = 6480: Me.Enabled = True
NodeTitle.Width = Me.Width - 480
PilotLight.left = Me.Width - 650
NodeInputBox.Width = Me.Width - 480
NodeInputBox.Height = Me.Height - 1750
End Sub

Private Sub Form_Unload(Cancel As Integer)
nodeEditFormLock = False
End Sub

Private Sub NodeInputBox_Change()
inputBoxContent = NodeInputBox.TextRTF
End Sub

Private Sub NodeInputBox_GotFocus()
If NodeInputBox.Text = "请输入节点内容..." Then NodeInputBox.Text = ""
End Sub

Private Sub NodeInputBox_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
Select Case KeyCode
    Case 27
        退出_Click
End Select
End Sub

Private Sub NodeInputBox_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 13
        If 换行缩进.Checked = True Then
            NodeInputBox.SelText = NodeInputBox.SelText & TEXTINDENT
        End If
End Select
End Sub

Private Sub NodeInputBox_LostFocus()
If NodeInputBox.Text = "" Then NodeInputBox.Text = "请输入节点内容..."
End Sub

Private Sub NodeTitle_GotFocus()
If NodeTitle.Text = "请输入节点标题..." Then NodeTitle.Text = ""
End Sub

Private Sub NodeTitle_KeyDown(KeyCode As Integer, Shift As Integer)
NodeInputBox_KeyDown KeyCode, Shift
End Sub

Private Sub NodeTitle_LostFocus()
If NodeTitle.Text = "" Then NodeTitle.Text = "请输入节点标题..."
End Sub

Private Sub TxtCheck_Timer()
If nodeEditLock = False Then
    PilotLight.BackColor = RGB(255, 0, 0)
ElseIf node(nodeEditAim).t = NodeTitle.Text And _
        node(nodeEditAim).content = inputBoxContent And _
        node(nodeEditAim).b = True Then
    PilotLight.BackColor = RGB(0, 255, 0)
ElseIf node(nodeEditAim).b = True Then
    PilotLight.BackColor = RGB(255, 165, 0)
Else
    PilotLight.BackColor = RGB(255, 0, 0)
End If
End Sub

Private Sub 保存_Click()
BehaviorIdSet
If nodeEditLock = True Then
    NodeEdit_ReviseNode nodeEditAim, NodeTitle.Text, NodeInputBox.TextRTF
Else
    If NodeEdit_ContentFilter(NodeInputBox.Text) = True Then
        NodeEdit_NewNode NodeTitle.Text, "", nodeEditPos.x, nodeEditPos.y
    Else
        NodeEdit_NewNode NodeTitle.Text, NodeInputBox.TextRTF, nodeEditPos.x, nodeEditPos.y
    End If
End If
NodeInputBox.SelStart = Len(NodeInputBox.Text)
End Sub

Private Sub 橙色_Click()
NodeInputBox.SelColor = RGB(255, 165, 0)
End Sub

Private Sub 红色_Click()
NodeInputBox.SelColor = RGB(255, 0, 0)
End Sub

Private Sub 换行缩进_Click()
If 换行缩进.Checked = True Then 换行缩进.Checked = False Else 换行缩进.Checked = True
End Sub

Private Sub 黄色_Click()
NodeInputBox.SelColor = RGB(255, 255, 0)
End Sub

Private Sub 加粗_Click()
If NodeInputBox.SelBold = True Then NodeInputBox.SelBold = False Else NodeInputBox.SelBold = True
End Sub

Private Sub 减小字号_Click()
NodeInputBox.SelFontSize = NodeInputBox.SelFontSize - 2
End Sub

Private Sub 居中对齐_Click()
NodeInputBox.SelAlignment = rtfCenter
End Sub

Private Sub 蓝色_Click()
NodeInputBox.SelColor = RGB(0, 0, 255)
End Sub

Private Sub 绿色_Click()
NodeInputBox.SelColor = RGB(0, 128, 0)
End Sub

Private Sub 青色_Click()
NodeInputBox.SelColor = RGB(0, 255, 255)
End Sub

Private Sub 倾斜_Click()
If NodeInputBox.SelItalic = True Then NodeInputBox.SelItalic = False Else NodeInputBox.SelItalic = True
End Sub

Private Sub 删除线_Click()
If NodeInputBox.SelStrikeThru = True Then NodeInputBox.SelStrikeThru = False Else NodeInputBox.SelStrikeThru = True

End Sub

Private Sub 退出_Click()
Unload Me
End Sub

Private Sub 下划线_Click()
If NodeInputBox.SelUnderline = True Then NodeInputBox.SelUnderline = False Else NodeInputBox.SelUnderline = True

End Sub

Private Sub 右对齐_Click()
NodeInputBox.SelAlignment = rtfRight
End Sub

Private Sub 增大字号_Click()
NodeInputBox.SelFontSize = NodeInputBox.SelFontSize + 2
End Sub

Private Sub 紫色_Click()
NodeInputBox.SelColor = RGB(128, 0, 128)
End Sub

Private Sub 自定义_Click()
With CommonDialog1
    .Flags = 1
    .ShowColor
    NodeInputBox.SelColor = .color
End With
End Sub

Private Sub 字体_Click()
With CommonDialog1
    .Flags = 1
    .FontName = NodeInputBox.SelFontName
    .FontBold = NodeInputBox.SelBold
    .FontSize = NodeInputBox.SelFontSize
    .FontItalic = NodeInputBox.SelItalic
    .FontUnderline = NodeInputBox.SelUnderline
    .FontStrikethru = NodeInputBox.SelStrikeThru
    .ShowFont
    NodeInputBox.SelFontName = .FontName
    NodeInputBox.SelBold = .FontBold
    NodeInputBox.SelFontSize = .FontSize
    NodeInputBox.SelItalic = .FontItalic
    NodeInputBox.SelUnderline = .FontUnderline
    NodeInputBox.SelStrikeThru = .FontStrikethru
End With
End Sub

Private Sub 左对齐_Click()
NodeInputBox.SelAlignment = rtfLeft
End Sub
