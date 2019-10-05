VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form NodeFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Node Find"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Note_NodeFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   6255
   StartUpPosition =   2  '屏幕中心
   Begin RichTextLib.RichTextBox nodeTmp 
      Height          =   495
      Left            =   5520
      TabIndex        =   0
      Top             =   -500
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      _Version        =   393217
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Note_NodeFind.frx":700A
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5880
      Top             =   720
   End
   Begin VB.TextBox FindText 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   120
      TabIndex        =   1
      Text            =   "输入查询内容后回车搜索..."
      ToolTipText     =   "输入查询内容后回车搜索..."
      Top             =   120
      Width           =   6000
   End
   Begin VB.Menu 选项 
      Caption         =   "选项"
      Begin VB.Menu 搜索 
         Caption         =   "搜索"
         Shortcut        =   ^S
      End
      Begin VB.Menu 区分大小写 
         Caption         =   "区分大小写"
         Shortcut        =   ^B
      End
      Begin VB.Menu 选中范围内搜索 
         Caption         =   "选中范围内搜索"
         Shortcut        =   ^R
      End
      Begin VB.Menu Cut1 
         Caption         =   "-"
      End
      Begin VB.Menu 新文件中输出 
         Caption         =   "新文件中输出"
         Shortcut        =   ^N
      End
   End
   Begin VB.Menu 输出 
      Caption         =   "输出"
      Visible         =   0   'False
      Begin VB.Menu 圆形阵列 
         Caption         =   "圆形阵列"
         Checked         =   -1  'True
      End
      Begin VB.Menu 椭圆阵列 
         Caption         =   "椭圆阵列"
      End
   End
End
Attribute VB_Name = "NodeFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FindText_GotFocus()
If FindText.Text = "输入查询内容后回车搜索..." Then FindText.Text = ""
End Sub

Private Sub FindText_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 13
        搜索_Click
    Case 27
        Unload Me
End Select
End Sub

Private Sub FindText_LostFocus()
If FindText.Text = "" Then FindText.Text = "输入查询内容后回车搜索..."
End Sub

Private Sub Timer1_Timer()
FindText.SetFocus
Timer1.Enabled = False
End Sub

Private Sub 区分大小写_Click()
If 区分大小写.Checked = True Then 区分大小写.Checked = False Else 区分大小写.Checked = True
End Sub

Private Sub 搜索_Click()
    FindNode FindText.Text, 区分大小写.Checked, 选中范围内搜索.Checked, 新文件中输出.Checked
End Sub

Private Sub 椭圆阵列_Click()
If 椭圆阵列.Checked = True Then 椭圆阵列.Checked = False:  圆形阵列.Checked = True Else 椭圆阵列.Checked = True: 圆形阵列.Checked = False
End Sub

Private Sub 新文件中输出_Click()
If 新文件中输出.Checked = True Then 新文件中输出.Checked = False: 输出.Visible = False Else 新文件中输出.Checked = True: 输出.Visible = True
End Sub

Private Sub 选中范围内搜索_Click()
If 选中范围内搜索.Checked = True Then 选中范围内搜索.Checked = False Else 选中范围内搜索.Checked = True
End Sub

Private Sub 圆形阵列_Click()
If 圆形阵列.Checked = True Then 圆形阵列.Checked = False: 椭圆阵列.Checked = True Else 圆形阵列.Checked = True: 椭圆阵列.Checked = False

End Sub
