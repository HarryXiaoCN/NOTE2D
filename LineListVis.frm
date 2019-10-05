VERSION 5.00
Begin VB.Form LineListVis 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "连接列表"
   ClientHeight    =   5220
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4080
   Icon            =   "LineListVis.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   4080
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox 连接列表 
      Appearance      =   0  'Flat
      Height          =   5250
      ItemData        =   "LineListVis.frx":700A
      Left            =   0
      List            =   "LineListVis.frx":7011
      TabIndex        =   0
      ToolTipText     =   "点击F5刷新清单"
      Top             =   0
      Width           =   4095
   End
   Begin VB.Menu 功能 
      Caption         =   "功能"
      Begin VB.Menu 刷新 
         Caption         =   "刷新"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "LineListVis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    连接列表更新锁 = True
    LineListVisUpdata
End Sub

Private Sub Form_Unload(Cancel As Integer)
    连接列表更新锁 = False
End Sub

Private Sub 连接列表_DblClick()
    Dim d As 二维坐标
    If 连接列表.Text <> "" Then
        d.x = Note.Width / 2 - (node(nodeLine(连接列表.ListIndex).source).x + node(nodeLine(连接列表.ListIndex).target).x) / 2
        d.y = Note.Height / 2 - (node(nodeLine(连接列表.ListIndex).source).y + node(nodeLine(连接列表.ListIndex).target).y) / 2
        Updata_AllNodeMove_Moving d.x, d.y, False
    End If
End Sub

Private Sub 刷新_Click()
    LineListVisUpdata
End Sub
