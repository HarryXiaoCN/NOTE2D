VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form NodeListVis 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "节点列表"
   ClientHeight    =   5220
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4080
   Icon            =   "NodeListVis.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   4080
   StartUpPosition =   3  '窗口缺省
   Begin RichTextLib.RichTextBox 转义文本 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"NodeListVis.frx":700A
   End
   Begin VB.ListBox 节点列表 
      Appearance      =   0  'Flat
      Height          =   5250
      ItemData        =   "NodeListVis.frx":70A7
      Left            =   0
      List            =   "NodeListVis.frx":70A9
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
Attribute VB_Name = "NodeListVis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    节点列表更新锁 = True
    NodeListVisUpdata
End Sub

Private Sub Form_Unload(Cancel As Integer)
    节点列表更新锁 = False
End Sub

Private Sub 节点列表_DblClick()
    Dim d As 二维坐标
    If 节点列表.Text <> "" Then
        d.x = Note.Width / 2 - node(节点列表.ListIndex).x
        d.y = Note.Height / 2 - node(节点列表.ListIndex).y
        Updata_AllNodeMove_Moving d.x, d.y, False
    End If
End Sub

Private Sub 刷新_Click()
    NodeListVisUpdata
End Sub
