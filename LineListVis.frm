VERSION 5.00
Begin VB.Form LineListVis 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "�����б�"
   ClientHeight    =   5220
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4080
   Icon            =   "LineListVis.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   4080
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ListBox �����б� 
      Appearance      =   0  'Flat
      Height          =   5250
      ItemData        =   "LineListVis.frx":700A
      Left            =   0
      List            =   "LineListVis.frx":7011
      TabIndex        =   0
      ToolTipText     =   "���F5ˢ���嵥"
      Top             =   0
      Width           =   4095
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu ˢ�� 
         Caption         =   "ˢ��"
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
    �����б������ = True
    LineListVisUpdata
End Sub

Private Sub Form_Unload(Cancel As Integer)
    �����б������ = False
End Sub

Private Sub �����б�_DblClick()
    Dim d As ��ά����
    If �����б�.Text <> "" Then
        d.x = Note.Width / 2 - (node(nodeLine(�����б�.ListIndex).source).x + node(nodeLine(�����б�.ListIndex).target).x) / 2
        d.y = Note.Height / 2 - (node(nodeLine(�����б�.ListIndex).source).y + node(nodeLine(�����б�.ListIndex).target).y) / 2
        Updata_AllNodeMove_Moving d.x, d.y, False
    End If
End Sub

Private Sub ˢ��_Click()
    LineListVisUpdata
End Sub
