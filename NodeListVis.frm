VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form NodeListVis 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "�ڵ��б�"
   ClientHeight    =   5220
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4080
   Icon            =   "NodeListVis.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   4080
   StartUpPosition =   3  '����ȱʡ
   Begin RichTextLib.RichTextBox ת���ı� 
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
   Begin VB.ListBox �ڵ��б� 
      Appearance      =   0  'Flat
      Height          =   5250
      ItemData        =   "NodeListVis.frx":70A7
      Left            =   0
      List            =   "NodeListVis.frx":70A9
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
Attribute VB_Name = "NodeListVis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    �ڵ��б������ = True
    NodeListVisUpdata
End Sub

Private Sub Form_Unload(Cancel As Integer)
    �ڵ��б������ = False
End Sub

Private Sub �ڵ��б�_DblClick()
    Dim d As ��ά����
    If �ڵ��б�.Text <> "" Then
        d.x = Note.Width / 2 - node(�ڵ��б�.ListIndex).x
        d.y = Note.Height / 2 - node(�ڵ��б�.ListIndex).y
        Updata_AllNodeMove_Moving d.x, d.y, False
    End If
End Sub

Private Sub ˢ��_Click()
    NodeListVisUpdata
End Sub
