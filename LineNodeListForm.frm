VERSION 5.00
Begin VB.Form LineNodeListForm 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ӽڵ��б���"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4695
   Icon            =   "LineNodeListForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4695
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox ���ӽڵ��б����� 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2820
      Left            =   0
      ScaleHeight     =   2790
      ScaleWidth      =   4665
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.ListBox ���ӽڵ��б� 
         Appearance      =   0  'Flat
         Height          =   2550
         Index           =   0
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "���Ľڵ�ȥ��ڵ㼯"
         Top             =   120
         Width           =   2175
      End
      Begin VB.ListBox ���ӽڵ��б� 
         Appearance      =   0  'Flat
         Height          =   2550
         Index           =   1
         Left            =   2400
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "���Ľڵ�Դ��ڵ㼯"
         Top             =   120
         Width           =   2175
      End
   End
End
Attribute VB_Name = "LineNodeListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Top = lineNodeListFormTop
    Me.left = lineNodeListFormLeft
    FormStick Me, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    lineNodeListFormTop = Me.Top
    lineNodeListFormLeft = Me.left
End Sub

Private Sub ���ӽڵ��б�_DblClick(Index As Integer)
    Dim d As ��ά����, tmp As Long
    If ���ӽڵ��б�(Index).Text Like "*:""*" Then
        tmp = Val(Mid(���ӽڵ��б�(Index).Text, 1, InStr(1, ���ӽڵ��б�(Index).Text, ":""")))
        angleOfView.X = Note.width / 2 - node(tmp).X
        angleOfView.Y = Note.height / 2 - node(tmp).Y
        MainCoordinateSystemDefinition
    End If
End Sub
