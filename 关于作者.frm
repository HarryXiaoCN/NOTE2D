VERSION 5.00
Begin VB.Form AboutNote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ڡ��ڵ�ʼǡ�"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4905
   Icon            =   "��������.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4905
   StartUpPosition =   2  '��Ļ����
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   1800
      Picture         =   "��������.frx":1084A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�������䣺xiaoharry@foxmail.com"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   2040
      Width           =   3900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Network Note"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   2565
   End
End
Attribute VB_Name = "AboutNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Caption = "�ڵ�ʼ� " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
