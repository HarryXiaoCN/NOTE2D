VERSION 5.00
Begin VB.Form AboutNote 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ڡ��ڵ�ʼǡ�"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4935
   Icon            =   "��������.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   4935
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox ˵���� 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   2520
      Width           =   4695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   3900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ڵ�ʼ�"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1440
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   1560
      Picture         =   "��������.frx":700A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "AboutNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim fP As String, outS As String
    On Error GoTo Er
    Me.Caption = "�ڵ�ʼ� " & App.Major & "." & App.Minor & "." & App.Revision
    fP = App.Path & "\����˵��.txt"
    If Dir(fP) <> "" Then
        ReadFile_ALL_HV fP, outS
        ˵����.Text = outS
    End If
Er:
End Sub

