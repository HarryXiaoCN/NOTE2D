VERSION 5.00
Begin VB.Form AboutNote 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于“节点笔记”"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4935
   Icon            =   "关于作者.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   4935
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox 说明书 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "作者邮箱：xiaoharry@foxmail.com"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Caption         =   "节点笔记"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Picture         =   "关于作者.frx":700A
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
    Me.Caption = "节点笔记 " & App.Major & "." & App.Minor & "." & App.Revision
    fP = App.Path & "\更新说明.txt"
    If Dir(fP) <> "" Then
        ReadFile_ALL_HV fP, outS
        说明书.Text = outS
    End If
Er:
End Sub

