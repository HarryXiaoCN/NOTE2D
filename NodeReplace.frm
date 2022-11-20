VERSION 5.00
Begin VB.Form NodeReplace 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Node Replace"
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2880
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox SearchOptional 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "替换节点标题"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   2400
      TabIndex        =   10
      Top             =   1920
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox SearchOptional 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "替换连接内容"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox SearchOptional 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "替换节点内容"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   8
      Top             =   1440
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox SearchOptional 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "仅替换选区内内容"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox InTextBox 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   1200
      TabIndex        =   6
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox InTextBox 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Top             =   480
      Width           =   3375
   End
   Begin VB.PictureBox TitleBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   4650
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      Begin VB.Label TitleLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "×"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Index           =   0
         Left            =   4200
         TabIndex        =   2
         Top             =   0
         Width           =   405
      End
      Begin VB.Label TitleLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "节点/连接内容有损替换"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "损失：节点内容格式将被初始化"
         Top             =   45
         Width           =   1890
      End
   End
   Begin VB.Label CMDLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "开始替换"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   4455
   End
   Begin VB.Label ContentLabel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "替换内容："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1050
   End
   Begin VB.Label ContentLabel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "查找内容："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1050
   End
   Begin VB.Shape 边框 
      Height          =   2520
      Left            =   0
      Top             =   360
      Width           =   4680
   End
End
Attribute VB_Name = "NodeReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private titleMousePosX As Single, titleMousePosY As Single

Private Sub CMDLabel_Click()
    Dim i As Long, sT As String, rT As String
    sT = InTextBox(0).text
    rT = InTextBox(1).text
    If SearchOptional(1).value = 1 Or SearchOptional(3).value = 1 Then
        For i = 0 To nSum
            With node(i)
                If .b Then
                    If (.select = True And SearchOptional(0).value = 1) Or SearchOptional(0).value = 0 Then
                        If SearchOptional(1).value = 1 Then
                            .content = Replace(.text, sT, rT)
                        End If
                        If SearchOptional(3).value = 1 Then
                            .t = Replace(.t, sT, rT)
                        End If
                    End If
                End If
            End With
        Next
    End If
    If SearchOptional(2).value = 1 Then
        For i = 0 To lSum
            With nodeLine(i)
                If .b Then
                    If (.select = True And SearchOptional(0).value = 1) Or SearchOptional(0).value = 0 Then
                        .content = Replace(.content, sT, rT)
                    End If
                End If
            End With
        Next
    End If
End Sub

Private Sub TitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    titleMousePosX = X
    titleMousePosY = Y
End Sub

Private Sub TitleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And titleMousePosX <> 0 And titleMousePosY <> 0 Then
        Me.left = Me.left + X - titleMousePosX
        Me.Top = Me.Top + Y - titleMousePosY
    End If
End Sub

Private Sub TitleLabel_Click(Index As Integer)
    If Index = 0 Then Me.Hide
End Sub

Private Sub TitleLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    titleMousePosX = X
    titleMousePosY = Y
End Sub

Private Sub TitleLabel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And titleMousePosX <> 0 And titleMousePosY <> 0 Then
        Me.left = Me.left + X - titleMousePosX
        Me.Top = Me.Top + Y - titleMousePosY
    End If
End Sub
