VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form NodeNetwork 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Node Network"
   ClientHeight    =   4920
   ClientLeft      =   14010
   ClientTop       =   1200
   ClientWidth     =   5385
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NodeNetwork.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5385
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer WskStateTimer 
      Interval        =   500
      Left            =   5040
      Top             =   3960
   End
   Begin VB.TextBox PortBox 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Text            =   "20000"
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "启动"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox InfoBox 
      Appearance      =   0  'Flat
      Height          =   3855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   600
      Width           =   5175
   End
   Begin MSWinsockLib.Winsock Wsk 
      Left            =   4800
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label WskStateLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   5175
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "端口："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "NodeNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private 大容量传输模式 As Boolean, 大容量字节数 As Long

Private Sub Command1_Click()
    If Command1.Caption = "启动" Then
        Wsk.LocalPort = Val(PortBox.Text)
        Wsk.Listen
    Else
        Wsk.Close
    End If
End Sub

Private Sub Wsk_Close()
    Wsk.Close
    Wsk.Listen
End Sub

Private Sub Wsk_ConnectionRequest(ByVal requestID As Long)
    Wsk.Close
    Wsk.Accept requestID
End Sub

Private Sub Wsk_DataArrival(ByVal bytesTotal As Long)
    Dim 消息 As String, 反馈 As String
    If 大容量传输模式 Then
        If 大容量字节数 <= bytesTotal Then
            Wsk.GetData 消息
            输出文本 时间合成(" - [消息]：", "实际接收：" & bytesTotal) & vbCrLf
            反馈 = 执行消息(消息)
            Wsk.SendData 反馈
            大容量传输模式 = False
        End If
    Else
        Wsk.GetData 消息
        输出文本 时间合成(" - [消息]：", 消息) & vbCrLf
        反馈 = 执行消息(消息)
        Wsk.SendData 反馈
    End If
End Sub
Private Function 输出文本(s As String)
    On Error GoTo Er
    InfoBox.SelStart = Len(InfoBox.Text)
    InfoBox.SelText = s
    Exit Function
Er:
    InfoBox.Text = 时间合成(" - [错误]：", Err.Description) & vbCrLf
    On Error Resume Next
End Function

Private Function 执行消息(s As String) As String
    On Error GoTo Er:
        If Mid(s, 1, 12) = "启动大容量命令传输模式:" Then
            大容量传输模式 = True
            大容量字节数 = Val(Mid(s, 13))
            执行消息 = "大容量命令传输模式准备就绪！准备接收：" & 大容量字节数 & "字节数据"
        ElseIf Mid(LCase(s), 1, 43) = "start extra long command transmission mode:" Then
            大容量传输模式 = True
            大容量字节数 = Val(Mid(s, 44))
            执行消息 = "Ready for long command transmission mode！Ready to receive:" & 大容量字节数 & " bytesTotal"
        Else
            执行消息 = CMD_In(s)
        End If
    Exit Function
Er:
    执行消息 = Err.Description
End Function

Private Sub Wsk_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    输出文本 时间合成(" - [错误]：", Description) & vbCrLf
    Wsk.Close
End Sub

Private Function 时间合成(t As String, s As String) As String
    时间合成 = Now() & t & s
End Function

Private Sub WskStateTimer_Timer()
    WskStateLabel.Caption = WSK状态转换(Wsk.state)
    If Wsk.state = 0 Then
        Command1.Caption = "启动"
    Else
        Command1.Caption = "关闭"
    End If
End Sub

Private Function WSK状态转换(state As Integer)
    Select Case state
        Case 0
            WSK状态转换 = "连接关闭"
        Case 1
            WSK状态转换 = "连接打开"
        Case 2
            WSK状态转换 = "侦听中..."
        Case 3
            WSK状态转换 = "连接挂起"
        Case 4
            WSK状态转换 = "解析域名"
        Case 5
            WSK状态转换 = "已识别主机"
        Case 6
            WSK状态转换 = "正在连接"
        Case 7
            WSK状态转换 = "已连接"
        Case 8
            WSK状态转换 = "同级人员正在关闭连接"
        Case 9
            WSK状态转换 = "错误"
    End Select
End Function
 
