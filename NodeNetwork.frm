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
      Name            =   "����"
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
   StartUpPosition =   2  '��Ļ����
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "�˿ڣ�"
      BeginProperty Font 
         Name            =   "����"
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
Private ����������ģʽ As Boolean, �������ֽ��� As Long

Private Sub Command1_Click()
    If Command1.Caption = "����" Then
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
    Dim ��Ϣ As String, ���� As String
    If ����������ģʽ Then
        If �������ֽ��� <= bytesTotal Then
            Wsk.GetData ��Ϣ
            ����ı� ʱ��ϳ�(" - [��Ϣ]��", "ʵ�ʽ��գ�" & bytesTotal) & vbCrLf
            ���� = ִ����Ϣ(��Ϣ)
            Wsk.SendData ����
            ����������ģʽ = False
        End If
    Else
        Wsk.GetData ��Ϣ
        ����ı� ʱ��ϳ�(" - [��Ϣ]��", ��Ϣ) & vbCrLf
        ���� = ִ����Ϣ(��Ϣ)
        Wsk.SendData ����
    End If
End Sub
Private Function ����ı�(s As String)
    On Error GoTo Er
    InfoBox.SelStart = Len(InfoBox.Text)
    InfoBox.SelText = s
    Exit Function
Er:
    InfoBox.Text = ʱ��ϳ�(" - [����]��", Err.Description) & vbCrLf
    On Error Resume Next
End Function

Private Function ִ����Ϣ(s As String) As String
    On Error GoTo Er:
        If Mid(s, 1, 12) = "���������������ģʽ:" Then
            ����������ģʽ = True
            �������ֽ��� = Val(Mid(s, 13))
            ִ����Ϣ = "�����������ģʽ׼��������׼�����գ�" & �������ֽ��� & "�ֽ�����"
        ElseIf Mid(LCase(s), 1, 43) = "start extra long command transmission mode:" Then
            ����������ģʽ = True
            �������ֽ��� = Val(Mid(s, 44))
            ִ����Ϣ = "Ready for long command transmission mode��Ready to receive:" & �������ֽ��� & " bytesTotal"
        Else
            ִ����Ϣ = CMD_In(s)
        End If
    Exit Function
Er:
    ִ����Ϣ = Err.Description
End Function

Private Sub Wsk_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ����ı� ʱ��ϳ�(" - [����]��", Description) & vbCrLf
    Wsk.Close
End Sub

Private Function ʱ��ϳ�(t As String, s As String) As String
    ʱ��ϳ� = Now() & t & s
End Function

Private Sub WskStateTimer_Timer()
    WskStateLabel.Caption = WSK״̬ת��(Wsk.state)
    If Wsk.state = 0 Then
        Command1.Caption = "����"
    Else
        Command1.Caption = "�ر�"
    End If
End Sub

Private Function WSK״̬ת��(state As Integer)
    Select Case state
        Case 0
            WSK״̬ת�� = "���ӹر�"
        Case 1
            WSK״̬ת�� = "���Ӵ�"
        Case 2
            WSK״̬ת�� = "������..."
        Case 3
            WSK״̬ת�� = "���ӹ���"
        Case 4
            WSK״̬ת�� = "��������"
        Case 5
            WSK״̬ת�� = "��ʶ������"
        Case 6
            WSK״̬ת�� = "��������"
        Case 7
            WSK״̬ת�� = "������"
        Case 8
            WSK״̬ת�� = "ͬ����Ա���ڹر�����"
        Case 9
            WSK״̬ת�� = "����"
    End Select
End Function
 
