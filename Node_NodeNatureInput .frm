VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form NodeInput 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Node Edit"
   ClientHeight    =   8865
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6240
   FillColor       =   &H80000002&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   Icon            =   "Node_NodeNatureInput .frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   6240
   StartUpPosition =   3  '����ȱʡ
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer TxtCheck 
      Interval        =   100
      Left            =   5760
      Top             =   8400
   End
   Begin RichTextLib.RichTextBox NodeInputBox 
      Height          =   7770
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Ctrl+S����ڵ�༭���ݣ�ESC���رսڵ�༭����"
      Top             =   960
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   13705
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Node_NodeNatureInput .frx":700A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox NodeTitle 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "������ڵ����..."
      Top             =   120
      Width           =   6000
   End
   Begin VB.Shape ɫѡ�� 
      BorderColor     =   &H00FFBF00&
      BorderWidth     =   2
      Height          =   255
      Left            =   4180
      Top             =   670
      Width           =   180
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   24
      Left            =   5880
      TabIndex        =   27
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   23
      Left            =   5640
      TabIndex        =   26
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   22
      Left            =   5400
      TabIndex        =   25
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   21
      Left            =   5160
      TabIndex        =   24
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   20
      Left            =   4920
      TabIndex        =   23
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   19
      Left            =   4680
      TabIndex        =   22
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   18
      Left            =   4440
      TabIndex        =   21
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFBF00&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   17
      Left            =   4200
      TabIndex        =   20
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   16
      Left            =   3960
      TabIndex        =   19
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   15
      Left            =   3720
      TabIndex        =   18
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   14
      Left            =   3480
      TabIndex        =   17
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   13
      Left            =   3240
      TabIndex        =   16
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   12
      Left            =   3000
      TabIndex        =   15
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   11
      Left            =   2760
      TabIndex        =   14
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   10
      Left            =   2520
      TabIndex        =   13
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   9
      Left            =   2280
      TabIndex        =   12
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   8
      Left            =   2040
      TabIndex        =   11
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   7
      Left            =   1800
      TabIndex        =   10
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   6
      Left            =   1560
      TabIndex        =   9
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   5
      Left            =   1320
      TabIndex        =   8
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   4
      Left            =   1080
      TabIndex        =   7
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   3
      Left            =   840
      TabIndex        =   6
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   2
      Left            =   600
      TabIndex        =   5
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   135
   End
   Begin VB.Label �ڵ���ɫԤ�� 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   135
   End
   Begin VB.Label PilotLight 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "  "
      Height          =   315
      Left            =   5835
      TabIndex        =   2
      ToolTipText     =   "��ɫ���޸��ѱ��棻��ɫ���ڵ��Ѵ��ڵ��޸�δ���棻��ɫ���½��ڵ㻹δ����"
      Top             =   0
      Width           =   150
   End
   Begin VB.Menu �ڵ� 
      Caption         =   "�ڵ�"
      Begin VB.Menu ���� 
         Caption         =   "����"
         Shortcut        =   ^S
      End
      Begin VB.Menu �������� 
         Caption         =   "��������"
         Checked         =   -1  'True
         Shortcut        =   ^K
      End
      Begin VB.Menu jdcut2 
         Caption         =   "-"
      End
      Begin VB.Menu ѡ��ͬ���޸� 
         Caption         =   "ѡ��ͬ���޸�"
         Begin VB.Menu �ڵ�ͬ������ 
            Caption         =   "�ڵ����"
            Index           =   0
         End
         Begin VB.Menu �ڵ�ͬ������ 
            Caption         =   "�ڵ�����"
            Index           =   1
         End
         Begin VB.Menu �ڵ�ͬ������ 
            Caption         =   "�ڵ���ɫ"
            Index           =   2
         End
         Begin VB.Menu �ڵ�ͬ������ 
            Caption         =   "�ڵ��С"
            Index           =   3
         End
      End
      Begin VB.Menu Cut1 
         Caption         =   "-"
      End
      Begin VB.Menu �˳� 
         Caption         =   "�˳�"
      End
   End
   Begin VB.Menu ��ʽ 
      Caption         =   "��ʽ"
      Begin VB.Menu �Ӵ� 
         Caption         =   "�Ӵ�"
         Shortcut        =   ^B
      End
      Begin VB.Menu ��б 
         Caption         =   "��б"
         Shortcut        =   ^Q
      End
      Begin VB.Menu �»��� 
         Caption         =   "�»���"
         Shortcut        =   ^U
      End
      Begin VB.Menu ɾ���� 
         Caption         =   "ɾ����"
         Shortcut        =   ^D
      End
      Begin VB.Menu Cut4 
         Caption         =   "-"
      End
      Begin VB.Menu �����ֺ� 
         Caption         =   "�����ֺ� [Ctrl+Shift+>]"
      End
      Begin VB.Menu ��С�ֺ� 
         Caption         =   "��С�ֺ� [Ctrl+Shift+<]"
      End
      Begin VB.Menu Cut5 
         Caption         =   "-"
      End
      Begin VB.Menu ����� 
         Caption         =   "����� [Ctrl+L]"
      End
      Begin VB.Menu �Ҷ��� 
         Caption         =   "�Ҷ��� [Ctrl+R]"
      End
      Begin VB.Menu ���ж��� 
         Caption         =   "���ж���  [Ctrl+E]"
      End
      Begin VB.Menu Cut7 
         Caption         =   "-"
      End
      Begin VB.Menu �������� 
         Caption         =   "��������"
         Shortcut        =   ^T
      End
      Begin VB.Menu Cut2 
         Caption         =   "-"
      End
      Begin VB.Menu ���� 
         Caption         =   "����"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu ��ɫ 
      Caption         =   "��ɫ"
      Begin VB.Menu ��ɫ 
         Caption         =   "��ɫ"
         Shortcut        =   {F1}
      End
      Begin VB.Menu ��ɫ 
         Caption         =   "��ɫ"
         Shortcut        =   {F2}
      End
      Begin VB.Menu ��ɫ 
         Caption         =   "��ɫ"
         Shortcut        =   {F3}
      End
      Begin VB.Menu ��ɫ 
         Caption         =   "��ɫ"
         Shortcut        =   {F4}
      End
      Begin VB.Menu ��ɫ 
         Caption         =   "��ɫ"
         Shortcut        =   {F5}
      End
      Begin VB.Menu ��ɫ 
         Caption         =   "��ɫ"
         Shortcut        =   {F6}
      End
      Begin VB.Menu ��ɫ 
         Caption         =   "��ɫ"
         Shortcut        =   {F7}
      End
      Begin VB.Menu Cut6 
         Caption         =   "-"
      End
      Begin VB.Menu �Զ��� 
         Caption         =   "��ɫ"
         Shortcut        =   {F8}
      End
   End
End
Attribute VB_Name = "NodeInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private inputBoxContent As String, synchronizationState As String
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label1.FontBold = True
End Sub

Private Sub Form_Load()
    Me.height = nodeInputFormHeight
    Me.width = nodeInputFormWidth
    Me.Top = nodeInputFormTop
    Me.left = nodeInputFormLeft
    nodeEditFormLock = True
    �ڵ�ѡ��ɫ = 17
    Me.BackColor = NodeInputBackColor
    If Note.ȫ��͸��2.Checked = True Then
        FormTransparent Me, 50
    ElseIf Note.ȫ��͸��2.Checked = True Then
        FormTransparent Me, 125
    ElseIf Note.ȫ��͸��2.Checked = True Then
        FormTransparent Me, 200
    End If
End Sub

Private Sub Form_Resize()
    If WindowState = 1 Then Exit Sub
    If Me.height < 4000 Then Me.Enabled = False: Me.height = 4000: Me.Enabled = True
    If Me.width < 6350 Then Me.Enabled = False: Me.width = 6350: Me.Enabled = True
    NodeTitle.width = Me.width - 480
    PilotLight.left = Me.width - 650
    NodeInputBox.width = Me.width - 480 '360
    NodeInputBox.height = Me.height - 1980 '1900
    nodeInputFormHeight = Me.height
    nodeInputFormWidth = Me.width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    nodeInputFormTop = Me.Top
    nodeInputFormLeft = Me.left
    nodeEditFormLock = False
End Sub

Private Sub NodeInputBox_Change()
    inputBoxContent = NodeInputBox.TextRTF
End Sub

Private Sub NodeInputBox_GotFocus()
    If NodeInputBox.Text = "������ڵ�����..." Then NodeInputBox.Text = ""
End Sub

Private Sub NodeInputBox_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
Select Case KeyCode
    Case 27
        �˳�_Click
End Select
End Sub

Private Sub NodeInputBox_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 13
        If ��������.Checked = True Then
            NodeInputBox.SelText = NodeInputBox.SelText & TEXTINDENT
        End If
End Select
End Sub

'Private Sub NodeInputBox_LostFocus()
'    If NodeInputBox.Text = "" Then NodeInputBox.Text = "������ڵ�����..."
'End Sub

Private Sub NodeTitle_Change()
    Me.Caption = "�ڵ�����" & NodeTitle.Text & synchronizationState
End Sub

Private Sub NodeTitle_GotFocus()
    If NodeTitle.Text = "������ڵ����..." Then NodeTitle.Text = ""
End Sub

Private Sub NodeTitle_KeyDown(KeyCode As Integer, Shift As Integer)
    NodeInputBox_KeyDown KeyCode, Shift
    If KeyCode = vbKeyTab Then
        NodeInputBox.SetFocus
        NodeInputBox.SelStart = 0
        NodeInputBox.SelLength = Len(NodeInputBox.Text)
    End If
End Sub

Private Sub NodeTitle_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyTab Then
        KeyAscii = 0
    End If
End Sub

'Private Sub NodeTitle_LostFocus()
'    If NodeTitle.Text = "" Then NodeTitle.Text = "������ڵ����..."
'End Sub

Private Sub TxtCheck_Timer()
On Error GoTo Er
    If nodeEditLock = False Then
        PilotLight.BackColor = RGB(255, 0, 0)
    ElseIf node(nodeEditAim).t = NodeTitle.Text And _
            node(nodeEditAim).content = inputBoxContent And _
            node(nodeEditAim).setColor = ɫѡ��.BorderColor And _
            node(nodeEditAim).b = True Then
        PilotLight.BackColor = RGB(0, 255, 0)
    ElseIf node(nodeEditAim).b = True Then
        PilotLight.BackColor = RGB(255, 165, 0)
    Else
        PilotLight.BackColor = RGB(255, 0, 0)
    End If
Er:
End Sub

Private Sub ��������_Click()
    ��������.Checked = Not ��������.Checked
End Sub

Private Sub ����_Click()
    Dim i As Long, t As String, c As String, size As Single, color As Long
    BehaviorIdSet
    ɫѡ��.BorderColor = NCF_NodeColorControl(NodeInputBox.Text, ɫѡ��.BorderColor)
    If nodeEditLock = True Then
        NodeEdit_ReviseNode nodeEditAim, NodeTitle.Text, NodeInputBox.TextRTF, ɫѡ��.BorderColor, node(nodeEditAim).setSize
    Else
        If NodeEdit_ContentFilter(NodeInputBox.Text) Then
            NodeInputBox.Text = ""
            NodeEdit_NewNode NodeTitle.Text, NodeInputBox.TextRTF, ɫѡ��.BorderColor, nodeDefaultSize, nodeEditPos.x, nodeEditPos.y
        Else
            NodeEdit_NewNode NodeTitle.Text, NodeInputBox.TextRTF, ɫѡ��.BorderColor, nodeDefaultSize, nodeEditPos.x, nodeEditPos.y
        End If
    End If
    If �ڵ�ͬ������(0).Checked = True Or �ڵ�ͬ������(1).Checked = True Or �ڵ�ͬ������(2).Checked = True Or �ڵ�ͬ������(3).Checked = True Then
        For i = 0 To nSum
            With node(i)
                If .b = True And .select = True Then
                    If �ڵ�ͬ������(0).Checked Then
                        t = NodeTitle.Text
                    Else
                        t = .t
                    End If
                    If �ڵ�ͬ������(1).Checked Then
                        c = NodeInputBox.TextRTF
                    Else
                        c = .content
                    End If
                    If �ڵ�ͬ������(3).Checked Then
                        size = node(nodeEditAim).setSize
                    Else
                        size = .setSize
                    End If
                    If �ڵ�ͬ������(2).Checked Then
                        color = node(nodeEditAim).setColor
                    Else
                        color = .setColor
                    End If
                    NodeEdit_ReviseNode i, t, c, color, size
                End If
            End With
        Next
    End If
    fictitiousRootNodeId = nodeEditAim
    needUpdataNodePrint = True
    fictitiousIndexName = NodeTitle.Text
    FictitiousCheck
    Note.SetFocus
End Sub

Private Sub ��ɫ_Click()
NodeInputBox.SelColor = RGB(255, 165, 0)
End Sub

Private Sub ��ɫ_Click()
NodeInputBox.SelColor = RGB(255, 0, 0)
End Sub

Private Sub ��������_Click()
    ��������.Checked = Not ��������.Checked
End Sub

Private Sub ��ɫ_Click()
    NodeInputBox.SelColor = RGB(255, 255, 0)
End Sub

Private Sub �Ӵ�_Click()
If NodeInputBox.SelBold = True Then NodeInputBox.SelBold = False Else NodeInputBox.SelBold = True
End Sub

Private Sub ��С�ֺ�_Click()
NodeInputBox.SelFontSize = NodeInputBox.SelFontSize - 2
End Sub

Private Sub �ڵ�ͬ������_Click(Index As Integer)
    �ڵ�ͬ������(Index).Checked = Not �ڵ�ͬ������(Index).Checked
    synchronizationState = " - "
    For i = 0 To 3
        If �ڵ�ͬ������(i).Checked Then
            synchronizationState = synchronizationState & Replace(�ڵ�ͬ������(i).Caption, "�ڵ�", "") & "/"
        End If
    Next
    If synchronizationState <> " - " Then
        synchronizationState = Mid(synchronizationState, 1, Len(synchronizationState) - 1) & "��ͬ���޸�"
        Me.Caption = "�ڵ�����" & NodeTitle.Text & synchronizationState
    Else
        synchronizationState = ""
        Me.Caption = "�ڵ�����" & NodeTitle.Text
    End If
End Sub

Public Sub �ڵ���ɫԤ��_Click(Index As Integer)
    ɫѡ��.BorderColor = �ڵ���ɫԤ��(Index).BackColor
    ɫѡ��.left = �ڵ���ɫԤ��(Index).left - 10
End Sub
Public Function ɫ��ƥ��(c As Long) As Integer
    Dim i As Integer
    For i = 0 To �ڵ���ɫԤ��.Count - 1
        If �ڵ���ɫԤ��(i).BackColor = c Then
            ɫ��ƥ�� = i: Exit Function
        End If
    Next
End Function
Private Sub ���ж���_Click()
NodeInputBox.SelAlignment = rtfCenter
End Sub

Private Sub ��ɫ_Click()
NodeInputBox.SelColor = RGB(0, 0, 255)
End Sub

Private Sub ��ɫ_Click()
NodeInputBox.SelColor = RGB(0, 128, 0)
End Sub

Private Sub ��ɫ_Click()
NodeInputBox.SelColor = RGB(0, 255, 255)
End Sub

Private Sub ��б_Click()
If NodeInputBox.SelItalic = True Then NodeInputBox.SelItalic = False Else NodeInputBox.SelItalic = True
End Sub

Private Sub ɾ����_Click()
If NodeInputBox.SelStrikeThru = True Then NodeInputBox.SelStrikeThru = False Else NodeInputBox.SelStrikeThru = True

End Sub

Private Sub �˳�_Click()
Unload Me
End Sub

Private Sub �»���_Click()
If NodeInputBox.SelUnderline = True Then NodeInputBox.SelUnderline = False Else NodeInputBox.SelUnderline = True

End Sub

Private Sub �Ҷ���_Click()
NodeInputBox.SelAlignment = rtfRight
End Sub

Private Sub �����ֺ�_Click()
NodeInputBox.SelFontSize = NodeInputBox.SelFontSize + 2
End Sub

Private Sub ��ɫ_Click()
NodeInputBox.SelColor = RGB(128, 0, 128)
End Sub

Private Sub �Զ���_Click()
With CommonDialog1
    .Flags = 1
    .ShowColor
    NodeInputBox.SelColor = .color
End With
End Sub

Private Sub ����_Click()
With CommonDialog1
    .Flags = 1
    .FontName = NodeInputBox.SelFontName
    .FontBold = NodeInputBox.SelBold
    .FontSize = NodeInputBox.SelFontSize
    .FontItalic = NodeInputBox.SelItalic
    .FontUnderline = NodeInputBox.SelUnderline
    .FontStrikethru = NodeInputBox.SelStrikeThru
    .ShowFont
    NodeInputBox.SelFontName = .FontName
    NodeInputBox.SelBold = .FontBold
    NodeInputBox.SelFontSize = .FontSize
    NodeInputBox.SelItalic = .FontItalic
    NodeInputBox.SelUnderline = .FontUnderline
    NodeInputBox.SelStrikeThru = .FontStrikethru
End With
End Sub

Private Sub �����_Click()
NodeInputBox.SelAlignment = rtfLeft
End Sub
