VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form NodeInput 
   AutoRedraw      =   -1  'True
   Caption         =   "Node Edit"
   ClientHeight    =   8865
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6240
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
   Icon            =   "Node_NodeNatureInput.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      Height          =   8000
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Ctrl+S����ڵ�༭���ݣ�ESC���رսڵ�༭����"
      Top             =   740
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   14129
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Node_NodeNatureInput.frx":1084A
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
      Text            =   "������ڵ����..."
      Top             =   120
      Width           =   6000
   End
   Begin VB.Label PilotLight 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "  "
      Height          =   315
      Left            =   5830
      TabIndex        =   2
      ToolTipText     =   "��ɫ���޸��ѱ��棻��ɫ���ڵ��Ѵ��ڵ��޸�δ���棻��ɫ���½��ڵ㻹δ����"
      Top             =   520
      Width           =   150
   End
   Begin VB.Menu �ڵ� 
      Caption         =   "�ڵ�"
      Begin VB.Menu ���� 
         Caption         =   "����"
         Shortcut        =   ^S
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
         Checked         =   -1  'True
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
Private inputBoxContent As String
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.FontBold = True
End Sub

Private Sub Form_Load()
nodeEditFormLock = True
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
If Me.Height < 9450 Then Me.Enabled = False: Me.Height = 9450: Me.Enabled = True
If Me.Width < 6480 Then Me.Enabled = False: Me.Width = 6480: Me.Enabled = True
NodeTitle.Width = Me.Width - 480
PilotLight.left = Me.Width - 650
NodeInputBox.Width = Me.Width - 480
NodeInputBox.Height = Me.Height - 1750
End Sub

Private Sub Form_Unload(Cancel As Integer)
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

Private Sub NodeInputBox_LostFocus()
If NodeInputBox.Text = "" Then NodeInputBox.Text = "������ڵ�����..."
End Sub

Private Sub NodeTitle_GotFocus()
If NodeTitle.Text = "������ڵ����..." Then NodeTitle.Text = ""
End Sub

Private Sub NodeTitle_KeyDown(KeyCode As Integer, Shift As Integer)
NodeInputBox_KeyDown KeyCode, Shift
End Sub

Private Sub NodeTitle_LostFocus()
If NodeTitle.Text = "" Then NodeTitle.Text = "������ڵ����..."
End Sub

Private Sub TxtCheck_Timer()
If nodeEditLock = False Then
    PilotLight.BackColor = RGB(255, 0, 0)
ElseIf node(nodeEditAim).t = NodeTitle.Text And _
        node(nodeEditAim).content = inputBoxContent And _
        node(nodeEditAim).b = True Then
    PilotLight.BackColor = RGB(0, 255, 0)
ElseIf node(nodeEditAim).b = True Then
    PilotLight.BackColor = RGB(255, 165, 0)
Else
    PilotLight.BackColor = RGB(255, 0, 0)
End If
End Sub

Private Sub ����_Click()
BehaviorIdSet
If nodeEditLock = True Then
    NodeEdit_ReviseNode nodeEditAim, NodeTitle.Text, NodeInputBox.TextRTF
Else
    If NodeEdit_ContentFilter(NodeInputBox.Text) = True Then
        NodeEdit_NewNode NodeTitle.Text, "", nodeEditPos.x, nodeEditPos.y
    Else
        NodeEdit_NewNode NodeTitle.Text, NodeInputBox.TextRTF, nodeEditPos.x, nodeEditPos.y
    End If
End If
NodeInputBox.SelStart = Len(NodeInputBox.Text)
End Sub

Private Sub ��ɫ_Click()
NodeInputBox.SelColor = RGB(255, 165, 0)
End Sub

Private Sub ��ɫ_Click()
NodeInputBox.SelColor = RGB(255, 0, 0)
End Sub

Private Sub ��������_Click()
If ��������.Checked = True Then ��������.Checked = False Else ��������.Checked = True
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
