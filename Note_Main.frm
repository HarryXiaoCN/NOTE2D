VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Note 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Note"
   ClientHeight    =   9120
   ClientLeft      =   6375
   ClientTop       =   3465
   ClientWidth     =   14760
   DrawWidth       =   2
   FillColor       =   &H00FFFF00&
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
   ForeColor       =   &H80000013&
   Icon            =   "Note_Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   9120
   ScaleWidth      =   14760
   Begin VB.Timer ActionTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   7920
   End
   Begin VB.PictureBox �ӽڵ���ͼ���� 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   1200
      ScaleHeight     =   7785
      ScaleWidth      =   12825
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   12855
      Begin VB.PictureBox �ӽڵ���ͼ������ 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   12825
         TabIndex        =   9
         Top             =   0
         Width           =   12855
         Begin VB.Label �ӽڵ���ͼ��ť 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   315
            Index           =   2
            Left            =   0
            TabIndex        =   12
            ToolTipText     =   "��Ԥ���ʼ�"
            Top             =   0
            Width           =   375
         End
         Begin VB.Label �ӽڵ���ͼ��ť 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   315
            Index           =   1
            Left            =   12000
            TabIndex        =   11
            ToolTipText     =   "���/��С���ӽڵ���ͼ"
            Top             =   0
            Width           =   375
         End
         Begin VB.Label �ӽڵ���ͼ��ť 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   0
            Left            =   12360
            TabIndex        =   10
            ToolTipText     =   "�ر��ӽڵ���ͼ"
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.PictureBox �ӽڵ���ͼ 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   10000
         Left            =   -2000
         ScaleHeight     =   9975
         ScaleWidth      =   9975
         TabIndex        =   8
         Top             =   -2000
         Width           =   10000
      End
   End
   Begin VB.PictureBox λͼ����� 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   2520
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox λͼ��ȡ������ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   1320
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox ��ӡ������ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      ScaleHeight     =   945
      ScaleWidth      =   1065
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox RTBtemp 
      Height          =   495
      Left            =   6600
      TabIndex        =   3
      Top             =   8400
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"Note_Main.frx":700A
   End
   Begin VB.Timer MapUpdataTimer 
      Interval        =   100
      Left            =   2040
      Top             =   8520
   End
   Begin VB.PictureBox GlobalView 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FAE6E6&
      FillColor       =   &H009AFA00&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2500
      Left            =   11000
      ScaleHeight     =   2475
      ScaleWidth      =   3585
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6500
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Timer PLC 
      Interval        =   500
      Left            =   120
      Top             =   8520
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer MainTime 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1440
      Top             =   8520
   End
   Begin RichTextLib.RichTextBox NodePrintBox 
      Height          =   1755
      Left            =   2640
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   3096
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Note_Main.frx":70B7
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
   Begin VB.Label ����������ʾ���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Menu �ļ� 
      Caption         =   "�ļ�"
      Begin VB.Menu �½��ʼ� 
         Caption         =   "�½�"
         Shortcut        =   ^N
      End
      Begin VB.Menu �򿪱ʼ� 
         Caption         =   "��"
         Shortcut        =   ^O
      End
      Begin VB.Menu ����ʼ� 
         Caption         =   "����"
         Shortcut        =   ^S
      End
      Begin VB.Menu ���Ϊ 
         Caption         =   "���Ϊ"
      End
      Begin VB.Menu wencut2 
         Caption         =   "-"
      End
      Begin VB.Menu ��ӡ 
         Caption         =   "��ӡ"
         Begin VB.Menu ��ӡ��PNGͼƬ 
            Caption         =   "��ӡ��PNGͼƬ"
         End
      End
      Begin VB.Menu Cut3 
         Caption         =   "-"
      End
      Begin VB.Menu �˳� 
         Caption         =   "�˳�"
      End
   End
   Begin VB.Menu �༭ 
      Caption         =   "�༭"
      Begin VB.Menu ���� 
         Caption         =   "����"
         Shortcut        =   ^Z
      End
      Begin VB.Menu ���� 
         Caption         =   "����"
         Shortcut        =   ^Y
      End
      Begin VB.Menu Cut1 
         Caption         =   "-"
      End
      Begin VB.Menu ���� 
         Caption         =   "����"
         Shortcut        =   ^C
      End
      Begin VB.Menu ���� 
         Caption         =   "����"
         Shortcut        =   ^X
      End
      Begin VB.Menu ճ�� 
         Caption         =   "ճ��"
         Shortcut        =   ^V
      End
      Begin VB.Menu bjcut4 
         Caption         =   "-"
      End
      Begin VB.Menu ӡ�� 
         Caption         =   "ӡ�ԣ�Ctrl+Shift+C��"
      End
      Begin VB.Menu Cut4 
         Caption         =   "-"
      End
      Begin VB.Menu ���� 
         Caption         =   "����"
         Shortcut        =   ^F
      End
      Begin VB.Menu �����滻 
         Caption         =   "�����滻"
         Shortcut        =   ^H
      End
      Begin VB.Menu Cut2 
         Caption         =   "-"
      End
      Begin VB.Menu ѡ�� 
         Caption         =   "ѡ��"
         Shortcut        =   ^R
      End
      Begin VB.Menu ȫѡ 
         Caption         =   "ȫѡ"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu ��ͼ 
      Caption         =   "��ͼ"
      Begin VB.Menu �ڵ�����ʾ 
         Caption         =   "�ڵ�����ʾ"
         Begin VB.Menu ��ʾȫ���ڵ��� 
            Caption         =   "��ʾȫ���ڵ���"
            Checked         =   -1  'True
         End
         Begin VB.Menu ��ʾ˳��ڵ��� 
            Caption         =   "��ʾ˳��ڵ���"
         End
         Begin VB.Menu ��ʾ����ڵ��� 
            Caption         =   "��ʾ����ڵ���"
         End
         Begin VB.Menu Cut6 
            Caption         =   "-"
         End
         Begin VB.Menu ʼ����ʾѡ���� 
            Caption         =   "ʼ����ʾѡ����"
         End
         Begin VB.Menu Cut5 
            Caption         =   "-"
         End
         Begin VB.Menu ��ʾ�ڵ����ID 
            Caption         =   "��ʾ�ڵ����ID"
         End
      End
      Begin VB.Menu �ڵ�������ʾ 
         Caption         =   "�ڵ�������ʾ"
         Begin VB.Menu ��ʾȫ������ 
            Caption         =   "��ʾȫ������"
            Checked         =   -1  'True
         End
         Begin VB.Menu ��ʾ˳������ 
            Caption         =   "��ʾ˳������"
         End
         Begin VB.Menu ��ʾ�������� 
            Caption         =   "��ʾ��������"
         End
         Begin VB.Menu Cut7 
            Caption         =   "-"
         End
         Begin VB.Menu ʼ����ʾѡ�� 
            Caption         =   "ʼ����ʾѡ��"
         End
      End
      Begin VB.Menu jmcut1 
         Caption         =   "-"
      End
      Begin VB.Menu ������ 
         Caption         =   "������"
         Begin VB.Menu ���� 
            Caption         =   "����"
         End
         Begin VB.Menu ͸��2 
            Caption         =   "͸��"
            Begin VB.Menu ȫ��͸�� 
               Caption         =   "ȫ��͸��"
            End
            Begin VB.Menu ȫ��͸�� 
               Caption         =   "ȫ��͸��"
            End
            Begin VB.Menu ȫ��͸�� 
               Caption         =   "ȫ��͸��"
            End
         End
         Begin VB.Menu ������ɫ 
            Caption         =   "������ɫ"
         End
         Begin VB.Menu bjcut1 
            Caption         =   "-"
         End
         Begin VB.Menu ����ɫ 
            Caption         =   "����ɫ"
         End
         Begin VB.Menu ����ͼ 
            Caption         =   "����ͼ"
         End
         Begin VB.Menu ɾ������ͼ 
            Caption         =   "ɾ������ͼ"
         End
         Begin VB.Menu Cut9 
            Caption         =   "-"
         End
         Begin VB.Menu �ص� 
            Caption         =   "�ص�"
         End
         Begin VB.Menu ���� 
            Caption         =   "����"
         End
         Begin VB.Menu �ʺ�Ȧ 
            Caption         =   "�ʺ�Ȧ"
            Checked         =   -1  'True
         End
         Begin VB.Menu �ʺ��� 
            Caption         =   "�ʺ���"
            Checked         =   -1  'True
         End
         Begin VB.Menu ������� 
            Caption         =   "�������"
         End
      End
      Begin VB.Menu ������� 
         Caption         =   "�������"
         Begin VB.Menu ͸�� 
            Caption         =   "͸��"
            Begin VB.Menu ȫ��͸��2 
               Caption         =   "ȫ��͸��"
            End
            Begin VB.Menu ȫ��͸��2 
               Caption         =   "ȫ��͸��"
            End
            Begin VB.Menu ȫ��͸��2 
               Caption         =   "ȫ��͸��"
            End
         End
         Begin VB.Menu ����ɫ2 
            Caption         =   "����ɫ"
         End
      End
      Begin VB.Menu ������� 
         Caption         =   "�������"
         Begin VB.Menu �ö� 
            Caption         =   "�ö�"
            Shortcut        =   ^T
         End
         Begin VB.Menu ��ǩ�� 
            Caption         =   "��ǩ��"
            Shortcut        =   ^L
         End
         Begin VB.Menu ͸��3 
            Caption         =   "͸��"
            Begin VB.Menu ȫ��͸��3 
               Caption         =   "ȫ��͸��"
            End
            Begin VB.Menu ȫ��͸��3 
               Caption         =   "ȫ��͸��"
            End
            Begin VB.Menu ȫ��͸��3 
               Caption         =   "ȫ��͸��"
            End
         End
      End
      Begin VB.Menu jmcut2 
         Caption         =   "-"
      End
      Begin VB.Menu ȫ����ͼ 
         Caption         =   "ȫ����ͼ"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu ���ӷ�ת 
         Caption         =   "���ӷ�ת(R)"
      End
      Begin VB.Menu Բ��ڵ� 
         Caption         =   "Բ��ڵ�(A)"
      End
      Begin VB.Menu �������� 
         Caption         =   "��������(C)"
      End
      Begin VB.Menu �����ɫ 
         Caption         =   "�����ɫ(N)"
      End
      Begin VB.Menu �����ڵ� 
         Caption         =   "�����ڵ�(W)"
      End
      Begin VB.Menu ���ؽڵ� 
         Caption         =   "���ؽڵ�(P)"
      End
      Begin VB.Menu ���껯�� 
         Caption         =   "���껯��(V)"
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu ���� 
         Caption         =   "����"
         Begin VB.Menu �����ı��ļ� 
            Caption         =   "��״TAB�ָ���ʽ�ı�"
         End
         Begin VB.Menu ����λͼ 
            Caption         =   "BMPλͼ"
         End
         Begin VB.Menu ����TXT���� 
            Caption         =   "TXT����"
         End
      End
      Begin VB.Menu ���� 
         Caption         =   "����"
         Begin VB.Menu �����ı��ļ� 
            Caption         =   "��״TAB�ָ�ʽ�ı�"
            Shortcut        =   ^E
         End
         Begin VB.Menu ����λͼ 
            Caption         =   "BMPλͼ"
            Shortcut        =   ^I
         End
         Begin VB.Menu ����TXT���� 
            Caption         =   "TXT����"
         End
      End
      Begin VB.Menu gncut6 
         Caption         =   "-"
      End
      Begin VB.Menu ѡ������ 
         Caption         =   "ѡ������(Shift+N)"
      End
      Begin VB.Menu ѡ������ 
         Caption         =   "ѡ������(Shift+L)"
      End
      Begin VB.Menu gncut7 
         Caption         =   "-"
      End
      Begin VB.Menu �ڵ��һ 
         Caption         =   "�ڵ��һ(Ctrl+Shift+O)"
      End
      Begin VB.Menu �ڵ���� 
         Caption         =   "�ڵ����"
      End
      Begin VB.Menu Cut11 
         Caption         =   "-"
      End
      Begin VB.Menu �ڵ��嵥 
         Caption         =   "�ڵ��嵥"
      End
      Begin VB.Menu �����嵥 
         Caption         =   "�����嵥"
      End
      Begin VB.Menu gncut3 
         Caption         =   "-"
      End
      Begin VB.Menu ����Ĭ�Ͻڵ��С 
         Caption         =   "����Ĭ�Ͻڵ��С"
      End
      Begin VB.Menu ����Ĭ�����ӿ�� 
         Caption         =   "����Ĭ�����ӿ��"
      End
      Begin VB.Menu gncut4 
         Caption         =   "-"
      End
      Begin VB.Menu ��ͼˢ�¼�� 
         Caption         =   "��ͼˢ�¼��"
      End
      Begin VB.Menu �Զ������� 
         Caption         =   "�Զ�������"
      End
      Begin VB.Menu gncut5 
         Caption         =   "-"
      End
      Begin VB.Menu RGBɫ��VBColor��ת���� 
         Caption         =   "RGBɫ��VBColor��ת����"
      End
      Begin VB.Menu ������ڵ��ļ�Ŀ¼ 
         Caption         =   "������ڵ��ļ�Ŀ¼"
      End
      Begin VB.Menu ������ӿ� 
         Caption         =   "������ӿ�"
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu ����̨ 
         Caption         =   "����̨"
         Shortcut        =   {F12}
      End
      Begin VB.Menu ���ڽڵ�ʼ� 
         Caption         =   "���ڽڵ�ʼ�"
      End
   End
End
Attribute VB_Name = "Note"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private saveNtxTimeNow As Single
Public pilotLightColor As Long, pilotLightX0 As Single, pilotLightY0 As Single, pilotLightX1 As Single, pilotLightY1 As Single
Public homeBackPicPath As String, updataSpeed As Long, Բ��뾶���� As Single
Private childNodeVisOld As ��ά����, childNodeVisPos As ��Ԫ��, childNodeFormOld As ��ά����

Private Sub ActionTimer_Timer()
    ObjectAction
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print KeyCode; ; Shift
    Select Case KeyCode
        Case vbKeyA
            If Shift = 0 Then
                Բ��ڵ�_Click
            End If
        Case 13 '�س�
            ȷ�ϴ�������ڵ�
        Case 27 'ESC
            DeselectObjcet
            If �ӽڵ���ͼ����.Visible Then
                �ӽڵ���ͼ��ť_Click 0
            End If
            fictitiousIndexLock = False
        Case vbKeyW
            If Shift = 0 Then �����ڵ�_Click
        Case vbKeyP
            If Shift = 0 Then ���ؽڵ�_Click
        Case vbKeyV
            If Shift = 0 Then ���껯��_Click
        Case vbKeyO
            If Shift = 3 Then
                ��һ��ѡ�нڵ�
            End If
        Case 46
            BehaviorIdSet
            DeleteSelectObjcet
        Case 107, 187 '+
            If Shift = 2 Then
                RollerEventHandling False
            ElseIf Shift = 0 Then
                �ڵ����Ӵ�С�仯 10, 1
            End If
        Case vbKeyC
            If Shift = 3 Then
                ӡ��_Click
            ElseIf Shift = 0 Then
                ��������_Click
            End If
        Case 109, 189 '-
            If Shift = 2 Then
                RollerEventHandling True
            ElseIf Shift = 0 Then
                �ڵ����Ӵ�С�仯 -10, -1
            End If
'        Case vbKey0
'            MainCoordinateSystemZero mouseV3Pos
        Case 37
            MapUpdata_AoVMove_Moving -10, 0
        Case 38
            MapUpdata_AoVMove_Moving 0, 10
        Case 39
            MapUpdata_AoVMove_Moving 10, 0
        Case 40
            MapUpdata_AoVMove_Moving 0, -10
        Case 192
            NoteControlDesk.Show
        Case vbKeyR
            ���ӷ�ת_Click
        Case vbKeyL
            If Shift = 1 Then
                ȡ��ѡ���ڵ���������
            End If
        Case vbKeyN
            If Shift = 1 Then
                ȡ��ѡ���ڵ����нڵ�
            ElseIf Shift = 0 Then
                �����ӽڵ���ɫ
            End If
        Case 48 To 57
            If Shift = 2 Then
                �ڵ�����ѡ������ KeyCode - 48
            ElseIf Shift = 0 Then
                �ڵ�����ѡ��ʹ�� KeyCode - 48
            ElseIf Shift = 1 Then
                �ڵ�����ѡ��ɾ�� KeyCode - 48
            End If
    End Select
End Sub
Private Sub ȷ�ϴ�������ڵ�()
    Dim i As Long, j As Long
    BehaviorIdSet
    If fictitiousIndexLock Then
        For i = 0 To UBound(fictitiousNote)
            With fictitiousNote(i)
                If .be Then
                    For j = 0 To UBound(.nodeLine)
                        If .nodeLine(j).direction = 1 Then
                            LineAdd fictitiousRootNodeId, NodeEdit_NewNode(.node(.nodeLine(j).target).t, .node(.nodeLine(j).target).content, .node(.nodeLine(j).target).setColor, .node(.nodeLine(j).target).setSize, .node(.nodeLine(j).target).realityX, .node(.nodeLine(j).target).realityY), .nodeLine(j).content, .nodeLine(j).size
                        ElseIf .nodeLine(j).direction = 2 Then
                            LineAdd NodeEdit_NewNode(.node(.nodeLine(j).Source).t, .node(.nodeLine(j).Source).content, .node(.nodeLine(j).Source).setColor, .node(.nodeLine(j).Source).setSize, .node(.nodeLine(j).Source).realityX, .node(.nodeLine(j).Source).realityY), fictitiousRootNodeId, .nodeLine(j).content, .nodeLine(j).size
                        ElseIf .nodeLine(j).direction = 3 Then
                            LineAdd fictitiousRootNodeId, .nodeLine(j).realityId, .nodeLine(j).content, .nodeLine(j).size
                        ElseIf .nodeLine(j).direction = 4 Then
                            LineAdd .nodeLine(j).realityId, fictitiousRootNodeId, .nodeLine(j).content, .nodeLine(j).size
                        End If
                    Next
                End If
            End With
        Next
        fictitiousIndexLock = False
    End If
End Sub
Private Sub ��һ��ѡ�нڵ�()
    Dim i As Long
    BehaviorIdSet
    For i = 0 To nSum
        With node(i)
            If .b Then
                If .select = True Or nodeTargetAim = i Then
                    �ڵ�ȥ�� i
                End If
            End If
        End With
    Next
End Sub
Private Function �ڵ�ȥ��(sN As Long)
    Dim i As Long
    For i = sN + 1 To nSum
        With node(i)
            If .b Then
                If .t = node(sN).t And .content = node(sN).content And .setColor = node(sN).setColor And .setSize = node(sN).setSize Then
                    ����ת�� i, sN
                    NodeDelete i
                End If
            End If
        End With
    Next
End Function
Private Function ����ת��(aN As Long, rN As Long)
    Dim i As Long
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                If .Source = aN Then
                    If .target = rN Then
                        LineDelete i
                    Else
                        LineReplace i, 0, rN, aN
                    End If
                    
                ElseIf .target = aN Then
                    If .Source = rN Then
                        LineDelete i
                    Else
                        LineReplace i, 1, rN, aN
                    End If
                End If
            End If
        End With
    Next
End Function
Private Sub �����ӽڵ���ɫ()
    If nodeTargetAim <> -1 Then
        NodeColorSelectForm.����ĸ�ڵ���� = nodeTargetAim
        NodeColorSelectForm.Show 1
    End If
End Sub
Private Sub Բ���ӽڵ�()
    Dim ���� As String, nidT As Long
    nidT = δѡ�д���(nodeTargetAim)
    If nidT <> -1 Then
        ���� = InBox("������Բ��뾶[1000,100000]��", Բ��뾶����)
        If promptBoxSelect = 0 Then
            Բ��뾶���� = ������ֵ(Val(����), 1000, 100000)
            NodeArray nidT, Բ��뾶����
        End If
    End If
End Sub
Private Sub ȡ��ѡ���ڵ����нڵ�()
    Dim i As Long
    For i = 0 To nSum
        With node(i)
            If .b Then
                .select = False
            End If
        End With
    Next
End Sub
Private Sub ȡ��ѡ���ڵ���������()
    Dim i As Long
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                .select = False
            End If
        End With
    Next
End Sub
Private Sub �ڵ�����ѡ��ɾ��(ByVal key As String)
    If nodeSelectKeyDic.Exists(key) Then
        nodeSelectKeyDic.Remove key
        lineSelectKeyDic.Remove key
    End If
End Sub
Private Sub �ڵ�����ѡ��ʹ��(ByVal key As String)
    Dim i As Long
    If nodeSelectKeyDic.Exists(key) Then
        For i = 0 To nSum
            With node(i)
                If .b Then
                    If InStr(1, nodeSelectKeyDic(key), "," & i & ",") > 0 Then
                        .select = True
                    Else
                        .select = False
                    End If
                End If
            End With
        Next
        For i = 0 To lSum
            With nodeLine(i)
                If .b Then
                    If InStr(1, lineSelectKeyDic(key), "," & i & ",") > 0 Then
                        .select = True
                    Else
                        .select = False
                    End If
                End If
            End With
        Next
    End If
End Sub
Private Sub �ڵ�����ѡ������(ByVal key As String)
    Dim i As Long, n As String, l As String
    n = ","
    For i = 0 To nSum
        With node(i)
            If .b Then
                If .select Then
                    n = n & i & ","
                End If
            End If
        End With
    Next
    nodeSelectKeyDic.Add key, n
    l = ","
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                If .select Then
                    l = l & i & ","
                End If
            End If
        End With
    Next
    lineSelectKeyDic.Add key, l
End Sub
Private Sub �ڵ����Ӵ�С�仯(������ As Single, ������ As Single)
    Dim i As Long
    For i = 0 To nSum
        With node(i)
            If .b Then
                If nodeTargetAim = i Or .select = True Then
                    .setSize = .setSize + ������
                    If .setSize > 500 Then
                        .setSize = 500
                    ElseIf .setSize < 50 Then
                        .setSize = 50
                    End If
                End If
            End If
        End With
    Next
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                If .select Then
                    .size = .size + ������
                    If .size > 10 Then
                        .size = 10
                    ElseIf .size < 1 Then
                        .size = 1
                    End If
                End If
            End If
        End With
    Next
End Sub
Private Sub �������ݸ���(���� As String)
    Dim i As Long
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                If .select Then
                    .content = ����
                End If
            End If
        End With
    Next
End Sub
Private Sub ���ƽڵ�Ϊ���ı�()
    Clipboard.Clear
    Clipboard.SetText nodeToTxt
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim dirPath As String
    PublicVarLoad2
    zoomFactor = 1: Բ��뾶���� = 1000
    If App.LogMode <> 0 Then
        HookMouse Me.hWnd
    End If
    notePrintNodeId = -1
    MeExeIdSet
    ProfilePath = Environ("USERPROFILE") & "\Documents\Note\"
    InstallPath = Environ("systemdrive") & "\ProgramData\Note\"
    LoadProfile
    Select Case ע���ע��("NodeNote", ".ntx")
        Case 0
            MsgBox "���ע��ʧ�ܣ����ù���Ա������У�"
        Case 1
            MsgBox "���ע��ɹ���"
    End Select
    MainCoordinateSystemDefinition
    �ӽڵ���ͼ����.Scale (0, �ӽڵ���ͼ����.height)-(�ӽڵ���ͼ����.width, 0)
    �ӽڵ���ͼ������.Scale (0, �ӽڵ���ͼ������.height)-(�ӽڵ���ͼ������.width, 0)
    If ��ǩ��.Checked = False Then NodePrint.Show
    If Command <> "" Then
        pilotLightColor = RGB(0, 255, 0)
        dirPath = Replace(Command, """", "")
'        MsgBox dirPath
        NoteFileRead dirPath
    Else
        pilotLightColor = RGB(255, 0, 0)
        newAddNote
    End If
    FictitiousNtxLoad
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'C:2;S:1;A:4
mainFormMouseState = True
Select Case Button
    Case 1
        nodeClickAim = NodeCheck(X, Y)
        If nodeClickAim = -1 Then  '�ƶ�����ϵ
            Select Case Shift
                Case 0
                    allNodeMoveLock = True
                    allNodeMoveStart.X = X: allNodeMoveStart.Y = Y
                    If lineAddLock = True Then lineAddLock = False
                Case 2
                    allNodeMoveStart.X = X: allNodeMoveStart.Y = Y
                    selectMoveLock = True
                Case 4
                    regionalSelectStart.X = X: regionalSelectStart.Y = Y
                    regionalSelectLock = True
            End Select
        Else
            Select Case Shift
                Case 0
                    If MultipointConnection = False Then
                        If lineAddLock = False Then
                            lineAddLock = True
                            lineAddStrat.X = X: lineAddStrat.Y = Y
                            lineAddSource = nodeClickAim
                        Else
                            BehaviorIdSet
                            If lineAddSource <> nodeClickAim Then
                                LineAdd lineAddSource, nodeClickAim, "", lineDefaultSize
                            End If
                            lineAddLock = False
                        End If
                        nodeMoveLock = True
                        nodeMoveStart.X = X: nodeMoveStart.Y = Y
                    End If
                Case 4
                    ChainSelection nodeClickAim, 0
            End Select
        End If
    Case 2
        NodeEditeStart X, Y
    Case 4
        Select Case Shift
            Case 4
                DirectSelect
        End Select
End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mousePos.X = X: mousePos.Y = Y
    mouseV3Pos.X = X: mouseV3Pos.Y = Y: mouseV3Pos.z = zoomFactor
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mainFormMouseState = False
allNodeMoveLock = False
nodeMoveLock = False
regionalSelectLock = False
selectMoveLock = False
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
NoteFileRead Data.Files(1)
End Sub

Private Sub Form_Resize()
If WindowState = 1 Then Exit Sub
MainCoordinateSystemDefinition
If Me.height < 4000 Then Me.Enabled = False: Me.height = 4000: Me.Enabled = True
If Me.width < 4000 Then Me.Enabled = False: Me.width = 4000: Me.Enabled = True
'PilotLight.left = Me.Width * zoomFactor - 240 * zoomFactor
Note.pilotLightX0 = Me.width - 300
Note.pilotLightX1 = Me.width - 230
Note.pilotLightY0 = Me.height - 300
Note.pilotLightY1 = Me.height - 200

GlobalView.left = Me.width * zoomFactor - GlobalView.width - 120 * zoomFactor
GlobalView.Top = GlobalView.height + 120 * zoomFactor
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveProfile
'UnHookMouse Me.hWnd
If App.LogMode <> 0 Then
    UnHookMouse Me.hWnd
End If
End
End Sub

Private Sub GlobalView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim dX As Single: Dim dY As Single
    If Button = 1 Or Button = 2 Then
        dX = Note.width / 2 - X
        dY = Note.height / 2 - Y
        MapUpdata_AoVMove_Moving dX, dY
        mouseMapPos.X = Note.width / 2
        mouseMapPos.Y = Note.height / 2
        mapMoveLock = True
    End If
End Sub

Private Sub GlobalView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If mapMoveLock = True And mapGetMousePosLock = False Then mouseMapPos.X = X: mouseMapPos.Y = Y
End Sub

Private Sub GlobalView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mapMoveLock = False
End Sub

Private Sub MainTime_Timer()
    If �ӽڵ���ͼ����.Visible = False Then
        Updata
    End If
End Sub

Private Sub MapUpdataTimer_Timer()
    If nSum > 0 And ȫ����ͼ.Checked = True And �ӽڵ���ͼ����.Visible = False Then
        MapUpdata
    End If
End Sub

Private Sub NodePrintBox_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub NodePrintBox_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub PLC_Timer()
    Select Case noteSaveCheck
        Case 0
            pilotLightColor = RGB(255, 0, 0)
        Case 1
            pilotLightColor = RGB(255, 165, 0)
        Case 2
            pilotLightColor = RGB(0, 255, 0)
    End Select
    On Error GoTo Er
        If Clipboard.GetText = "" Then ճ��.Enabled = False Else ճ��.Enabled = True
        If bHLSum < 1 Then ����.Enabled = False Else ����.Enabled = True
        If redoSum < 1 Then ����.Enabled = False Else ����.Enabled = True
        If ��������Ҫ��ʾ��ʾ����ʱ > 0 Then
            ��������Ҫ��ʾ��ʾ����ʱ = ��������Ҫ��ʾ��ʾ����ʱ - 1
            ����������ʾ����.Caption = Format(zoomFactor, "0.0") & "X"
            If ����������ʾ����.Visible = False Then ����������ʾ����.Visible = True
        ElseIf ����������ʾ����.Visible Then
            ����������ʾ����.Visible = False
        End If
        �˵��������
        If saveNtxTime <> 0 Then
            saveNtxTimeNow = saveNtxTimeNow + 0.5
            If saveNtxTimeNow > saveNtxTime And pilotLightColor = RGB(255, 165, 0) Then
                ����ʼ�_Click
                saveNtxTimeNow = 0
            End If
        End If
Er:
End Sub
Private Function �˵��������()
If ��ʾ����ڵ���.Checked = True Then ��ʾ˳��ڵ���.Checked = False: ��ʾȫ���ڵ���.Checked = False
If ��ʾ˳��ڵ���.Checked = True Then ��ʾ����ڵ���.Checked = False: ��ʾȫ���ڵ���.Checked = False
If ��ʾȫ���ڵ���.Checked = True Then ��ʾ����ڵ���.Checked = False: ��ʾ˳��ڵ���.Checked = False

If ��ʾ��������.Checked = True Then ��ʾ˳������.Checked = False: ��ʾȫ������.Checked = False
If ��ʾ˳������.Checked = True Then ��ʾ��������.Checked = False: ��ʾȫ������.Checked = False
If ��ʾȫ������.Checked = True Then ��ʾ��������.Checked = False: ��ʾ˳������.Checked = False

End Function


Private Sub RGBɫ��VBColor��ת����_Click()
    RGBTOVBColorForm.Show
End Sub

Private Sub ���½ڵ㻯(txt As String)
    Dim i As Long, �س����� As Long, �б� As Long, sT As String, idT As Long
    �س����� = 1
    For i = 1 To Len(txt)
        sT = Mid(txt, i, 1)
        �б� = �б� + 1
        idT = NodeEdit_NewNode(�س�ת��(sT), "", &HFFBF00, nodeDefaultSize, �б� * imageToNtx_StepX, Me.height + �س����� * imageToNtx_StepY)
        If i > 1 Then LineAdd idT - 1, idT, "", lineDefaultSize
        If sT = vbLf Then �س����� = �س����� + 1: �б� = 0
    Next
End Sub
Private Function �س�ת��(s As String) As String
    Select Case s
        Case vbCr
            �س�ת�� = "\n"
        Case vbLf
            �س�ת�� = "\r"
        Case "\n"
            �س�ת�� = vbCr
        Case "\r"
            �س�ת�� = vbLf
        Case Else
            �س�ת�� = s
    End Select
End Function

Private Sub ����ʼ�_Click()
    Dim filePath As String
    On Error GoTo Er
    If Dir(ntxPath) = "" Then
        filePath = �Ի���ѡȡ�����ļ�·��("�ڵ�ʼ� (*.ntx)|*.ntx|�����ļ� (*.*)|*.*")
        If filePath <> "" Then NoteFileWrite_203 filePath
    Else
        NoteFileWrite_203 ntxPath
    End If

Exit Sub
Er:
    MsgBox "����ʼ�ʧ�ܣ�ԭ��" & Err.Description, 16, "����ʼ�"
End Sub

Private Sub ����ɫ_Click()
    With CommonDialog1
        .Flags = 1
        .color = Me.BackColor
        .ShowColor
        Me.BackColor = .color
    End With
End Sub

Private Sub ����ɫ2_Click()
With CommonDialog1
    .Flags = 1
    .color = NodeInput.BackColor
    .ShowColor
    NodeInputBackColor = .color
End With
NodeInput.BackColor = NodeInputBackColor
End Sub

Private Sub ����ͼ_Click()
    On Error GoTo Er
        With CommonDialog1
            .Flags = cdlOFNHideReadOnly
            .Filter = "ͼƬ�ļ� (*.bmp;*.png;*jpg)|*.bmp;*.png;*jpg|�����ļ� (*.*)|*.*"
            .FilterIndex = 1
            .ShowOpen
            homeBackPicPath = .filename
            ���ر���ͼ homeBackPicPath
        End With
Er:
End Sub
Public Sub ���ر���ͼ(fP As String)
    On Error GoTo Er
        Me.Picture = LoadPicture(fP)
    Exit Sub
Er:
    MsgBox "���ر���ͼʧ�ܣ�ԭ��" & Err.Description
End Sub
Private Sub ��ǩ��_Click()
If ��ǩ��.Checked = False Then
    ��ǩ��.Checked = True
    Unload NodePrint
Else
    ��ǩ��.Checked = False
    nodePrintBeLock = False
End If
End Sub

Private Sub �����ڵ�_Click()
    Dim rootNid As Long, tempValue As Single
    rootNid = δѡ�д���(nodeTargetAim)
    If rootNid <> -1 Then
        If NCF_NodeValueControl(���ı�ת��(node(rootNid).content), tempValue) Then
            ClearNode_ToTreeTxtLock
            NodeWave rootNid, tempValue, False, 0, 0
        End If
    End If
End Sub

Private Sub �ʺ�Ȧ_Click()
If �ʺ�Ȧ.Checked = False Then �ʺ�Ȧ.Checked = True Else �ʺ�Ȧ.Checked = False
End Sub

Private Sub �ʺ���_Click()
    �ʺ���.Checked = Not �ʺ���.Checked
    If �ʺ���.Checked Then
        �������.Checked = False
    End If
End Sub

Private Sub ����_Click()
NodeFind.Show
End Sub

Private Sub ����_Click()
RedoSet
RevokeBehavior
End Sub


Private Sub �򿪱ʼ�_Click()
    Dim filePath As String
    filePath = �Ի���ѡȡ���ļ�·��("�ڵ�ʼ� (*.ntx)|*.ntx|�����ļ� (*.*)|*.*")
    If filePath <> "" Then
        newAddNote
        NoteFileRead filePath
    End If
End Sub
Private Function �Ի���ѡȡ���ļ�·��(������ As String) As String
    Dim filePath As String
    With CommonDialog1
        .CancelError = True
        On Error GoTo ErrHandler
            .Flags = cdlOFNHideReadOnly
            .Filter = ������
            .FilterIndex = 1
            .ShowOpen
            �Ի���ѡȡ���ļ�·�� = .filename
    End With
    Exit Function
ErrHandler:
End Function
Private Function �Ի���ѡȡ�����ļ�·��(�������� As String) As String
    Dim filePath As String
    With CommonDialog1
        .CancelError = True
        On Error GoTo ErrHandler
            .Flags = cdlOFNOverwritePrompt
            .Filter = ��������
            .FilterIndex = 1
            .ShowSave
            �Ի���ѡȡ�����ļ�·�� = .filename
    End With
    Exit Function
ErrHandler:
End Function

Private Sub ������ڵ��ļ�Ŀ¼_Click()
    On Error GoTo Er
        Shell "cmd.exe /c start """" """ & fictitiousNtxPath & """"
    Exit Sub
Er:
    MsgBox "������ڵ��ļ�Ŀ¼ʧ�ܣ�ԭ��" & Err.Description, 16, "����"
End Sub

Private Sub ������ӿ�_Click()
    NodeNetwork.Show
End Sub

Private Sub ��ӡ��PNGͼƬ_Click()
    Dim fP As String
    On Error GoTo Er:
        fP = �Ի���ѡȡ�����ļ�·��("ͼƬ (*.png)|*.png")
        If fP <> "" Then
            NotePrint ��ӡ������
            SavePicture ��ӡ������.image, fP
        End If
    Exit Sub
Er:
    MsgBox "��ӡʧ�ܣ�ʧ��ԭ��" & Err.Description, 16, "��ӡ����"
End Sub

Private Sub ����TXT����_Click()
    Dim filePath As String, nIdTemp As Long, outS As String
    On Error GoTo Er
        nIdTemp = δѡ�д���(nodeTargetAim)
        If nIdTemp <> -1 Then
            filePath = �Ի���ѡȡ�����ļ�·��("�ı��ļ� (*.txt)|*.txt")
            If filePath <> "" Then
                ClearNode_ToTreeTxtLock
                �ڵ�ת���� outS, nIdTemp
                SaveFile_All filePath, outS
                MsgBox "�����ɹ���", 64, "����TXT����"
            End If
        End If
    Exit Sub
Er:
    MsgBox "��������ԭ��" & Err.Description, 16, "����TXT����"
End Sub

Private Function �ڵ�ת����(outS As String, sNid As Long)
    Dim i As Long
    node(sNid).toTreeTxtLock = True
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                If .Source = sNid And node(.target).toTreeTxtLock = False Then
                    outS = outS & �س�ת��(node(.target).t)
                    �ڵ�ת���� outS, .target
                    Exit Function
                End If
            End If
        End With
    Next
End Function

Private Sub ����λͼ_Click()
    Dim filePath As String, nIdTemp As Long
    On Error GoTo Er
        nIdTemp = δѡ�д���(nodeTargetAim)
        If nIdTemp <> -1 Then
            filePath = �Ի���ѡȡ�����ļ�·��("ͼƬ�ļ� (*.bmp)|*.bmp")
            If filePath <> "" Then
                If NoteToImage(filePath, λͼ�����) Then
                    MsgBox "�����ɹ���", 64, "�����ʼǵ�λͼ�ļ�"
                End If
            End If
        End If
    Exit Sub
Er:
    MsgBox "��������ԭ��" & Err.Description, 16, "�����ʼǵ�λͼ�ļ�"
End Sub

Private Sub �����ı��ļ�_Click()
    Dim filePath As String, nIdTemp As Long
On Error GoTo Er
    nIdTemp = δѡ�д���(nodeTargetAim)
    If nIdTemp <> -1 Then
        filePath = �Ի���ѡȡ�����ļ�·��("�ı��ļ� (*.txt)|*.txt")
        If filePath <> "" Then
            NoteToTreeTXT filePath, nIdTemp
            MsgBox "�����ɹ���", 64, "�����ʼǵ��ı��ļ�"
        End If
    Else
        MsgBox "δѡ�нڵ㣬����������Ч��", 64, "�����ʼǵ��ı��ļ�"
    End If
    Exit Sub
Er:
    MsgBox "��������ԭ��" & Err.Description, 16, "�����ʼǵ��ı��ļ�"
End Sub
Public Function δѡ�д���(n As Long) As Long
    Dim i As Long
    If n = -1 Then
        δѡ�д��� = -1
        For i = 0 To nSum
            If node(i).b Then
                If node(i).select = True Then
                    δѡ�д��� = i
                End If
            End If
        Next
    Else
        δѡ�д��� = n
    End If
End Function

Private Sub ����TXT����_Click()
    Dim fP As String, txt As String
    On Error GoTo Er
        fP = �Ի���ѡȡ���ļ�·��("�ı��ļ� (*.txt)|*.txt")
        If fP <> "" Then
            ReadFile_ALL_HV fP, txt
            BehaviorIdSet
            ���½ڵ㻯 txt
        End If
    Exit Sub
Er:
    MsgBox "����ʧ�ܣ�ԭ��" & Err.Description, 16, "����TXT����"
End Sub

Private Sub ����λͼ_Click()
    Dim fP As String
    On Error GoTo Er
    fP = �Ի���ѡȡ���ļ�·��("ͼƬ�ļ� (*.bmp;*.png;*.jpg)|*.bmp;*.png;*.jpg")
    If fP <> "" Then
        λͼ��ȡ������.Picture = LoadPicture(fP)
        BehaviorIdSet
        λͼ�ڵ㻯 λͼ��ȡ������
    End If
    Exit Sub
Er:
    MsgBox "����ʧ�ܣ�ԭ��" & Err.Description, 16, "����λͼ"
End Sub

Private Sub �����ı��ļ�_Click()
    On Error GoTo Er:
        ����TXT�ļ�·�� = �Ի���ѡȡ���ļ�·��("�ı��ĵ� (*.txt)|*.txt")
        If ����TXT�ļ�·�� <> "" Then
            BehaviorIdSet
            TreeTXTToNtx
        End If
    Exit Sub
Er:
    MsgBox "����ʧ�ܣ�ԭ��" & Err.Description, 16, "�����ı��ļ�"
End Sub

Private Sub ����_Click()
    CopyObject False
End Sub

Private Function nodeToTxt() As String
    Dim i As Long
    For i = 0 To nSum
        With node(i)
            If .b = True Then
                If .select = True Or i = nodeTargetAim Then
                    RTBtemp.TextRTF = .content
                    nodeToTxt = nodeToTxt & Chr(34) & .t & """ : """ & RTBtemp.Text & """ , "
                End If
            End If
        End With
    Next
    nodeToTxt = Mid(nodeToTxt, 1, Len(nodeToTxt) - 2)
End Function

Private Sub ���ڽڵ�ʼ�_Click()
    AboutNote.Show
End Sub

Private Sub ��ͼˢ�¼��_Click()
    Dim sT As String
    sT = InBox("�������ͼˢ�¼�������ԽС����Ҫ��Խ�ߣ����������[10~100]��", updataSpeed)
    If promptBoxSelect = 0 Then
        updataSpeed = ������ֵ(Val(sT), 10, 100)
        MainTime.interval = updataSpeed
    End If
End Sub

Private Sub ����_Click()
    BehaviorIdSet
    CopyObject True
End Sub

Private Sub �ڵ��һ_Click()
    ��һ��ѡ�нڵ�
End Sub

Private Sub �ڵ����_Click()
    �ڵ����.Checked = Not �ڵ����.Checked
End Sub

Private Sub �ڵ��嵥_Click()
    NodeListVis.Show
End Sub


Private Sub �ص�_Click()
    �ص�.Checked = Not �ص�.Checked
End Sub

Private Sub ����_Click()
    ����.Checked = Not ����.Checked
End Sub

Private Sub ����̨_Click()
NoteControlDesk.Show
End Sub

Private Sub ���ӷ�ת_Click()
    ConnectionReversal
End Sub

Private Sub ��������_Click()
    Dim sT As String
    sT = InBox("������ѡ�����ӵ���ʾ���ݣ�")
    If promptBoxSelect = 0 Then
        �������ݸ��� sT
    End If
End Sub

Private Sub �����嵥_Click()
    LineListVis.Show
End Sub

Private Sub ���Ϊ_Click()
    Dim filePath As String
    On Error GoTo Er
    filePath = �Ի���ѡȡ�����ļ�·��("�ڵ�ʼ� (*.ntx)|*.ntx|�����ļ� (*.*)|*.*")
    If filePath <> "" Then NoteFileWrite_203 filePath
    Exit Sub
Er:
    MsgBox "���Ϊ�ʼ�ʧ�ܣ�ԭ��" & Err.Description, 16, "���Ϊ�ʼ�"
End Sub

Private Sub �������_Click()
    �������.Checked = Not �������.Checked
    If �������.Checked Then
        �ʺ���.Checked = False
    End If
End Sub

Private Sub ȫ��͸��_Click()
If ȫ��͸��.Checked = True Then
    ȫ��͸��.Checked = False: FormTransparent Me, 255
Else
    ȫ��͸��.Checked = True: FormTransparent Me, 50
    ȫ��͸��.Checked = False: ȫ��͸��.Checked = False
End If
End Sub

Private Sub ȫ��͸��_Click()
If ȫ��͸��.Checked = True Then
    ȫ��͸��.Checked = False: FormTransparent Me, 255
Else
    ȫ��͸��.Checked = True: FormTransparent Me, 125
    ȫ��͸��.Checked = False: ȫ��͸��.Checked = False
End If
End Sub

Private Sub ȫ��͸��_Click()
If ȫ��͸��.Checked = True Then
    ȫ��͸��.Checked = False: FormTransparent Me, 255
Else
    ȫ��͸��.Checked = True: FormTransparent Me, 200
    ȫ��͸��.Checked = False: ȫ��͸��.Checked = False
End If
End Sub

Private Sub ȫ��͸��2_Click()
If ȫ��͸��2.Checked = True Then
    ȫ��͸��2.Checked = False: FormTransparent NodeInput, 255
Else
    ȫ��͸��2.Checked = True: FormTransparent NodeInput, 50
    ȫ��͸��2.Checked = False: ȫ��͸��2.Checked = False
End If
End Sub

Private Sub ȫ��͸��2_Click()
If ȫ��͸��2.Checked = True Then
    ȫ��͸��2.Checked = False: FormTransparent NodeInput, 255
Else
    ȫ��͸��2.Checked = True: FormTransparent NodeInput, 125
    ȫ��͸��2.Checked = False: ȫ��͸��2.Checked = False
End If
End Sub

Private Sub ȫ��͸��2_Click()
If ȫ��͸��2.Checked = True Then
    ȫ��͸��2.Checked = False: FormTransparent NodeInput, 255
Else
    ȫ��͸��2.Checked = True: FormTransparent NodeInput, 200
    ȫ��͸��2.Checked = False: ȫ��͸��2.Checked = False
End If
End Sub
Private Sub ȫ��͸��3_Click()
If ȫ��͸��3.Checked = True Then
    ȫ��͸��3.Checked = False: FormTransparent NodePrint, 255
Else
    ȫ��͸��3.Checked = True: FormTransparent NodePrint, 50
    ȫ��͸��3.Checked = False: ȫ��͸��3.Checked = False
End If
End Sub

Private Sub ȫ��͸��3_Click()
If ȫ��͸��3.Checked = True Then
    ȫ��͸��3.Checked = False: FormTransparent NodePrint, 255
Else
    ȫ��͸��3.Checked = True: FormTransparent NodePrint, 125
    ȫ��͸��3.Checked = False: ȫ��͸��3.Checked = False
End If
End Sub

Private Sub ȫ��͸��3_Click()
If ȫ��͸��3.Checked = True Then
    ȫ��͸��3.Checked = False: FormTransparent NodePrint, 255
Else
    ȫ��͸��3.Checked = True: FormTransparent NodePrint, 200
    ȫ��͸��3.Checked = False: ȫ��͸��3.Checked = False
End If
End Sub
Private Sub ȫ����ͼ_Click()
If ȫ����ͼ.Checked = True Then
    ȫ����ͼ.Checked = False: GlobalView.Visible = False
Else
    ȫ����ͼ.Checked = True: GlobalView.Visible = True
End If
End Sub

Private Sub ȫѡ_Click()
AllSelection
End Sub

Private Sub ɾ������ͼ_Click()
    Me.Picture = Nothing
    homeBackPicPath = ""
End Sub

Private Sub ����Ĭ�Ͻڵ��С_Click()
    Dim sT As String
    sT = InBox("����ÿ���½��ڵ�Ĵ�С[50,500]��", nodeDefaultSize)
    If promptBoxSelect = 0 Then
        nodeDefaultSize = ������ֵ(Val(sT), 50, 500)
    End If
End Sub

Private Sub ����Ĭ�����ӿ��_Click()
    Dim sT As String
    sT = InBox("�������½����ӿ��[1,10]��", lineDefaultSize)
    If promptBoxSelect = 0 Then
        lineDefaultSize = ������ֵ(Val(sT), 1, 10)
    End If
End Sub

Private Sub �����ɫ_Click()
    �����ӽڵ���ɫ
End Sub

Private Sub ʼ����ʾѡ����_Click()
If ʼ����ʾѡ����.Checked = True Then ʼ����ʾѡ����.Checked = False Else ʼ����ʾѡ����.Checked = True
End Sub

Private Sub ʼ����ʾѡ��_Click()
If ʼ����ʾѡ��.Checked = True Then ʼ����ʾѡ��.Checked = False Else ʼ����ʾѡ��.Checked = True
End Sub

Private Sub �˳�_Click()
End
End Sub

Private Sub λͼ�����_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print X; ; Y
End Sub

Private Sub ������ɫ_Click()
    With CommonDialog1
        .Flags = 1
        .color = Me.ForeColor
        .ShowColor
        Me.ForeColor = .color
    End With
End Sub

Private Sub ��ʾ�ڵ����ID_Click()
If ��ʾ�ڵ����ID.Checked = True Then ��ʾ�ڵ����ID.Checked = False Else ��ʾ�ڵ����ID.Checked = True
End Sub

Private Sub ��ʾ����ڵ���_Click()
If ��ʾ����ڵ���.Checked = False Then
    ��ʾ����ڵ���.Checked = True: ��ʾȫ���ڵ���.Checked = False: ��ʾ˳��ڵ���.Checked = False
Else
    ��ʾ����ڵ���.Checked = False
End If

End Sub

Private Sub ��ʾ��������_Click()
If ��ʾ��������.Checked = False Then
    ��ʾ��������.Checked = True: ��ʾȫ������.Checked = False: ��ʾ˳������.Checked = False
Else
    ��ʾ��������.Checked = False
End If

End Sub

Private Sub ��ʾȫ������_Click()
If ��ʾȫ������.Checked = False Then
    ��ʾȫ������.Checked = True: ��ʾ˳������.Checked = False: ��ʾ��������.Checked = False
Else
    ��ʾȫ������.Checked = False
End If

End Sub

Private Sub ��ʾ˳��ڵ���_Click()
If ��ʾ˳��ڵ���.Checked = False Then
    ��ʾ˳��ڵ���.Checked = True: ��ʾȫ���ڵ���.Checked = False: ��ʾ����ڵ���.Checked = False
Else
    ��ʾ˳��ڵ���.Checked = False
End If

End Sub

Private Sub ��ʾ˳������_Click()
If ��ʾ˳������.Checked = False Then
    ��ʾ˳������.Checked = True: ��ʾȫ������.Checked = False: ��ʾ��������.Checked = False
Else
    ��ʾ˳������.Checked = False
End If

End Sub

Private Sub ��ʾȫ���ڵ���_Click()
If ��ʾȫ���ڵ���.Checked = False Then
    ��ʾȫ���ڵ���.Checked = True: ��ʾ˳��ڵ���.Checked = False: ��ʾ����ڵ���.Checked = False
Else
    ��ʾȫ���ڵ���.Checked = False
End If

End Sub

Private Sub ���ؽڵ�_Click()
    NodePixel
End Sub

Private Sub �½��ʼ�_Click()
    newAddNote
End Sub

Private Sub ѡ��_Click()
    SelectDisplayObjcet
End Sub

Private Sub ѡ������_Click()
    ȡ��ѡ���ڵ����нڵ�
End Sub

Private Sub ѡ������_Click()
    ȡ��ѡ���ڵ���������
End Sub

Private Sub ӡ��_Click()
    On Error GoTo Er
    ���ƽڵ�Ϊ���ı�
Er:
End Sub

Private Sub �����滻_Click()
    NodeReplace.Show
End Sub

Private Sub Բ��ڵ�_Click()
    Բ���ӽڵ�
End Sub

Private Sub ճ��_Click()
    BehaviorIdSet
    PasteObject
End Sub

Private Sub �ö�_Click()
If �ö�.Checked = True Then
    �ö�.Checked = False: FormStick NodePrint, False
Else
    �ö�.Checked = True: FormStick NodePrint, True
End If
Me.SetFocus
End Sub

Private Sub ����_Click()
RedoBehavior
End Sub

Private Sub �ӽڵ���ͼ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    childNodeVisOld.X = X
    childNodeVisOld.Y = Y
End Sub

Private Sub �ӽڵ���ͼ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And childNodeVisOld.X <> 0 And childNodeVisOld.Y <> 0 Then
        �ӽڵ���ͼ.left = �ӽڵ���ͼ.left + X - childNodeVisOld.X
        �ӽڵ���ͼ.Top = �ӽڵ���ͼ.Top + Y - childNodeVisOld.Y
    End If
End Sub

Private Sub �ӽڵ���ͼ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    �ӽڵ���ͼ_MouseMove Button, Shift, X, Y
End Sub

Private Sub �ӽڵ���ͼ������_DblClick()
    �ӽڵ���ͼ��ť_Click 1
End Sub

Private Sub �ӽڵ���ͼ������_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    childNodeFormOld.X = X
    childNodeFormOld.Y = Y
End Sub

Private Sub �ӽڵ���ͼ������_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And childNodeFormOld.X <> 0 And childNodeFormOld.Y <> 0 Then
        If �ӽڵ���ͼ��ť(1).Caption = "-" Then
            �ӽڵ���ͼ��ť_Click 1
        End If
        �ӽڵ���ͼ����.left = �ӽڵ���ͼ����.left + X - childNodeFormOld.X
        �ӽڵ���ͼ����.Top = �ӽڵ���ͼ����.Top + Y - childNodeFormOld.Y
    End If
End Sub
Private Sub �ӽڵ���ͼ������_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    �ӽڵ���ͼ������_MouseMove Button, Shift, X, Y
End Sub
Private Sub �ӽڵ���ͼ��ť_Click(Index As Integer)
    Select Case Index
        Case 0
            nodeTargetAim = -1
            childNodeVisLock = False
            �ӽڵ���ͼ����.Visible = False
        Case 1
            If �ӽڵ���ͼ��ť(1).Caption = "��" Then
                With childNodeVisPos
                    .xE = �ӽڵ���ͼ����.width
                    .xS = �ӽڵ���ͼ����.left
                    .yE = �ӽڵ���ͼ����.height
                    .yS = �ӽڵ���ͼ����.Top
                End With
                �ӽڵ���ͼ����.Top = Me.height
                �ӽڵ���ͼ����.left = 0
                �ӽڵ���ͼ����.height = Me.height
                �ӽڵ���ͼ����.width = Me.width
                �ӽڵ���ͼ��ť(1).Caption = "-"
                �ӽڵ���ͼ������.left = (�ӽڵ���ͼ����.width - �ӽڵ���ͼ������.width) / 2
            Else
                With childNodeVisPos
                    �ӽڵ���ͼ����.Top = .yS
                    �ӽڵ���ͼ����.left = .xS
                    �ӽڵ���ͼ����.height = .yE
                    �ӽڵ���ͼ����.width = .xE
                    �ӽڵ���ͼ��ť(1).Caption = "��"
                    �ӽڵ���ͼ������.left = 0
                End With
            End If
        Case 2
            On Error Resume Next
                Shell App.EXEName & ".exe " & childNodeVisNtxPath, vbNormalFocus
    End Select
End Sub


Private Sub �Զ�������_Click()
    Dim sT As String
    On Error GoTo Er
    sT = InBox("�������Զ�����ʱ����(��λ:��,����0�����Զ����棡)", saveNtxTime)
    If promptBoxSelect = 0 Then
        saveNtxTime = Val(sT)
    End If
    Exit Sub
Er:
    MsgBox "����ʧ�ܣ�ԭ��" & Err.Description, 16, "�����Զ�������"
    saveNtxTime = 0
End Sub

Private Sub ����_Click()
On Error GoTo Er
With CommonDialog1
    .Flags = 1
    .FontName = Me.Font.name
    .FontBold = Me.Font.Bold
    .FontSize = MainFormFontSize
    .FontItalic = Me.Font.Italic
    .FontUnderline = Me.Font.Underline
    .FontStrikethru = Me.Font.Strikethrough
    .ShowFont
    Me.Font.name = .FontName '��������
    Me.Font.Bold = .FontBold  '�Ӵ֣�
    MainFormFontSize = .FontSize '�����С
    Me.Font.Italic = .FontItalic '��б��
    Me.Font.Underline = .FontUnderline '�»��ߣ�
    Me.Font.Strikethrough = .FontStrikethru 'ɾ����
End With
Er:
End Sub

Private Sub ���껯��_Click()
    On Error GoTo Er
    nodeAttributedToIntegers = Val(InBox("�����뻯����������", nodeAttributedToIntegers))
    If promptBoxSelect = 0 Then
        NodePositionVague nodeAttributedToIntegers
    End If
    Exit Sub
Er:
    MsgBox "���껯��ʧ�ܣ�ԭ��" & Err.Description, 16, "���껯��"
End Sub
