VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Note 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Note"
   ClientHeight    =   9120
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   14760
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
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   9120
   ScaleWidth      =   14760
   StartUpPosition =   2  '��Ļ����
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
      Height          =   555
      Left            =   2640
      TabIndex        =   1
      Top             =   8400
      Visible         =   0   'False
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   979
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Note_Main.frx":1084A
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
   Begin VB.Label PilotLight 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "  "
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   14480
      TabIndex        =   2
      ToolTipText     =   "��ɫ���޸��ѱ��棻��ɫ���ʼ��Ѵ��ڵ��޸�δ���棻��ɫ���½��ʼǻ�δ����"
      Top             =   120
      Width           =   120
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
      Begin VB.Menu Cut4 
         Caption         =   "-"
      End
      Begin VB.Menu ���� 
         Caption         =   "����"
         Shortcut        =   ^F
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
      Begin VB.Menu ȫ����ͼ 
         Caption         =   "ȫ����ͼ"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
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
         Begin VB.Menu ����ɫ 
            Caption         =   "����ɫ"
         End
         Begin VB.Menu ������ɫ 
            Caption         =   "������ɫ"
         End
         Begin VB.Menu Cut9 
            Caption         =   "-"
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
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu ���� 
         Caption         =   "����"
         Begin VB.Menu �ڵ�����TXT�ı� 
            Caption         =   "�ڵ�����TXT�ı�"
         End
      End
      Begin VB.Menu ���ӷ�ת 
         Caption         =   "���ӷ�ת(R)"
      End
      Begin VB.Menu Cut11 
         Caption         =   "-"
      End
      Begin VB.Menu �Զ������� 
         Caption         =   "�Զ�������"
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu ����̨ 
         Caption         =   "����̨"
         Shortcut        =   {F12}
         Visible         =   0   'False
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
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
Select Case KeyCode
    Case 27
        DeselectObjcet
    Case 46
        BehaviorIdSet
        DeleteSelectObjcet
    Case 107 '+
        RollerEventHandling False
    Case 109 '-
        RollerEventHandling True
    Case vbKey0
        MainCoordinateSystemZero mouseV3Pos
    Case 37
        MapUpdata_AoVMove_Moving -10, 0
    Case 38
        MapUpdata_AoVMove_Moving 0, 10
    Case 39
        MapUpdata_AoVMove_Moving 10, 0
    Case 40
        MapUpdata_AoVMove_Moving 0, -10
    Case vbKeyR
        ���ӷ�ת_Click
End Select
End Sub

Private Sub Form_Load()
Dim dirPath As String
zoomFactor = 1
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
If ��ǩ��.Checked = False Then NodePrint.Show
If Command <> "" Then
    dirPath = Replace(Command, """", "")
    NoteFileRead dirPath
Else
    newAddNote
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'C:2;S:1;A:4
Select Case Button
    Case 1
        nodeClickAim = NodeCheck(x, y)
        If nodeClickAim = -1 Then  '�ƶ�����ϵ
            Select Case Shift
                Case 0
                    allNodeMoveLock = True
                    allNodeMoveStart.x = x: allNodeMoveStart.y = y
                    If lineAddLock = True Then lineAddLock = False
                Case 2
                    allNodeMoveStart.x = x: allNodeMoveStart.y = y
                    selectMoveLock = True
                Case 4
                    regionalSelectStart.x = x: regionalSelectStart.y = y
                    regionalSelectLock = True
            End Select
        Else
            Select Case Shift
                Case 0
                    If MultipointConnection = False Then
                        If lineAddLock = False Then
                            lineAddLock = True
                            lineAddStrat.x = x: lineAddStrat.y = y
                            lineAddSource = nodeClickAim
                        Else
                            BehaviorIdSet
                            If lineAddSource <> nodeClickAim Then
                                LineAdd lineAddSource, nodeClickAim
                            End If
                            lineAddLock = False
                        End If
                        nodeMoveLock = True
                        nodeMoveStart.x = x: nodeMoveStart.y = y
                    End If
                Case 4
                    ChainSelection nodeClickAim, 0
            End Select
        End If
    Case 2
        NodeEditeStart x, y
    Case 4
        Select Case Shift
            Case 4
                DirectSelect
        End Select
End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
mousePos.x = x: mousePos.y = y
mouseV3Pos.x = x: mouseV3Pos.y = y: mouseV3Pos.z = zoomFactor
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
allNodeMoveLock = False
nodeMoveLock = False
regionalSelectLock = False
selectMoveLock = False
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
NoteFileRead Data.Files(1)
End Sub

Private Sub Form_Resize()
If WindowState = 1 Then Exit Sub
MainCoordinateSystemDefinition
If Me.Height < 10000 Then Me.Enabled = False: Me.Height = 10000: Me.Enabled = True
If Me.Width < 15000 Then Me.Enabled = False: Me.Width = 15000: Me.Enabled = True
PilotLight.left = Me.Width * zoomFactor - 240 * zoomFactor
GlobalView.left = Me.Width * zoomFactor - GlobalView.Width - 120 * zoomFactor
GlobalView.Top = GlobalView.Height + 120 * zoomFactor
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveProfile
'UnHookMouse Me.hWnd
If App.LogMode <> 0 Then
    UnHookMouse Me.hWnd
End If
End
End Sub

Private Sub GlobalView_KeyDown(KeyCode As Integer, Shift As Integer)
Form_KeyDown KeyCode, Shift
End Sub

Private Sub GlobalView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim dx As Single: Dim dy As Single
If Button = 1 Or Button = 2 Then
    dx = Note.Width / 2 - x
    dy = Note.Height / 2 - y
    MapUpdata_AoVMove_Moving dx, dy
    mouseMapPos.x = Note.Width / 2
    mouseMapPos.y = Note.Height / 2
    mapMoveLock = True
End If
End Sub

Private Sub GlobalView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If mapMoveLock = True And mapGetMousePosLock = False Then mouseMapPos.x = x: mouseMapPos.y = y
End Sub

Private Sub GlobalView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
mapMoveLock = False
End Sub

Private Sub MainTime_Timer()
Updata

End Sub

Private Sub MapUpdataTimer_Timer()
If nSum > 0 And ȫ����ͼ.Checked = True Then MapUpdata
End Sub

Private Sub PLC_Timer()
Select Case noteSaveCheck
    Case 0
        PilotLight.BackColor = RGB(255, 0, 0)
    Case 1
        PilotLight.BackColor = RGB(255, 165, 0)
    Case 2
        PilotLight.BackColor = RGB(0, 255, 0)
End Select
On Error GoTo Er
If Clipboard.GetText = "" Then ճ��.Enabled = False Else ճ��.Enabled = True
If bHLSum < 1 Then ����.Enabled = False Else ����.Enabled = True
If redoSum < 1 Then ����.Enabled = False Else ����.Enabled = True
�˵��������
If saveNtxTime <> 0 Then
    saveNtxTimeNow = saveNtxTimeNow + 0.5
    If saveNtxTimeNow > saveNtxTime And PilotLight.BackColor = RGB(255, 165, 0) Then
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


Private Sub ����ʼ�_Click()
Dim filePath As String
' ���á�CancelError��Ϊ True
If Dir(ntxPath) = "" Then
    Note.CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' ���ñ�־
'    Note.CommonDialog1.Flags = cdlOFNHideReadOnly
    Note.CommonDialog1.Flags = cdlOFNOverwritePrompt
    ' ���ù�����
    Note.CommonDialog1.Filter = "�ڵ�ʼ� (*.ntx)|*.ntx|All Files (*.*)|*.*"
    ' ָ��ȱʡ�Ĺ�����
    Note.CommonDialog1.FilterIndex = 1
    ' ��ʾ���򿪡��Ի���
    Note.CommonDialog1.ShowSave
    ' ��ʾѡ���ļ�������
    filePath = Note.CommonDialog1.FileName
    NoteFileWrite_201 filePath
Else
    NoteFileWrite_201 ntxPath
End If
Exit Sub
ErrHandler:
' �û����ˡ�ȡ������ť
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

Private Sub ��ǩ��_Click()
If ��ǩ��.Checked = False Then
    ��ǩ��.Checked = True
    Unload NodePrint
Else
    ��ǩ��.Checked = False
    nodePrintBeLock = False
End If
End Sub

Private Sub �ʺ�Ȧ_Click()
If �ʺ�Ȧ.Checked = False Then �ʺ�Ȧ.Checked = True Else �ʺ�Ȧ.Checked = False
End Sub

Private Sub �ʺ���_Click()
If �ʺ���.Checked = False Then �ʺ���.Checked = True Else �ʺ���.Checked = False: �������.Checked = False
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
' ���á�CancelError��Ϊ True
newAddNote
Note.CommonDialog1.CancelError = True
On Error GoTo ErrHandler
' ���ñ�־
Note.CommonDialog1.Flags = cdlOFNHideReadOnly
' ���ù�����
Note.CommonDialog1.Filter = "�ڵ�ʼ� (*.ntx)|*.ntx|All Files (*.*)|*.*"
' ָ��ȱʡ�Ĺ�����
Note.CommonDialog1.FilterIndex = 1
' ��ʾ���򿪡��Ի���
Note.CommonDialog1.ShowOpen
' ��ʾѡ���ļ�������
filePath = Note.CommonDialog1.FileName
NoteFileRead filePath
Exit Sub
ErrHandler:
' �û����ˡ�ȡ������ť
End Sub

Private Sub ����_Click()
CopyObject False
End Sub

Private Sub ���ڽڵ�ʼ�_Click()
AboutNote.Show
End Sub

Private Sub ����_Click()
BehaviorIdSet
CopyObject True
End Sub

Private Sub �ڵ�����TXT�ı�_Click()
Dim filePath As String
Note.CommonDialog1.CancelError = True
On Error GoTo ErrHandler
' ���ñ�־
'    Note.CommonDialog1.Flags = cdlOFNHideReadOnly
Note.CommonDialog1.Flags = cdlOFNOverwritePrompt
' ���ù�����
Note.CommonDialog1.Filter = "�ı��ĵ� (*.txt)|*.txt"
' ָ��ȱʡ�Ĺ�����
Note.CommonDialog1.FilterIndex = 1
' ��ʾ���򿪡��Ի���
Note.CommonDialog1.ShowSave
' ��ʾѡ���ļ�������
filePath = Note.CommonDialog1.FileName
NodesToTxt filePath
Exit Sub
ErrHandler:
' �û����ˡ�ȡ������ť
End Sub

Private Sub ����̨_Click()
NoteControlDesk.Show
End Sub

Private Sub ���ӷ�ת_Click()
ConnectionReversal
End Sub

Private Sub ���Ϊ_Click()
Dim filePath As String
Note.CommonDialog1.CancelError = True
On Error GoTo ErrHandler
' ���ñ�־
'    Note.CommonDialog1.Flags = cdlOFNHideReadOnly
Note.CommonDialog1.Flags = cdlOFNOverwritePrompt
' ���ù�����
Note.CommonDialog1.Filter = "�ڵ�ʼ� (*.ntx)|*.ntx|All Files (*.*)|*.*"
' ָ��ȱʡ�Ĺ�����
Note.CommonDialog1.FilterIndex = 1
' ��ʾ���򿪡��Ի���
Note.CommonDialog1.ShowSave
' ��ʾѡ���ļ�������
filePath = Note.CommonDialog1.FileName
NoteFileWrite_201 filePath
Exit Sub
ErrHandler:
' �û����ˡ�ȡ������ť
End Sub

Private Sub �������_Click()
If �������.Checked = True Then
    �������.Checked = False
ElseIf �ʺ���.Checked = True Then
    �������.Checked = True
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

Private Sub ʼ����ʾѡ����_Click()
If ʼ����ʾѡ����.Checked = True Then ʼ����ʾѡ����.Checked = False Else ʼ����ʾѡ����.Checked = True
End Sub

Private Sub ʼ����ʾѡ��_Click()
If ʼ����ʾѡ��.Checked = True Then ʼ����ʾѡ��.Checked = False Else ʼ����ʾѡ��.Checked = True
End Sub

Private Sub �˳�_Click()
End
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

Private Sub �½��ʼ�_Click()
newAddNote
End Sub

Private Sub ѡ��_Click()
SelectDisplayObjcet
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

Private Sub �Զ�������_Click()
On Error GoTo Er
saveNtxTime = Val(InputBox("�������Զ�����ʱ����(��λ:��,����0�����Զ����棡)", "�ڵ�ʼ��Զ�����ʱ��������"))
Exit Sub
Er:
MsgBox "����ʧ�ܣ�", , "���棡"
saveNtxTime = 0
End Sub

Private Sub ����_Click()
On Error GoTo Er
With CommonDialog1
    .Flags = 1
    .FontName = Me.Font.Name
    .FontBold = Me.Font.Bold
    .FontSize = MainFormFontSize
    .FontItalic = Me.Font.Italic
    .FontUnderline = Me.Font.Underline
    .FontStrikethru = Me.Font.Strikethrough
    .ShowFont
    Me.Font.Name = .FontName '��������
    Me.Font.Bold = .FontBold  '�Ӵ֣�
    MainFormFontSize = .FontSize '�����С
    Me.Font.Italic = .FontItalic '��б��
    Me.Font.Underline = .FontUnderline '�»��ߣ�
    Me.Font.Strikethrough = .FontStrikethru 'ɾ����
End With
Er:
End Sub
