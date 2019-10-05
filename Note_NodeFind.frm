VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form NodeFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Node Find"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Note_NodeFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   6255
   StartUpPosition =   2  '��Ļ����
   Begin RichTextLib.RichTextBox nodeTmp 
      Height          =   495
      Left            =   5520
      TabIndex        =   0
      Top             =   -500
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      _Version        =   393217
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Note_NodeFind.frx":700A
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5880
      Top             =   720
   End
   Begin VB.TextBox FindText 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   120
      TabIndex        =   1
      Text            =   "�����ѯ���ݺ�س�����..."
      ToolTipText     =   "�����ѯ���ݺ�س�����..."
      Top             =   120
      Width           =   6000
   End
   Begin VB.Menu ѡ�� 
      Caption         =   "ѡ��"
      Begin VB.Menu ���� 
         Caption         =   "����"
         Shortcut        =   ^S
      End
      Begin VB.Menu ���ִ�Сд 
         Caption         =   "���ִ�Сд"
         Shortcut        =   ^B
      End
      Begin VB.Menu ѡ�з�Χ������ 
         Caption         =   "ѡ�з�Χ������"
         Shortcut        =   ^R
      End
      Begin VB.Menu Cut1 
         Caption         =   "-"
      End
      Begin VB.Menu ���ļ������ 
         Caption         =   "���ļ������"
         Shortcut        =   ^N
      End
   End
   Begin VB.Menu ��� 
      Caption         =   "���"
      Visible         =   0   'False
      Begin VB.Menu Բ������ 
         Caption         =   "Բ������"
         Checked         =   -1  'True
      End
      Begin VB.Menu ��Բ���� 
         Caption         =   "��Բ����"
      End
   End
End
Attribute VB_Name = "NodeFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FindText_GotFocus()
If FindText.Text = "�����ѯ���ݺ�س�����..." Then FindText.Text = ""
End Sub

Private Sub FindText_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 13
        ����_Click
    Case 27
        Unload Me
End Select
End Sub

Private Sub FindText_LostFocus()
If FindText.Text = "" Then FindText.Text = "�����ѯ���ݺ�س�����..."
End Sub

Private Sub Timer1_Timer()
FindText.SetFocus
Timer1.Enabled = False
End Sub

Private Sub ���ִ�Сд_Click()
If ���ִ�Сд.Checked = True Then ���ִ�Сд.Checked = False Else ���ִ�Сд.Checked = True
End Sub

Private Sub ����_Click()
    FindNode FindText.Text, ���ִ�Сд.Checked, ѡ�з�Χ������.Checked, ���ļ������.Checked
End Sub

Private Sub ��Բ����_Click()
If ��Բ����.Checked = True Then ��Բ����.Checked = False:  Բ������.Checked = True Else ��Բ����.Checked = True: Բ������.Checked = False
End Sub

Private Sub ���ļ������_Click()
If ���ļ������.Checked = True Then ���ļ������.Checked = False: ���.Visible = False Else ���ļ������.Checked = True: ���.Visible = True
End Sub

Private Sub ѡ�з�Χ������_Click()
If ѡ�з�Χ������.Checked = True Then ѡ�з�Χ������.Checked = False Else ѡ�з�Χ������.Checked = True
End Sub

Private Sub Բ������_Click()
If Բ������.Checked = True Then Բ������.Checked = False: ��Բ����.Checked = True Else Բ������.Checked = True: ��Բ����.Checked = False

End Sub
