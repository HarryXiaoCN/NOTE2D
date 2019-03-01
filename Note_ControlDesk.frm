VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form NoteControlDesk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Desk"
   ClientHeight    =   4320
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   9015
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   9015
   StartUpPosition =   3  '窗口缺省
   Begin RichTextLib.RichTextBox CMDOutBox 
      Height          =   2055
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   3625
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"Note_ControlDesk.frx":0000
   End
   Begin RichTextLib.RichTextBox CMDInBox 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   3625
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"Note_ControlDesk.frx":009D
   End
   Begin VB.Menu 运行 
      Caption         =   "运行"
      Begin VB.Menu 启动 
         Caption         =   "启动"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "NoteControlDesk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDInBox_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 192 Or KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub 启动_Click()
Dim cmd As String
Dim fOut
'On Error GoTo Er:
cmd = CMDInBox.Text
fOut = CMD_In(cmd)
CMDOutBox.Text = CMDOutBox.Text & CMDInBox.Text & vbCrLf & " 运行成功！:)   返回值:" & fOut & vbCrLf
Exit Sub
Er:
CMDOutBox.Text = CMDOutBox.Text & CMDInBox.Text & vbCrLf & " 运行失败！:(" & vbCrLf
End Sub
