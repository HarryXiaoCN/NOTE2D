VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form NoteControlDesk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Desk"
   ClientHeight    =   4320
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9015
   Icon            =   "Note_ControlDesk.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   9015
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer CDMouseUpdataTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   3960
   End
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
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Note_ControlDesk.frx":700A
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
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Note_ControlDesk.frx":70A7
   End
   Begin VB.Menu 运行 
      Caption         =   "运行"
      Begin VB.Menu 启动 
         Caption         =   "启动"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu 清空 
      Caption         =   "清空"
      Begin VB.Menu 全部清空 
         Caption         =   "全部清空"
         Shortcut        =   {F1}
      End
      Begin VB.Menu 空行清空 
         Caption         =   "空行清空"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "NoteControlDesk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CDMouseUpdataTimer_Timer()
    Me.Caption = 控制台名字 & " - MousePos: " & Format(mousePos.x, "0.00") & "," & Format(mousePos.y, "0.00")
End Sub

Private Sub CMDInBox_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27, 192
            Unload Me
        Case 13
            If Shift = 0 Then
                启动_Click
            End If
        Case 33
            cursorPos = cursorPos - 1
            If cursorPos < 0 Then cursorPos = 0
            CMDInBox.Text = inputRecord(cursorPos)
        Case 34
            cursorPos = cursorPos + 1
            If cursorPos > UBound(inputRecord) Then cursorPos = UBound(inputRecord)
            CMDInBox.Text = inputRecord(cursorPos)
    End Select
End Sub

Private Sub CMDInBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub 空行清空_Click()
    Dim t() As String, i As Long
    t = Split(CMDOutBox.Text, vbCrLf)
    CMDOutBox.Text = ""
    For i = 0 To UBound(t)
        If t(i) <> "" Then
            CMDOutBox.Text = CMDOutBox.Text & t(i) & vbCrLf
        End If
    Next
End Sub

Private Sub 启动_Click()
    Dim cmd As String
    Dim fOut As String
On Error GoTo Er:
    BehaviorIdSet
    ReDim Preserve inputRecord(UBound(inputRecord) + 1)
    cursorPos = UBound(inputRecord)
    inputRecord(cursorPos) = CMDInBox.Text
    cmd = CMDInBox.Text
    fOut = CMD_In(cmd)
    CMDOutBox.Text = CMDOutBox.Text & CMDInBox.Text & vbCrLf & fOut & vbCrLf
    CMDInBox.Text = ""
    CMDOutBox.SelStart = Len(CMDOutBox.TextRTF)
    Exit Sub
Er:
    CMDOutBox.Text = CMDOutBox.Text & CMDInBox.Text & vbCrLf & " 运行失败！原因：" & Err.Description & vbCrLf & "返回值：" & fOut & vbCrLf
End Sub

Private Sub 全部清空_Click()
    CMDOutBox.Text = ""
End Sub
