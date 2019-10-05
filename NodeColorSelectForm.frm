VERSION 5.00
Begin VB.Form NodeColorSelectForm 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选择子节点颜色"
   ClientHeight    =   435
   ClientLeft      =   5670
   ClientTop       =   8880
   ClientWidth     =   8085
   Icon            =   "NodeColorSelectForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   435
   ScaleWidth      =   8085
   Begin VB.TextBox 深度 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1200
      TabIndex        =   26
      Text            =   "1"
      ToolTipText     =   "范围[1,100]"
      Top             =   80
      Width           =   615
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   0
      Left            =   2040
      TabIndex        =   24
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   1
      Left            =   2280
      TabIndex        =   23
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   2
      Left            =   2520
      TabIndex        =   22
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   3
      Left            =   2760
      TabIndex        =   21
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   4
      Left            =   3000
      TabIndex        =   20
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   5
      Left            =   3240
      TabIndex        =   19
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   6
      Left            =   3480
      TabIndex        =   18
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   7
      Left            =   3720
      TabIndex        =   17
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   8
      Left            =   3960
      TabIndex        =   16
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   9
      Left            =   4200
      TabIndex        =   15
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   10
      Left            =   4440
      TabIndex        =   14
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   11
      Left            =   4680
      TabIndex        =   13
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   12
      Left            =   4920
      TabIndex        =   12
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   13
      Left            =   5160
      TabIndex        =   11
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   14
      Left            =   5400
      TabIndex        =   10
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   15
      Left            =   5640
      TabIndex        =   9
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   16
      Left            =   5880
      TabIndex        =   8
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFBF00&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   17
      Left            =   6120
      TabIndex        =   7
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   18
      Left            =   6360
      TabIndex        =   6
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   19
      Left            =   6600
      TabIndex        =   5
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   20
      Left            =   6840
      TabIndex        =   4
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   21
      Left            =   7080
      TabIndex        =   3
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   22
      Left            =   7320
      TabIndex        =   2
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   23
      Left            =   7560
      TabIndex        =   1
      Top             =   120
      Width           =   135
   End
   Begin VB.Label 节点颜色预览 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   24
      Left            =   7800
      TabIndex        =   0
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "子节点深度："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "NodeColorSelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public 锁定母节点序号 As Long
Private 锁定深度 As Long

Private Sub 节点颜色预览_Click(Index As Integer)
    
    锁定深度 = Val(深度.Text)
    锁定深度 = 限制数值(锁定深度, 1, 100)
    深度节点颜色设置 节点颜色预览(Index).BackColor, 1, 锁定深度, 锁定母节点序号
'    Unload Me
End Sub

Private Sub 深度节点颜色设置(c As Long, ByVal d As Long, maxD As Long, ByVal nId As Long)
    Dim i As Long, dT As Long
    dT = d
    For i = 0 To lSum
        With nodeLine(i)
            If .b Then
                d = dT
                If .source = nId Then
                    If d >= maxD Then
                        node(.target).setColor = c
                    Else
                        d = d + 1
                        深度节点颜色设置 c, d, maxD, .target
                    End If
                End If
            End If
        End With
    Next
End Sub
