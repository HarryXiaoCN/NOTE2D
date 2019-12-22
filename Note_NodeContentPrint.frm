VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form NodePrint 
   AutoRedraw      =   -1  'True
   Caption         =   "Node Content"
   ClientHeight    =   7995
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   6000
   Icon            =   "Note_NodeContentPrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   6000
   StartUpPosition =   3  '窗口缺省
   Begin RichTextLib.RichTextBox NodePrintBox 
      Height          =   8000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   14102
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Note_NodeContentPrint.frx":700A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "NodePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
nodePrintBeLock = True
If Note.全高透明3.Checked = True Then
    FormTransparent Me, 50
ElseIf Note.全半透明3.Checked = True Then
    FormTransparent Me, 125
ElseIf Note.全低透明3.Checked = True Then
    FormTransparent Me, 200
End If
End Sub

Private Sub Form_Resize()
If WindowState = 1 Then Exit Sub
If Me.height < 4000 Then Me.Enabled = False: Me.height = 4000: Me.Enabled = True
If Me.width < 6240 Then Me.Enabled = False: Me.width = 6240: Me.Enabled = True
NodePrintBox.height = Me.height
NodePrintBox.width = Me.width
End Sub

Private Sub Form_Unload(Cancel As Integer)
nodePrintBeLock = False
End Sub
