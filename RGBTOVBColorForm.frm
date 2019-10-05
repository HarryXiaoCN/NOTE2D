VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RGBTOVBColorForm 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "RGB To VBColor"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin MSComDlg.CommonDialog ColorCDL 
      Left            =   120
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox RGBText 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   3
      Left            =   720
      TabIndex        =   7
      Top             =   1920
      Width           =   3855
   End
   Begin VB.TextBox RGBText 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   2
      Left            =   720
      TabIndex        =   6
      Top             =   1440
      Width           =   3855
   End
   Begin VB.TextBox RGBText 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   720
      TabIndex        =   5
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox RGBText 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   720
      TabIndex        =   4
      Top             =   480
      Width           =   3855
   End
   Begin VB.PictureBox TitleBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   4650
      TabIndex        =   1
      Top             =   0
      Width           =   4680
      Begin VB.Label TitleLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "RGB To VBColor"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   45
         Width           =   1485
      End
      Begin VB.Label TitleLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "¡Á"
         BeginProperty Font 
            Name            =   "ºÚÌå"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Index           =   0
         Left            =   4200
         TabIndex        =   2
         Top             =   0
         Width           =   405
      End
   End
   Begin VB.Label ColorVis 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "µã»÷´ò¿ªÑÕÉ«Ñ¡Ôñ¶Ô»°¿ò"
      Top             =   2400
      Width           =   4455
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "V B£º"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   1965
      Width           =   570
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BÖµ£º"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   1480
      Width           =   570
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "GÖµ£º"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1005
      Width           =   585
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "RÖµ£º"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   520
      Width           =   570
   End
   Begin VB.Shape Shape1 
      Height          =   2710
      Left            =   0
      Top             =   360
      Width           =   4680
   End
End
Attribute VB_Name = "RGBTOVBColorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private titleMousePosX As Single, titleMousePosY As Single, singleColor As Boolean, manyColor As Boolean

Private Sub ColorVis_Click()
    With ColorCDL
        .Flags = 1
        .ShowColor
        RGBText(3).Text = .color
    End With
    Me.SetFocus
End Sub

Private Sub RGBText_Change(Index As Integer)
    If Index < 3 And manyColor = False Then
        singleColor = True
        If Val(RGBText(Index).Text) < 0 Then RGBText(Index).Text = 0
        If Val(RGBText(Index).Text) > 255 Then RGBText(Index).Text = 255
        RGBText(3).Text = RGB(Val(RGBText(0)), Val(RGBText(1)), Val(RGBText(2)))
        Label(3).ForeColor = Val(RGBText(3).Text)
        ColorVis.BackColor = Label(3).ForeColor
        singleColor = False
    ElseIf Index = 3 And singleColor = False Then
        Dim rC As Long, gC As Long, rGBC As Long
        manyColor = True
        If Len(RGBText(3).Text) > 8 Then RGBText(3).Text = 16777215
        rGBC = Val(RGBText(3).Text)
        If rGBC < 0 Then rGBC = 0: RGBText(3).Text = 0
        If rGBC > 16777216 Then rGBC = 16777215: RGBText(3).Text = 16777215
        ColorVis.BackColor = rGBC
        Label(3).ForeColor = ColorVis.BackColor
        rC = ColorVis.BackColor Mod 256
        RGBText(0).Text = rC
        gC = ((ColorVis.BackColor - rC) Mod 65536) / 256
        RGBText(1).Text = gC
        RGBText(2).Text = (ColorVis.BackColor - rC - gC * 256) / 65536
        manyColor = False
    End If
End Sub

Private Sub TitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    titleMousePosX = X
    titleMousePosY = Y
End Sub

Private Sub TitleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And titleMousePosX <> 0 And titleMousePosY <> 0 Then
        Me.left = Me.left + X - titleMousePosX
        Me.Top = Me.Top + Y - titleMousePosY
    End If
End Sub

Private Sub TitleLabel_Click(Index As Integer)
    If Index = 0 Then Unload Me
End Sub

Private Sub TitleLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    titleMousePosX = X
    titleMousePosY = Y
End Sub

Private Sub TitleLabel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And titleMousePosX <> 0 And titleMousePosY <> 0 Then
        Me.left = Me.left + X - titleMousePosX
        Me.Top = Me.Top + Y - titleMousePosY
    End If
End Sub
