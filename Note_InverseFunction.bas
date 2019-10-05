Attribute VB_Name = "Note_InverseFunction"
Public Function RedoBehavior()
Dim i As Long: Dim currentBehavior() As String
Dim redoBehaviorId As String
currentBehavior = Split(redolist(redoSum - 1), ",")
redoBehaviorId = currentBehavior(0)
For i = redoSum - 1 To 0 Step -1
    currentBehavior = Split(redolist(i), ",")
    If redoBehaviorId = currentBehavior(0) Then
        Select Case currentBehavior(1)
            Case "Revoke_LineAdd_Save"
                Redo_LineAdd_Save Val(currentBehavior(2))
            Case "Revoke_LineDelete"
                Redo_LineDelete Val(currentBehavior(2))
            Case "Revoke_NodeEdit_NewNode"
                Redo_NodeEdit_NewNode Val(currentBehavior(2))
            Case "Revoke_NodeEdit_ReviseNode"
                Redo_NodeEdit_ReviseNode Val(currentBehavior(2)), currentBehavior(3), currentBehavior(4), currentBehavior(5), Val(currentBehavior(6))
            Case "Revoke_NodeDelete"
                Redo_NodeDelete Val(currentBehavior(2))
            Case "Revoke_LineReplace"
                Redo_LineReplace Val(currentBehavior(2)), Val(currentBehavior(3)), Val(currentBehavior(4)), Val(currentBehavior(5))
        End Select
        redoSum = i
    Else
        redoSum = i + 1
        Exit For
    End If
    bHLSum = bHLSum + 1
Next
End Function
Public Function Redo_LineReplace(ByVal lid As Long, ByVal f As Long, ByVal newN As Long, ByVal oldN As Long)
    If f = 0 Then
        nodeLine(lid).source = newN
    Else
        nodeLine(lid).target = newN
    End If
End Function
Public Function Redo_LineAdd_Save(ByVal lid As Long)
nodeLine(lid).b = True
End Function
Public Function Redo_LineDelete(ByVal lid As Long)
nodeLine(lid).b = False
End Function
Public Function Redo_NodeEdit_NewNode(ByVal nId As Long)
node(nId).b = True
End Function
Public Function Redo_NodeEdit_ReviseNode(ByVal nId As Long, ByVal t As String, ByVal content As String, ByVal setC As String, ByVal setS As Single)
With node(nId)
    .t = t
    .content = content
    .setColor = setC
    .setSize = setS
    If nId = nodeEditAim Then
        NodeInput.NodeTitle.Text = .t
        NodeInput.NodeInputBox.TextRTF = .content
    End If
End With
End Function
Public Function Redo_NodeDelete(ByVal nId As Long)
node(nId).b = False
End Function

Public Function RevokeBehavior()
Dim i As Long: Dim currentBehavior
Dim RevokeBehaviorId As String
currentBehavior = Split(behaviorList(bHLSum - 1), ",")
RevokeBehaviorId = currentBehavior(0)
For i = bHLSum - 1 To 0 Step -1
    currentBehavior = Split(behaviorList(i), ",")
    If RevokeBehaviorId = currentBehavior(0) Then
        Select Case currentBehavior(1)
            Case "LineAdd_Save"
                Revoke_LineAdd_Save Val(currentBehavior(2))
            Case "LineDelete"
                Revoke_LineDelete Val(currentBehavior(2))
            Case "NodeEdit_NewNode"
                Revoke_NodeEdit_NewNode Val(currentBehavior(2))
            Case "NodeEdit_ReviseNode"
                Revoke_NodeEdit_ReviseNode Val(currentBehavior(2)), currentBehavior(3), currentBehavior(4), Val(currentBehavior(5)), Val(currentBehavior(6))
            Case "NodeDelete"
                Revoke_NodeDelete Val(currentBehavior(2))
            Case "LineReplace"
                Revoke_LineReplace Val(currentBehavior(2)), currentBehavior(3), currentBehavior(4), currentBehavior(5)
        End Select
        bHLSum = i
    Else
        bHLSum = i + 1
        Exit For
    End If
Next
End Function
Public Function Revoke_LineReplace(ByVal lid As Long, ByVal f As Long, ByVal newN As Long, ByVal oldN As Long)
    RedoListAdd "Revoke_LineReplace", lid, f, newN, oldN
    If f = 0 Then
        nodeLine(lid).source = oldN
    Else
        nodeLine(lid).target = oldN
    End If
End Function
Public Function Revoke_LineAdd_Save(ByVal lid As Long)
RedoListAdd "Revoke_LineAdd_Save", lid
nodeLine(lid).b = False
End Function
Public Function Revoke_LineDelete(ByVal lid As Long)
RedoListAdd "Revoke_LineDelete", lid
nodeLine(lid).b = True
End Function
Public Function Revoke_NodeEdit_NewNode(ByVal nId As Long)
RedoListAdd "Revoke_NodeEdit_NewNode", nId
node(nId).b = False
End Function
Public Function Revoke_NodeEdit_ReviseNode(ByVal nId As Long, ByVal t As String, ByVal content As String, setC As Long, setS As Single)
With node(nId)
    RedoListAdd "Revoke_NodeEdit_ReviseNode", nId, .t, .content, .setColor, .setSize
    .t = t
    .content = content
    .setColor = setC
    .setSize = setS
    If nId = nodeEditAim Then
        NodeInput.NodeTitle.Text = .t
        NodeInput.NodeInputBox.TextRTF = .content
    End If
End With
End Function
Public Function Revoke_NodeDelete(ByVal nId As Long)
    RedoListAdd "Revoke_NodeDelete", nId
    node(nId).b = True
End Function
