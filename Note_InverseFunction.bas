Attribute VB_Name = "Note_InverseFunction"
Public Function RedoBehavior()
Dim i As Long: Dim currentBehavior
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
                Redo_NodeEdit_ReviseNode Val(currentBehavior(2)), currentBehavior(3), currentBehavior(4)
            Case "Revoke_NodeDelete"
                Redo_NodeDelete Val(currentBehavior(2))
        End Select
        redoSum = i
    Else
        redoSum = i + 1
        Exit For
    End If
    bHLSum = bHLSum + 1
Next
End Function

Public Function Redo_LineAdd_Save(ByVal lid As Long)
nodeLine(lid).b = True
End Function
Public Function Redo_LineDelete(ByVal lid As Long)
nodeLine(lid).b = False
End Function
Public Function Redo_NodeEdit_NewNode(ByVal nid As Long)
node(nid).b = True
End Function
Public Function Redo_NodeEdit_ReviseNode(ByVal nid As Long, ByVal t As String, ByVal content As String)
With node(nid)
    .t = t
    .content = content
    If nid = nodeEditAim Then
        NodeInput.NodeTitle.Text = .t
        NodeInput.NodeInputBox.TextRTF = .content
    End If
End With
End Function
Public Function Redo_NodeDelete(ByVal nid As Long)
node(nid).b = False
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
                Revoke_NodeEdit_ReviseNode Val(currentBehavior(2)), currentBehavior(3), currentBehavior(4)
            Case "NodeDelete"
                Revoke_NodeDelete Val(currentBehavior(2))
        End Select
        bHLSum = i
    Else
        bHLSum = i + 1
        Exit For
    End If
Next
End Function

Public Function Revoke_LineAdd_Save(ByVal lid As Long)
RedoListAdd "Revoke_LineAdd_Save", lid
nodeLine(lid).b = False
End Function
Public Function Revoke_LineDelete(ByVal lid As Long)
RedoListAdd "Revoke_LineDelete", lid
nodeLine(lid).b = True
End Function
Public Function Revoke_NodeEdit_NewNode(ByVal nid As Long)
RedoListAdd "Revoke_NodeEdit_NewNode", nid
node(nid).b = False
End Function
Public Function Revoke_NodeEdit_ReviseNode(ByVal nid As Long, ByVal t As String, ByVal content As String)
With node(nid)
    RedoListAdd "Revoke_NodeEdit_ReviseNode", nid, .t, .content
    .t = t
    .content = content
    If nid = nodeEditAim Then
        NodeInput.NodeTitle.Text = .t
        NodeInput.NodeInputBox.TextRTF = .content
    End If
End With
End Function
Public Function Revoke_NodeDelete(ByVal nid As Long)
RedoListAdd "Revoke_NodeDelete", nid
node(nid).b = True
End Function
