Attribute VB_Name = "PromptBox"
Public promptBoxSelect As Integer
Public promptBoxText As String
Public promptBoxDefaultText As String, promptBoxPromptText As String

Public Function InBox(ByVal prompt As String, Optional ByVal defaultText As String) As String
    promptBoxDefaultText = defaultText
    promptBoxPromptText = prompt
    InjectionLoaded
    PromptForm.Show 1
    InBox = promptBoxText
End Function

Private Sub InjectionLoaded()
    On Error GoTo Er
        PromptForm.������.Text = promptBoxDefaultText
        PromptForm.��ʾ�ı�.Caption = promptBoxPromptText
        PromptForm.������.SelLength = Len(������.Text)
Er:
End Sub
