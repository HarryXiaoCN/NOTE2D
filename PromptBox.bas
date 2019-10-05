Attribute VB_Name = "PromptBox"
Public promptBoxSelect As Integer
Public promptBoxText As String
Public promptBoxDefaultText As String, promptBoxPromptText As String

Public Function InBox(ByVal prompt As String, Optional ByVal defaultText As String) As String
    promptBoxDefaultText = defaultText
    promptBoxPromptText = prompt
    PromptForm.Show 1
    InBox = promptBoxText
End Function
