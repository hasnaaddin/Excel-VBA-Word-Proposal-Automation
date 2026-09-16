Attribute VB_Name = "Module1"
Option Explicit

Sub GenerateProposal()

    Dim wdApp As Object
    Dim wdDoc As Object
    Dim TemplatePath As String

    'Word template is stored in the same folder as this Excel workbook
    TemplatePath = ThisWorkbook.Path & "\Septic_Proposal_Template.docx"

    'Check that the Word template exists
    If Dir(TemplatePath) = "" Then
        MsgBox "Word template could not be found:" & vbCrLf & _
               TemplatePath, vbExclamation, "Template Not Found"
        Exit Sub
    End If

    'Open Microsoft Word
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True

    'Open the proposal template
    Set wdDoc = wdApp.Documents.Open(TemplatePath)

    MsgBox "Proposal template opened successfully.", _
           vbInformation, "Proposal Generator"

End Sub

