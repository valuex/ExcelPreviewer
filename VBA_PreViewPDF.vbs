' worksheet function

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Static PreRowNum As Integer
If (Target.Count = 1) Then
    ThisRowNum = Target.Row
    ThisValue = Cells(ThisRowNum, 1).Value
    If (ThisRowNum <> PreRowNum) Then
        Call PreviewPDF(ThisValue)
    End If
    PreRowNum = ThisRowNum
End If
End Sub

