Private Sub Workbook_Open()
Range("E2").Value = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)
End Sub
