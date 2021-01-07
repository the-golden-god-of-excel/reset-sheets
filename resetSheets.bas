Sub resetSheets()
    Dim currentWorksheet As Worksheet
    For Each currentWorksheet In ActiveWorkbook.Worksheets
        currentWorksheet.Activate
        
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 1
        
        currentWorksheet.Range("A1").Select
    Next currentWorksheet
    
    If ActiveSheet.Index < ThisWorkbook.Sheets.Count Then
        Sheets(ActiveSheet.Index + 1).Activate
    Else
        Sheets(1).Activate
    End If
End Sub
