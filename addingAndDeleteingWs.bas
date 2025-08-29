Attribute VB_Name = "addingAndDeleteingWs"
Sub addWorksheetWithAName(wsName As String)

    'add worksheet with a specific name
    'if another worksheet with the same name in the same workbook exists - it will be deleted first
    
    deleteWorksheetWithAName (wsName)
    
    ThisWorkbook.Worksheets.Add after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    
    ThisWorkbook.ActiveSheet.Name = wsName
    
    
End Sub
Sub deleteWorksheetWithAName(wsName As String)

    'loop through the worksheets and delete the one with name wsName
    
    For Each ws In ThisWorkbook.Worksheets
        If LCase(ws.Name) = LCase(wsName) Then
            Application.DisplayAlerts = False
                ws.Delete
            Application.DisplayAlerts = True
            Exit Sub
        End If
    Next ws

End Sub
