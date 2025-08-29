Attribute VB_Name = "gettingLastRowAndCols"
Function getColumnNumber(wsName As String, columnName As String) As Integer

    'function to get the column index, requires worksheet name and column name
    'assumes headers are always on first row

    Dim ws As Worksheet
    Dim lastCol As Integer
    
    Set ws = ThisWorkbook.Worksheets(wsName)
    
    lastCol = getLastColumn(wsName)
    
    For i = lastCol To 1 Step -1
        
        If LCase(ws.Cells(1, i).Value) = LCase(columnName) Then
        
            getColumnNumber = i
            Exit Function
        End If
        
    Next i
    

End Function

Function getLastColumn(wsName As String)

    'function to get the last column as number of a worksheet
    'assumes headers are always on first row
    
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets(wsName)
    
    getLastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column


End Function

Function getLastRow(wsName As String, columnName As String) As Long

    'returns the last populated row of a specific column in a specific worksheet

    Dim ws As Worksheet
    Dim targetCol As Integer
    
    Set ws = ThisWorkbook.Worksheets(wsName)
    
    targetCol = getColumnNumber(wsName, columnName)
    
    getLastRow = ws.Cells(Rows.Count, targetCol).End(xlUp).Row
    
End Function
