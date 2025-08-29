Attribute VB_Name = "getColumnNumberByName_fn"
Function getColumnNumberByName(wsName As String, columnName As String, Optional rowNum As Long) As Integer

    'function to get the column index, requires worksheet name and column name

    Dim ws As Worksheet
    Dim lastCol As Integer
    
    Set ws = ThisWorkbook.Worksheets(wsName)
    
    If rowNum = 0 Then rowNum = 1
    
    lastCol = ws.Cells(rowNum, Columns.Count).End(xlToLeft).Column
    
    For i = lastCol To 1 Step -1
        
        If LCase(ws.Cells(rowNum, i).Value) = LCase(columnName) Then
        
            getColumnNumberByName = i
            Exit Function
        End If
        
    Next i
    

End Function
