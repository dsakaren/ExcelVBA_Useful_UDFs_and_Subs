Attribute VB_Name = "checkIfItemIsInArray"
Function isInArray(el As Variant, arr) As Boolean
    
    Dim i As Integer
    
    For i = LBound(arr) To UBound(arr) - 1
        If arr(i) = el Then
            isInArray = True
            Exit Function
        End If
    Next i
    
    isInArray = False
    
    
End Function
