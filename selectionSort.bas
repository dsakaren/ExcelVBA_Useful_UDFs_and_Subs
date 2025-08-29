Attribute VB_Name = "selectionSort"

Function selectionSort(arr)
    
    Dim i As Integer, j As Integer
    Dim temp As Integer
    Dim minNumber As Integer, minIndex As Integer
    Dim nextNumber As Integer, nextIndex As Integer
    
    For i = LBound(arr) To UBound(arr)
    
        minNumber = arr(i)
        minIndex = i
        
        For j = LBound(arr) + 1 + i To UBound(arr)
            nextNumber = arr(j)
            If nextNumber < minNumber Then
                minNumber = nextNumber
                minIndex = j
                
            End If
            
        
        Next j
        
        temp = arr(i)
        arr(i) = arr(minIndex)
        arr(minIndex) = temp
        
    Next i
    
    selectionSort = arr
    
    
End Function

Sub testSelectionSort()

    'test the algorhitm
    'populates random number in column A and sorts them in ascending order in column B

    Dim nums(7) As Integer
    Dim sortedArray As Variant
    
    
    For i = 0 To 7
        
        nums(i) = Application.WorksheetFunction.RandBetween(1, 100)
        Range("A" & i + 1).Value = nums(i)
        
    Next
    
    sortedArray = selectionSort(nums)
    
    For i = 0 To 7
        
        Range("B" & i + 1).Value = sortedArray(i)
        
    Next
    
End Sub

