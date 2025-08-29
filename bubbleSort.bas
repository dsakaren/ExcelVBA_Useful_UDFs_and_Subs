Attribute VB_Name = "bubbleSort"
Function bubbleSort(arr)
    
    'receives an array and sorts it ascending
    
    Dim isSorted As Boolean
    Dim swapCounter As Integer
    Dim i As Integer, temp As Integer
    
    isSorted = False
    swapCounter = 0
    
    Do While isSorted = False
        isSorted = True
        
        For i = LBound(arr) + 1 To UBound(arr) - swapCounter
            If arr(i) < arr(i - 1) Then
                temp = arr(i)
                arr(i) = arr(i - 1)
                arr(i - 1) = temp
                isSorted = False
            End If
            
        Next i
    
    Loop
      
    bubbleSort = arr
    
    
End Function
Sub testBubbleSorting()

    'tests the sorting algorithm
    'populates random numbers between 1 and 100 in range A1:A7
    'sorts them in range B1:B7

    Dim nums(7) As Integer
    Dim sortedArray As Variant
    
    
    For i = 0 To 7
        
        nums(i) = Application.WorksheetFunction.RandBetween(1, 100)
        Range("A" & i + 1).Value = nums(i)
        
    Next
    
    sortedArray = bubbleSort(nums)
    
    For i = 0 To 7
        
        Range("B" & i + 1).Value = sortedArray(i)
        
    Next
    
End Sub




