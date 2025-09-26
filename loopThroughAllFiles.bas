Attribute VB_Name = "loopThroughAllFiles"
Sub loopThroughFiles()
    
    Application.ScreenUpdating = False
    
    Dim folderPath As String
    Dim fileName As String
    
    folderPath = "C:\Users\coolk\Documents\VBA\UFC Data\"
    
    fileName = Dir(folderPath & "results_*.xlsx")
    
    Do While fileName <> ""
        
        Set wb = Workbooks.Open(folderPath & fileName)
        
        '''do stuff
        
        wb.Close
        
        fileName = Dir
    Loop
    
    Application.ScreenUpdating = True
    
End Sub
