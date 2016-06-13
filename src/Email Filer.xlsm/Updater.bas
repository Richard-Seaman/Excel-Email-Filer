Attribute VB_Name = "Updater"
Public myWorkbook As Workbook
Public updaterWorkbook As Workbook


Sub checkForUpdate()

    ' is it already the latest?
    If Sheets("INFO").Cells(6, 3) = Sheets("INFO").Cells(7, 3) Then
        MsgBox ("This is already the latest version, no need to update")
        End
    End If
        
    
    Set myWorkbook = Excel.ActiveWorkbook
    
    myworkbookname = Excel.ActiveWorkbook.Name
    myworkbookpath = Excel.ActiveWorkbook.FullName
    
    If IsFile(myworkbookpath) Then
    
        MsgBox ("This excel sheet will close and another will open, say yes if asked to save changes.")
    
        updaterPath = Sheets("INFO").Cells(9, 3)
                
        Set updaterWorkbook = Workbooks.Open(updaterPath)
        
        updaterWorkbook.Sheets(1).Cells(15, 3) = myworkbookpath
        
        myWorkbook.Close
        
    Else
    
        MsgBox ("Could not find updater." & vbNewLine & "Make sure the path on the INFO tab is correct and that the updater exists!")
    
    End If
    

    Set myWorkbook = Nothing
    Set updaterWorkbook = Nothing
    

End Sub
