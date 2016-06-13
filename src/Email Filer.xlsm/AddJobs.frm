VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddJobs 
   Caption         =   "Email Filer"
   ClientHeight    =   11715
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7875
   OleObjectBlob   =   "AddJobs.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddJobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const divider As String = " - "
Private Const addJobText As String = "Add Job's Rules Worksheet"
Private Const updateJobtext As String = "Update Job's Rules Worksheet"

' Colours for button - refer to MSAccess colours at link below:
' http://www.endprod.com/colors/
Private Const blueColour As Double = 13464600
Private Const greenColour As Double = 3329330


Private Sub UserForm_Initialize()

    Dim ws As Worksheet
    
    ' disable the add button to begin with
    AddJobButton.Visible = False

    ' populate the listbox with the available job sheets from the index

    For Each ws In indexWB.Worksheets
    
        ' Exclude the info and make sure it's a job rule sheet (by checking for "P5000" in the name)
        If ws.Name <> "INFO" And InStr(1, LCase(ws.Name), LCase("P5000"), vbTextCompare) Then
                        
            ' Grab the project name from the worksheet
            projectName = ws.Cells(3, 2)
            
            ' The formatting of this is very important
            ' When adding a sheet from the indexWB to the filerWB, we'll need to just use the project number (sheetname) so we'll search for the divider and split the string
            listBoxItemName = ws.Name & divider & projectName
                
            ' Add the project name to the listbox
            JobListBox.AddItem (listBoxItemName)
            
            
        End If
        
    Next ws

End Sub


Private Sub JobListBox_Click()

    Dim selectedProjectNumber As String
    
    ' enable the add button
    AddJobButton.Visible = True
    
    ' Grab the project number from the listbox (number only, must trim the divider and the name)
    selectedProjectNumber = Left(JobListBox.Text, InStr(JobListBox.Text, divider) - 1)
    
    ' Check if the sheet exists (will either Add or Update)
    If SheetExists(selectedProjectNumber, filerWB) Then
        
        ' Job sheet already exists in filerWB, set to update
        AddJobButton.Caption = updateJobtext
        AddJobButton.BackColor = blueColour     ' Blue
        
    Else
        ' Job sheet doesn't exist in filerWB, set to add
        AddJobButton.Caption = addJobText
        AddJobButton.BackColor = greenColour    ' Green

    End If

End Sub

Private Sub AddJobButton_Click()

    Dim selectedProjectNumber As String
    Dim workSheetName As String
    Dim lastWorkSheetName As String
    
    ' Grab the project number from the listbox (number only, must trim the divider and the name)
    selectedProjectNumber = Left(JobListBox.Text, InStr(JobListBox.Text, divider) - 1)
    selectedIndex = JobListBox.ListIndex
    
    ' make sure the sheet still exists in the index (it should since it was added to the listbox)
    If SheetExists(selectedProjectNumber, indexWB) Then

        ' Check if it already exists in the filer wb
        ' Shouldn't be in the listbox if it's already in this workbook, but check anyway
        If Not SheetExists(selectedProjectNumber, filerWB) Then
        
            ' ADD THE PROJECT SHEET
        
            ' Figure out where to add the sheet
            ' Initally set to inserting after dashboard
            sheetNameToInsertAfter = "Dashboard"
            lastWorkSheetName = "Dashboard"
            For Each ws In filerWB.Worksheets
            
                ' Get the current worksheet name
                workSheetName = ws.Name
                
                ' ignore the dashboard and the INFO
                If workSheetName <> "Dashboard" And workSheetName <> "INFO" Then
                
                    ' check if the current WS name is greater than the selected
                    ' (comes after it in the alphabet)
                    If ReturnLargerProjectNumber(workSheetName, selectedProjectNumber) = workSheetName Then
                    
                        ' If so, insert the selectedProject sheet after the previous worksheet's name
                        sheetNameToInsertAfter = lastWorkSheetName
                        Exit For
                                        
                    ElseIf workSheetName = filerWB.Worksheets(filerWB.Worksheets.Count).Name Then
                    
                        ' if it's the last worksheet, insert it after it
                        sheetNameToInsertAfter = workSheetName
                        Exit For
                        
                    End If
                
                End If
                
                ' Remember the name for next time
                lastWorkSheetName = workSheetName
            
            Next ws

            ' Add the job's rules WS from the index to the filerWB
            indexWB.Sheets(selectedProjectNumber).Copy After:=filerWB.Sheets(sheetNameToInsertAfter)
            
            ' Remove it from the list
            ' JobListBox.RemoveItem (selectedIndex)
            
            ' Hide the button if none left
            If JobListBox.ListCount = 0 Then
                AddJobButton.Visible = False
            End If

        Else
        
            ' UPDATE THE PROJECT SHEET
            
            Dim filerJobWS As Worksheet
            Dim indexJobWS As Worksheet
            
            ' make sure it still exists in the index
            If SheetExists(selectedProjectNumber, indexWB) Then
                
                ' grab the two sheets from each of the workbooks
                Set filerJobWS = filerWB.Sheets(selectedProjectNumber)
                Set indexJobWS = indexWB.Sheets(selectedProjectNumber)
                
                ' Cycle through each of the rules in the indexWB and make sure they're present in the filerWB
                
                For indexRow = ruleStartRow To 1000
                
                    ' Grab the data
                    indexSubjectMustContain = indexJobWS.Cells(indexRow, subjectColumn)
                    indexBodyMustContain = indexJobWS.Cells(indexRow, bodyColumn)
                    indexEmailMustContain1 = indexJobWS.Cells(indexRow, emailColumn1)
                    indexEmailMustContain2 = indexJobWS.Cells(indexRow, emailColumn2)
                    indexEmailMustContain3 = indexJobWS.Cells(indexRow, emailColumn3)
                    
                    ' If no more rules on the index, stop
                    If indexSubjectMustContain = "" And indexBodyMustContain = "" And indexEmailMustContain1 = "" And indexEmailMustContain2 = "" And indexEmailMustContain3 = "" Then
                        Exit For
                    End If
                    
                    ' Cycle through the filerWB rules for this job and check whether the current index rule is present (if not add it)
                    ruleFound = False
                    For filerRow = ruleStartRow To 1000
                    
                        ' Grab the data
                        filerSubjectMustContain = filerJobWS.Cells(filerRow, subjectColumn)
                        filerBodyMustContain = filerJobWS.Cells(filerRow, bodyColumn)
                        filerEmailMustContain1 = filerJobWS.Cells(filerRow, emailColumn1)
                        filerEmailMustContain2 = filerJobWS.Cells(filerRow, emailColumn2)
                        filerEmailMustContain3 = filerJobWS.Cells(filerRow, emailColumn3)
                        
                        ' If no more rules on the filerWB, stop
                        If filerSubjectMustContain = "" And filerBodyMustContain = "" And filerEmailMustContain1 = "" And filerEmailMustContain2 = "" And filerEmailMustContain3 = "" Then
                            Exit For
                        End If
                        
                        ' Check if it's a match
                        If filerSubjectMustContain = indexSubjectMustContain And filerBodyMustContain = indexBodyMustContain And filerEmailMustContain1 = indexEmailMustContain1 And filerEmailMustContain2 = indexEmailMustContain2 And filerEmailMustContain3 = indexEmailMustContain3 Then
                            ruleFound = True
                            Exit For
                        End If
                    
                    
                    Next filerRow
                    
                    ' If we haven't found the rule, add it
                    ' The row number will be the filerRow as we exited the for when we found nothing
                    ' We don't have to worry about exiting the for loop early when the rule is found as we only do this when it isn't found
                    If Not ruleFound Then
                        
                        filerJobWS.Cells(filerRow, subjectColumn) = indexSubjectMustContain
                        filerJobWS.Cells(filerRow, bodyColumn) = indexBodyMustContain
                        filerJobWS.Cells(filerRow, emailColumn1) = indexEmailMustContain1
                        filerJobWS.Cells(filerRow, emailColumn2) = indexEmailMustContain2
                        filerJobWS.Cells(filerRow, emailColumn3) = indexEmailMustContain3
                    
                    End If
                
                
                Next indexRow
                
                MsgBox ("The rules for " & selectedProjectNumber & " have been updated from the index." & vbNewLine & "Any rules on the index that weren't on your job's sheet have been added." & vbNewLine & vbNewLine & "Make sure you review them before filing your mail.")
                
            Else
            
                MsgBox ("Job sheet could not be updated as it doesn't exist in the Index")
            
            End If
            
            ' Clean up
            Set filerJobWS = Nothing
            Set indexJobWS = Nothing
            
            
        End If

    End If

    ' Click the listbox so that it updates (button colour etc.)
    Call JobListBox_Click

End Sub

Private Sub CloseButton_Click()

    Unload Me

End Sub

Private Sub UserForm_Terminate()
    
    ' Called when the Add Job Form is closed
    
    ' There is a bug that when a sheet is added from a form, even though it appears to be the active sheet (in the GUI)
    ' if the user scrolls down or enters data, it still thinks it's on the sheet that created the form (Dashboard)
    
    ' to fix this we need to select the dashboard and then reselect the required sheet AFTER the form is unloaded
    ' https://social.msdn.microsoft.com/Forums/office/en-US/ce24066c-2eaf-4590-948b-2f954cb3e456/add-sheet-via-form-using-vba-cause-wrong-activesheet?forum=exceldev
       
     ' Note: above fix doesn't work, but chanign the form to non modal does
    
    ' Close the index and clean up
    indexWB.Close False
    Set indexWB = Nothing
    Set filerWB = Nothing
    
    Unload Me
    

End Sub




