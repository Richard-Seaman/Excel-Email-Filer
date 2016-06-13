Attribute VB_Name = "Jobs"
' There's a known bug that makes the macro stop when you open a workbook in excel
' Seemingly, you need to check if the shift button is pressed

' Global variables

Public filerWB As Workbook
Public indexWB As Workbook

' See link: https://support.microsoft.com/en-us/kb/555263
' That's what the below two blocks of code are for

'Declare API
Public Declare Function GetKeyState Lib "User32" _
(ByVal vKey As Integer) As Integer
Const SHIFT_KEY = 16

Public Function ShiftPressed() As Boolean
'Returns True if shift key is pressed
    ShiftPressed = GetKeyState(16) < 0
End Function


Public Function IsFile(ByVal fName As String) As Boolean
'Returns TRUE if the provided name points to an existing file.
'Returns FALSE if not existing, or if it's a folder
    On Error Resume Next
    IsFile = ((GetAttr(fName) And vbDirectory) <> vbDirectory)
End Function

Public Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

     If wb Is Nothing Then Set wb = ThisWorkbook
     On Error Resume Next
     Set sht = wb.Sheets(shtName)
     On Error GoTo 0
     SheetExists = Not sht Is Nothing
 End Function

Public Function ReturnLargerProjectNumber(numberString1 As String, numberString2 As String) As String

    If StrComp(numberString1, numberString2, vbTextCompare) = 1 Then
        ReturnLargerProjectNumber = numberString1
    Else
        ReturnLargerProjectNumber = numberString2
    End If

End Function

Sub addJobsEng()
        
    Dim indexFilePath As String
    ' Grab the index file path from the filer
    indexFilePath = Sheets("INFO").Cells(4, 3)
    Call showAddJobsForm(indexFilePath)
    
End Sub

Sub addJobsArch()
    
    Dim indexFilePath As String
    ' Grab the index file path from the filer
    indexFilePath = Sheets("INFO").Cells(5, 3)
    Call showAddJobsForm(indexFilePath)
    
End Sub

Private Sub showAddJobsForm(indexFilePath As String)

    ' This sub is called from either the "Add M&E Jobs" or "Add Arch Jobs" buttons
        
    If Not FolderExists("\\bdp\Dublin\Workgrp\M&E") Then
        MsgBox ("Needs to be connected to the BDP network" & vbNewLine & "Additional files required.")
        End
    End If
    
    ' Check if index exists
    If Not IsFile(indexFilePath) Then
        MsgBox ("Could not find the index workbook. make sure the path is correct on the INFO tab. Fix and run again.")
        End
    End If

    ' Set the current WB and open up the index
    Set filerWB = ActiveWorkbook
    
    ' Don't show the indexWB opening
    ' (Note: if we open in a background application of excel there seems to be an issue copying the worksheets over, so just open within current application and hide)
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
    End With
            
    
    ' There's a known bug that makes the macro stop when you open a workbook in excel
    ' Seemingly, you need to check if the shift button is pressed
    ' See link: https://support.microsoft.com/en-us/kb/555263
    Do While ShiftPressed()
        DoEvents
    Loop
    Set indexWB = Workbooks.Open(indexFilePath, False, True)
    
    ' Hide the index
    indexWB.Windows(1).Visible = False
    
    ' Switch back to the filerWB
    filerWB.Activate
    
    ' Then show the workbook again
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    
    ' Show the form, now that the two workbooks are available
    AddJobs.Show
    
    
End Sub

Sub setupJobsEng()

    Dim indexFilePath As String
    ' Grab the index file path from the filer
    indexFilePath = Sheets("INFO").Cells(4, 3)
    Call openIndex(indexFilePath)
    
End Sub

Sub setupJobsArch()

    Dim indexFilePath As String
    ' Grab the index file path from the filer
    indexFilePath = Sheets("INFO").Cells(5, 3)
    Call openIndex(indexFilePath)
    
End Sub

Private Sub openIndex(indexFilePath As String)

    ' This sub is called by the "Set Up M&E Jobs" and "Set Up Arch Jobs" buttons
        
    If Not FolderExists("\\bdp\Dublin\Workgrp\M&E") Then
        MsgBox ("Needs to be connected to the BDP network" & vbNewLine & "Additional files required.")
        End
    End If
    
    ' It opens the corresponding index (if it exists) on the Dashboard

    ' Check if index exists
    If Not IsFile(indexFilePath) Then
        MsgBox ("Could not find the index workbook. make sure the path is correct on the INFO tab. Fix and run again.")
        End
    End If
    
    ' Open the workbook
    Workbooks.Open indexFilePath, False, False
    
    ' Show the dashboard
    ActiveWorkbook.Sheets("Dashboard").Activate
    
    ' Select the project name cell
    Cells(8, 3).Select

End Sub

