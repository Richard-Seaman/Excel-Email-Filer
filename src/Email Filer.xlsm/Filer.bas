Attribute VB_Name = "Filer"
' Use global variables so all modules can access
' NB: order of Sub calls is important so that varibales are set before they're accessed

Dim objNS As Outlook.Namespace
    
Dim userFolder As Outlook.MAPIFolder                ' Entire mailbox
Dim newformaFolder As Outlook.MAPIFolder            ' Newforma folders

Dim inboxFolder As Outlook.MAPIFolder               ' Inbox Folder
Dim sentFolder As Outlook.MAPIFolder                ' Sent Folder

Dim projectNewformaFolder As Outlook.MAPIFolder     ' Where to file it
Dim projectUserFolder As Outlook.MAPIFolder         ' Where the user wants it

Dim inboxItems As Outlook.Items                     ' Inbox items (mail)
Dim sentItems As Outlook.Items                      ' Sent items (mail)

Dim oFolder As Outlook.MAPIFolder                   ' a mail folder
Dim oMail As Outlook.MailItem                       ' a mail item
Dim cMailCopy As Outlook.MailItem                   ' a mail item (copied from oMail)


' The folder name for me was "Mailbox - Seaman, Richard" rather than "richard.seaman@bdp.com"
' The user must use whatever the folder name is
Dim mymailbox As String                             ' main mailbox name - entered on the dashboard
Dim mainNewformaFolderName As String                ' main mailbox name - entered on the dashboard

Dim fileInbox As Boolean                            ' whether to file the user's Inbox mail items - entered on the dashboard
Dim fileSent As Boolean                             ' whether to file the user's Sent mail items - entered on the dashboard

Dim fileCategories As Boolean                       ' whether to file emails with a category
Dim fileFlagged As Boolean                          ' whether to file emails that are flagged
Dim fileUnread As Boolean                           ' whether to file unread mail items - entered on the dashboard
Dim minNumberOfWeeksOld As Integer                  ' Don't file email if less than X weeks old - entered on the dashboard

Dim dashboard As Worksheet                          ' the dashboard worksheet


' Used to keep track of the currently found folder (from the function at bottom) so that we can set it within the main Sub
Dim currentFoundFolder As Outlook.MAPIFolder

' Change the max wait time here (in seconds)
Private Const maximumWaitForVaultRestore As Integer = 45


' Rule variable positions to allow them to be easily changed in future
Public Const ruleStartRow As Integer = 25
Public Const subjectColumn As Integer = 3
Public Const bodyColumn As Integer = 4
Public Const emailColumn1 As Integer = 5
Public Const emailColumn2 As Integer = 6
Public Const emailColumn3 As Integer = 7
Public Const senderEmailColumn As Integer = 8
Public Const subjectColumnNot As Integer = 9
Public Const bodyColumnNot As Integer = 10
Public Const emailColumn1Not As Integer = 11
Public Const emailColumn2Not As Integer = 12
Public Const emailColumn3Not As Integer = 13
Public Const senderEmailColumnNot As Integer = 14

Dim progress As Shape
Dim summary As Shape
Dim numberEmailsFiled As Integer

Dim tryAgain As Boolean


Public Function FolderExists(strFolderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (GetAttr(strFolderPath) And vbDirectory) = vbDirectory
    On Error GoTo 0
End Function


Sub hideCover()

    ' Make sure they've saved a copy
    If MsgBox("Are you sure you have saved your own copy of this and are not going to overwrite the original?" & vbNewLine & vbNewLine & "Note: Microsoft Office 2013 required to use", vbYesNo) = vbYes Then
        ' used to hide the initial cover
        ActiveSheet.Shapes("CoverSHape").Visible = False
    End If
    
End Sub

Sub showCover()
    ' Used to show the inital cover again (do this before saving it)
    ActiveSheet.Shapes("CoverSHape").Visible = True
End Sub

Sub hideSummary()
    ' Used to hide the summary
    ActiveSheet.Shapes("Summary").Visible = False
End Sub

Sub hideProgress()
    ' Used to initally hide (can't find the property to do thi (For manual use only)
    ActiveSheet.Shapes("Progress").Visible = False

End Sub

Sub fileEmails()

    If Not FolderExists("\\bdp\Dublin\Workgrp\M&E") Then
        MsgBox ("Needs to be connected to the BDP network" & vbNewLine & "Additional files required.")
        End
    End If

    ' Check that some jobs have been added
    If ActiveWorkbook.Worksheets.Count <= 2 Then
        MsgBox ("You haven't added any jobs yet!" & vbNewLine & "Use the Add Jobs button on the dashboard")
        End
    End If

    ' Warn the User
    If MsgBox("You are about to file your emails, are you sure you've configured the Dashboard correctly and reviewed all your project rules?", vbYesNo) = vbNo Then End
        
    ' Grab the dashboard
    Set dashboard = Sheets("Dashboard")
    
    ' Check whether to file unread mail
    fileUnread = False
    If LCase(dashboard.Cells(3, 3)) <> "y" And LCase(dashboard.Cells(3, 3)) <> "n" Then
        MsgBox ("You must enter Y or N for whether to file unread email or not" & vbNewLine & "Refer to the Dashboard.")
        End
    End If
    If LCase(dashboard.Cells(3, 3)) = "y" Then fileUnread = True
    
    ' Check the minium mail age
    minNumberOfWeeksOld = 0
    userEnteredMinNumberOfWeeksOld = dashboard.Cells(4, 3)
    If IsNumeric(userEnteredMinNumberOfWeeksOld) Then
        If userEnteredMinNumberOfWeeksOld < 0 Then
            MsgBox ("The don't file email that is less than X weeks old value must be positive." & vbNewLine & "Refer to the Dashboard.")
            End
        End If
        
        ' Convert whatever was entered to an integer
        minNumberOfWeeksOld = Int(userEnteredMinNumberOfWeeksOld)
        
    Else
        MsgBox ("The don't file email that is less than X weeks old value must be a number." & vbNewLine & "Refer to the Dashboard.")
        End
    End If
    
    
    ' Check wheter to include inbox and/or sent
    fileInbox = False
    If LCase(dashboard.Cells(8, 3)) <> "y" And LCase(dashboard.Cells(8, 3)) <> "n" Then
        MsgBox ("You must enter Y or N for whether to file emails from your inbox." & vbNewLine & "Refer to the Dashboard.")
        End
    End If
    If LCase(dashboard.Cells(8, 3)) = "y" Then fileInbox = True
    
    fileSent = False
    If LCase(dashboard.Cells(9, 3)) <> "y" And LCase(dashboard.Cells(9, 3)) <> "n" Then
        MsgBox ("You must enter Y or N for whether to file emails from your sent folder." & vbNewLine & "Refer to the Dashboard.")
        End
    End If
    If LCase(dashboard.Cells(9, 3)) = "y" Then fileSent = True
    
    If Not fileInbox And Not fileSent Then
        MsgBox ("You must apply to the Inbox folder or the Sent folder (or both), currently neither included." & vbNewLine & "Refer to the Dashboard.")
    End If
    
    ' Assume we don't file categories
    fileCategories = False
    If LCase(dashboard.Cells(10, 3)) <> "y" And LCase(dashboard.Cells(10, 3)) <> "n" Then
        MsgBox ("You must enter Y or N for whether to file emails with a category assigned." & vbNewLine & "Refer to the Dashboard.")
        End
    End If
    If LCase(dashboard.Cells(10, 3)) = "y" Then fileCategories = True
    
    ' Assume we don't file flags
    fileFlagged = False
    If LCase(dashboard.Cells(11, 3)) <> "y" And LCase(dashboard.Cells(11, 3)) <> "n" Then
        MsgBox ("You must enter Y or N for whether to file emails with a flag assigned." & vbNewLine & "Refer to the Dashboard.")
        End
    End If
    If LCase(dashboard.Cells(11, 3)) = "y" Then fileFlagged = True
        
    ' Grab the folder names from the worksheet
    mymailbox = dashboard.Cells(5, 3)
    mainNewformaFolderName = dashboard.Cells(6, 3)
        
    ' Namespace, whatever that is...
    Set objNS = GetNamespace("MAPI")
    
    ' Before we start setting folders, we need to make sure they exist
    ' Otherwise we'll get an error (and people won't know what to do)
    ' Refer to checkIfMailFolderExistsInFolder function at bottom
        
        
    ' Check that user folders exist (it won't if the user entered it wrong)
    
    ' Cycle through each folder in the namespace and check if the user folder exists
    ' (we have to cycle through these as objNS is a NameSpace not a MAPIFolder so we can't use the function below)
    folderFound = False
    For Each oFolder In objNS.Folders
        
        ' If the folder name matches the one searched for, exit
        If oFolder.Name = mymailbox Then
            folderFound = True
            Set userFolder = objNS.Folders(mymailbox)
            Exit For
        End If
        
    Next oFolder
    
    If Not folderFound Then
        ' Folder not found
        MsgBox ("Could not find your mailbox, make sure you have entered it correctly (as it appears in your Outlook)" & vbNewLine & vbNewLine & "Note: there are two possible formats" & vbNewLine & "Refer to the example on the dashboard")
        End
    End If
            
    
    ' Check that the Newforma folders exists
    If checkIfMailFolderExistsInFolder(mainNewformaFolderName, userFolder) Then
        ' Main Newforma Folder definately exists, okay to Set from CurrentFoundFolder
        Set newformaFolder = currentFoundFolder
    Else
        ' Folder not found
        MsgBox ("Could not find the parent Newforma folder, make sure you have entered it correctly (as it appears in your Outlook)" & vbNewLine & vbNewLine & "Refer to the example on the dashboard")
        End
    End If
            
        
    ' Inbox & sent folders are there by default (so don't need to check if they exist??)
    Set inboxFolder = objNS.GetDefaultFolder(olFolderInbox)
    Set sentFolder = objNS.GetDefaultFolder(olFolderSentMail)
    
    ' Grab the items from the inbox / sent folders
    Set inboxItems = inboxFolder.Items
    Set sentItems = sentFolder.Items
    
    ' Count the number of fileable items (for output at end)
    ' count mail items only
    totalNumberOfMailItems = 0
    If fileInbox Then
        For Each MailItem In inboxItems
            totalNumberOfMailItems = totalNumberOfMailItems + 1
        Next MailItem
    End If
    If fileSent Then
        For Each MailItem In sentItems
            totalNumberOfMailItems = totalNumberOfMailItems + 1
        Next MailItem
    End If
    
    ' Show/Hide the shapes
    Set progress = dashboard.Shapes("Progress")
    Set summary = dashboard.Shapes("Summary")
    progress.Visible = True
    summary.Visible = False
    
    progress.TextFrame.Characters.Text = "Emails Filing, Please Wait..." & vbNewLine & vbNewLine & "(This may take some time, be patient!)"
    summary.TextFrame.Characters.Text = ""
    
    Excel.Application.ScreenUpdating = True
    progress.TextFrame.Characters.Text = "Emails Filing, Please Wait..." & vbNewLine & vbNewLine & "(This may take some time, be patient!)"
    Application.wait (Now + TimeValue("00:00:02"))
    'Excel.Application.ScreenUpdating = False
    
        
    ' Cycle through all sheets and call the fileJobEmails sub (if it's not a job sheet it will be skipped)
    Dim currentSheet As Worksheet
    numberEmailsFiled = 0
    
    For Each ws In ActiveWorkbook.Worksheets
        ' Exclude the dashboard and info
        If ws.Name <> "Dashboard" And ws.Name <> "INFO" Then
            Set currentSheet = ws
            Call fileJobEmails(currentSheet)
        End If
    Next ws
    
    ' About to start restoring emails, warn the user
    progress.TextFrame.Characters.Text = "Finished filing emails" & vbNewLine & "Restoring filed emails from the vault so they can be synchronised with Newforma." & vbNewLine & vbNewLine & "(You will see some emails opening and closing, just ignore them)"
    Excel.Application.ScreenUpdating = True
    progress.TextFrame.Characters.Text = "Finished filing emails" & vbNewLine & "Restoring filed emails from the vault so they can be synchronised with Newforma." & vbNewLine & vbNewLine & "(You will see some emails opening and closing, just ignore them)"
    Application.wait (Now + TimeValue("00:00:02"))
    'Excel.Application.ScreenUpdating = False
       
    ' Restore Newforma emails
    
    ' MODIFIED 26/05/16
    ' Automatically restoring was causing too many issues on first runs
    ' Changed so that you must press Restore after you file
    ' Can revert back once people have used it a few times
    
    ' Call restoreEmailsInNewformaFolders
    
    ' Hide the blue shape, show the green shape
    progress.Visible = False
    summary.Visible = True
    
    If numberEmailsFiled > 1 Then
        ' More than 1, plural
        summary.TextFrame.Characters.Text = "Finished" & vbNewLine & vbNewLine & numberEmailsFiled & " emails were filed." & vbNewLine & "Make sure you check them to ensure they're in the right folders." & vbNewLine & vbNewLine & "(Click to dismiss)"
    ElseIf numberEmailsFiled = 1 Then
        ' Only 1, singular
        summary.TextFrame.Characters.Text = "Finished" & vbNewLine & vbNewLine & numberEmailsFiled & " email was filed." & vbNewLine & "Make sure you check it to ensure it's in the right folder." & vbNewLine & vbNewLine & "(Click to dismiss)"
    Else
        ' None
        summary.TextFrame.Characters.Text = "Finished" & vbNewLine & vbNewLine & "No emails were filed." & vbNewLine & vbNewLine & "(Click to dismiss)"
    End If
    
    
    ' Record the number of mails filed
    Dim strFile_Path As String
    strFile_Path = "K:\M&E\Calculations\APPLICATIONS\EmailFilerLog.txt"
    
    ' Export in format ready for import to excel
    Open strFile_Path For Append As #1
    Write #1, Excel.Application.UserName,
    Write #1, totalNumberOfMailItems,
    Write #1, numberEmailsFiled,
    Write #1, Format(Now(), "yyyy/MM/dd"),
    Write #1, Environ$("computername")
    Close #1
    
 
    
    ' Clean up
    Set currentSheet = Nothing
    Set objNS = Nothing
    Set inboxFolder = Nothing
    Set sentFolder = Nothing
    Set newformaDestFolder = Nothing
    Set userDestFolder = Nothing
    Set userFolder = Nothing
    Set newformaFolder = Nothing
    Set inboxItems = Nothing
    Set sentItems = Nothing
    Set oFolder = Nothing
    Set oMail = Nothing
    Set cMailCopy = Nothing
    
    
End Sub

Sub restoreEmailsInNewformaFolders()
        
    If Not FolderExists("\\bdp\Dublin\Workgrp\M&E") Then
        MsgBox ("Needs to be connected to the BDP network" & vbNewLine & "Additional files required.")
        End
    End If

    ' Note: This sub is called by the FileEmails sub but can also be called independantly

    ' SAME CODE AS FOR FILE EMAILS
    ' (makes sure folders are present)

    ' Grab the dashboard
    Set dashboard = Sheets("Dashboard")

    ' Grab the folder names from the worksheet
    mymailbox = dashboard.Cells(5, 3)
    mainNewformaFolderName = dashboard.Cells(6, 3)
        
    ' Namespace, whatever that is...
    Set objNS = GetNamespace("MAPI")
    
    ' Before we start setting folders, we need to make sure they exist
    ' Otherwise we'll get an error (and people won't know what to do)
    ' Refer to checkIfMailFolderExistsInFolder function at bottom
        
    ' Check that user folders exist (it won't if the user entered it wrong)
    
    ' Cycle through each folder in the namespace and check if the user folder exists
    ' (we have to cycle through these as objNS is a NameSpace not a MAPIFolder so we can't use the function below)
    folderFound = False
    For Each oFolder In objNS.Folders
        
        ' If the folder name matches the one searched for, exit
        If oFolder.Name = mymailbox Then
            folderFound = True
            Set userFolder = objNS.Folders(mymailbox)
            Exit For
        End If
        
    Next oFolder
    
    If Not folderFound Then
        ' Folder not found
        MsgBox ("Newforma mail could not be unarchived" & vbNewLine & "Could not find your mailbox, make sure you have entered it correctly (as it appears in your Outlook)" & vbNewLine & vbNewLine & "Note: there are two possible formats" & vbNewLine & "Refer to the example on the dashboard")
        Exit Sub
    End If
    
    ' Check that the Newforma folders exists
    If checkIfMailFolderExistsInFolder(mainNewformaFolderName, userFolder) Then
        ' Main Newforma Folder definately exists, okay to Set from CurrentFoundFolder
        Set newformaFolder = currentFoundFolder
    Else
        ' Folder not found
        MsgBox ("Newforma mail could not be unarchived" & vbNewLine & "Could not find the parent Newforma folder, make sure you have entered it correctly (as it appears in your Outlook)" & vbNewLine & vbNewLine & "Refer to the example on the dashboard")
        Exit Sub
    End If
    
    
    ' Show/Hide the shapes
    Set progress = dashboard.Shapes("Progress")
    Set summary = dashboard.Shapes("Summary")
    progress.Visible = True
    summary.Visible = False
    
    
    ' CODE FOR RESTORING EMAILS STARTS HERE
    
    ' There's a trick for restoring emails from the vault:
    ' http://stackoverflow.com/questions/17198895/do-for-all-open-emails-and-move-to-a-folder
    ' Basically:
    ' open the mail (wait for it to unarchive)
    ' make a copy of the unarchived version
    ' delete the original archived version
        
    ' Keep a count of the number of mail unarchived
    unarchivedCount = 0
    
    ' cycle through each of the Newforma folders
    For Each oFolder In newformaFolder.Folders
    
        
        ' Update the progress
        progress.TextFrame.Characters.Text = "Restoring archived emails from " & oFolder.Name & vbNewLine & vbNewLine & "(You will see some emails opening and closing, just ignore them)"
        forceScreenToUpdateAndPause ("00:00:01")
    
        For i = oFolder.Items.Count To 1 Step -1
            
            Set oMail = oFolder.Items(i)
        
            If oMail.Subject = "RE: PM12003 Lucan CC - Project Directory - Rev3 16.11.2015" Then
                breakpoint = True
            End If
            
            ' Check if it's an archived item
            archive = (InStr(1, oMail.MessageClass, "EnterpriseVault", vbTextCompare) > 0)
            If archive Then
                
                ' Display the mail (this extracts it from the vault)
                oMail.Display
                
                ' Sometimes enterprise vault takes a few seconds to open the archived mail
                ' In this case, the Set myInspectors will throw an error (because there's no CurrentItem)
                ' We need to catch when this happens, pause the program, and then try again
                ' This should give the vault a chance to open the mail
                ' We do this by looping and pausing whenever an error is detected
                
                numberOfTimesWaited = 0
                
                On Error GoTo WaitWhileVaultExtracts
                tryAgain = False
                                
                ' This code makes a copy of the current unarchived version
                ' if it hits an error, it waits and trys again (see error handler at bottom of sub)
                ' until the max number of attempts have been made, after which it skips this mail
tryAgain:
                Set myInspectors = Outlook.Application.ActiveInspector.CurrentItem
                If tryAgain Then
                    tryAgain = False
                    GoTo tryAgain
                End If
                Set myCopiedInspectors = myInspectors.Copy
                If tryAgain Then
                    tryAgain = False
                    GoTo tryAgain
                End If
                myCopiedInspectors.Move oFolder
                If tryAgain Then
                    tryAgain = False
                    GoTo tryAgain
                End If
                myInspectors.Close olDiscard
                If tryAgain Then
                    tryAgain = False
                    GoTo tryAgain
                End If
                Set myCopiedInspectors = Nothing
                
                ' Delete the original archived version
                oMail.Delete
                
                ' Increment the counter
                unarchivedCount = unarchivedCount + 1
                        
            End If
           
        Next
        
    Next oFolder
    
    ' Give a summary of the emails restored
    summaryTextToUse = ""
    If unarchivedCount > 1 Then
        ' More than 1, plural
        summaryTextToUse = "Finished restoring from Vault" & vbNewLine & vbNewLine & unarchivedCount & " emails were restored and are ready for upload to Newforma."
    ElseIf unarchivedCount = 1 Then
        ' Only 1, singular
        summaryTextToUse = "Finished restoring from Vault" & vbNewLine & vbNewLine & unarchivedCount & " email was restored and is ready for upload to Newforma."
    Else
        ' None
        summaryTextToUse = "Finished restoring from Vault" & vbNewLine & vbNewLine & "No emails were restored."
    End If

    ' Give a summary of the emails restored
    progress.TextFrame.Characters.Text = summaryTextToUse
    forceScreenToUpdateAndPause ("00:00:03")
    
    ' Need to hide the progress here in case this sub was called on its own
    progress.Visible = False
    
    
    ' Not all of these are used but clean up anyway
    Set currentSheet = Nothing
    Set objNS = Nothing
    Set inboxFolder = Nothing
    Set sentFolder = Nothing
    Set newformaDestFolder = Nothing
    Set userDestFolder = Nothing
    Set userFolder = Nothing
    Set newformaFolder = Nothing
    Set inboxItems = Nothing
    Set sentItems = Nothing
    Set oFolder = Nothing
    Set oMail = Nothing
    Set cMailCopy = Nothing
    Set oMail = Nothing
    Set newformaFolder = Nothing
    
    ' Exit before it reaches the error handler
    
    Exit Sub
    
WaitWhileVaultExtracts:
    
    ' This Error handler can only be executed once before it Resumes to the next line (from the statement that caused it)
    ' Otherwise, the next statement that causes the error won't be caught, as it thinks it's still handling the first error
    ' to get around this, we resume everytime but set a tryAgain variable to true
    ' When this is true, the code goes back to before the statements that caused the error and tries again
    
    ' Wait a second, then try again
    If numberOfTimesWaited >= maximumWaitForVaultRestore Then
        ' if we've already waited awhile, move onto the next mail
        Resume Next
    End If
    Application.wait (Now + TimeValue("00:00:01"))
    numberOfTimesWaited = numberOfTimesWaited + 1
    
    ' Set try again to true and resume
    tryAgain = True
    Resume

End Sub


Sub fileJobEmails(currentWorksheet As Worksheet)

    ' This sub needs to be applied to all sheets
    
    ' Note: Exit Sub rather than End is used so that it returns to the main fileEmails sub (and skips this sheet)
        
        
    ' Make sure this sheet is a Job's Rules sheet
    ' If it is, Cells(2,1) will be "Project Number"
    If currentWorksheet.Cells(2, 1) <> "Project Number" Then Exit Sub ' (skips this sheet)
    
    
    Dim projectNewformaFolderName As String
    Dim projectUserFolderName As String
    
    ' grab the name (needed for MsgBoxes and progress)
    projectName = currentWorksheet.Cells(3, 2)
        
    ' Define the specific Newforma Folder (name)
    projectNewformaFolderName = currentWorksheet.Cells(5, 2)
    projectUserFolderName = currentWorksheet.Cells(7, 2)
    
    ' Filing to a newforma folder is now optional
    ' This allows non job specific rules to be set up (marketing, admin etc.)
    fileToNewformaFolder = False
    If projectNewformaFolderName <> "" Then fileToNewformaFolder = True
    
    ' Filing to an additional user folder is optional
    fileToUserFolder = False
    If projectUserFolderName <> "" Then fileToUserFolder = True
    
    ' If neither folder entered:
    If Not fileToNewformaFolder And Not fileToUserFolder Then
        MsgBox ("Emails not filed for " & projectName & vbNewLine & vbNewLine & "You must specific a Newforma folder or a User folder." & vbNewLine & vbNewLine & "Please correct this error and run again.")
        Exit Sub '(skip this sheet)
    End If
              
              
    ' Check newforma folder exists (only if one was entered)
    If fileToNewformaFolder Then
    
        ' Check that the Project Newforma folder exists
        If checkIfMailFolderExistsInFolder(projectNewformaFolderName, newformaFolder) Then
            ' Specified Newforma Folder definately exists, okay to Set from CurrentFoundFolder
            Set projectNewformaFolder = currentFoundFolder
        Else
            ' The user entered a newforma folder but it wasn't found. Don't file the mail until this is corrected
            MsgBox ("Emails not filed for " & projectName & vbNewLine & vbNewLine & "Could not find the specified Newforma folder '" & projectNewformaFolderName & "', make sure you have it in your Outlook" & vbNewLine & vbNewLine & "Please correct this error and run again.")
            Exit Sub ' (skips this sheet)
        End If
    
    End If
    
    
    ' Check user's specified folder exists (only if one was enetered)
    If fileToUserFolder Then
    
        If checkIfMailFolderExistsInFolder(projectUserFolderName, userFolder) Then
            ' Specified User Folder definately exists, okay to Set from CurrentFoundFolder
            Set projectUserFolder = currentFoundFolder
        Else
            ' The user entered a user folder but it wasn't found. Don't file the mail until this is corrected
            MsgBox ("Emails not filed for " & projectName & vbNewLine & vbNewLine & "Could not find the specified user folder '" & projectUserFolderName & "', make sure you have it in your Outlook" & vbNewLine & vbNewLine & "Please correct this error and run again.")
            Exit Sub ' (skips this sheet)
        End If
    
    End If
         
         
    ' Make an array of the items so we don't have to duplicate the code below for each of the items
    ' Also easier to add items in future if we ever need to
    Dim itemsArray(1) As Outlook.Items
    Set itemsArray(0) = inboxItems
    Set itemsArray(1) = sentItems
    
    ' Also need a corresponding Boolean array for whther to include the items at each index
    Dim includeArray(1) As Boolean
    includeArray(0) = fileInbox
    includeArray(1) = fileSent
    
    Dim currentItems As Outlook.Items
    
    ' Cycle through each of the items (currently only Inbox and Sent)
        
    ' Count how many filed
    numberEmailsFiledForCurrentJob = 0
    
    ' Only report every so many emails (otherwise the application breaks add up and make it take much longer to run)
    reportInterval = 10
    numberOfEmailsFiledSincelastProgressUpdate = 0
        
    ' Update the progress shape for this job
    progress.TextFrame.Characters.Text = "Current Job: " & projectName & vbNewLine & "Number of emails filed: 0" & vbNewLine & vbNewLine & "(You can still use Outlook, just don't delete anything!)"
    forceScreenToUpdateAndPause ("00:00:01")
    
    For currentItemsIndex = 0 To UBound(itemsArray) Step 1
    
        Set currentItems = itemsArray(currentItemsIndex)
        includeCurrentItems = includeArray(currentItemsIndex)
        
        ' Only continue if we're filing the current items
        If includeCurrentItems Then
                
            ' Cycle backwards through the mail items so that when we move we don't skip any (the index still works)
            For i = currentItems.Count To 1 Step -1
            
                ' Make sure it's a mail item
                If TypeName(currentItems(i)) = "MailItem" Then
                
                    ' Set the current mail item
                    Set oMail = currentItems(i)
                                        
                    ' Default data (in case error grabbing the actual data
                    currentMailAge = Format(Now, "dd/MM/yyyy hh:mm:ss")
                    currentMailUnread = True
                    currentMailSubject = ""
                    currentMailBody = ""
                    currentMailSenderEmail = ""
                    currentMailCategory = ""
                    currentMailSenderCategory = ""
                    currentMailFlag = ""
                    
                    ' Try to Grab the actual data from the mail
                    On Error Resume Next
                    currentMailAge = oMail.SentOn
                    currentMailUnread = oMail.UnRead
                    currentMailSubject = oMail.Subject
                    currentMailBody = oMail.Body
                    currentMailSenderEmail = oMail.SenderEmailAddress
                    currentMailCategory = oMail.Categories
                    currentMailFlag = oMail.FlagStatus
                                    
                    ' Grab the recipient addresses
                    ' This is quite tricky, but afterwards the array "currentMailRecipients" will have all the recipient email addresses in it
                    Dim person As Outlook.Recipient
                    Dim currentMailRecipients() As String
                     
                    ' Erase the array to ensure there are no emails carried over from the previous mail
                    Erase currentMailRecipients
                      
                    currentMailRecipientIndex = 0
                    For Each person In oMail.Recipients
                        currentAddress = person.Address
                        ReDim Preserve currentMailRecipients(currentMailRecipientIndex + 1) As String
                        currentMailRecipients(currentMailRecipientIndex) = currentAddress
                        currentMailRecipientIndex = currentMailRecipientIndex + 1
                    Next person
                    
                    ' Add an empty email ("") to the array to ensure the array isn't empty (which can cause errors in the loops below)
                    ' It's okay to add a blank, as emails won't be contained in it and blanks aren't checked anyway
                    ReDim Preserve currentMailRecipients(currentMailRecipientIndex + 1) As String
                    currentMailRecipients(currentMailRecipientIndex) = ""
                    currentMailRecipientIndex = currentMailRecipientIndex + 1
                         
                    ' Check whether the mail is unread
                    okayToFileMailDueToReadStatus = fileUnread
                    
                    ' If it's not okay to file unread mail
                    If Not okayToFileMailDueToReadStatus Then
                        ' Check if the current mail is unread (if not, then file away)
                        If Not currentMailUnread Then
                            okayToFileMailDueToReadStatus = True
                        End If
                    End If
                    
                    ' Check whether the mail is not old enough to be filed
                    Dim earliestDate As Date
                    today = Format(Now, "dd/MM/yyyy hh:mm:ss")
                    earliestDate = Format(DateAdd("d", -7 * minNumberOfWeeksOld, today), "dd/MM/yyyy hh:mm:ss")
                    
                    mailIsOldEnough = False
                    If currentMailAge < earliestDate Then
                        ' Mail is old enough
                        mailIsOldEnough = True
                    End If
                    
                    ' Check whether we can file it based on it's category
                    ' assume we can
                    mailCategoryCanBeFiled = True
                    ' if we're not filing emails with categories
                    If Not fileCategories Then
                        ' Check if the mail has a category, if so, don't file it
                        If currentMailCategory <> "" Then
                            mailCategoryCanBeFiled = False
                        End If
                    End If
                    
                    ' Check whether we can file it based on it's flag
                    ' assume we can
                    mailFlagCanBeFiled = True
                    ' if we're not filing emails with flags
                    If Not fileFlagged Then
                        ' Check if the mail has a flag, if so, don't file it (if it's a complete flag, then it's okay to file)
                        If currentMailFlag <> 0 Then
                            mailFlagCanBeFiled = False
                        End If
                    End If
                    
                    ' Only continue if the mail's unread status is okay and it's old enough and we can file it based on its category and flag
                    If okayToFileMailDueToReadStatus And mailIsOldEnough And mailCategoryCanBeFiled And mailFlagCanBeFiled Then
                                
                               
                        ' Cycle through each of the rules
                        For ruleRow = ruleStartRow To 100000
                        
                            ' Grab the data
                            subjectMustContain = currentWorksheet.Cells(ruleRow, subjectColumn)
                            bodyMustContain = currentWorksheet.Cells(ruleRow, bodyColumn)
                            emailMustContain1 = currentWorksheet.Cells(ruleRow, emailColumn1)
                            emailMustContain2 = currentWorksheet.Cells(ruleRow, emailColumn2)
                            emailMustContain3 = currentWorksheet.Cells(ruleRow, emailColumn3)
                            emailSenderMustContain = currentWorksheet.Cells(ruleRow, senderEmailColumn)
                            
                            subjectMustNotContain = currentWorksheet.Cells(ruleRow, subjectColumnNot)
                            bodyMustNotContain = currentWorksheet.Cells(ruleRow, bodyColumnNot)
                            emailMustNotContain1 = currentWorksheet.Cells(ruleRow, emailColumn1Not)
                            emailMustNotContain2 = currentWorksheet.Cells(ruleRow, emailColumn2Not)
                            emailMustNotContain3 = currentWorksheet.Cells(ruleRow, emailColumn3Not)
                            emailSenderMustNotContain = currentWorksheet.Cells(ruleRow, senderEmailColumnNot)
                            
                            ' If no data entered (under AND columns), stop (no more rules)
                            If subjectMustContain = "" And bodyMustContain = "" And emailMustContain1 = "" And emailMustContain2 = "" And emailMustContain3 = "" And emailSenderMustContain = "" Then
                                Exit For
                            End If
                            
                            ' Assume we do file the current mail
                            fileCurrentMail = True
                            
                            ' Cycle through each condition in turn, if any of them fail don't file it
                            ' (by assuming true and then setting to false, even if one fails it will stop it filing)
                            
                            ' Note: InStr() returns an integer so we can't use "Not InStr(...)" to check if it's not in it, we have to use "InStr(...) = 0"
                            
                            ' Using Like instead of InStr() allows the use of wildcards...
                            
                            ' Use the string to compare first
                            ' Followed by the string to search for, wrapped in wildcards...
                            
                            
                            ' SUBJECT CONDITIONS
                            
                            ' Check the subject condition
                            ' must contain
                            If subjectMustContain <> "" And Not LCase(currentMailSubject) Like "*" & LCase(subjectMustContain) & "*" Then
                                ' Does not meet the condition
                                fileCurrentMail = False
                            End If
                            ' must not contain
                            If subjectMustNotContain <> "" And LCase(currentMailSubject) Like "*" & LCase(subjectMustNotContain) & "*" Then
                                ' Does not meet the condition
                                fileCurrentMail = False
                            End If
                            
                            
                            ' BODY CONDITIONS
                            
                            ' Check the body condition (if we haven't failed yet)
                            ' must contain
                            If fileCurrentMail And bodyMustContain <> "" And Not LCase(currentMailBody) Like "*" & LCase(bodyMustContain) & "*" Then
                                ' Does not meet the condition
                                fileCurrentMail = False
                            End If
                            ' must not contain
                            If fileCurrentMail And bodyMustNotContain <> "" And LCase(currentMailBody) Like "*" & LCase(bodyMustNotContain) & "*" Then
                                ' Does not meet the condition
                                fileCurrentMail = False
                            End If
                            
                            
                            ' ANY EMAIL CONDITIONS
                            
                            ' Check the any email must contain conditions
                            Dim emailMustContainArray(2) As String
                            emailMustContainArray(0) = emailMustContain1
                            emailMustContainArray(1) = emailMustContain2
                            emailMustContainArray(2) = emailMustContain3
                                                        
                            ' Cycle through each email must contain string
                            For Each mustContainEmailString In emailMustContainArray
                            
                                ' If we haven't failed yet and the string isn't empty
                                If fileCurrentMail And mustContainEmailString <> "" Then
                                
                                    ' The email can be in any of the emails, sender or recipient
                                    foundInAtLeastOneEmail = False
                                                
                                    If LCase(currentMailSenderEmail) Like "*" & LCase(mustContainEmailString) & "*" Then
                                        ' Found it
                                        foundInAtLeastOneEmail = True
                                    End If
                                    
                                    ' If we haven't found one yet
                                    If Not foundInAtLeastOneEmail Then
                                    
                                        ' Cycle through each of the recipient addresses
                                        
                                        If Not IsEmpty(currentMailRecipients) Then
                                        
                                            For Each recipientEmail In currentMailRecipients
                                            
                                                If LCase(recipientEmail) Like "*" & LCase(mustContainEmailString) & "*" Then
                                                    ' Found it
                                                    foundInAtLeastOneEmail = True
                                                    ' Don't bother checking anymore recipient addresses
                                                    Exit For
                                                End If
                                            Next
                                            
                                        End If
                                    
                                    End If
                                    
                                    ' If we didn't find it in any of the email addresses, only then set fileCurrentMail to false
                                    If Not foundInAtLeastOneEmail Then fileCurrentMail = False
                                    
                                    
                                End If
                            
                            Next
                            
                            ' Check the any email must NOT contain conditions
                            Dim emailMustNotContainArray(2) As String
                            emailMustNotContainArray(0) = emailMustNotContain1
                            emailMustNotContainArray(1) = emailMustNotContain2
                            emailMustNotContainArray(2) = emailMustNotContain3
                            
                            ' Cycle through each email must contain string
                            For Each mustNotContainEmailString In emailMustNotContainArray
                            
                                ' If we haven't failed yet and the string isn't empty
                                If fileCurrentMail And mustNotContainEmailString <> "" Then
                                
                                    ' The email can be in any of the emails, sender or recipient
                                    foundInAtLeastOneEmail = False
                                                
                                    If LCase(currentMailSenderEmail) Like "*" & LCase(mustNotContainEmailString) & "*" Then
                                        ' Found it
                                        foundInAtLeastOneEmail = True
                                    End If
                                    
                                    ' If we haven't found one yet
                                    If Not foundInAtLeastOneEmail Then
                                    
                                        ' Cycle through each of the recipient addresses
                                        
                                        If Not IsEmpty(currentMailRecipients) Then
                                        
                                            For Each recipientEmail In currentMailRecipients
                                            
                                                If LCase(recipientEmail) Like "*" & LCase(mustNotContainEmailString) & "*" Then
                                                    ' Found it
                                                    foundInAtLeastOneEmail = True
                                                    ' Don't bother checking anymore recipient addresses
                                                    Exit For
                                                End If
                                            
                                            Next
                                            
                                        End If
                                        
                                    End If
                                    
                                    ' If we find it in any of the email addresses, only then set fileCurrentMail to false
                                    If foundInAtLeastOneEmail Then fileCurrentMail = False
                                    
                                End If
                            
                            Next
                            
                            
                            ' SENDER CONDITIONS
                            
                            ' Check the sender email must contain conditon
                            ' must contain
                            If fileCurrentMail And emailSenderMustContain <> "" And Not LCase(currentMailSenderEmail) Like "*" & LCase(emailSenderMustContain) & "*" Then
                                ' Does not meet the condition
                                fileCurrentMail = False
                            End If
                            ' must not contain
                            If fileCurrentMail And emailSenderMustNotContain <> "" And LCase(currentMailSenderEmail) Like "*" & LCase(emailSenderMustNotContain) & "*" Then
                                ' Does not meet the condition
                                fileCurrentMail = False
                            End If
                            
                            
                            ' AT THIS POINT WE KNOW WHETHER THE FILE IS TO BE FILED OR NOT
                            
                            ' File the current mail if it meets the conditions
                            If fileCurrentMail Then
                            
                                ' Only file to user folder if folder entered
                                If fileToUserFolder Then
                                    ' Create a copy
                                    Set cMailCopy = oMail.Copy
                                    ' Move the copy to the user's folder
                                    cMailCopy.Move projectUserFolder
                                End If
                                
                                ' Only file to newforma folder if folder entered
                                If fileToNewformaFolder Then
                                    ' Create a new copy
                                    Set cMailCopy = oMail.Copy
                                    ' Move the copy to newforma
                                    cMailCopy.Move projectNewformaFolder
                                End If
                                
                                ' Delete the original (if it's been moved to one of the folders)
                                If fileToUserFolder Or fileToNewformaFolder Then
                                
                                    oMail.Delete
                                    
                                    ' Increment the counter & update the progress
                                    numberEmailsFiled = numberEmailsFiled + 1                                                       ' Overall count
                                    numberEmailsFiledForCurrentJob = numberEmailsFiledForCurrentJob + 1                             ' Count for this job
                                    numberOfEmailsFiledSincelastProgressUpdate = numberOfEmailsFiledSincelastProgressUpdate + 1     ' Count since last progress update
                                    
                                    If numberOfEmailsFiledSincelastProgressUpdate >= reportInterval Then
                                    
                                        ' Output the progress
                                        progress.TextFrame.Characters.Text = "Current Job: " & projectName & vbNewLine & "Number of emails filed: " & numberEmailsFiledForCurrentJob & vbNewLine & vbNewLine & "(You can still use Outlook, just don't delete anything!)"
                                        forceScreenToUpdateAndPause ("00:00:01")
                                        
                                        ' Reset the count
                                        numberOfEmailsFiledSincelastProgressUpdate = 0
                                        
                                    End If
                                
                                
                                End If
                                
                                ' Don't check anymore rules as we've already filed it
                                Exit For
                            
                            End If
                            
                        ' next rule
                        Next ruleRow
                    
                    End If
                
                
                End If
                
                                  
            ' Next mail item in the current Items
            Next i
        
        End If
    
    ' next Items
    Next currentItemsIndex
    
    ' Report total number of emails filed for this job
    progress.TextFrame.Characters.Text = "Current Job: " & projectName & vbNewLine & "Number of emails filed: " & numberEmailsFiledForCurrentJob & vbNewLine & vbNewLine & "(You can still use Outlook, just don't delete anything!)"
    forceScreenToUpdateAndPause ("00:00:01")
    
               
    ' clean up
    Set currentItems = Nothing
    Set itemsArray(0) = Nothing
    Set itemsArray(1) = Nothing

End Sub


Private Sub forceScreenToUpdateAndPause(wait As String)

    ' This sub forces the sheet to update
    ' It should be called AFTER the change has been made
    
    ' NB: the string must be in the format "00:00:02"
    ' THIS IS NOT CHECKED!
    ' The value is always hardcoded so it won't be an error unless we mess up

    Excel.Application.ScreenUpdating = True
    DoEvents    ' This is the line that makes all the difference!
    Application.wait (Now + TimeValue(wait))
    Excel.Application.ScreenUpdating = False

End Sub

Private Sub testWildcard()
' THIS IS A MANUAL TEST ONLY
    ' Using Like instead of InStr() allows the use of wildcards...
    ' Use the string to compare first
    ' Followed by the string to search for, wrapped in wildcards...
    If Not "Ennistymon - Primary" Like "*Ennistymon*" Then
        MsgBox ("true")
    End If
    
    If LCase("Ennistymon - Primary") Like LCase("*ENNISTYMON*") Then
        MsgBox ("true")
    End If
    
End Sub

Function checkIfMailFolderExistsInFolder(folderName As String, parentFolder As Outlook.MAPIFolder) As Boolean

    Dim oFolder As Outlook.MAPIFolder               ' a mail folder
    
    ' Initially false
    folderFound = False
    Set currentFoundFolder = Nothing
        
    ' Don't check Deleted Items Folder
    If parentFolder.Name <> "Deleted Items" Then
    
        ' Cycle through each folder
        For Each oFolder In parentFolder.Folders
            
            ' If the folder name matches the one searched for, exit
            If oFolder.Name = folderName Then
                folderFound = True
                Set currentFoundFolder = oFolder
                Exit For
            End If
        
            ' If we haven't yet found the folder
            ' Check the sub folders with a recursive call
            If Not folderFound And oFolder.Folders.Count > 0 Then
                folderFound = checkIfMailFolderExistsInFolder(folderName, oFolder)
            End If
            
            ' If found in the sub folders, then exit
            If folderFound Then Exit For
        
        Next oFolder
    
    End If
    
    ' Return the answer
    checkIfMailFolderExistsInFolder = folderFound
    
End Function
