On Error Resume Next
Set objFS = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

Dim KRYPTONHome
KRYPTONHome = objFS.GetParentFolderName(WScript.ScriptFullName)

Dim Controller_Log
Controller_Log= KRYPTONHome & "\kryptonexecution.log"

Const blnUseGridEmailFeatures = True
If blnUseGridEmailFeatures Then
    sendKRYPTONEmail= "no"
Else
    sendKRYPTONEmail= "yes"
End If

Dim emailRecipients
emailRecipients= ""

'Email constants
Const fromEmail	= "krypton.thinksys@gmail.com"
Const password	= "Thinksys@2345"
Dim emailSubject: emailSubject = ""

Const maxTestExecutionTime= 1800		'in seconds
Const HIDDEN_WINDOW = 1
Const FOR_READING = 1
Const FOR_WRITING = 2
Const Quote= """"

Dim absoluteHostNodes
Dim processId
Dim testResultLocation
Dim logFileName
Dim strTestSuiteName
Dim strTestEnvironment
'Dim strNotification

'First thing first, check if enough arguments have been passed
Set objArgs = WScript.Arguments
For Each argument In objArgs
    
    If InStr(1,LCase(argument),"hostlist=")>0 Then
        absoluteHostNodes= Split(argument, "=")(1)
    End If
    
    If InStr(1,LCase(argument),"pid=")>0 Then
        processId= Split(argument, "=")(1)
    End If
    
    If InStr(1,LCase(argument),"loglocation=")>0 Then
        testResultLocation= Split(argument, "=")(1)
    End If
    
    If InStr(1,LCase(argument),"logxmlname=")>0 Then
        logFileName= Split(argument, "=")(1)
    End If
    
    If InStr(1,LCase(argument),"testsuite=")>0 Then
        strTestSuiteName= Split(argument, "=")(1)
    End If
    
    If InStr(1,LCase(argument),"env=")>0 Then
        strTestEnvironment= Split(argument, "=")(1)
    End If
    
Next

'Wait until execution for current suite is over on all machines
WScript.Echo "Checking execution status on hosts " & strHostList
Call WaitForTestSuiteExecutionToComplete(absoluteHostNodes, "Krypton.exe", processId, maxTestExecutionTime)

'Initiate generation of test case report and send email
    'Email and Report for atteching will create only if 
    'EmailNotification is true in parameter ini.
strNotification=GetEmailNotification()
 If Trim(LCase(strNotification)) = LCase("TRUE") Then  
  WScript.Echo "Generating HTML Report from " + testResultLocation
  Call GenerateHtmlReport(testResultLocation, processId, logFileName, strTestSuiteName, strTestEnvironment)
 End if
Public Function WaitForTestSuiteExecutionToComplete(strHostList, strProcessName, processId, maxTimeOut)
    
    'Decode list of hosts into an array
    Dim arrHostList
    Dim blnExecutionComplete
    blnExecutionComplete = True
    
    On Error Resume Next
    strHostList= Replace(strHostList, " ", "")
    
    'Add localhost to hosts list, this ensures that no local process is running with same process id
    strHostList = strHostList & ",."
    
    arrHostList= Split(strHostList, ",")
    intTimer= 1
    
    'Start a loop for arrays upper bound
    Do
        blnExecutionComplete= True
        For i=0 To UBound(arrHostList)
                    
            'Retrieve instances of wexectrl.exe running on remote box
            Set objHMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
            arrHostList(i) & "\root\cimv2")
            
            Set colHProcesses = objHMIService.ExecQuery ("SELECT * FROM Win32_Process WHERE Name = '" & _
            strProcessName & "'")

            If Not IsEmpty(colHProcesses) Then
                If colHProcesses.Count> 0 Then
                    For Each objHProcess In colHProcesses
                        If InStr(1,LCase(objHProcess.CommandLine), LCase(processId))> 0 Then
                            blnExecutionComplete= False
                            WScript.Echo "Elapsed timeout " & intTimer                   
                            WScript.Echo "Test execution still running on " & arrHostList(i)
                        End If
                    Next
                End If
            End If
        Next
        
        'Exit loop if you found an available host
        If blnExecutionComplete Or intTimer > maxTimeOut Then
            Exit Do
        End If
        WScript.Sleep 1000
        intTimer= intTimer+1
    Loop
End Function


Public Function GenerateHtmlReport(resultsLocation, processId, xmlFileName, testSuiteName, testEnvironment)
    Dim listXmls
    listXmls = ""
    On Error Resume Next
    
    'Create a temp txt file temp folder of the user, overwriting if there is anyone already
    txtFileForXMLEntries = KRYPTONHome & "\KryptonXMLLists_" & processId & ".txt"
    objFS.CreateTextFile txtFileForXMLEntries, True
    
    'Retrive collection of all xml log files
    Set objHTMLFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objHTMLFSO.GetFolder(resultsLocation)
    Set colSubfolders = objFolder.SubFolders
    
    For Each objSubfolder In colSubfolders
        If InStr(objSubfolder.Name, processId)> 0 Then
            xmlPath = objSubfolder.Path & "\1\" & xmlFileName
            Call AppendLineToFile(txtFileForXMLEntries, xmlPath)
            listXmls = listXmls & "," & xmlPath
        End If
    Next
    
    listXmls = Right(listXmls, Len(listXmls)-1)
    
    'Generate html file combining all xml files
    htmlSummaryLocation = testSuiteName & "_" & processId
    htmlDetailedLocation = testSuiteName & "_" & processId
    
    ReportGenerationCommand= Quote & KRYPTONHome & "\KryptonCreateReport.exe" & Quote & _
    " " & Quote & "xmlfiles=" & txtFileForXMLEntries & Quote & _
    " " & Quote & "htmlfilename=" & htmlSummaryLocation & Quote & _
    " " & Quote & "emailFor=end" & Quote & _
    " " & Quote & "environment=" & testEnvironment & Quote & _
    " " & Quote & "testsuite=" & testSuiteName & Quote & _
    " " & Quote & "emailrequired=" & sendKRYPTONEmail & Quote
    
    Call AppendLineToFile(Controller_Log, ReportGenerationCommand)
    
    WshShell.Run ReportGenerationCommand, 0, True
    
    'Extract exact path of html files generated
    For Each htmlFile In objFolder.Files 
        If InStr(1, LCase(Trim(htmlFile.Name)), LCase(Trim(htmlSummaryLocation))) Then
            htmlDetailedLocation = htmlFile.Path
            htmlSummaryLocation = Replace(htmlDetailedLocation, ".html", "s.html")
            Exit For
        End If
    Next
    
    WScript.Echo "HTML Report generation completed"
    
    WScript.Echo "Generating summary report now..."
    summaryContents= GenerateSummaryReport(htmlSummaryLocation, htmlDetailedLocation)
    WScript.Echo "Summary Report generation completed"
    
    If blnUseGridEmailFeatures Then
    	WScript.Echo "Sending grid specific email..."
         htmlDetailedLocation = Replace(htmlDetailedLocation, ".html", "smail.html")
        Call SendFinishedExecutionEmail(testSuiteName, summaryContents, htmlDetailedLocation)
        WScript.Echo "Email sent complete."
    End If
    
    Set objHTMLFSO= Nothing 


    Set obj = CreateObject("Scripting.FileSystemObject") 'Calls the File System Object  
    obj.DeleteFile(htmlDetailedLocation)
    obj.DeleteFile(htmlSummaryLocation)
    Set obj = Nothing
End Function


Public Function GenerateSummaryReport(htmlSummaryLocation, htmlDetailedLocation)
    On Error Resume Next
    Set htmlFile = objFS.OpenTextFile(htmlSummaryLocation)
    entireHTMLContents= htmlFile.ReadAll
    htmlFile.Close
    startLocation = 0

    'Remove detailed steps from the report
    i= 0
    Do 
    	i= i+1
        startLocation = InStr(1, entireHTMLContents, "<div id='tblID")
        If startLocation <= 0 Then
            Exit Do
        End If
        endLocation = InStr(startLocation, entireHTMLContents, "</div>") +6
        text_to_replace = Mid(entireHTMLContents, startLocation, endLocation-startLocation)
        entireHTMLContents= Replace(entireHTMLContents, text_to_replace, "")
        entireHTMLContents= Replace(entireHTMLContents, "style='table-layout:fixed'", "")
    	If i > 500 Then
    		Exit Do
    	End If
    Loop
    
    'Add serial number to the summary report instead of expand link
    intCounter = 1
    Do 
        expandLocation= InStr(1, entireHTMLContents, ">[+]<")
        If expandLocation <= 0 Then
            Exit Do
        End If
        entireHTMLContents = Replace(entireHTMLContents, ">[+]<", ">" & intCounter & "<",1,1)
        intCounter = intCounter + 1
        If intCounter > 500 Then
    		Exit Do
    	End If
    Loop
    
    'Remove all links
    entireHTMLContents = Replace(entireHTMLContents, "<a>", "")
    entireHTMLContents = Replace(entireHTMLContents, "<A>", "")
    entireHTMLContents = Replace(entireHTMLContents, "<\a>", "")
    entireHTMLContents = Replace(entireHTMLContents, "<\A>", "")

    'Extract total  test case count  
    totalExecuted = GetTotalCount("Test Case(s) Executed:",entireHTMLContents)
	
    'Extract passed test case count   
    totalPassed = GetTotalCount("Total Passed:",entireHTMLContents) 
	
    'Extract failed test case count
    totalFailed = GetTotalCount("Total Failed:",entireHTMLContents)
	
    'Extract warning test case count
    totalWarning = GetTotalCount("Total Warning:",entireHTMLContents)
	
    'Update pass count to be sum of pass and warning
    totalPassed = totalPassed + totalWarning

    'Create email subject based on numbers
    If CInt(totalFailed) > 0 Then
        emailSubject = "[" & totalFailed & "/" & totalExecuted & "] Failed! '$testsuite$' test suite has failed"
    Else
        emailSubject = "[" & totalExecuted & "] Passed! '$testsuite$' test suite execution completed."
    End If
    
    GenerateSummaryReport= entireHTMLContents
End Function
'Get Total Count
 public Function GetTotalCount(strToBeSearched,entireHTMLContents)
    Dim totalMathcFound
	totalMathcFound = strToBeSearched
    startLocation = InStr(1, entireHTMLContents, totalMathcFound)
    startLocation = InStr(startLocation, entireHTMLContents, ">") +1
    endLocation = InStr(startLocation, entireHTMLContents, "<")

    totalMathcFound = Mid(entireHTMLContents, startLocation, endLocation-startLocation)
    totalMathcFound = CInt(totalMathcFound)
	GetTotalCount = totalMathcFound
  End Function


Public Function SendFinishedExecutionEmail(testSuiteName, htmlSummaryFile, htmlDetailedFile)
    On Error Resume Next
    
    Set objMessage = CreateObject("CDO.Message")
    objMessage.Sender = fromEmail    
    EmailRecipient= GetEmailRecipients()
	
    objMessage.To       = EmailRecipient
    
    If emailSubject = "" Then
        emailSubject = "Test Suite: " & testSuiteName & " finished."
    Else
        emailSubject = Replace(emailSubject, "$testsuite$", testSuiteName)
    End If

    
    objMessage.Subject = emailSubject
    objMessage.From    = "Krypton Automation" 
    
    'Set email configurations
    Set emailConfig = objMessage.Configuration
    emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver")       = "smtp.gmail.com"
    emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")   = 465
    emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")        = 2
    emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl")       = True
    emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername")     = fromEmail
    emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword")     = password
    emailConfig.Fields.Update
    
    'The line below shows how to send using HTML included directly in your script
    objMessage.HTMLBody = htmlSummaryFile
    
    'The line below shows how to send a webpage from a remote site
    'objMessage.CreateMHTMLBody "http://www.paulsadowski.com/wsh/"
    
    'The line below shows how to send a webpage from a file on your machine
    
    objMessage.AddAttachment htmlDetailedFile
    
    
    objMessage.Send
End Function

Public Function GetEmailNotification()
    Dim emailnotificatio
    emailnotificatio= ""
    
   ' On Error Resume Next

   Set objFS1 = CreateObject("Scripting.FileSystemObject")
   
    BaseLocation = GetProjectFolderPath()
    Set ParametersFile = objFS.OpenTextFile(BaseLocation & "\Parameters.ini")
    arrnotifications = Split(ParametersFile.ReadAll, vbNewLine)
   
    
    For Each notification In arrnotifications
        suiteName = Split(notification, ":")(0)
        
        EmailNotificationValue = Split(notification, ":")(1)

         If Trim(LCase(suiteName)) = LCase("EmailNotification") Then
            emailnotificatio=EmailNotificationValue
            GetEmailNotification=emailnotificatio
         Exit Function
         End if
                 
    Next
    GetEmailNotification= emailnotificatio
End Function

Public Function GetEmailRecipients()
    Dim fileText, AllRecipients
    AllRecipients=""
    ProjectPath=GetProjectFolderPath()
    Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(ProjectPath & "\EmailNotification.ini",1)
    fileText = objFileToRead.ReadAll()   
    objFileToRead.Close
    Set objFileToRead = Nothing

    Dim startIndex, length
    startIndex = InStr(fileText, ":")
    length = Len(fileText)
 
    AllRecipients = Mid(fileText, startIndex +1 , length - startIndex + 1)           
    GetEmailRecipients= AllRecipients
    End Function

	
Public Function AppendLineToFile(strFilePath, strLineContents)
    Const FOR_APPENDING = 8
    On Error Resume Next
    Set objAFS = CreateObject("Scripting.FileSystemObject")
    
    If Not objAFS.FileExists(strFilePath) Then
        objAFS.CreateTextFile strFilePath, True
    End If
    
    ' Writing String Content to End of Existing Text File
    Set objATS = objAFS.OpenTextFile(strFilePath, FOR_APPENDING)
    objATS.WriteLine strLineContents
    objATS.Close
    
    Set objATS= Nothing
    Set objAFS= Nothing
End Function


Public Function GetProjectFolderPath()
    Set objFS = CreateObject("Scripting.FileSystemObject")
    InstallationPath= objFS.GetParentFolderName(WScript.ScriptFullName)
    
    Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(InstallationPath & "\root.ini",1)
    Dim ProjectPath
    do while not objFileToRead.AtEndOfStream
         Dim tempProjectPath 
         tempProjectPath = objFileToRead.ReadLine()
         if (InStr(tempProjectPath, "ProjectPath") > 0) Then
            ProjectPath = tempProjectPath
         End If
         'Do something with the line
    loop
    objFileToRead.Close
    Set objFileToRead = Nothing

    Dim startIndex, length
    startIndex = InStr(ProjectPath, ":")
    length = Len(ProjectPath)

    Dim ActualProjectPath, RelativePath
    ActualProjectPath = Mid(ProjectPath, startIndex +1 , length - startIndex + 1)
    ActualProjectPath = Trim(ActualProjectPath)
    RelativePath = objFS.GetAbsolutePathName(ActualProjectPath)

    GetProjectFolderPath = RelativePath
End Function
