Rem Start Date: Sep 02, 2010
Rem Business: To be able to start automated execution on remote machines from one central location
Rem Purpose: To read execution queue text file in C:\QA_Automation folder and execute each statement one by one
'MsgBox "Hey"
'See if there is anything in the queue first
Set objFS = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

Const HIDDEN_WINDOW = 1    '1 - Window is shown minimized
'Const HIDDEN_WINDOW = 3    '3 - Window is shown maximized
'Const HIDDEN_WINDOW = 5    '5 - Window is shown in normal view
'Const HIDDEN_WINDOW = 12   '12 - Window is hidden and not displayed to the user

Const blnWaitForReportGeneration= False

Dim strComputer
Dim objWMIService
Dim colProcesses
Dim objFS
Dim intProcessedQueueLength
Dim arrExecutionCommands
Dim strExecutionCommand
Dim iNumberOfLinesToDelete
Dim arrHostNodes
Dim strFreeAvailableHost
Dim testResultLocation: testResultLocation= ""
Dim uniqueTestResultLocation: uniqueTestResultLocation= ""
Dim strTestEnvironment: strTestEnvironment= ""
Dim strHostNodes
strHostNodes = ""

Dim HomeLocation 
HomeLocation = objFS.GetParentFolderName(WScript.ScriptFullName)
xStudioIniFile = HomeLocation & "\xStudio.ini"

Dim blnIsLocalExecution: blnIsLocalExecution = False
Dim blnUseGridEmailFeatures: blnUseGridEmailFeatures= True
Dim blnStartOfTestSuite: blnStartOfTestSuite= True				'Indicates if test suite execution is just starting now

If blnUseGridEmailFeatures Then
    sendKRYPTONEmail= "no"
Else
    sendKRYPTONEmail= "yes"
End If

Dim emailRecipients
emailRecipients= ""

'Define maximum time that a test case will take when executing, this serves as a timeout when waiting for execution to finish
Const maxTestExecutionTime= 1800		'in seconds

'Email constants
Const fromEmail	= "krypton.thinksys@gmail.com"
Const password	= "Thinksys@123"
Dim emailSubject
emailSubject = ""

'Stores count of existing processes before launching new instance of Krypton
Dim intExistingProcessCount: intExistingProcessCount= 0

Dim processId: processId = ""
Dim lastKnownCommand


Const FOR_READING = 1
Const FOR_WRITING = 2
Const Quote= """"
Const MaxFFSessions= 1
Dim KRYPTONHome
Dim strTestSuiteName
Const TIMEOUT= 10

On Error Resume Next


Set objArgs= WScript.Arguments

strUserName = ""
strPassword = ""

'First thing first, check if enough arguments have been passed
For Each argument In objArgs
    
    'Msgbox argument
    If InStr(1,LCase(argument),"queue")>0 Then
        strQueueFile= Split(argument, "=")(1)
    End If
    
    If InStr(1,LCase(argument),"kryptonhome")>0 Then
        KRYPTONHome= Split(argument, "=")(1)
    End If
    
    If InStr(1,LCase(argument),"nodes")>0 Then
        strHostNodes= strHostNodes & Trim(Split(argument, "=")(1)) & ","
    End If
    
    If InStr(1,LCase(argument),"username")>0 Then
        strUserName= Split(argument, "=")(1)
    End If
    
    If InStr(1,LCase(argument),"password")>0 Then
        strPassword= Split(argument, "=")(1)
    End If
    
Next

If LCase(strHostNodes) = "localhost" Or LCase(strHostNodes) = "localhost," Then
    strHostNodes = "localhost[W]"
    blnIsLocalExecution = True
End If

'MsgBox blnIsLocalExecution

Dim UNPROCESSED_QUEUE
UNPROCESSED_QUEUE= KRYPTONHome & "\ExecutionQueue_NotProcessed.txt"

Dim Controller_Log
Controller_Log= KRYPTONHome & "\kryptonexecution.log"

If objArgs.Count>=2 Then
    arrHostNodes= Split(strHostNodes,",")
Else
    WScript.Echo "Not enough arguments passed, exiting..."
    WScript.Quit
End If

intProcessedQueueLength= 0

If Not objFS.FileExists(strQueueFile) Then
    WScript.Quit
End If

'Start reading the queue, line by line, execute a line and remove from queue
strTestSuiteName = ""
processId = Day(Now()) & Month(Now()) & Right(Year(Now()),2) & Hour(Now()) & _
Minute(Now()) & Second(Now())

'This parameter tell if the execution we are going to launch is on windows os or some other one
Dim blnIsWindows
Dim blnIsSaucelabs
Dim platformType

Dim totalTestsLaunched
totalTestsLaunched = 0

Do 
    iNumberOfLinesToDelete = 0
    
    Set objTS = objFS.OpenTextFile(strQueueFile, FOR_READING)
    If objTS.AtEndOfStream Then
        Exit Do
    End If
    
    strExecutionCommand = objTS.ReadLine
    WScript.Echo "Command Retrieved: " & strExecutionCommand
    
    If Trim(strExecutionCommand)="" Then
        iNumberOfLinesToDelete= iNumberOfLinesToDelete+1
        
    ElseIf InStr(strExecutionCommand, "[TEST_SUITE")>0 Then
        blnStartOfTestSuite = True
        xStudioNextAgentFile= ""
        emailSubject= ""
        
        strTestSuiteName = Replace(strExecutionCommand, "[TEST_SUITE:", "")
        strTestSuiteName = Replace(strTestSuiteName, "]", "")
        strTestSuiteName= Trim(strTestSuiteName)
        
        iNumberOfLinesToDelete= iNumberOfLinesToDelete+1
        processId = Lpad(CStr(Day(Now())),"0",2) & Lpad(CStr(Month(Now())),"0",2) & CStr(Year(Now())) & Lpad(CStr(Hour(Now())),"0",2) & _
        Lpad(CStr(Minute(Now())),"0",2) & Lpad(CStr(Second(Now())),"0",2)
        
        'Send email notifying start of execution
        Call SentStartExecutionEmail(strTestSuiteName)
        
    ElseIf Trim(strExecutionCommand)="[FINISH]" Then
        
        If objFS.FileExists(xStudioNextAgentFile) Then
            objFS.DeleteFile xStudioNextAgentFile, True
        End If
        
        iNumberOfLinesToDelete= iNumberOfLinesToDelete+1
        
        'Wait for execution to be finished
        WScript.Echo "Waiting for execution to finish."
        
        'To allow files to be copied to network location, must be removed later on
        
        
        absoluteHostNodes = Replace(strHostNodes, "[W]", "")
        absoluteHostNodes = Replace(absoluteHostNodes, "[L]", "")
        absoluteHostNodes = Replace(absoluteHostNodes, "[M]", "")
        absoluteHostNodes = Replace(absoluteHostNodes, "[U]", "")
        
        'Extract test result location from last executed command
        testResultLocation =Split(lastKnownCommand, "logdestinationfolder=")(1)
        testResultLocation =Split(testResultLocation, Quote)(0)
        
        'Extract test environment from last executed command
        strTestEnvironment =Split(lastKnownCommand, "environment=")(1)
        strTestEnvironment =Split(strTestEnvironment, Quote)(0)
        
        'Extract xml log file name from last executed command
        logFileName =Split(lastKnownCommand, "logfilename=")(1)
        logFileName =Split(logFileName, Quote)(0)
        
        'Wait until execution for current suite is over on all machines
        
        strReportGenerationScript=KRYPTONHome & "\GenerateHTMLandSendEmail.vbs"
        strReportGenerationScript = "cscript " & Quote & strReportGenerationScript & Quote & _
        " " & Quote & "hostlist=" & Trim(absoluteHostNodes) & Quote & _
        " " & Quote & "pid=" & processId & Quote & _
        " " & Quote & "loglocation=" & testResultLocation & Quote & _
        " " & Quote & "logxmlname=" & logFileName & Quote & _
        " " & Quote & "testsuite=" & strTestSuiteName & Quote & _
        " " & Quote & "env=" & strTestEnvironment & Quote
        
        Call AppendLineToFile(Controller_Log, strReportGenerationScript)
        
        WshShell.Run strReportGenerationScript, 0, blnWaitForReportGeneration
        
    Else
        'Extract test result location
        testResultLocation =Split(strExecutionCommand, "logdestinationfolder=")(1)
        testResultLocation =Split(testResultLocation, Quote)(0)
        
        
        
        'Extract test manager information
        testManagerType =Split(strExecutionCommand, "managertype=")(1)
        testManagerType =Trim(Split(testManagerType, Quote)(0))
        WScript.Echo testManagerType
        
        'Extract test suite id information
        testSuiteId =Split(strExecutionCommand, "testsuiteid=")(1)
        testSuiteId =Split(testSuiteId, Quote)(0)
        
        'Extract browser information from the command
        browser =Split(strExecutionCommand, "browser=")(1)
        browser =Split(browser, Quote)(0)
        
        'Update process id to inclued browser information
        If blnStartOfTestSuite Then
            processId = CStr(processId) & "[" & browser & "]"
        End If
        
        'Handle conditional launching in case of xStudio is test manager
        If LCase(testManagerType) = "xstudio" Then
            
            'Create a new folder for result destination if starting test suite execution
            If blnStartOfTestSuite Then
                uniqueTestResultLocation = testResultLocation & "\" & strTestSuiteName & "_" & CStr(processId)
                objFS.CreateFolder(uniqueTestResultLocation)
                WScript.Echo uniqueTestResultLocation
                
                'Retrieve information from xStudio ini file
                xStudioHome = GetParameterFromIniFile(xStudioIniFile, "xStudioLocation", HomeLocation)
                xStudioNextAgentExe = GetParameterFromIniFile(xStudioIniFile, "xStudioNextAgentApp", HomeLocation)				
                xStudioTestAgentLocation = GetParameterFromIniFile(xStudioIniFile, "xStudioAgentApp", HomeLocation)
                xStudioIntegrationLocation = GetParameterFromIniFile(xStudioIniFile, "xStudioIntegrationApp", HomeLocation)
                xStudioNextAgentFile=GetParameterFromIniFile(xStudioIniFile, "xStudioNextAgentFile", "")
                
                'Following file contains id of test agent to be used
                'It will be create when we'll call getNextAgent.exe available in bin folder of xStudio
                
                xStudioNextAgentFile = HomeLocation & "\" & processId & ".txt"
                
                xStudioServerAgentId= GetParameterFromIniFile(xStudioIniFile, "xStudioDefaultTestAgentId", HomeLocation)
                xStudioSUTId = GetParameterFromIniFile(xStudioIniFile, "xStudioSUTId", HomeLocation)
                xStudioConfigurationId = GetParameterFromIniFile(xStudioIniFile, "xStudioConfigurationId", HomeLocation)
                xStudioCategoryId = GetParameterFromIniFile(xStudioIniFile, "xStudioCategoryId", HomeLocation)
                
                WScript.echo xStudioHome
                WScript.echo xStudioNextAgentExe
                WScript.echo xStudioTestAgentLocation
                WScript.echo xStudioIntegrationLocation
                WScript.echo xStudioNextAgentFile
                WScript.echo xStudioSUTId
                WScript.echo xStudioConfigurationId
                WScript.echo xStudioCategoryId
                
                'Create envrionment variable for result destination folder
                Call setSystemEnvVariable("xStudioResultDestination", uniqueTestResultLocation)
                
                Call setSystemEnvVariable("xstudio_home", xStudioHome)
                Call setSystemEnvVariable("krypton_home", HomeLocation)
                command= Quote & xStudioNextAgentExe & Quote & " " & processId & " " & Quote & uniqueTestResultLocation & Quote
                AppendLineToFile Controller_Log, command
                WshShell.Run command,0,True
                
                'This next agent application will create a txt file containing id of test agent to be used, wait for it
                For i=0 To 20
                    If objFS.FileExists(xStudioNextAgentFile) Then
                        Set AgentIdFile = objFS.OpenTextFile(xStudioNextAgentFile)
                        xStudioServerAgentId= AgentIdFile.ReadLine
                        AgentIdFile.Close
                        Exit For
                    Else
                        WScript.Sleep 1000
                    End If
                Next
                
                
                xStudioLaunchCommand = Quote & xStudioIntegrationLocation & Quote &_
                " --campaignId " & testSuiteId &_
                " --agents " & xStudioServerAgentId & ":1:2" &_
                " --sutId " & xStudioSUTId &_
                " --configurations " & xStudioCategoryId & ":" & xStudioConfigurationId &_
                " --sessionName " & processId &_
                " --monitoringAgentId " & xStudioServerAgentId
                
                Err.Clear
                AppendLineToFile Controller_Log, xStudioLaunchCommand
                WshShell.Exec xStudioLaunchCommand
                WScript.Echo Err.Description
            End If
            'blnStartOfTestSuite= False
            
            'Update test result location in execution command
            strExecutionCommand = Replace(strExecutionCommand, "logdestinationfolder=" & testResultLocation, "logdestinationfolder=" & uniqueTestResultLocation)
            
        Else
            uniqueTestResultLocation= testResultLocation
        End If
        
        blnStartOfTestSuite= False
        
        'For the rest, assume that it is a statement to execute on one the available box
        'Get available box
        WScript.Echo "Checking for freely available machine out of: " & strHostNodes
        
        strFreeAvailableHost= GetAvailableHost(strHostNodes, "Krypton.exe", browser, strExecutionCommand)
        
        'Push current available host to last of the queue
        arrHostNodes = Split(strHostNodes, ",")
        For Each hostnode In arrHostNodes
            If InStr(1, LCase(hostnode), LCase(strFreeAvailableHost)) > 0 Then
                strHostNodes = Replace(strHostNodes, hostnode, "")
                strHostNodes = strHostNodes & "," & hostnode
                strHostNodes = Replace(strHostNodes, ",,", ",")
            End If
        Next
        lastFreeHost = strFreeAvailableHost
        
        WScript.Echo "Host:" & strFreeAvailableHost & " was found to be free out of " & strHostNodes
        
        'Append machine id and process id to Krypton commands
        strExecutionCommand = strExecutionCommand & " " & Quote & "RCProcessId=" & processId & Quote
        
        If LCase(strFreeAvailableHost) = "localhost" Then
            strExecutionCommand = strExecutionCommand & " " & Quote & "rcmachineid=" & WshShell.ExpandEnvironmentStrings("%COMPUTERNAME%") & Quote
        Else
            If InStr(1, strFreeAvailableHost, "saucelabs") > 1 Then
                
				strExecutionCommand = strExecutionCommand & " " & Quote & "rcmachineid=" & strFreeAvailableHost & Quote
            Else
                strExecutionCommand = strExecutionCommand & " " & Quote & "rcmachineid=" & strFreeAvailableHost & Quote
            End If
        End If
        
        'Append platform type information
        strExecutionCommand = strExecutionCommand & " " & Quote & "platform=" & platformType & Quote
        
        'Remove extra information from host name
        strFreeAvailableHost = Split(strFreeAvailableHost, "[")(0)
        
        'Add remote host parameters if execution need to be done on a remote machine
        If Not blnIsWindows And Not blnIsSaucelabs Then
            strExecutionCommand = strExecutionCommand & " " & Quote & "runremoteexecution=true" & Quote
            strExecutionCommand = strExecutionCommand & " " & Quote & "runonremotebrowserurl=http://" & strFreeAvailableHost & ":4444" & Quote
            blnIsLocalExecution = True
        ElseIf blnIsSaucelabs Then
            strExecutionCommand = strExecutionCommand & " " & Quote & "runremoteexecution=true" & Quote
            strExecutionCommand = strExecutionCommand & " " & Quote & "runonremotebrowserurl=http://" & strFreeAvailableHost & ":80" & Quote
            blnIsLocalExecution = True
        Else
            strExecutionCommand = strExecutionCommand & " " & Quote & "runremoteexecution=false" & Quote
            
        End If
        
        
         strExecutionCommand = strExecutionCommand & " " & Quote & "emailnotification=false" & Quote  
        lastKnownCommand= strExecutionCommand
        
        
        
        If blnIsLocalExecution Then
            ExecutionString = strExecutionCommand
        Else
            ExecutionString = QUOTE & KRYPTONHome & "\psexec.exe" & QUOTE & _
            " \\" & strFreeAvailableHost & _
            " -u " & Quote & strUserName & Quote & " -p " & Quote & strPassword & Quote & _
            " -accepteula -i " & GetUserSessionId(strFreeAvailableHost) & " -d -n " & TIMEOUT & " " & strExecutionCommand
        End If
        
        WScript.Echo "Launching test execution on " & strFreeAvailableHost & ". Following statement will be executed now: " & vbNewLine & ExecutionString
        
        AppendLineToFile Controller_Log, ExecutionString
        
        WshShell.Run ExecutionString, 0, False		'hidden window
        
        
        totalTestsLaunched = totalTestsLaunched +1
        WScript.Sleep 1000
        
        'Once the execution has been launched, wait for it to start executing and then delete from execution queue.
        'If it appears that execution was not started, but rather, save to unprocessed execution queue
        If blnIsLocalExecution Then
            ExecutionStartedOnHost = "."
        Else
            ExecutionStartedOnHost = strFreeAvailableHost
        End If
        
        If Not totalTestsLaunched < UBound(Split(strHostNodes, ",")) Then
            If Not blnIsExecutionStarted(ExecutionStartedOnHost, "Krypton.exe", strExecutionCommand, 20) Then
                AppendLineToFile UNPROCESSED_QUEUE, Now() & vbTab & strExecutionCommand
            End If
        End If
        
        iNumberOfLinesToDelete= iNumberOfLinesToDelete+1
    End If
    
    'Close existing file
    objTS.Close
    Set objTS =Nothing
    
    'Open file in write mode and Delete statements from execution queue
    'Delete executed command from queue
    Set objTS = objFS.OpenTextFile(strQueueFile, FOR_READING)
    strContents = objTS.ReadAll
    objTS.Close
    
    arrLines = Split(strContents, vbNewLine)
    Set objTS = objFS.OpenTextFile(strQueueFile, FOR_WRITING)
    For i=0 To UBound(arrLines)
        If i > (iNumberOfLinesToDelete - 1) And Trim(arrLines(i))<> "" Then
            objTS.WriteLine arrLines(i)
        End If
    Next
    
    objTS.Close
    Set objTS= Nothing
Loop


Public Function GetAvailableHost(strHostList, strProcess, ByRef browser, ByRef strExecutionCommand)
    
    'Host availability means the following
    '1. Host is online and user is logged in (session id is more than one)
    '2. Test execution process (e.g. Krypton.exe) is not already running (linked with browser)
    '3. Specific browser process is not already running (coupled with browser argument)
    
    'Decode list of hosts into an array
    strHostList= Replace(strHostList, " ", "")
    Dim arrHostList: 		arrHostList= Split(strHostList, ",")
    Dim blnAvailableHost: 	blnAvailableHost= False
    Dim strAvailableHost: 	strAvailableHost= ""
    
    Dim currentSauceLabsInstances
    currentSauceLabsInstances= 0
    
    On Error Resume Next
    
    'Start a loop for arrays upper bound
    Do
        
        'Out of the list available, check which machine is completely free
        For i=0 To UBound(arrHostList)
            
            strHostName = arrHostList(i)
            WScript.Echo "Checking availability of " & strHostName
            
            'Check if host is windows or otherwise
            If InStr(1, UCase(strHostName), "[W]") >0 Then
                platformType = "Windows"
                blnIsWindows = True
            ElseIf InStr(1, UCase(strHostName), "[M]") >0 Then
                platformType = "Mac"
                blnIsWindows = False
            ElseIf InStr(1, UCase(strHostName), "[L]") >0 Then
                platformType = "Linux"
                blnIsWindows = False
            ElseIf InStr(1, UCase(strHostName), "[U]") >0 Then
                platformType = "Unix"
                blnIsWindows = False
            Else
                platformType = "Windows"
                blnIsWindows = True
            End If
            
            'Check if host is a saucelabs cloud computer
            If InStr(1, UCase(strHostName), "SAUCELABS") >0 Then
                blnIsSaucelabs = True
            Else
                blnIsSaucelabs = False
            End If
            
            
            'For windows
            strHostName = Replace(strHostName, "[W]", "")
            strHostName = Replace(strHostName, "[w]", "")
            
            'For Mac
            strHostName = Replace(strHostName, "[M]", "")
            strHostName = Replace(strHostName, "[m]", "")
            
            'For Linux
            strHostName = Replace(strHostName, "[L]", "")
            strHostName = Replace(strHostName, "[l]", "")
            
            'For Unix
            strHostName = Replace(strHostName, "[U]", "")
            strHostName = Replace(strHostName, "[u]", "")
            
            'For execution on a window based computer, look if Krypton process is currently running on that machine
            If blnIsWindows And Not blnIsSaucelabs Then                
                
                'Check if a user is logged on the workstation
                userSession = 0
                WScript.Echo "Checking user session id on " & strHostName
                userSession = GetUserSessionId(strHostName)
                WScript.Echo "User session id for host " & strHostName & " was found to be " & userSession
                
                If userSession < 0 Then
                    blnAvailableHost= False
                Else
                    blnAvailableHost= True
                End If
                
                'Check if application process is running on box, if not, box is deemed to be free
                If blnAvailableHost Then
                    Set objHMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
                    strHostName & "\root\cimv2")
                    
                    Set colHProcesses = objHMIService.ExecQuery ("SELECT * FROM Win32_Process WHERE Name = '" & _
                    strProcess & "'")
                    
                    WScript.Echo "Process count for " & strProcess & " application on host " & strHostName & " was found to be " & colHProcesses.Count
                    
                    If colHProcesses.Count> 0 Then
                        blnAvailableHost= False
                        
                    Else
                        strAvailableHost= strHostName
                        blnAvailableHost= True
                        Exit For
                    End If
                End If
                
                'For a non-window machine, check following
                '1. Extract list of all Krypton processes running on controller, if there are none, launch execution
                '2. Check if any of existing Krypton process is already running that specific remote host
            Else
                
                strAvailableHost = strHostName
                blnAvailableHost= False
                
                'Check if application process is running on box, if not, box is deemed to be free
                Set objHMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
                "." & "\root\cimv2")
                
                Set colHProcesses = objHMIService.ExecQuery ("SELECT * FROM Win32_Process WHERE Name = '" & _
                strProcess & "'")
                
               
                
                If colHProcesses.Count= 0 Then
                    blnAvailableHost= True
                    Exit For
                Else
                    
                    'Iterate through each process command line and see if Krypton is already launched on that specific computer
                    commandsForKRYPTON = ""
                    For Each objItem In colHProcesses
                        strCommand= objItem.commandLine
                        commandsForKRYPTON = commandsForKRYPTON & strCommand
                    Next
                    
                    'MsgBox commandsForKRYPTON
                    'By default, assume here that host will be available
                    If InStr(1, LCase(commandsForKRYPTON), LCase(strHostName)) > 0 Then
                        blnAvailableHost = False
                    Else
                        blnAvailableHost = True
                        Exit For
                    End If
                End If
            End If
        Next
        
        
        'Exit loop if you found an available host
        If blnAvailableHost Then
            GetAvailableHost= strAvailableHost
            Exit Do
        End If
    Loop
End Function


'This method will retrieve list of all free hosts at once
Public Function GetAvailableHosts(strHostList, strProcess, ByRef browser, ByRef strExecutionCommand)
    
    'Host availability means the following
    '1. Host is online and user is logged in (session id is more than one)
    '2. Test execution process (e.g. Krypton.exe) is not already running (linked with browser)
    '3. Specific browser process is not already running (coupled with browser argument)
    
    'Decode list of hosts into an array
    strHostList= Replace(strHostList, " ", "")
    Dim arrHostList: 		arrHostList= Split(strHostList, ",")
    Dim blnAvailableHost: 	blnAvailableHost= False
    Dim strAvailableHost: 	strAvailableHost= ""
    
    Dim currentSauceLabsInstances
    currentSauceLabsInstances= 0
    
    On Error Resume Next
    
    'Start a loop for arrays upper bound
    Do
        
        'Out of the list available, check which machine is completely free
        For i=0 To UBound(arrHostList)
            
            strHostName = arrHostList(i)
            WScript.Echo "Checking availability of " & strHostName
            
            'Check if host is windows or otherwise
            If InStr(1, UCase(strHostName), "[W]") >0 Then
                platformType = "Windows"
                blnIsWindows = True
            ElseIf InStr(1, UCase(strHostName), "[M]") >0 Then
                platformType = "Mac"
                blnIsWindows = False
            ElseIf InStr(1, UCase(strHostName), "[L]") >0 Then
                platformType = "Linux"
                blnIsWindows = False
            ElseIf InStr(1, UCase(strHostName), "[U]") >0 Then
                platformType = "Unix"
                blnIsWindows = False
            Else
                platformType = "Windows"
                blnIsWindows = True
            End If
            
            'Check if host is a saucelabs cloud computer
            If InStr(1, UCase(strHostName), "SAUCELABS") >0 Then
                blnIsSaucelabs = True
            Else
                blnIsSaucelabs = False
            End If
            
            'For windows
            strHostName = Replace(strHostName, "[W]", "")
            strHostName = Replace(strHostName, "[w]", "")
            
            'For Mac
            strHostName = Replace(strHostName, "[M]", "")
            strHostName = Replace(strHostName, "[m]", "")
            
            'For Linux
            strHostName = Replace(strHostName, "[L]", "")
            strHostName = Replace(strHostName, "[l]", "")
            
            'For Unix
            strHostName = Replace(strHostName, "[U]", "")
            strHostName = Replace(strHostName, "[u]", "")
            
            'For execution on a window based computer, look if Krypton process is currently running on that machine
            If blnIsWindows And Not blnIsSaucelabs Then                
                
                'Check if a user is logged on the workstation
                userSession = 0
                WScript.Echo "Checking user session id on " & strHostName
                userSession = GetUserSessionId(strHostName)
                WScript.Echo "User session id for host " & strHostName & " was found to be " & userSession
                
                If userSession < 0 Then
                    blnAvailableHost= False
                Else
                    blnAvailableHost= True
                End If
                
                'Check if application process is running on box, if not, box is deemed to be free
                If blnAvailableHost Then
                    Set objHMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
                    strHostName & "\root\cimv2")
                    
                    Set colHProcesses = objHMIService.ExecQuery ("SELECT * FROM Win32_Process WHERE Name = '" & _
                    strProcess & "'")
                    
                    WScript.Echo "Process count for " & strProcess & " application on host " & strHostName & " was found to be " & colHProcesses.Count
                    
                    If colHProcesses.Count> 0 Then
                        blnAvailableHost= False
                        
                    Else
                        strAvailableHost= strAvailableHost & strHostName & ","
                        blnAvailableHost= True
                        'Exit For
                    End If
                End If
                
                'For a non-window machine, check following
                '1. Extract list of all Krypton processes running on controller, if there are none, launch execution
                '2. Check if any of existing Krypton process is already running that specific remote host
            Else
                blnAvailableHost= False
                
                'Check if application process is running on box, if not, box is deemed to be free
                Set objHMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
                "." & "\root\cimv2")
                
                Set colHProcesses = objHMIService.ExecQuery ("SELECT * FROM Win32_Process WHERE Name = '" & _
                strProcess & "'")
                
                'WScript.Echo "Found " & colHProcesses.Count & " instances of " & strProcess
                
                If colHProcesses.Count= 0 Then
                    blnAvailableHost= True
                    strAvailableHost = strAvailableHost & strHostName & ","
                    'Exit For
                Else
                    
                    'Iterate through each process command line and see if Krypton is already launched on that specific computer
                    commandsForKRYPTON = ""
                    For Each objItem In colHProcesses
                        strCommand= objItem.commandLine
                        commandsForKRYPTON = commandsForKRYPTON & strCommand
                    Next
                    
                    'MsgBox commandsForKRYPTON
                    'By default, assume here that host will be available
                    If InStr(1, LCase(commandsForKRYPTON), LCase(strHostName)) > 0 Then
                        blnAvailableHost = False
                    Else
                        strAvailableHost = strAvailableHost & strHostName & ","
                        blnAvailableHost = True
                        'Exit For
                    End If
                    
                End If
            End If
        Next
        
        
        'Exit loop if you found an available host
        If strAvailableHost <> "" Then
            GetAvailableHost= strAvailableHost
            Exit Do
        End If
    Loop
End Function


Public Function GetUserSessionId(strSessionHostName)
    
    On Error Resume Next
    Err.Clear
    
    Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strSessionHostName & "\root\cimv2")
    
    
    If Err.Number <> 0 Then: WScript.Echo Err.Description: GetUserSessionId= -1: Exit Function: End If
    'WScript.Echo "Connected to host: " & strSessionHostName
    
    
    Set colProcesses = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_Process Where Name = 'explorer.exe'")
   
    If Err.Number <> 0 Then: GetUserSessionId= -1: Exit Function: End If
    
    If colProcesses.Count = 0 Then
        intSessionId = -1
    Else
        For Each objItem In colProcesses
            intSessionId = objItem.SessionId
            Exit For
        Next
    End If
    
    Set colProcesses = Nothing
    Set objWMIService = Nothing
    GetUserSessionId = intSessionId
End Function

Sub LogError(Details, intSerial)
    Dim fs : Set fs = CreateObject("Scripting.FileSystemObject")
    Dim logFile : Set logFile = fs.OpenTextFile("c:\qa_automation\errors.log", 8, True)
    logFile.WriteLine "[" & intSerial & "]" & Now() & ": Error: " & Details.Number & " Details: " & Details.Description
    Err.Clear
End Sub

Public Function blnIsExecutionStarted(strHostName, strProcessName, strCommand, intTimeout)
    
    'This method checks if execution has been started on remote box or not
    'Following must be checked when using this
    '1. If intended process has been started on the workstation
    '2. If the process is running with correct command arguments
    
    On Error Resume Next
    intTimer= 1							'in seconds
    
    Set objHMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
    strHostName & "\root\cimv2")
    Do  
        
        'Check if correct process is running of remote machine
        Set colHProcesses = objHMIService.ExecQuery ("SELECT * FROM Win32_Process WHERE Name = '" & _
        strProcessName & "'")
        
        'Check if a csript has been launched on remote host for the statement
        If colHProcesses.Count= 0 Then
            blnIsExecutionStarted= False
        Else
            For Each objHProcess In colHProcesses
                
                'This checks if the process has been launched with correct command
                If InStr(1,LCase(strCommand), LCase(objHProcess.CommandLine))> 0 Then
                    blnIsExecutionStarted= True
                    AppendLineToFile Controller_Log, "Found Execution running for above command after timer value " & intTimer
                    Exit Function
                Else
                    blnIsExecutionStarted= False
                End If
            Next
        End If
        
        'Making sure it does not go into infinite looping
        If intTimer> intTimeout Then
            AppendLineToFile Controller_Log, "The Execution failed to start within timeout of " & intTimer
            Exit Function
        End If
        
        intTimer= intTimer+1
        WScript.Sleep 1000
    Loop
    
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


Public Function GenerateHtmlReport(resultsLocation, processId, xmlFileName, testSuiteName, testEnvironment)
    Dim listXmls
    listXmls = ""
    
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
        If InStr(1, LCase(htmlFile.Name), LCase(htmlSummaryLocation)) Then
            htmlDetailedLocation = htmlFile.Path
            htmlSummaryLocation = Replace(htmlDetailedLocation, ".html", "s.html")
            Exit For
        End If
    Next
    
    WScript.Echo "HTML Report generation completed"
    
    WScript.Echo "Generating summary report now..."
    Call GenerateSummaryReport(htmlSummaryLocation, htmlDetailedLocation)
    WScript.Echo "Summary Report generation completed"
    
    If blnUseGridEmailFeatures Then
        WScript.Echo "Sending grid specific email..."
        Call SendFinishedExecutionEmail(testSuiteName, htmlSummaryLocation, htmlDetailedLocation)
        WScript.Echo "Email sent complete."
    End If
    
    Set objHTMLFSO= Nothing
End Function


Public Function SentStartExecutionEmail(testSuiteName)
    
    EmailSendCommand= KRYPTONHome & "\KryptonCreateReport.exe" & _
    " " & Quote & "emailFor=start" & Quote & _
    " " & Quote & "testsuite=" & testSuiteName & Quote & _
    " " & Quote & "emailrequired=" & sendKRYPTONEmail & Quote
    
    WshShell.Run ReportGenerationCommand, 0, False
    
End Function

Public Function GenerateSummaryReport(htmlSummaryLocation, htmlDetailedLocation)
    On Error Resume Next
    Set htmlFile = objFS.OpenTextFile(htmlDetailedLocation)
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
    
    'Extract total  test case count
    totalTestCaseString = "Test Case(s) Executed:"
    startLocation = InStr(1, entireHTMLContents, totalTestCaseString)
    startLocation = InStr(startLocation, entireHTMLContents, ">") +1
    startLocation = InStr(startLocation, entireHTMLContents, ">") +1
    endLocation   = InStr(startLocation, entireHTMLContents, "<")
    
    totalExecuted = Mid(entireHTMLContents, startLocation, endLocation-startLocation)
    totalExecuted = CInt(totalExecuted)
    
    
    'Extract passed test case count
    totalPassedString = "Total Passed:"
    startLocation = InStr(1, entireHTMLContents, totalPassedString)
    startLocation = InStr(startLocation, entireHTMLContents, ">") +1
    startLocation = InStr(startLocation, entireHTMLContents, ">") +1
    endLocation   = InStr(startLocation, entireHTMLContents, "<")
    
    totalPassed = Mid(entireHTMLContents, startLocation, endLocation-startLocation)
    totalPassed = CInt(totalPassed)
   
    
    'Extract failed test case count
    totalFailedString = "Total Failed:"
    startLocation = InStr(1, entireHTMLContents, totalFailedString)
    startLocation = InStr(startLocation, entireHTMLContents, ">") +1
    startLocation = InStr(startLocation, entireHTMLContents, ">") +1
    endLocation   = InStr(startLocation, entireHTMLContents, "<")
    
    totalFailed = Mid(entireHTMLContents, startLocation, endLocation-startLocation)
    totalFailed = CInt(totalFailed)
    
    
    'Extract warning test case count
    totalWarningString = "Total Warning:"
    startLocation = InStr(1, entireHTMLContents, totalWarningString)
    startLocation = InStr(startLocation, entireHTMLContents, ">") +1
    startLocation = InStr(startLocation, entireHTMLContents, ">") +1
    endLocation   = InStr(startLocation, entireHTMLContents, "<")
    
    totalWarning = Mid(entireHTMLContents, startLocation, endLocation-startLocation)
    totalWarning = CInt(totalWarning)
   
    
    'Update pass count to be sum of pass and warning
    totalPassed = totalPassed + totalWarning
    
    'Create email subject based on numbers
    If CInt(totalFailed) > 0 Then
        emailSubject = "[" & totalFailed & "/" & totalExecuted & "] Failed! '$testsuite$' test suite has failed"
    Else
        emailSubject = "[" & totalPassed & "] Passed! '$testsuite$' test suite execution completed."
    End If
    
   
    Set tempFile= objFS.CreateTextFile(htmlSummaryLocation, True)
    tempFile.Write entireHTMLContents
    tempFile.Close
End Function

Public Function SendFinishedExecutionEmail(testSuiteName, htmlSummaryFile, htmlDetailedFile)
    On Error Resume Next
    
    Set objMessage = CreateObject("CDO.Message")
    objMessage.Sender = fromEmail
    Dim ProjectPath  
    EmailRecipient= GetEmailRecipients()
    
    objMessage.To       = EmailRecipient
    
    If emailSubject = "" Then
        emailSubject = "Test Suite: " & testSuiteName & " finished."
    Else
        emailSubject = Replace(emailSubject, "$testsuite$", testSuiteName)
    End If
    
    
    objMessage.Subject = emailSubject
    objMessage.From    = "Krypton Automation"
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
    
    
    objMessage.CreateMHTMLBody htmlSummaryFile
    objMessage.AddAttachment htmlDetailedFile
    
    objMessage.Send
End Function


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
            
            If colHProcesses.Count> 0 Then
                For Each objHProcess In colHProcesses
                    If InStr(1,LCase(objHProcess.CommandLine), LCase(processId))> 0 Then
                        blnExecutionComplete= False
                    End If
                Next
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


Public Function GetEmailRecipients()
    Dim fileText, AllRecipients
    AllRecipients=""
   ProjectPath= GetProjectFolderPath()
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


Public Function setSystemEnvVariable(VarName, VarValue)
    Dim WshShell
    Dim WshUserEnv
    Set WshShell =CreateObject("WScript.Shell")
    Set WshUserEnv = WshShell.Environment("SYSTEM")
    WshUserEnv(VarName) = VarValue
    Set WshUserEnv= Nothing
    Set WshShell=	Nothing
End Function

Public Function getSystemEnvVariable(VarName)
    Dim WshShell
    Dim WshUserEnv
    Set WshShell =CreateObject("WScript.Shell")
    Set WshUserEnv = WshShell.Environment("SYSTEM")
    getSystemEnvVariable= WshUserEnv(VarName)
    Set WshUserEnv= Nothing
    Set WshShell=	Nothing
End Function

Function Lpad (MyValue, MyPadChar, MyPaddedLength)
    On Error Resume Next
    Lpad = MyValue
    Lpad = String(MyPaddedLength - Len(MyValue),MyPadChar) & MyValue
End Function

Function StartProcessIfNotRunningAlready(processExecutable)
    On Error Resume Next
    
    'Extract Process name
    Dim processfso
    Set processfso = WScript.CreateObject("Scripting.Filesystemobject")
    strComputer = "."
    
    processName = processfso.GetFileName(processExecutable)
    
    'Check if process is already running or not, if not, start only then
    Set objHMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
    strComputer & "\root\cimv2")
    
    Set colHProcesses = objHMIService.ExecQuery ("SELECT * FROM Win32_Process WHERE Name = '" & _
    processName & "'")
    
    'If process is not running already, start it now, in a hidden window
    If colHProcesses.Count= 0 Then
        Set objStartup = objHMIService.Get("Win32_ProcessStartup")
        Set objConfig = objStartup.SpawnInstance_
        objConfig.ShowWindow = HIDDEN_WINDOW
        Set objProcess = GetObject("winmgmts:root\cimv2:Win32_Process")
        errReturn = objProcess.Create(processExecutable, Null, objConfig, intProcessID)
    End If
    
End Function


Public Function GetParameterFromIniFile(iniFileLocation, parameterName, homeLocation)
    Dim parameterValue
    parameterValue= ""
    
    'On Error Resume Next
    Set objMyFS = CreateObject("Scripting.FileSystemObject")
    Set iniFile = objMyFS.OpenTextFile(iniFileLocation)
    
    arrParameters = Split(iniFile.ReadAll, vbNewLine)
    
    iniFile.Close
    Set iniFile = Nothing
    Set objMyFS = Nothing
    
    For Each parameter In arrParameters
        
        name = Split(parameter, ":")(0)
        name = Trim(name)
        
        If LCase(name) = LCase(parameterName) Then
            parameterValue= Split(parameter, ":")(1)
            parameterValue=Trim(parameterValue)
            Exit For
        End If
    Next
    GetParameterFromIniFile= Replace(parameterValue, "./", homeLocation & "\")
End Function