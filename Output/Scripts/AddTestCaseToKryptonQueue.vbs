
'First Thing first, check for passed arguments, should be at least one
'Argument passed to this script will indicate what batch file to executed

Set objArgs = WScript.Arguments
On Error Resume Next

Const Quote= """"

Dim KRYPTONHome
Dim strBatchtoExecute
Dim objFSO
Dim objExecutionQueueFile
Dim ExecutionQueue
Dim strTestSetsQueue
Dim blnControllerScriptRunning
Dim Arguments
Dim testSuiteLocation
Dim testSuiteName
Dim blnIsLocalExecution
blnIsLocalExecution = False

'Any extra argument that are not recognized by this script, will be passed to controller script as it is
Dim controllerArguments
controllerArguments= ""

Dim strControllerScript
strTestSetsQueue = ""

'Quit script if there is no argument provided
If objArgs.Count<1 Then
    WScript.Quit
End If

Set objFSO= CreateObject("Scripting.FileSystemObject")
KRYPTONHome = objFSO.GetParentFolderName(WScript.ScriptFullName)

'Parameters from arguments
Dim TestCaseId
Dim TestEnv
Dim Host
Dim IsGroup
Dim strNodes
strNodes = ""
Dim strUserName
Dim strPassword
strUserName= ""
strPassword= ""
Dim Browser

'Default execution queue
ExecutionQueue= objFSO.GetParentFolderName(WScript.ScriptFullName) & "\ExecutionQueue.txt"

Dim ExecutionLogFileName
ExecutionLogFileName= objFSO.GetParentFolderName(WScript.ScriptFullName) & "\kryptonexecution.log"


For Each argument In objArgs
    If InStr(1,LCase(argument),"testsuitename=")>0 Then
        strBatchtoExecute= Split(argument, "=")(1)
        testSuiteName= strBatchtoExecute
        
    ElseIf InStr(1,LCase(argument),"testsuitelocation=")>0 Then
        testSuiteLocation= Split(argument, "=")(1)
    ElseIf InStr(1,LCase(argument),"kryptonhome=")>0 Then
        KRYPTONHome= Split(argument, "=")(1)
    ElseIf InStr(1,LCase(argument),"host=")>0 Then
        Host= Trim(Split(argument, "=")(1))
        ExecutionQueue= KRYPTONHome & "\ExecutionQueue_" & Host & ".txt"
        If LCase(Host) = "localhost" Then
            blnIsLocalExecution = True
        End If
    ElseIf InStr(1,LCase(argument),"isgroup=")>0 Then
        IsGroup= Split(argument, "=")(1)
    ElseIf InStr(1,LCase(argument),"username=")>0 Then
        strUserName= Split(argument, "=")(1)
    ElseIf InStr(1,LCase(argument),"password=")>0 Then
        strPassword= Split(argument, "=")(1)
    ElseIf InStr(1,LCase(argument),"command=")>0 Then
        Arguments= Right(argument, Len(argument)-InStr(argument, "="))
    Else
        'build set of argument that will be passed directly to controller script
        'nodes will also be passed as part of this argument only
        controllerArguments = controllerArguments & " " & Quote & argument & Quote
    End If
    
Next

strControllerScript=KRYPTONHome & "\StartExecutionFromController.vbs"

'Check if QA automation folder exists, if not, quit
If Not objFSO.FolderExists(KRYPTONHome) Then
    WScript.Quit
End If

'If keyword delete is passed as an argument, this means to delete the existing queue
If UCase(strBatchtoExecute)= "{DELETE}" Then
    objFSO.DeleteFile KRYPTONHome & "\ExecutionQueue*.txt", True
    objFSO.DeleteFile KRYPTONHome & "\KRYPTONXMLLists_*.txt", True
    WScript.Quit
End If

'Create Queue file if it does not exists already
If Not objFSO.FileExists(ExecutionQueue) Then
    objFSO.CreateTextFile ExecutionQueue, True
End If

'Open Queue file in append mode, insert a statement for argument passed to this script and close file
Set objExecutionQueueFile= objFSO.OpenTextFile(ExecutionQueue, 8)	'8 means append mode

'Also adding nodes along with test suite itself
objExecutionQueueFile.WriteLine "[TEST_SUITE: " & testSuiteName & "]"
strBatchtoExecute= Replace(strBatchtoExecute, "{QUOTE}", Quote)
Arguments= Replace(Arguments, "$$", Quote)

'Fill test case id array from test suite file
testSuiteFileName = testSuiteLocation & "\" & testSuiteName & ".suite"

'MsgBox testSuiteFileName
Set objTextFile = objFSO.OpenTextFile(testSuiteFileName, 1)
testCaseIds= objTextFile.ReadAll()
'MsgBox testCaseIds

arrTestCaseIds= Split(testCaseIds, vbNewLine) 

'Additional parameters can be passed on at starting of the test suite file
'Format will be as below

Dim additionalParameters
additionalParameters = ""

For Each TestCaseId In arrTestCaseIds
    Dim blnAdditionalParameters
    blnAdditionalParameters = False
    
    If Left(TestCaseId, 1) = "[" Then
        blnAdditionalParameters = True
        additionalParameters = Mid(TestCaseId, 2, Len(TestCaseId)-2)
        TestCaseId = ""
    End If 
    
    'Handle situation where test suite id is also provided in *.suite file
    testSuiteId = ""
    If UBound(Split(TestCaseId,vbTab)) >=1 Then
        testSuiteId = Split(TestCaseId,vbTab)(0)
        TestCaseId = Split(TestCaseId,vbTab)(1)
    End If
    
    KRYPTONCommands = Quote & KRYPTONHome & "\Krypton.exe" & Quote &_
    " " & Quote & "testsuite=" & Trim(testSuiteName) & Quote &_
    " " & Quote & "testsuiteid=" & Trim(testSuiteId) & Quote &_
    " " & Quote & "testcaseid=" & Trim(TestCaseId) & Quote &_
    " " & Arguments
    
    KRYPTONCommands = KRYPTONCommands & " " & Quote & "endexecutionwaitrequired=false" & Quote
    KRYPTONCommands = KRYPTONCommands & " " & Quote & "closebrowseroncompletion=true" & Quote
    KRYPTONCommands = KRYPTONCommands & " " & additionalParameters
    
    If Not(Trim(TestCaseId) = "") Then
        objExecutionQueueFile.WriteLine KRYPTONCommands
    End If
Next
objExecutionQueueFile.WriteLine "[FINISH]"

objExecutionQueueFile.Close
Set objExecutionQueueFile= Nothing
Set objFSO= Nothing


'Check if controller script is already running, if so, quit as existing running script will take care of everything
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colProcesses = objWMIService.ExecQuery _
("SELECT * FROM Win32_Process WHERE Name = " & "'wscript.exe' OR Name = 'cscript.exe'")

blnControllerScriptRunning= False

If colProcesses.Count> 0 Then
    For Each objProcess In colProcesses
        If InStr(1,LCase(objProcess.CommandLine),LCase(ExecutionQueue))>0 Then
            blnControllerScriptRunning= True
            Exit For
        End If
    Next
End If

hostParameters = " -accepteula"
If Not blnIsLocalExecution Then
    hostParameters = " -u " & QUOTE & strUserName & QUOTE & " -p " & QUOTE & strPassword & QUOTE & _
    " -accepteula -i " & _
    GetUserSessionId(".")
End If

'If controller script is not running, start it
If Not blnControllerScriptRunning Then
    Set WshShell = CreateObject("WScript.Shell")
    ExecutionString = QUOTE & KRYPTONHome & "\psexec.exe" & QUOTE & _
    hostParameters & _
    " -d -n 10" & " cscript " & Quote & strControllerScript & Quote & _
    " " & Quote & "queue=" & Trim(ExecutionQueue) & Quote & _
    " " & Quote & "kryptonhome=" & KRYPTONHome & Quote & _
    " " & Quote & "username=" & strUserName & Quote & _
    " " & Quote & "password=" & strPassword & Quote & _
    " " &   " " & Trim(controllerArguments)
    
    Call AppendLineToFile(ExecutionLogFileName, "Controller Launch Command: " & vbNewLine & ExecutionString)
    
    WshShell.Run ExecutionString,1,False
    
    'Wait for controller script to actually start execution
    For i=0 To 10
        Set colProcesses = objWMIService.ExecQuery _
        ("SELECT * FROM Win32_Process WHERE Name = " & "'wscript.exe' OR Name = 'cscript.exe'")
        
        If colProcesses.Count> 0 Then
            For Each objProcess In colProcesses
                If InStr(1,LCase(objProcess.CommandLine),LCase(ExecutionQueue))>0 Then
                    Exit For
                End If
            Next
        End If
    WScript.Sleep 1000    
    Next
    
    Set WshShell= Nothing
End If

Set colProcesses= Nothing
Set objWMIService= Nothing

Public Function DecodeBatchToTestSet(strBatchName)
    Dim BaseLocationForBatches
    Dim objDecodeFso
    Dim objDecodeBatchFile
    On Error Resume Next
    
    BaseLocationForBatches = batchLocation
    
    Set objDecodeFso = CreateObject("Scripting.FileSystemObject")
    If Not objDecodeFso.FileExists(BaseLocationForBatches & "\" & strBatchName) Then
        Exit Function
    End If
    
    Set objDecodeBatchFile = objDecodeFso.OpenTextFile(BaseLocationForBatches & "\" & strBatchName, 1)
    
    Do Until objDecodeBatchFile.AtEndOfStream
        strNextLine = LCase(Trim(objDecodeBatchFile.Readline))
        
        If InStr(1, strNextLine, ".bat") > 0 And Not Left(strNextLine, 2) = "--" Then
            strNextLine = Replace(strNextLine, "cscript", "")
            strNextLine = Replace(strNextLine, "c:\qa_automation\executetestset.vbs", "")
            strNextLine = Replace(strNextLine, "executetestset.vbs", "")
            strNextLine = Trim(Replace(strNextLine, Quote, ""))
            Call DecodeBatchToTestSet(strNextLine)
        Else
            If Not Left(strNextLine, 2) = "--" Then
                strNextLine = Replace(strNextLine, "c:\qa_automation\executetestset.vbs", "executetestset.vbs")
                strNextLine = Replace(strNextLine, "executetestset.vbs", "c:\qa_automation\executetestset.vbs")
                strTestSetsQueue = strTestSetsQueue & vbNewLine & strNextLine
            End If
        End If
    Loop
    DecodeBatchToTestSet= strTestSetsQueue
End Function

Public Function SetUserEnvVariable(VarName, VarValue)
    Dim WshShell
    Dim WshUserEnv
    Set WshShell =CreateObject("WScript.Shell")
    Set WshUserEnv = WshShell.Environment("USER")
    WshUserEnv(VarName) = VarValue
    Set WshUserEnv= Nothing
    Set WshShell=	Nothing
End Function

Public Function GetUserSessionId(strSessionHostName)
    
    On Error Resume Next
    
    Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strSessionHostName & "\root\cimv2")
    
    Set colProcesses = objWMIService.ExecQuery _
    ("Select SessionId from Win32_Process Where Name = 'explorer.exe'")
    
    If colProcesses.Count = 0 Then
        intSessionId = 0
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