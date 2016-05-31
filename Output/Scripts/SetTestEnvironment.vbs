Dim targetFolder
Dim objFSO
Dim strHostFilePath
Dim strEnvName
Dim browserType
browserType = ""

Const quote = """"
Const sTestEnvVarName= "TestEnv"
usage = "Usage is as follows" & vbNewLine & _
		"SetTestEnvironment.vbs " & "QA1" & " " & quote & "\\match.corp\files\qa$\HostFiles\QA1\hosts" & quote

On Error Resume Next

'>>>First Thing First, validate if correct arguments have been passed
Set objArgs = WScript.Arguments

If objArgs.Count<2 Then
    WScript.Echo usage
    WScript.Quit
End If

strEnvName= UCase(objArgs(0))
strHostFilePath = Trim(objArgs(1))

If objArgs.Count >= 3 Then
	browserType = Trim(objArgs(2))
End If

If UCase(browserType) = "IE" Then 
	Call DeleteCookies()
End If

Set objFSO = CreateObject("Scripting.FileSystemObject")
If Not objFSO.FileExists(strHostFilePath) Then
	WScript.Echo "Cannot locate host file at path: " & 	strHostFilePath & vbNewLine & _
				 "Please check that the file exists and appropriate permissions are granted."
End If

Err.Clear

'Retrieve the Target path where host file need to be copied
targetFolder= GetSystemEnvVariable("windir") & "\system32\drivers\etc"

'Copy host file to system folder
objFSO.CopyFile strHostFilePath, targetFolder & "\hosts", True

'Set environment variable on local system
Call SetUserEnvVariable(sTestEnvVarName, strEnvName)
Set objFSO= Nothing


Public Function getSystemEnvVariable(VarName)
    Dim WshShell
    Dim WshUserEnv
    Set WshShell =CreateObject("WScript.Shell")
    getSystemEnvVariable = WshShell.ExpandEnvironmentStrings("%"&VarName&"%")
    Set WshShell=	Nothing
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

Public Function DeleteCookies()
    
    'To clear all history data
    On Error Resume Next
    Set WshShell = CreateObject("WScript.Shell")
	WshShell.Run "DeleteIECookies.exe",0,True
	
    'brwsrVersion = GetIEBrowserVersion()
    'MsgBox brwsrVersion
    'If brwsrVersion >= 7 Then
	    'WshShell.Run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255", 0, False
	    'WshShell.Run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255", 0, True
	    'WshShell.Run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2", 0, True
    'Else
    	'WshShell.Run "DeleteIE6Cookies.bat",0,True
    'End If
    Set WshShell =Nothing
End Function


Public Function GetIEBrowserVersion()

	On Error Resume Next
	CurrentMode = Reporter.Filter
	Reporter.Filter= 3
	
	Dim WshShell
	Dim browserVersion
	Set WshShell = CreateObject("wscript.Shell")
	browserVersion= WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Version")
	browserVersion= Left(browserVersion, 1)
	browserVersion= Replace(browserVersion, " ", "")
	GetIEBrowserVersion= browserVersion
	Set WshShell= Nothing
	Reporter.Filter= CurrentMode 
End Function