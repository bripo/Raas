'CRM RaaS Prerequistes sample test script
Option Explicit

Dim bEnableVerboseMode
Dim bSkipWMIRobustness

bEnableVerboseMode = false 'set to true for verbose logging

bSkipWMIRobustness = false 'by default set to false, it skips WMI 10K robustness loop (set to true only for advanced diagnostics with TAM and/or PFE)

const RemoteWMIRobustnessMaxCount = 10000

Dim oLogQuery
Dim oIIS
Dim oWMI
Dim oShell
Dim oReg

Dim sServerName
Dim colServices
Dim objEvent
Dim objService
Dim bIISFound
Dim bCRMFound
Dim sRegValue
Dim REG_NET_Framework4
Dim REG_POWERSHELL_2
Dim REG_IIS_VersionString
Dim REG_IIS_ScriptAndTools
Dim REG_IIS_MANAGEMENT_SCRIPT_AND_TOOLS
Dim REG_HTTP_Logging
Dim REG_MaxIdleTime
Dim REG_MaxDisconnectionTime
Dim REG_UNINSTALL
Dim REG_MSCRM
Dim strEntry1a
Dim strEntry1b
Dim strEntry2
Dim strEntry3
Dim strEntry4
Dim strEntry5
Dim bSuccessfulOutput
Dim strCheckAdminShare_Admin_Result
Dim CheckFailureCount
Dim strRegValueIISLocal
Dim strRegValueIISRemote
Dim dwScriptAndTools
Dim strCommand
Dim ObjFileSys
Dim objFile
Dim objFileFail
Dim args
Dim sArgs
Dim j
Dim v_remote_webAdmin
Dim str_NgenOut86
Dim str_NgenOut64
Dim str_wmiRepository
Dim objNetwork
Dim objUserName
Dim bIsRoamigUser
Dim bPassedEvtLogApp
Dim bPassedEvtLogSys
Dim bPassedEvtLogMultiple
Dim iCptRobustness
Const HKEY_CURRENT_USER = &H80000001
const HKEY_LOCAL_MACHINE = &H80000002
Dim colLoggedEvents
Dim Users
Dim UserProfiles
Dim Profile

Dim objWMIService 
Dim colItems 
Dim objItem
Dim strUser 
Dim strLocalMachineDateTime
Dim strLocalMachineTimeZone
Dim strRemoteMachineDateTime
Dim strRemoteMachineTimeZone
Dim iTimeDiffRemoteServer

Dim oSQLConn
Dim rs
Dim bSQLCheckSysAdmin
Dim bSQLServiceFound
Dim sSQLInstanceName

Dim StdOut
Dim bFolderRedirectionDetected
Dim arrValueNames
Dim arrValueTypes
Dim strValue
Dim i 

Dim bPendingRebootDetected
Dim iRequiredUpdates

Dim oLoggedOnUsers
Dim objLoggedOnUser
Dim sComputerName
Dim iIISLogsFailures
Dim bWMITestFailed
Dim LocalListAllInstalledSoftware
Dim IncompatibleInstalledSoftware
Dim RemoteListAllInstalledSoftware

Dim bFoundDomainUser
bIsRoamigUser = false
iIISLogsFailures = 0
Dim iTotalFailures
Dim iToolsMachinePerformanceWarnings

Set StdOut = WScript.StdOut
Set objNetwork = CreateObject("WScript.Network")

REG_NET_Framework4 = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\TargetVersion"    
REG_POWERSHELL_2= "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\PowerShell\1\PowerShellEngine\PowerShellVersion"
REG_HTTP_Logging="SOFTWARE\Microsoft\InetStp\Components\"
REG_IIS_VersionString = "SOFTWARE\Microsoft\InetStp\VersionString"
REG_IIS_ScriptAndTools = "SOFTWARE\Microsoft\InetStp\Components"
REG_IIS_MANAGEMENT_SCRIPT_AND_TOOLS = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\InetStp\Components\ManagementScriptingTools"
REG_MaxIdleTime = "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\MaxIdleTime"
REG_MaxDisconnectionTime = "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services\MaxDisconnectionTime"
REG_MSCRM = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\MSCRM\configdb"
REG_UNINSTALL = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
strEntry1a = "DisplayName" 
strEntry1b = "QuietDisplayName" 
strEntry2 = "InstallDate" 
strEntry3 = "VersionMajor" 
strEntry4 = "VersionMinor" 
strEntry5 = "EstimatedSize" 



Function CheckForIncompatibleSoftware (aListAllInstalledSoftware)

    Dim cptInstalledSoftware
    Dim InstalledSoftware
    Dim aInstalledSoftware()
    cptInstalledSoftware = 1

    Redim aInstalledSoftware(1)

    Err.Clear

    Dim sIssue 
    sIssue = ""

    For Each InstalledSoftware in aListAllInstalledSoftware

        If InStr (LCase(InstalledSoftware), "preview") or InStr (LCase(InstalledSoftware), "beta") or InStr (LCase(InstalledSoftware), "ctp") or InStr (LCase(InstalledSoftware), "alpha") Then
            CheckFailureCount = CheckFailureCount +1
            sIssue = sIssue & "* " & InstalledSoftware & VbCrLf

            aInstalledSoftware (cptInstalledSoftware) = InstalledSoftware
            cptInstalledSoftware = cptInstalledSoftware +1
            Redim preserve aInstalledSoftware(cptInstalledSoftware)
        End If

    Next

    If sIssue <> "" Then
        WScript.Echo "----------------------------------------------------------------------------"
        objFile.Writeline "----------------------------------------------------------------------------"
        objFileFail.Writeline "----------------------------------------------------------------------------"
        WScript.Echo "/!\ Potential INCOMPATIBLE BETA/Preview/CTP/Alfa software with RaaS found on tools machine: <" & sComputerName & "> : Affected software bellow - FAILED." & VbCrLf & sIssue & VbCrLf
        objFile.Writeline "/!\ Potential INCOMPATIBLE software with RaaS found on tools machine: <" & sComputerName & "> : Affected software bellow - FAILED." & VbCrLf & sIssue & VbCrLf
        objFileFail.Writeline "/!\ Potential INCOMPATIBLE software with RaaS found on tools machine: <" & sComputerName & "> : Affected software bellow - FAILED." & VbCrLf & sIssue & VbCrLf
        WScript.Echo "----------------------------------------------------------------------------"
        objFile.Writeline "----------------------------------------------------------------------------"
        objFileFail.Writeline "----------------------------------------------------------------------------"

    End If

    sIssue = ""
    For Each InstalledSoftware in aListAllInstalledSoftware

        If InStr (LCase(InstalledSoftware), "visual studio") or InStr (LCase(InstalledSoftware), "exchange") or InStr (LCase(InstalledSoftware), "sql server") or InStr (LCase(InstalledSoftware), "sharepoint") or InStr (LCase(InstalledSoftware), "crm") Then

            If InStr(LCase(InstalledSoftware), "outlook") = 0 Then


                CheckFailureCount = CheckFailureCount +1
                sIssue = sIssue & "* " & InstalledSoftware & VbCrLf

                aInstalledSoftware (cptInstalledSoftware) = InstalledSoftware
                cptInstalledSoftware = cptInstalledSoftware +1
                Redim preserve aInstalledSoftware(cptInstalledSoftware)
            End If
        End If
    Next

    If sIssue <> "" Then
        WScript.Echo "----------------------------------------------------------------------------"
        objFile.Writeline "----------------------------------------------------------------------------"
        objFileFail.Writeline "----------------------------------------------------------------------------"
        WScript.Echo "/!\ Potential INCOMPATIBLE server workload or developper tools with RaaS client found on tools machine:  <" & sComputerName & "> : Affected software bellow - FAILED." & VbCrLf & sIssue & VbCrLf
        objFile.Writeline "/!\ Potential INCOMPATIBLE server workload or developper tools with RaaS client found on tools machine:  <" & sComputerName & "> : Affected software bellow - FAILED." & VbCrLf & sIssue & VbCrLf
        objFileFail.Writeline "/!\ Potential INCOMPATIBLE server workload or developper tools with RaaS client found on tools machine:  <" & sComputerName & "> : Affected software bellow - FAILED." & VbCrLf & sIssue & VbCrLf
        WScript.Echo "----------------------------------------------------------------------------"
        objFile.Writeline "----------------------------------------------------------------------------"
        objFileFail.Writeline "----------------------------------------------------------------------------"

    End If



    CheckForIncompatibleSoftware = aInstalledSoftware

End Function


'Function to List All Installed Software
Function ListAllInstalledSoftware(sComputerName)
    Dim objReg
    Dim intRet1
    Dim strValue1, strValue2, strValue3, strValue4, strValue5
    Dim intValue3, intValue4, intValue5
    Dim Result
    Dim arrSubkeys
    Dim strSubkey
    Dim aInstalledSoftware()
    Dim cptInstalledSoftware
    Dim InstalledSoftware
    cptInstalledSoftware = 1

    Err.Clear

    Set objReg = GetObject("winmgmts://" & sComputerName & "/root/default:StdRegProv") 
    objReg.EnumKey HKEY_LOCAL_MACHINE, REG_UNINSTALL, arrSubkeys 
    
    Redim aInstalledSoftware(1)

    If Err.number = 0 Then

        WScript.Echo "Enumerating applications on: <" & sComputerName & ">" 
        objFile.Writeline "Enumerating applications on: <" & sComputerName & ">" 
        objFileFail.Writeline "Enumerating applications on: <" & sComputerName & ">" 
    
        For Each strSubkey In arrSubkeys 
            

            intRet1 = objReg.GetStringValue(HKEY_LOCAL_MACHINE, REG_UNINSTALL & strSubkey, strEntry1a, strValue1) 
            If intRet1 <> 0 Then 
                objReg.GetStringValue HKEY_LOCAL_MACHINE, REG_UNINSTALL & strSubkey, strEntry1b, strValue1 
            End If 
            If strValue1 <> "" Then 
                'WScript.Echo strValue1
                Result = Result & strValue1
            End If 
            objReg.GetStringValue HKEY_LOCAL_MACHINE, REG_UNINSTALL & strSubkey, strEntry2, strValue2 
            If strValue2 <> "" Then 
                'WScript.Echo "Install Date: " & strValue2
                Result = Result & " | Install Date: " & strValue2
            End If 
            objReg.GetDWORDValue HKEY_LOCAL_MACHINE, REG_UNINSTALL & strSubkey, strEntry3, intValue3 
            objReg.GetDWORDValue HKEY_LOCAL_MACHINE, REG_UNINSTALL & strSubkey, strEntry4, intValue4 
            If intValue3 <> "" Then 
                'WScript.Echo "Version: " & intValue3 & "." & intValue4
                Result = Result &  " | Version: " & intValue3 & "." & intValue4
            End If 

            objReg.GetDWORDValue HKEY_LOCAL_MACHINE, REG_UNINSTALL & strSubkey, strEntry5, intValue5 
            If intValue5 <> "" Then 
                'WScript.Echo "Estimated Size: " & Round(intValue5/1024, 3) & " megabytes" 
                'Result = Result & " | Estimated Size: " & Round(intValue5/1024, 3) & " megabytes" 
            End If
            If Result <> "" Then
               aInstalledSoftware (cptInstalledSoftware) =  Result
               ' Result = Result & VbCrLf
               cptInstalledSoftware = cptInstalledSoftware +1
               ReDim preserve aInstalledSoftware(cptInstalledSoftware)
               Result = ""
           End If


        Next

        WScript.Echo "Total applications installed on: <" & sComputerName & "> : " & cptInstalledSoftware
        objFile.Writeline "Total applications installed on: <" & sComputerName & "> : " & cptInstalledSoftware
        objFileFail.Writeline "Total applications installed on: <" & sComputerName & "> : " & cptInstalledSoftware

    Else
        CheckFailureCount = CheckFailureCount +1
        WScript.Echo "List All Installed Software on: <" & sComputerName & "> - FAILED."
        Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description

        objFile.Writeline "List All Installed Software on: <" & sComputerName & "> - FAILED."
        objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
        objFileFail.Writeline "List All Installed Software on: <" & sComputerName & "> - FAILED."
        objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description        
    End If
    Set objReg = Nothing
    Err.Clear


    objFile.Writeline "---------------------------------------------"

    For Each InstalledSoftware in aInstalledSoftware

        'WScript.Echo InstalledSoftware
        objFile.Writeline InstalledSoftware

    Next
    
    objFile.Writeline "---------------------------------------------"
   
    ListAllInstalledSoftware = aInstalledSoftware

End function


'Function to execute  command and read output from command 
Function Exec(objShell,cmd)
    Dim execShell	
	Set execShell = objShell.Exec(cmd)
	execShell.StdIn.Close 
	Exec=execShell.StdOut.ReadAll
	Set execShell =Nothing
End function

Function ExecEx(objShell,cmd)
    Dim execShell
	Set execShell = objShell.Exec(cmd)
	execShell.StdIn.Close	
    ExecEx=execShell.ExitCode
    'resultString=execShell.StdOut.ReadAll
	Set execShell =Nothing
End function

'Sub to check processor architecture and OS bit on tools machine

Sub CheckArchitecture(objWMI)
	Dim ColSettings
	Dim ObjProcessor
	Set ColSettings = ObjWMI.ExecQuery ("SELECT * FROM Win32_Processor") 
	For Each ObjProcessor In ColSettings 
    		Select Case ObjProcessor.Architecture 
        		Case 0 
            			WScript.Echo "Processor Architecture Used by the Platform: x86" 
        		Case 6 
            			WScript.Echo "Processor Architecture Used by the Platform: Itanium-Based System" 
        		Case 9 
            			WScript.Echo "Processor Architecture Used by the Platform: x64" 
    		End Select 
    		
    		WScript.Echo "Processor: " & ObjProcessor.DataWidth & "-Bit" 
    		WScript.Echo "Operating System: " & ObjProcessor.AddressWidth & "-Bit" 
   		    WScript.Echo  "Current Clock Speed: " & objProcessor.CurrentClockSpeed
		    Wscript.Echo "Maximum Clock Speed: " & objProcessor.MaxClockSpeed   
		
            Err.Clear
   		    objFile.Writeline "Current Clock Speed: " & objProcessor.CurrentClockSpeed
            If Err.number <> 0 Then
                Wscript.Echo "Could not write to trace log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
                CheckFailureCount = CheckFailureCount +1        
            End If
            Err.Clear

		    objFile.Writeline "Maximum Clock Speed: " & objProcessor.MaxClockSpeed   
		    objFile.Writeline "Number of Cores: " & objProcessor.NumberofCores   
		
    		If ObjProcessor.Architecture = 0 AND ObjProcessor.AddressWidth = 32 Then 
       			WScript.Echo "This Machine has 32 Bit Processor and Running 32 Bit OS" 
			    objFile.writeline "This Machine has 32 Bit Processor and Running 32 Bit OS"
    		End If 
    		If (ObjProcessor.Architecture = 6 OR ObjProcessor.Architecture = 9) AND ObjProcessor.DataWidth = 64 AND ObjProcessor.AddressWidth = 32 Then 
        		WScript.Echo "This Machine has 64-Bit Processor and Running 32-Bit OS" 
			objFile.Writeline "This Machine has 64-Bit Processor and Running 32-Bit OS" 
   		End If 
    		If (ObjProcessor.Architecture = 6 OR ObjProcessor.Architecture = 9) AND ObjProcessor.DataWidth = 64 AND ObjProcessor.AddressWidth = 64 Then 
       			 WScript.Echo "This Machine has 64-Bit Processor and Running 64-Bit OS"
			 objFile.Writeline "This Machine has 64-Bit Processor and Running 64-Bit OS"
   		 End If 
		If objProcessor.MaxClockSpeed < 1950 then
			iToolsMachinePerformanceWarnings = iToolsMachinePerformanceWarnings + 1
            Err.Clear
			WScript.Echo "Processor Frequency is not at the required level (at least 2 Ghz) - WARNING."
			objFile.Writeline "Processor Frequency is not at the required level (atleast 2 Ghz) - WARNING."
            If Err.number <> 0 Then
                Wscript.Echo "Could not write to trace log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
            End If
            Err.Clear
			objFileFail.Writeline "Processor Frequency is not at the required level (atleast 2 Ghz) - WARNING."
            If Err.number <> 0 Then
                Wscript.Echo "Could not write to trace error log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
            End If
            Err.Clear
		End If

		If objProcessor.NumberofCores  < 2 then
			iToolsMachinePerformanceWarnings = iToolsMachinePerformanceWarnings + 1
            Err.Clear
			WScript.Echo "Number of Processor Cores is not at the required level (atleast 2) - WARNING."
			objFile.Writeline "Number of Processor Cores is not at the required level (atleast 2) - WARNING."
            If Err.number <> 0 Then
                Wscript.Echo "Could not write to trace log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
            End If
            Err.Clear
			objFileFail.Writeline "Number of Processor Cores is not at the required level (atleast 2) - WARNING."
            If Err.number <> 0 Then
                Wscript.Echo "Could not write to trace error log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
            End If
		End If
		
	Next 
	Set ObjProcessor = Nothing:    Set ColSettings = Nothing
	Err.Clear
End Sub

'Sub to check OS version on local tools machine
Sub CheckOS(objWMI)
	Dim colOperatingSystems
	Dim objOperatingSystem
	Dim str_version,str_majorversion
	Set colOperatingSystems = objWMI.ExecQuery _
    	("Select * from Win32_OperatingSystem")

	For Each objOperatingSystem in colOperatingSystems
		
    		Wscript.Echo objOperatingSystem.Caption & _
    		"  " & objOperatingSystem.Version
		Wscript.Echo "Locale: " & objOperatingSystem.Locale
		objFile.Writeline objOperatingSystem.Caption & _
    		"  " & objOperatingSystem.Version
		objFile.Writeline "Locale: " & objOperatingSystem.Locale
		str_version= Split(objOperatingSystem.Version,".")
		str_majorversion=str_version(0)
		
		If CInt(str_majorversion) < 6 then
			CheckFailureCount = CheckFailureCount + 1
			Wscript.echo "Tools machine operating System not supported for CRM RaaS - FAILED."
            Err.Clear
			objFile.Writeline "Tools machine operating System not supported for CRM RaaS - FAILED."
            objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
            If Err.number <> 0 Then
                Wscript.Echo "Could not write to trace log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
            End If
            Err.Clear
			objFileFail.Writeline "Tools machine operating System not supported for CRM RaaS - FAILED."
            If Err.number <> 0 Then
                Wscript.Echo "Could not write to trace error log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
            End If
		End If
	Next
	
	Set colOperatingSystems=Nothing
	Err.Clear
End Sub

'Sub to get local machine's properties (RAM,Name,TimeZone)
Sub GetComputerProperties(objWMI)
	Dim colProperties
	Dim objComp
	Dim i_Mem
	Set colProperties = objWMI.ExecQuery ("Select * from Win32_ComputerSystem")
	For Each objComp in colProperties 
   		 Wscript.Echo "System Name: " & objComp.Name
    		 Wscript.Echo "Time Zone: " & objComp.CurrentTimeZone
   		 Wscript.Echo "Total Physical Memory: " & objComp.TotalPhysicalMemory
		 objFile.Writeline "System Name: " & objComp.Name
    		 objFile.Writeline "Time Zone: " & objComp.CurrentTimeZone
   		 objFile.Writeline "Total Physical Memory: " & objComp.TotalPhysicalMemory
	
		if objComp.TotalPhysicalMemory < 3800000000 then
		'if objComp.TotalPhysicalMemory < 4080218931 then

			CheckFailureCount = CheckFailureCount + 1
			WScript.Echo "Minimum of 4 GB RAM is required - FAILED."
            Err.Clear
			objFile.Writeline "Minimum of 4 GB RAM is required - FAILED."
            If Err.number <> 0 Then
                Wscript.Echo "Could not write to trace log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
            End If
            Err.Clear
			objFileFail.Writeline "Minimum of 4 GB RAM is required - FAILED."
            If Err.number <> 0 Then
                Wscript.Echo "Could not write to trace error log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
            End If

		End if
		 
	Next
	Err.Clear
End Sub

'Sub to check free disk space on tools machine
Sub CheckDiskspace(objWMI)
	Dim colDisks
	Dim objDisk
	Dim b_spacefound
	b_spacefound=false
	Set colDisks = objWMI.ExecQuery("Select * from Win32_LogicalDisk")
	For Each objDisk in colDisks
		'if local drive
		if objDisk.DriveType=3 then
    			Wscript.Echo "Free Disk Space on " & objDisk.Caption & " is " & objDisk.FreeSpace
			if objDisk.FreeSpace > 5368709120 then
				b_spacefound=true
			End if
		End if 
	Next
	If b_spacefound=false then
			CheckFailureCount = CheckFailureCount + 1
            Err.Clear
			WScript.Echo "Minimum of 5 GB free Disk space is required - FAILED."
			objFile.Writeline "Minimum of 5 GB free Disk space is required - FAILED."
            If Err.number <> 0 Then
                Wscript.Echo "Could not write to trace log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
            End If
            Err.Clear
			objFileFail.Writeline "Minimum of 5 GB Disk free space is required - FAILED."
            If Err.number <> 0 Then
                Wscript.Echo "Could not write to trace error log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
            End If
	End if
	Err.Clear
End Sub

'Sub to check port open on remote server
Sub CheckPort(objShell,sServer,nPort)
	Dim strOutPortCheck
	Dim  strOut
	strOutPortCheck=Exec(objShell,"powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile .\testport.ps1 " & sServer & " " & nPort & "; exit $LASTEXITCODE")

	If InStr(strOutPortCheck, "False")>0 or InStr(strOutPortCheck, "Connection to Port Timed Out")>0 or InStr(strOutPortCheck,": «")>0 Then
		CheckFailureCount = CheckFailureCount + 1

		WScript.Echo "Connectivity check on port " & nPort & " on server:<" & sServer & "> - FAILED."
        Err.Clear
		objFile.Writeline "Connectivity check on port " & nPort & " on server:<" & sServer & "> - FAILED."
        If Err.number <> 0 Then
            Wscript.Echo "Could not write to trace log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
        End If

        Err.Clear
		objFileFail.Writeline "Connectivity check on port " & nPort & " on server:<" & sServer & "> - FAILED."
        If Err.number <> 0 Then
            Wscript.Echo "Could not write to trace error log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
        End If

        CheckFailureCount = CheckFailureCount + 1
		Err.Clear
	Else
		WScript.Echo "Connectivity check - port " & nPort & " is open on server:<" & sServer & "> - OK."
		objFile.Writeline "Connectivity check - port " & nPort & " is open on server:<" & sServer & "> - OK."
	End if	
End Sub

'Sub to check If current user is local admin on computer

Function CheckLocalAdmin(strComputer)
Dim objNetwork
Dim objGroup
Dim bfound
Dim objUser

Set objNetwork = CreateObject("Wscript.Network")
'strComputer = objNetwork.ComputerName
strUser = objNetwork.UserName
bfound=false

'English
Set objGroup = GetObject("WinNT://" & strComputer & "/Administrators")
For Each objUser in objGroup.Members
    If objUser.Name = strUser Then
	    bfound=true        
    End If
Next
'French
If not bfound Then
    Set objGroup = GetObject("WinNT://" & strComputer & "/Administrateurs")
    For Each objUser in objGroup.Members
        If objUser.Name = strUser Then
	        bfound=true        
        End If
    Next
End If
'Spanish
If not bfound Then
    Set objGroup = GetObject("WinNT://" & strComputer & "/Administradores")
    For Each objUser in objGroup.Members
        If objUser.Name = strUser Then
	        bfound=true        
        End If
    Next
End If
'German
If not bfound Then
    Set objGroup = GetObject("WinNT://" & strComputer & "/Administratoren")
    For Each objUser in objGroup.Members
        If objUser.Name = strUser Then
	        bfound=true        
        End If
    Next
End If
'Italian
If not bfound Then
    Set objGroup = GetObject("WinNT://" & strComputer & "/Amministratori")
    For Each objUser in objGroup.Members
        If objUser.Name = strUser Then
	        bfound=true        
        End If
    Next
End If
'Swedish
If not bfound Then
    Set objGroup = GetObject("WinNT://" & strComputer & "/Administratörer")
    For Each objUser in objGroup.Members
        If objUser.Name = strUser Then
	        bfound=true        
        End If
    Next
End If
'Rusian
If not bfound Then
    Set objGroup = GetObject("WinNT://" & strComputer & "/??????????????")
    For Each objUser in objGroup.Members
        If objUser.Name = strUser Then
	        bfound=true        
        End If
    Next
End If
'Norwegian
If not bfound Then
    Set objGroup = GetObject("WinNT://" & strComputer & "/Administratorer")
    For Each objUser in objGroup.Members
        If objUser.Name = strUser Then
	        bfound=true        
        End If
    Next
End If
'Nethernlands
If not bfound Then
    Set objGroup = GetObject("WinNT://" & strComputer & "/beheerders")
    For Each objUser in objGroup.Members
        If objUser.Name = strUser Then
	        bfound=true        
        End If
    Next
End If
'Finland
If not bfound Then
    Set objGroup = GetObject("WinNT://" & strComputer & "/ylläpitäjät")
    For Each objUser in objGroup.Members
        If objUser.Name = strUser Then
	        bfound=true        
        End If
    Next
End If
'Hongary
If not bfound Then
    Set objGroup = GetObject("WinNT://" & strComputer & "/rendszergazdák")
    For Each objUser in objGroup.Members
        If objUser.Name = strUser Then
	        bfound=true        
        End If
    Next
End If
'Lituany
If not bfound Then
    Set objGroup = GetObject("WinNT://" & strComputer & "/administratoriai")
    For Each objUser in objGroup.Members
        If objUser.Name = strUser Then
	        bfound=true        
        End If
    Next
End If
'Malta
If not bfound Then
    Set objGroup = GetObject("WinNT://" & strComputer & "/amministraturi")
    For Each objUser in objGroup.Members
        If objUser.Name = strUser Then
	        bfound=true        
        End If
    Next
End If
'Slovenia
If not bfound Then
    Set objGroup = GetObject("WinNT://" & strComputer & "/administratorji")
    For Each objUser in objGroup.Members
        If objUser.Name = strUser Then
	        bfound=true        
        End If
    Next
End If



If strComputer = "." Then
    strComputer = "local tool machine"
End If
If bfound=true then
	Wscript.Echo "Currently logged on user:<" & strUser  & "> is member of Windows Administrators group on:<" & strComputer & "> - OK."
	objFile.Writeline "Currently logged on user:<" & strUser & " is member of Windows Administrators group on:<" & strComputer & "> - OK."
    CheckLocalAdmin = true
Else
	CheckFailureCount = CheckFailureCount + 1
	Wscript.Echo "Currently logged on user:<" & strUser & " is NOT member of Windows Administrators group on:<" & strComputer & "> - FAILED."
    Err.Clear
	objFile.Writeline "Currently logged on user:<" & strUser & " is NOT member of Windows Administrators group on:<" & strComputer & "> - FAILED."
    If Err.number <> 0 Then
        Wscript.Echo "Could not write to trace log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
    End If
    Err.Clear
	objFileFail.Writeline "Currently logged on user:<" & strUser & " is NOT member of Windows Administrators group on:<" & strComputer & "> - FAILED."
    If Err.number <> 0 Then
        Wscript.Echo "Could not write to trace error log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
    End If
    CheckLocalAdmin = false
    Err.Clear	
End If

End Function

'Sub to see If server can be pinged
Sub CheckPing(objShell,sServer)
    Dim strOutPing,strResult, ExitCode

	ExitCode=ExecEx(objShell,"ping -n 2 " & sServer)

	If ExitCode = 0 then
		WScript.Echo "Ping successful to server:<" & sServer & "> - OK."
		objFile.Writeline "Ping successful to server:<" & sServer & "> - OK."
	Else
		CheckFailureCount = CheckFailureCount + 1
		WScript.Echo "Ping Failed to Server <" & sServer & "> - FAILED."
        Err.Clear
		objFile.Writeline "Ping Failed to Server <" & sServer & "> - FAILED."
        If Err.number <> 0 Then
            Wscript.Echo "Could not write to trace log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
        End If
        Err.Clear
		objFileFail.Writeline "Ping Failed to Server <" & sServer & "> - FAILED."
        If Err.number <> 0 Then
            Wscript.Echo "Could not write to trace error log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
        End If
	End if
	Err.Clear
End Sub
'Function to check for access to Admin$ admin share on remote server

Function CheckAdminShare_Admin(strComputer,Share)
	Err.Clear
	Dim strServerShare
	Dim objFSO
	Dim objDirectory
	strServerShare = "\\" + strComputer+ "\" + Share + "$"
	strCheckAdminShare_Admin_Result= 1
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	On error resume next
	
	Set objDirectory = objFSO.GetFolder(strServerShare)

	If err.number<> 0 then
		bSuccessfulOutput=false
		CheckFailureCount = CheckFailureCount + 1
		strCheckAdminShare_Admin_Result= 0
		WScript.Echo "Failure - Admin Share " & Share & "$ cannot be accessed on server:<" & strComputer & "> - FAILED."
		Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
		objFile.Writeline "Failure - Admin Share " & Share & "$ cannot be accessed on server:<" & strComputer & "> - FAILED."
		objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
        Err.Clear
		objFileFail.Writeline "Failure - Admin Share " & Share & "$ cannot be accessed on server:<" & strComputer & "> - FAILED."
        If Err.number <> 0 Then
            Wscript.Echo "Could not write to trace log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
        End If
        Err.Clear
		objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
		Err.Clear
        If Err.number <> 0 Then
            Wscript.Echo "Could not write to trace error log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
        End If
	Else
		WScript.Echo "Admin Share " & Share & "$ can be accessed on server:<" & strComputer & "> - OK."
        Err.Clear
		objFile.Writeline "Admin Share " & Share & "$ can be accessed on server:<" & strComputer & "> - OK."
        If Err.number <> 0 Then
            Wscript.Echo "Could not write to trace log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
        End If
	End If

	Set objDirectory = Nothing
	Set objFSO = Nothing
	On error goto 0
	 CheckAdminShare_Admin = strCheckAdminShare_Admin_Result
End Function


'Procedure to check If PowerShell 2.0 is installed on local tools machine

Sub CheckPowerShell2Installed(objShell)
	Dim strRegValue
	On Error resume next
	strRegValue = objShell.RegRead(REG_POWERSHELL_2)
	
	If Err.Number <> 0 Then
		bSuccessfulOutput=false
        CheckFailureCount = CheckFailureCount + 1
	   	WScript.Echo "Checking for Powershell Installed on local tools machine - FAILED."
        Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
        Err.Clear
		objFile.Writeline "Checking for Powershell Installed on local tools machine - FAILED."
        If Err.number <> 0 Then
            Wscript.Echo "Could not write to trace log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
        End If
        objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
        Err.Clear
		objFileFail.Writeline "Checking for Powershell Installed on local tools machine - FAILED."
        objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description 
        If Err.number <> 0 Then
            Wscript.Echo "Could not write to trace error log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
        End If 
		Err.Clear      
	End If


	If strRegValue < "2.0" Then
	    bSuccessfulOutput=false
            CheckFailureCount = CheckFailureCount + 1
	    WScript.Echo "PowerShell 2.0 not found on local tools machine - FAILED."
	    objFile.Writeline "PowerShell 2.0 not found on local tools machine - FAILED."
	    objFileFail.Writeline "PowerShell 2.0 not found on local tools machine - FAILED."
	Else
		WScript.Echo "PowerShell 2.0 installed on local tools machine - OK."
		objFile.Writeline "PowerShell 2.0 installed on local tools machine - OK."   
	End If

	On error goto 0

End Sub

'Procedure to check If Remote Registry Service is running on remote server

Function CheckRemoteService(objcolServices,strService,strComputer)
	Dim objService 
	For Each objService In objcolServices
   	 	If objService.Name = strService Then
       		If objService.State<>"Running" then 
                bSuccessfulOutput=false
                CheckFailureCount = CheckFailureCount + 1
	        	WScript.Echo strService & " Service not running on server:<" &  strComputer  & "> - FAILED."
		        objFile.Writeline strService & " Service not running on server:<" &  strComputer  & "> - FAILED."
		        objFileFail.Writeline strService & " Service not running on server:<" &  strComputer  & "> - FAILED."
                CheckRemoteService = false
			Else
		 		WScript.Echo strService & " Service is running on server:<" &  strComputer  & "> - OK."
				objFile.Writeline strService & " Service is running on server:<" &  strComputer  & "> - OK."
                CheckRemoteService = true
			End if
    	End If
	Next
End Function

'Procedure to run a shell command
Sub RunCommand(objShell,strCommand)
 	objShell.Run strCommand,1,true 
End Sub


Sub CheckSysAdminRole (sSQLInstanceName)
	On error resume next
	
	Dim bSuccededQueryedSQL
	bSuccededQueryedSQL = false

	Set oSQLConn = CreateObject("ADODB.Connection")
	Set rs = CreateObject("ADODB.Recordset")

	oSQLConn.Open "Provider=SQLOLEDB;Data Source=" & sSQLInstanceName & ";Trusted_Connection=Yes;Initial Catalog=master"

	If Err.number <> 0 Then
		bSQLCheckSysAdmin = false
		Wscript.Echo "Failed connecting to SQL instance <" & sSQLInstanceName & "> to verify if current user has SQL sysadmin - FAILED."
		Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
		objFile.Writeline "Failed connecting to SQL instance <" & sSQLInstanceName & "> to verify if current user has SQL sysadmin - FAILED."
		objFileFail.Writeline "Failed connecting to SQL instance <" & sSQLInstanceName & "> to verify if current user has SQL sysadmin - FAILED."
		objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
		objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
		Err.Clear
	Else
		Set rs = oSQLConn.Execute ("SELECT IS_SRVROLEMEMBER ('sysadmin') ")
		If Err.number <> 0 Then
			bSQLCheckSysAdmin = false
			CheckFailureCount = CheckFailureCount + 1
			Wscript.Echo "Failed executing IS_SRVROLEMEMBER procedure on <" & sSQLInstanceName & "> to verify if current user has SQL sysadmin - FAILED."
			Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
			objFile.Writeline "Failed executing IS_SRVROLEMEMBER procedure on <" & sSQLInstanceName & "> to verify if current user has SQL sysadmin - FAILED."
			objFileFail.Writeline "Failed executing IS_SRVROLEMEMBER procedure on <" & sSQLInstanceName & "> to verify if current user has SQL sysadmin - FAILED."
			objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
			objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
			Err.Clear
		Else
			bSuccededQueryedSQL = true
		End If
		
	End If

	If bSQLCheckSysAdmin = true Then	
		If bSuccededQueryedSQL = true and not rs.EOF Then
			rs.MoveFirst
			If rs(0).Value = 1 Then
				Wscript.Echo "Current user has SQL sysadmin role on <" & sSQLInstanceName & ">  - OK."
				objFile.Writeline "Current user has SQL sysadmin role on <" & sSQLInstanceName & "> - OK."
				oSQLConn.Close
				Set rs = Nothing
				Set oSQLConn = Nothing
			Else
				CheckFailureCount = CheckFailureCount + 1
				Wscript.Echo "Current user does not have SQL sysadmin role on <" & sSQLInstanceName & "> - FAILED."
				objFile.Writeline "Current user does not have SQL sysadmin role on <" & sSQLInstanceName & "> - FAILED."
				objFileFail.Writeline "Current user does not have SQL sysadmin role on <" & sSQLInstanceName & "> - FAILED."
			End If		
		End If
	End If

	Err.Clear

End Sub

'Sub to check If iis logging module is installed on server
Sub CheckIISLoggingInstalled(objReg,strComputer)
	Dim strRegValue
	On Error resume next
	'strRegValue = objShell.RegRead(REG_HTTP_Logging)
	objReg.GetDWORDValue HKEY_LOCAL_MACHINE, REG_HTTP_Logging,"HttpLogging",strRegValue
	
	If Err.Number <> 0  Or strRegValue<> 1 Then
		bSuccessfulOutput=false
		CheckFailureCount = CheckFailureCount + 1
	    WScript.Echo "IIS logging is not installed on server:<" &  strComputer  & "> - FAILED."
		objFile.Writeline "IIS logging is not installed on server:<" &  strComputer  & "> - FAILED."
		objFileFail.Writeline "IIS Logging is not installed on server:<" &  strComputer  & "> - FAILED."
    Else
		WScript.Echo "IIS logging is installed on server:<" &  strComputer  & "> - OK." 
	    objFile.Writeline "IIS logging is installed on server:<" &  strComputer  & "> - OK."  
	End If
	
End Sub

'Sub to check If IIS logging enabled on website, log format is set to w3c  and then check If logs exist
Sub CheckIISLogsEnabledandFormat(strServer)
Dim oIIS
Dim oSites
Dim oSite
Dim strS1
Dim v_logsCheck
Dim oSection
Dim strShare
Dim strS

Set oIIS = GetObject("winmgmts:{impersonationLevel=impersonate,authenticationLevel=pktPrivacy}!\\" + strServer + "\root\WebAdministration")

Set oSites = oIIS.InstancesOf("Site")

For Each oSite In oSites            
     
     Set oSection = oIIS.Get("HttpLoggingSection.Path=" & "'MACHINE/WEBROOT/APPHOST',Location='" & oSite.Name &  "'")
	
     If OSection.DontLog = false then
		WScript.Echo "IIS logging is enabled on website '" & oSite.Name & "' on server:<" & strServer & "> - OK."
		objFile.Writeline "IIS logging is enabled on website '" & oSite.Name & "' on server:<" & strServer & "> - OK."
		
        If oSite.LogFile.LogFormat = 2 then
			wscript.echo "IIS logging format is set to w3c on website '" & oSite.Name & "' on server:<" & strServer & "> - OK."
			objFile.Writeline "IIS logging format is set to w3c on website '" & oSite.Name & "' on server:<" & strServer & "> - OK."
			strS=oSite.LogFile.Directory
			strS=Replace(oSite.LogFile.Directory,"%SystemDrive%", "C:")
			strS1=Replace(strS,":","$")
            strShare= "\\" & strServer & "\" & strS1 & "\W3SVC" & oSite.ID
			v_logsCheck=CheckIISLogsExist(strShare)
			
            If v_logsCheck=true then
				wscript.echo "IIS logs exist for website '" & oSite.Name & "' at " & strS & "\W3SVC" & oSite.ID & " on server:<" & strServer & "> - OK."
				objFile.Writeline "IIS logs exist for website '" & oSite.Name & "' at " & strS & "\W3SVC" & oSite.ID  & " on server:<" & strServer & "> - OK."
			Else
				iIISLogsFailures = iIISLogsFailures + 1
				wscript.echo "IIS logs DO NOT exist for website '" & oSite.Name & "' at " & strS & "\W3SVC" & oSite.ID & " on server:<" & strServer & "> - WARNING."
				objFile.Writeline "IIS logs DO NOT exist for website '" & oSite.Name & "' at " & strS & "\W3SVC" & oSite.ID & " on server:<" & strServer & "> - WARNING."
				objFileFail.Writeline "IIS logs DO NOT exist for website '" & oSite.Name & "' at " & strS & "\W3SVC" & oSite.ID & " on server:<" & strServer & "> - WARNING."
			End If
			CheckIISFields oSite,strServer
     		Else
			iIISLogsFailures = iIISLogsFailures + 1
	        wscript.echo "IIS logging format is NOT set to w3c for website '" & oSite.Name & "' on server:<" & strServer & "> - WARNING."
			objFile.Writeline "IIS logging format is NOT set to w3c for website '" & oSite.Name & "' on server:<" & strServer & "> - WARNING."
			objFileFail.Writeline "IIS logging format is NOT set to w3c for website '" & oSite.Name & "' on server:<" & strServer & "> - WARNING."
     		End If
     Else
		iIISLogsFailures = iIISLogsFailures + 1
	    WScript.Echo "IIS logging has been disabled on website '" & oSite.Name & "' on server:<" & strServer & "> - WARNING."
		objFile.Writeline "IIS logging has been disabled on website '" & oSite.Name & "' on server:<" & strServer & "> - WARNING."
		objFileFail.Writeline "IIS logging has been disabled on website '" & oSite.Name & "' on server:<" & strServer & "> - WARNING."
     End If
	Set oSection=Nothing     
Next
Set oSites = Nothing
Set oIIS = Nothing

End Sub

'Sub to check If IIS logging fields date, time, sc-status, uri stem, time taken are enabled
Sub CheckIISFields(oSite,strServer)
Dim iFlag
Dim iFailCount
iFailCount = 0
iFlag = oSite.LogFile.LogExtFileFlags
	
	If ((iFlag AND 1)=0) then
		CheckFailureCount = CheckFailureCount + 1
		iFailCount=iFailCount+1
		Wscript.Echo "IIS required logging field date NOT Enabled for website '" & oSite.Name & "' on server:<" & strServer & "> - FAILED."
		objFile.Writeline "IIS required logging field date Date NOT Enabled for website '" & oSite.Name & "' on server:<" & strServer & "> - FAILED."
		objFileFail.Writeline "IIS required logging field date Date NOT Enabled for website '" & oSite.Name & "' on server:<" & strServer & "> - FAILED."
	
	End If
	If ((iFlag AND 2)=0) then
		iFailCount=iFailCount+1
		CheckFailureCount = CheckFailureCount + 1
		Wscript.Echo "IIS required logging field Time NOT Enabled for website '" & oSite.Name & "' on server:<" & strServer & "> - FAILED."
		objFile.Writeline "IIS required logging field date Time NOT Enabled for website '" & oSite.Name & "' on server:<" & strServer & "> - FAILED."
		objFileFail.Writeline "IIS required logging field Time NOT Enabled for website '" & oSite.Name & "' on server:<" & strServer & "> - FAILED."
	
	End If
	
	If ((iFlag AND 256)=0) then
		iFailCount=iFailCount+1
		CheckFailureCount = CheckFailureCount + 1
		Wscript.Echo "IIS required logging field Uri Stem NOT Enabled for website '" & oSite.Name & "' on server:<" & strServer & "> - FAILED."
		objFile.Writeline "IIS required logging field Uri Stem NOT Enabled for website '" & oSite.Name & "' on server:<" & strServer & "> - FAILED."
		objFileFail.Writeline "IIS required logging field Uri Stem NOT Enabled for website '" & oSite.Name & "' on server:<" & strServer & "> - FAILED."
	End If
	
	If ((iFlag AND 1024)=0) then
		iFailCount=iFailCount+1
		CheckFailureCount = CheckFailureCount + 1
		Wscript.Echo "IIS required logging field Status NOT Enabled for website '" & oSite.Name & "' on server:<" & strServer & "> - FAILED."
		objFile.Writeline "IIS required logging field Status NOT Enabled for website '" & oSite.Name & "' on server:<" & strServer & "> - FAILED."
		objFileFail.Writeline "IIS required logging field Status NOT Enabled for website '" & oSite.Name & "' on server:<" & strServer & "> - FAILED."
	     
	End If
	
	If ((iFlag AND 16384)=0) then
		iFailCount=iFailCount+1
		CheckFailureCount = CheckFailureCount + 1
		Wscript.Echo "IIS required logging field Time Taken NOT Enabled for website '" & oSite.Name & "' on server:<" & strServer & "> - FAILED."
		objFile.Writeline "IIS required logging field Time Taken NOT Enabled for website '" & oSite.Name & "' on server:<" & strServer & "> - FAILED."
		objFileFail.Writeline "IIS required logging field Time Taken NOT Enabled for website '" & oSite.Name & "' on server:<" & strServer & "> - FAILED."
	End If
	
	If iFailCount=0 then
		Wscript.Echo "IIS required logging fields enabled for website '" & oSite.Name & "' on server:<" & strServer & "> - OK."
		objFile.Writeline "IIS required logging fields enabled for website '" & oSite.Name & "' on server:<" & strServer & "> - OK."
	End If




End Sub 


'Function to check If iis logs for website exist within past 7 days

Function CheckIISLogsExist(strShare)
	On error resume next
	Dim oRecordSet
	Dim oRecord
	Dim v_LogExist
	Dim strQuery
	Dim InputFormat
	Set oLogQuery = CreateObject("MSUtil.LogQuery")
        
	' Create InputFormat object
	Set InputFormat = CreateObject("MSUtil.LogQuery.FileSystemInputFormat")

	' Create query 
	strQuery = "SELECT Path,Name,creationtime FROM '" & strShare & "\*.log' Where CreationTime > SUB(TO_LOCALTIME(SYSTEM_TIMESTAMP()), TIMESTAMP('0000-01-07 00:00', 'yyyy-MM-dd HH:mm'))  order by CreationTime Desc"
	' Execute query 
	Set oRecordSet = oLogQuery.Execute ( strQuery, InputFormat )
	v_LogExist=False
	If err.number=0 then

	' Loop records
	Do while NOT oRecordSet.atEnd
      
		' Get record
		Set oRecord = oRecordSet.getRecord
	   	v_LogExist=true

		' next record
		oRecordSet.moveNext

	loop

	End if
	' Close RecordSet
	oRecordSet.close
	Set oRecordSet=Nothing
	Set InputFormat=Nothing
	Set oLogQuery=Nothing
	Err.Clear
 	CheckIISLogsExist= v_LogExist
End Function


On Error Resume Next

bWMITestFailed = false
bIISFound = false
bIISFound = false
bSuccessfulOutput=true
CheckFailureCount = 0


strCheckAdminShare_Admin_Result=1
'Create shell object
Set oShell = CreateObject("WScript.Shell")

If Err.Number <> 0 Then
	bSuccessfulOutput=false
	WScript.Echo "Creating a WScript.Shell object. - FAILED."
   	Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
	Wscript.Echo "Exiting <Press enter key to quit>."
    WScript.StdIn.ReadLine
   	WScript.Quit
End If


'Test If we are running as cscript (not wscript) If it is run it as cscipt
If InStr(1, WScript.FullName, "WScript.exe", vbTextCompare) <> 0 Then
        If Err.Number <> 0 Then
    		bSuccessfulOutput=false
            	CheckFailureCount = CheckFailureCount + 1
	        WScript.Echo "Failure - accessing WScript"
            	Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
	        Wscript.Echo "Exiting <Press enter key to quit>."
            	WScript.StdIn.ReadLine
        End If
        oShell.Run "%comspec% /c cscript /nologo """ & WScript.ScriptFullName & """", 1, False
        WScript.Quit(0)
End If

'Creating Log File

Set objFileSys = CreateObject("Scripting.FileSystemObject")
Set objFile = objFileSys.OpenTextFile("CRMRaasPreReqsScriptLog.txt", 8,1)
Set objFileFail= objFileSys.OpenTextFile("CRMRaasPreReqsFailuresLog.txt", 8,1)
objFile.writeLine " "
Err.Clear
If Err.number <> 0 Then
    Wscript.Echo "Could not write to trace log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
End If
objFile.Writeline "==========================================================================================="
objFile.writeLine " "
objFile.writeLine "				Log Started On " & Now
objFile.writeLine " "
objFile.Writeline "==========================================================================================="
objFile.writeLine " "

objFileFail.writeLine " "
Err.Clear
If Err.number <> 0 Then
    Wscript.Echo "Could not write to trace error log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
End If
objFileFail.Writeline "==========================================================================================="
objFileFail.writeLine " "
objFileFail.writeLine "				Log Started On " & Now
objFileFail.writeLine " "
objFileFail.Writeline "==========================================================================================="
objFileFail.writeLine " "

Wscript.Echo "Starting CRM RaaS prerequisites sample test script (Ctrl-C to stop script at any point in time)."
Wscript.Echo ""
objFile.Writeline "Starting CRM RaaS prerequisites sample test script (Ctrl-C to stop script at any point in time)."
objFile.WriteLine " "

Set args = Wscript.Arguments

'Dumping script arguments
For j=0 to args.Count-1
	sArgs = sArgs & args(j)
Next

Wscript.Echo "Script arguments :<" & sArgs & ">."
Wscript.Echo ""
objFile.Writeline "Script arguments :<" & sArgs & ">."
objFile.Writeline "Script arguments :<" & sArgs & ">."


If args.Count = 1 Then
	sArgs = args(0)
	
	If LCase(sArgs)  = "sysadmin"  Then
	
	'If LCase(sArgs[1]) = "sysadmin"  Then
	
			Wscript.Echo "Performing sysadmin check only."
			Wscript.Echo "----------------------------------------"
			objFile.Writeline "Performing sysadmin check only."
			objFile.Writeline "---------------------------"
	
	
			Wscript.Echo ""
			Wscript.Echo "*** Your input is required you we will now be performing sysadmin role check for this machine. Please enter SQL Instance name (for example Server\InstanceName,1433):" 
			sSQLInstanceName = WScript.StdIn.ReadLine
			If sSQLInstanceName =  "" Then
				sSQLInstanceName = sServerName
			End If
			bSQLCheckSysAdmin = true

			' perform sysadmin role check
			CheckSysAdminRole ( sSQLInstanceName )	
	
			
			Wscript.Echo "Exiting sysadmin check mode."
			Wscript.Echo "----------------------------------------"
			objFile.Writeline "Exiting sysadmin check mode."
			objFile.Writeline "---------------------------"	

			objFile.Close
			objFileFail.Close

			Set objFileFail=Nothing
			Set objFile=Nothing
			Set objFileSys=Nothing
			Set args=Nothing
			Set oShell=Nothing
			

			Wscript.Quit
			
	End If
	
End If

'Perform local tests
Wscript.Echo "1 - Performing tool machine local tests."
Wscript.Echo "----------------------------------------"

objFile.Writeline "1 - Performing local tests."
objFile.Writeline "---------------------------"


'Check If current user is local admin on tools machine
CheckLocalAdmin (".")

'List all installed software
LocalListAllInstalledSoftware = ListAllInstalledSoftware (".")

Err.Clear

Wscript.Echo "Now testing local call to WMI CIMv2 object..."
objFile.Writeline "Now testing local call to WMI CIMv2 object..."
objFileFail.Writeline "Now testing local call to WMI CIMv2 object..."


'Create local WMI CIMv2 object
Set oWMI = GetObject("winmgmts:root\CIMv2")

If Err.Number <> 0 Then
	bSuccessfulOutput=false
    CheckFailureCount = CheckFailureCount + 1
	Wscript.Echo "Could not access WMI CIMv2' on tools machine. - FAILED."
	Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
	objFile.Writeline "Could not access WMI CIMv2' on tools machine. - FAILED."
	objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
	objFileFail.Writeline "Could not access WMI CIMv2' on tools machine. - FAILED."
	objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
    Err.Clear
Else
	Wscript.Echo "Local CIMv2 WMI call - OK."
	objFile.Writeline "Local CIMv2 WMI call - OK."
End If


    'Get tools machine time on machine because if there is a significant time difference WMI may not work
    Set colItems = oWMI.ExecQuery("Select * from Win32_OperatingSystem",,48)
    If Err.number = 0 Then
        For Each objItem in colItems
            Wscript.Echo "Local machine CurrentTimeZone is: " & objItem.CurrentTimeZone
            Wscript.Echo "Local machine LocalDateTime is: " & objItem.LocalDateTime
            objFile.Writeline "Local machine CurrentTimeZone is: " & objItem.CurrentTimeZone
            objFile.Writeline "Local machine LocalDateTime is: " & objItem.LocalDateTime
            strLocalMachineTimeZone = objItem.CurrentTimeZone
            strLocalMachineDateTime = objItem.LocalDateTime
        Next

    Else

        CheckFailureCount = CheckFailureCount + 1
    	Wscript.Echo "An error occured when retriving time tools machine (step1):<" &  sServerName  & "> using WMI - FAILED."
	    Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
	    objFile.Writeline "An error occured when retriving time tools machine (step1):<" &  sServerName  & "> using WMI - FAILED."
        objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
	    objFileFail.Writeline "An error occured when retriving time tools machine (step1):<" &  sServerName  & "> using WMI - FAILED."
        objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
        Err.Clear

    End If



'Check if the user is using a Domain account (not a local account)
Err.Clear

Set oLoggedOnUsers = oWMI.ExecQuery ("select * from Win32_LoggedOnUser") 
If Err.number <> 0 Then
	bSuccessfulOutput=false
    CheckFailureCount = CheckFailureCount + 1
	Wscript.Echo "Could not retrieve the list of logged on users to verify if current user is logged on as a domain account. - FAILED."
	Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
	objFile.Writeline "Could not retrieve the list of logged on users to verify if current user is logged on as a domain account. - FAILED."
	objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
	objFileFail.Writeline "Could not retrieve the list of logged on users to verify if current user is logged on as a domain account. - FAILED."
	objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
    Err.Clear
End If

sComputerName = oShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )

bFoundDomainUser = false

For Each objLoggedOnUser in oLoggedOnUsers 
    If InStr (objLoggedOnUser.Antecedent, sComputerName)  = 0 and InStr (objLoggedOnUser.Antecedent, strUser) > 0 Then
        bFoundDomainUser = true
        If bEnableVerboseMode = true Then
            Wscript.Echo "Found logged on user session:" & objLoggedOnUser.Antecedent & " for user:" & strUser
        End If
    End If
Next 

Set oLoggedOnUsers = Nothing

If bFoundDomainUser = false Then
	bSuccessfulOutput=false
    CheckFailureCount = CheckFailureCount + 1
	Wscript.Echo "Could not find a logged on domain session for the current user:<"& strUser &">, please verify you are logged on with a DOMAIN account - FAILED."
	objFile.Writeline "Could not find a logged on domain session for the current user:<"& strUser &">, please verify you are logged on with a DOMAIN account - FAILED."
	objFileFail.Writeline "Could not find a logged on domain session for the current user:<"& strUser &">, please verify you are logged on with a DOMAIN account - FAILED."
    objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
    objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
    Err.Clear
Else
    Wscript.Echo "Current user:<"& strUser &"> is logged on with a domain account - OK."
    objFile.Writeline "Current user:<"& strUser &"> is logged on with a domain account - OK."
End If

Err.Clear


'Check for folder redirection
If bEnableVerboseMode = true Then
    Wscript.Echo "Checking for no Windows profile folder redirection..."
End If

bFolderRedirectionDetected = false

Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
oReg.EnumValues HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", arrValueNames, arrValueTypes


If Err.number <> 0 Then
    bFolderRedirectionDetected = true
    Wscript.Echo "User Shell Folders registry keys not found, this is unexpected (check that the domain user being used has local administrative priviledges)... - FAILED."
    objFile.Writeline  "User Shell Folders registry keys not found, this is unexpected (check that the domain user being used has local administrative priviledges)... - FAILED."
    objFileFail.Writeline "User Shell Folders registry keys not found, this is unexpected (check that the domain user being used has local administrative priviledges)... - FAILED."
    objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
    objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
End If

If bFolderRedirectionDetected = false and not IsNull(arrValueNames) Then

	For i=0 To UBound(arrValueNames)
	    oReg.GetStringValue HKEY_CURRENT_USER,"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", arrValueNames(i),strValue
        If InStr(strValue, "\\") > 0 Then
    		Wscript.Echo "Unsupported RaaS configuration detected Windows folder redirection detected for this user profile:<" & arrValueNames(i) & "> Value:<" & strValue & "> - FAILED."
            objFile.Writeline  "Unsupported RaaS configuration detected Windows folder redirection detected for this user profile:<" & arrValueNames(i) & "> Value:<" & strValue & "> - FAILED."
            objFileFail.Writeline  "Unsupported RaaS configuration detected Windows folder redirection detected for this user profile:<" & arrValueNames(i) & "> Value:<" & strValue & "> - FAILED."
            objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
            objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
            bFolderRedirectionDetected = true
        End If
 		
	Next

Else
	Wscript.Echo "Windows profile folder redirection - no User Shell Folders registry key found, this is unexpected (check that the domain user being used has local administrative priviledges)... - FAILED."
    objFile.Writeline "Windows profile folder redirection - no User Shell Folders registry key found, this is unexpected (check that the domain user being used has local administrative priviledges)... - FAILED."
    objFileFail.Writeline "Windows profile folder redirection - no User Shell Folders registry key found, this is unexpected (check that the domain user being used has local administrative priviledges)... - FAILED."
    objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
    objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
    bFolderRedirectionDetected = true
End If

Err.Clear


If bFolderRedirectionDetected = false Then
		Wscript.Echo "Windows profile folder redirection check for user <"& strUser &"> - OK."
		objFile.WriteLine "Windows profile folder redirection check for user <"& strUser &"> - OK."
Else
		Wscript.Echo "Windows profile folder redirection check for user <"& strUser &"> - FAILED."
		objFile.WriteLine "Windows profile folder redirection check for user <"& strUser &"> - FAILED."
        objFileFail.WriteLine "Windows profile folder redirection check for user <"& strUser &"> - FAILED."
        objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
        objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
        CheckFailureCount = CheckFailureCount + 1
End If
Err.Clear

Set oReg = Nothing



'Check for machine pending reboot for update to be applied
Wscript.Echo "Checking for machine pending reboot for updates to be applied..."
iRequiredUpdates = 0

bPendingRebootDetected = false

Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
oReg.EnumValues HKEY_LOCAL_MACHINE , "SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired", arrValueNames, arrValueTypes


If Err.number <> 0 Then
    bPendingRebootDetected = true
    Wscript.Echo "User Auto Update\RebootRequired registry keys not found, this is unexpected (check that the domain user being used has local administrative priviledges)... - FAILED."
    objFile.Writeline  "User Auto Update\RebootRequired registry keys not found, this is unexpected (check that the domain user being used has local administrative priviledges)... - FAILED."
    objFileFail.Writeline "UserAuto Update\RebootRequired registry keys not found, this is unexpected (check that the domain user being used has local administrative priviledges)... - FAILED."
    objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
    objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
End If

If bPendingRebootDetected = false Then
    If not IsNull (arrValueNames) Then
	    For i=0 To UBound(arrValueNames)
	        oReg.GetStringValue HKEY_LOCAL_MACHINE ,"SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired", arrValueNames(i),strValue
    		    Wscript.Echo "Windows pending reboot for:<" & arrValueNames(i) & "> Value:<" & strValue & "> - FAILED."
                objFile.Writeline  "Windows pending reboot for:<" & arrValueNames(i) & "> Value:<" & strValue & "> - FAILED."
                objFileFail.Writeline  "Windows pending reboot for:<" & arrValueNames(i) & "> Value:<" & strValue & "> - FAILED."
                iRequiredUpdates = iRequiredUpdates +1
                bPendingRebootDetected = true
	    Next
    End If
Else
	Wscript.Echo "Windows pending reboot check - FAILED."
    objFile.Writeline "Windows pending reboot check - FAILED."
    objFileFail.Writeline "Windows pending reboot check - FAILED."
    bPendingRebootDetected = true
End If

Err.Clear


If bPendingRebootDetected = false Then
		Wscript.Echo "No pending reboot for updates to be applied - OK."
		objFile.WriteLine "No pending reboot for updates to be applied - OK."
Else
		Wscript.Echo "There are <"& iRequiredUpdates &"> pending reboot, please apply updates by restarting the system and rerun script - FAILED."
		objFile.WriteLine "There are <"& iRequiredUpdates &"> pending reboot, please apply updates by restarting the system and rerun script - FAILED."
        objFileFail.WriteLine "There are <"& iRequiredUpdates &"> pending reboot, please apply updates by restarting the system and rerun script - FAILED."
        CheckFailureCount = CheckFailureCount + 1
End If
Err.Clear

Set oReg = Nothing

'Check if the logged on user is using a roaming profile (not supported)
If bEnableVerboseMode = true Then
    Wscript.Echo "Checking if the logged on user is using a roaming profile (not supported)..."
End If

objUserName = objNetwork.UserName

Set UserProfiles = GetObject("winmgmts:root\cimv2").ExecQuery("select * from win32_userprofile where RoamingConfigured = 'true' ")
For Each Profile in UserProfiles
    Set Users = GetObject("winmgmts:root\cimv2").ExecQuery("select * from Win32_UserAccount where SID = '" & Profile.SID & "'")

    For Each User in Users
        If InStr(User, objUserName)> -1 Then
            bIsRoamigUser = True
        End If
    Next
Next

If bIsRoamigUser = true Then
    Err.Clear
	Wscript.Echo "Using Windows a roaming profile is not supported, please user a user account without a Windows roaming profile - FAILED."
    objFileFail.Writeline "Using Windows a roaming profile is not supported, please user a user account without a Windows roaming profile - FAILED."
    Err.Clear
    If Err.number <> 0 Then
        Wscript.Echo "Could not write to trace error log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
    End If
Else
    Wscript.Echo "No Windows roaming profile found - OK."
    Err.Clear
    objFile.Writeline "No Windows roaming profile found - OK."
    Err.Clear
    If Err.number <> 0 Then
        Wscript.Echo "Could not write to trace log file, please verify the current user is a domain user running with adminsitrative priviledges on the tools machine - FAILED."
    End If
End If

Set UserProfiles = Nothing
Set objNetwork = Nothing
Set Users = Nothing
Err.Clear



'Check for .NET Framework 4 is installed
sRegValue = oShell.RegRead(REG_NET_Framework4)

If Err.Number <> 0 Then
		bSuccessfulOutput=false
        CheckFailureCount = CheckFailureCount + 1
	    WScript.Echo "Checking for .NET Framework 4.0 registry key ("& REG_NET_Framework4 &") - FAILED."
        Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
		objFile.Writeline "Checking for .NET Framework 4.0 registry key ("& REG_NET_Framework4 &") - FAILED."
		objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
		objFileFail.Writeline "Checking for .NET Framework 4.0 registry key ("& REG_NET_Framework4 &") - FAILED."
		objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
    	Err.Clear
End If


If sRegValue <> "4.0.0" Then
    bSuccessfulOutput=false
    CheckFailureCount = CheckFailureCount + 1
    WScript.Echo ".NET Framework 4.0 not found ("& sRegValue &") - FAILED."
    objFile.Writeline ".NET Framework 4.0 not found ("& sRegValue &") - FAILED."
    objFileFail.Writeline ".NET Framework 4.0 not found ("& sRegValue &") - FAILED."
    Err.Clear
Else
	WScript.Echo ".NET Framework 4.0 installed - OK."
	objFile.Writeline ".NET Framework 4.0 installed - OK."

	'Checking if .NET 4 installed then do ngen call
	strCommand="%windir%\Microsoft.NET\Framework\v4.0.30319\ngen executeQueuedItems"
	Wscript.echo "Running ngen 3 times on local tools machine. This might take a while...."
	objFile.writeline "Running ngen 3 times on local tools machine..."
	Exec oShell,strCommand
    Exec oShell,strCommand
    Exec oShell,strCommand
	
	Exec oShell,strCommand
	str_NgenOut86=Exec(oShell,strCommand)

	If Instr(str_NgenOut86,"All compilation targets are up to date") > 0 then
		wscript.Echo "X86 All compilation targets are up to date on local tools Machine - OK."
		objFile.WriteLine "X86 All compilation targets are up to date on local tools Machine - OK."
	Else
		wscript.echo "%windir%\Microsoft.NET\Framework\v4.0.30319\ngen executeQueuedItems Needed to be run more times. Please refer to page 7 of CRM RaaS prerequisites document - FAILED."
		objFile.Writeline "%windir%\Microsoft.NET\Framework\v4.0.30319\ngen executeQueuedItems Needed to be run more times. Please refer to page 7 of CRM RaaS prerequisites document - FAILED."
		objFileFail.Writeline "%windir%\Microsoft.NET\Framework\v4.0.30319\ngen executeQueuedItems Needed to be run more times. Please refer to page 7 of CRM RaaS prerequisites document - FAILED."
		CheckFailureCount = CheckFailureCount + 1
	End If
	strCommand="%windir%\Microsoft.NET\Framework64\v4.0.30319\ngen executeQueuedItems"
	'RunCommand oShell,strCommand
	'RunCommand oShell,strCommand
	'RunCommand oShell,strCommand
	Exec oShell,strCommand
	Exec oShell,strCommand

	resultCode=ExecEx(oShell,strCommand)
	if resultCode = 0 then
		wscript.Echo "X64 All compilation targets are up to date on local tools Machine - OK."
		objFile.WriteLine "X64 All compilation targets are up to date on local tools machine - OK."
	Else
		wscript.echo "'%windir%\Microsoft.NET\Framework64\v4.0.30319\ngen executeQueuedItems' needed to be run more times on local Tools machine. Please refer to page 7 of CRM RaaS prerequisites document - FAILED."
		objFile.Writeline "'%windir%\Microsoft.NET\Framework64\v4.0.30319\ngen executeQueuedItems' needed to be run more times on local Tools machine. Please refer to page 7 of CRM RaaS prerequisites document - FAILED."
		objFileFail.Writeline "'%windir%\Microsoft.NET\Framework64\v4.0.30319\ngen executeQueuedItems' needed to be run more times on local Tools machine. Please refer to page 7 of CRM RaaS prerequisites document - FAILED."
		CheckFailureCount = CheckFailureCount + 1
	End If
End If

sRegValue = 0
Err.Clear
'Check If REG_MaxIdleTime is set to disconect the terminal server session automaticly if inactive for for 4 hours
sRegValue = oShell.RegRead(REG_MaxIdleTime)

If Err.Number = 0 and sRegValue > 0 and sRegValue < 240000 Then
    ' If found it means the session will be closed automaticly
    bSuccessfulOutput=false
    CheckFailureCount = CheckFailureCount + 1
    WScript.Echo "Checking for MaxIdleTime ("& REG_MaxIdleTime &") a value of: <"&sRegValue&"> was found(bellow 4 hours), data collection requires at least four hours having the user disconnected may result in data collection failure. Please contact your domain administrator in order to get the tools machine temporarily out of policy for the time of this data collection - FAILED."
    objFile.Writeline "Checking for MaxIdleTime ("& REG_MaxIdleTime &") a value of: <"&sRegValue&"> was found(bellow 4 hours), data collection requires at least four hours having the user disconnected may result in data collection failure. Please contact your domain administrator in order to get the tools machine temporarily out of policy for the time of this data collection - FAILED."
    objFileFail.Writeline "Checking for MaxIdleTime ("& REG_MaxIdleTime &") a value of: <"&sRegValue&"> was found(bellow 4 hours), data collection requires at least four hours having the user disconnected may result in data collection failure. Please contact your domain administrator in order to get the tools machine temporarily out of policy for the time of this data collection - FAILED."
Else
    WScript.Echo "No bellow 4 hours session MaxIdleTime currently set - OK."
    objFile.Writeline "No bellow 4 hours session MaxIdleTime currently set - OK."
End If

Err.Clear

sRegValue = 0
Err.Clear
'Check If REG_MaxDisconnectionTime is set to close the terminal server session automaticly if disconnected for 4 hours
sRegValue = oShell.RegRead(REG_MaxDisconnectionTime)

If Err.Number = 0 and sRegValue > 0 and sRegValue < 240000 Then
    ' If found it means the session will be closed automaticly
    bSuccessfulOutput=false
    CheckFailureCount = CheckFailureCount + 1
    WScript.Echo "Checking for MaxDisconnectionTime ("& REG_MaxDisconnectionTime &") a value of: <"&sRegValue&"> was found(bellow 4 hours), data collection requires at least four hours having the user disconnected may result in data collection failure. Please contact your domain administrator in order to get the tools machine temporarily out of policy for the time of this data collection - FAILED."
    objFile.Writeline "Checking for MaxDisconnectionTime ("& REG_MaxDisconnectionTime &") a value of: <"&sRegValue&"> was found(bellow 4 hours), data collection requires at least four hours having the user disconnected may result in data collection failure. Please contact your domain administrator in order to get the tools machine temporarily out of policy for the time of this data collection - FAILED."
    objFileFail.Writeline "Checking for MaxDisconnectionTime ("& REG_MaxDisconnectionTime &") a value of: <"&sRegValue&"> was found(bellow 4 hours), data collection requires at least four hours having the user disconnected may result in data collection failure. Please contact your domain administrator in order to get the tools machine temporarily out of policy for the time of this data collection - FAILED."
Else
    WScript.Echo "No bellow 4 hours session MaxDisconnectionTime currently set - OK."
    objFile.Writeline "No bellow 4 hours session MaxDisconnectionTime currently set - OK."
End If

Err.Clear


Dim ExitCode
'Verify WMI' Repository on local Tools Machine
strCommand= "cmd /C winmgmt.exe /verifyrepository"
ExitCode=ExecEx(oShell,strCommand)

if  ExitCode = 0 then
	wscript.Echo "WMI repository is consistent on local tools machine - OK."
	objFile.Writeline "WMI repository is consistent on local tools machine - OK."
Else
	
	wscript.Echo "WMI repository is NOT consistent on local tools machine. Please refer to Page 7 of CRM RaaS Prerquisites document and http://msdn.microsoft.com/en-us/library/aa394525(v=vs.85).aspx - FAILED."
	objFile.Writeline "WMI repository is NOT consistent on local tools machine. Please refer to Page 7 of CRM RaaS Prerquisites document and http://msdn.microsoft.com/en-us/library/aa394525(v=vs.85).aspx - FAILED."
	objFileFail.Writeline "WMI repository is NOT consistent on local tools machine. Please refer to Page 7 of CRM RaaS Prerquisites document and http://msdn.microsoft.com/en-us/library/aa394525(v=vs.85).aspx - FAILED."
	CheckFailureCount = CheckFailureCount + 1
End If
Err.Clear

'Check for IIS Management script and tools are installed
sRegValue = oShell.RegRead(REG_IIS_MANAGEMENT_SCRIPT_AND_TOOLS)

If Err.Number <> 0 Then
		bSuccessfulOutput=false
	    WScript.Echo "IIS Management script and tools are not installed on local tools machine registry key ("& REG_IIS_MANAGEMENT_SCRIPT_AND_TOOLS &") - FAILED."
        Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
	    objFile.Writeline "IIS Management script and tools are not installed on local tools machine - registry key ("& REG_IIS_MANAGEMENT_SCRIPT_AND_TOOLS &") - FAILED."
	    objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
	    objFileFail.Writeline "IIS Management script and tools are not installed on local tools machine registry key ("& REG_IIS_MANAGEMENT_SCRIPT_AND_TOOLS &") - FAILED."
	    objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
        CheckFailureCount = CheckFailureCount + 1

        If Err.number = -2147217394 Then
            'scripts and tools are likely not installed on the CRM server itself

	        WScript.Echo "Please verify that IIS Management script and tools are installed specificly on the targeted server itself."
	        objFile.Writeline "Please verify that IIS Management script and tools are installed specificly on the targeted server itself."
	        objFileFail.Writeline "Please verify that IIS Management script and tools are installed specificly on the targeted server itself."

        End If

        Err.Clear
End If

If sRegValue <> "1" Then
	bSuccessfulOutput=false
    CheckFailureCount = CheckFailureCount + 1
    WScript.Echo "IIS Management script and tools are not installed on local tools machine - FAILED."
    objFile.Writeline "IIS Management script and tools are not installed on local tools machine - FAILED."
     objFileFail.Writeline "IIS Management script and tools are not installed on local tools machine - FAILED."
	Err.Clear
Else
	WScript.Echo "IIS Management script and tools installed on local tools machine - OK." 
	objFile.Writeline "IIS Management script and tools installed on local tools machine - OK." 
End If

'Check Log Parser create object status
'Create Log Parser COM object
Set oLogQuery = CreateObject("MSUtil.LogQuery")

'Check Log Parser create object status
If Err.Number <> 0 Then
	bSuccessfulOutput=false
    CheckFailureCount = CheckFailureCount + 1
	WScript.Echo "Log Parser needs to be installed http://www.bing.com/search?q=log+parser+download - (Error number:" & Err.Number & " - Error description : " & Err.Description & ") - FAILED."
	Wscript.Echo "Exiting <Press enter key to quit>."
	objFile.Writeline "Log Parser needs to be installed http://www.bing.com/search?q=log+parser+download - (Error number:" & Err.Number & " - Error description : " & Err.Description & ") - FAILED."
    objFileFail.Writeline "Log Parser needs to be installed http://www.bing.com/search?q=log+parser+download - (Error number:" & Err.Number & " - Error description : " & Err.Description & ") - FAILED."
    objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
    objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
    Err.Clear
Else
	WScript.Echo "Log Parser is installed - OK."
	objFile.Writeline "Log Parser is installed - OK."
End If


Wscript.Echo ""
Wscript.Echo "2 - Analyzing tool machine configuration"
Wscript.Echo "-----------------------------------------"
Wscript.Echo ""
CheckArchitecture(oWMI)
CheckOS(oWMI)
GetComputerProperties(oWMI)
CheckDiskspace(oWMI)

'Check for incompatible software
IncompatibleInstalledSoftware = CheckForIncompatibleSoftware (LocalListAllInstalledSoftware)

Err.Clear

'Create local WMI WebAdministration object
Set oIIS = GetObject("winmgmts:root\WebAdministration")

If Err.Number <> 0 Then
	bSuccessfulOutput=false
    CheckFailureCount = CheckFailureCount + 1
    Wscript.Echo "'IIS Management Scripts and Tools' not intalled on tools machine. - FAILED."
	Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
	objFile.Writeline "'IIS Management Scripts and Tools' not intalled on tools machine. - FAILED."
	objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
	objFile.Writeline "'IIS Management Scripts and Tools' not intalled on tools machine. - FAILED."
	objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
    Err.Clear
Else
	Wscript.Echo "Local WebAdministration WMI call - OK."
	objFile.Writeline "Local WebAdministration WMI call - OK."
End If

'Call procedure to check If powershell 2.0 is installed on tools machine 
CheckPowerShell2Installed oShell

iTotalFailures = CheckFailureCount

Wscript.Echo ""

Wscript.Echo  "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"

If  CheckFailureCount = 0 then 
	Wscript.Echo ""
	Wscript.Echo "Summary for tool machine checks - CRM RaaS prerequisites completed with - SUCCESS."
	Wscript.Echo ""
	objFile.Writeline ""
	objFile.Writeline "Summary for tool machine checks - CRM RaaS prerequisites uisites completed with SUCCESS."
	objFile.Writeline ""
	objFile.Writeline "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
Else
	Wscript.Echo ""
	Wscript.Echo "Summary for tool machine checks - CRM RaaS prerequisites completed with: " & CheckFailureCount & " failure(s) - FAILED."
	Wscript.Echo ""
    Wscript.Echo "Please review and send the CRMRaasPreReqsFailuresLog.txt log file to your PFE as quickly as possible."

    If UBound(IncompatibleInstalledSoftware) >1 Then
        Wscript.Echo ""
        Wscript.Echo "/!\ We identifed that this tools machine has potentialy incompatible software with RaaS client."
        objFile.Writeline "/!\ We identifed that this tools machine has potentialy incompatible software with RaaS client."
        objFileFail.Writeline "/!\ We identifed that this tools machine has potentialy incompatible software with RaaS client."
        Wscript.Echo ""
    End If

	objFile.Writeline ""
	objFile.Writeline "Summary for tool machine checks - CRM RaaS prerequisites completed with: " & CheckFailureCount & " failure(s) - FAILED."
	objFile.Writeline ""
	objFile.Writeline "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
	objFileFail.Writeline "Summary for tool machine checks - CRM RaaS prerequisites completed with: " & CheckFailureCount & " failure(s) - FAILED."
	objFileFail.Writeline ""
	objFileFail.Writeline "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
End If

If iToolsMachinePerformanceWarnings <> 0 Then
	Wscript.Echo ""
	Wscript.Echo "Tools machine performance warnings: " &  iToolsMachinePerformanceWarnings & "."
	Wscript.Echo ""
	objFile.Writeline ""
	objFile.Writeline "Tools machine performance warnings: " &  iToolsMachinePerformanceWarnings & "."
	objFile.Writeline ""
	objFileFail.Writeline ""
	objFileFail.Writeline "Tools machine performance warnings: " &  iToolsMachinePerformanceWarnings & "."
	objFileFail.Writeline ""
End If

Set args = Wscript.Arguments

If args.Count=0 then
    Wscript.Echo "-----------------------------------------------------------------------"
    Wscript.Echo "<!> Note: No remote server tests performed as no arguments found for script hence ran for local machine only.To run script against remote servers run the script with arguments"
    Wscript.Echo "Script Usage:"
    Wscript.echo "cscript VerifyCRMRaaSPreReq-V.vbs <server1> <server2> <server3>......."
    Wscript.Echo "-----------------------------------------------------------------------"


    objFile.Writeline "/!\No remote server tests performed as no arguments found for script hence ran for local machine only.To run script against remote servers run the script with arguments."

    objFileFail.Writeline "/!\No remote server tests performed as no arguments found for script hence ran for local machine only.To run script against remote servers run the script with arguments."

Else 

'Perform remote tests
Wscript.Echo        ""
Wscript.Echo        "3 - Performing remote tests"
Wscript.Echo        "----------------------------"
objFile.Writeline   "3 - Performing remote tests"
objFile.Writeline   "----------------------------"


For j=0 to args.Count-1

    sServerName=Trim(args(j))
    Wscript.Echo "Starting test phase for server:<" &  sServerName  & ">."
    objFile.Writeline "Connecting to server:<" &  sServerName  & ">..."
    Err.Clear
    bIISFound = false
    bSuccessfulOutput=true
    CheckFailureCount = 0 
    v_remote_webAdmin=false

    Wscript.Echo "Sending ping to server:<" &  sServerName & ">..."
    objFile.Writeline "Sending ping to server:<" &  sServerName & ">..."
    CheckPing oShell,sServerName
    Err.Clear
    Wscript.Echo "Checking ports 135, 139, 445 to server:<" &  sServerName & ">..."
    objFile.Writeline "Checking ports 135, 139,445 to server:<" &  sServerName & ">..."

    CheckPort oShell,sServerName,135
    Err.Clear
    CheckPort oShell,sServerName,139
    Err.Clear
    CheckPort oShell,sServerName,445
    Err.Clear

    'Check If current user is local admin on remote server
    CheckLocalAdmin (sServerName)

    RemoteListAllInstalledSoftware = ListAllInstalledSoftware (sServerName)

    Err.Clear
    'Create remote WMI CIMv2 object
    Set oWMI = Nothing

    Set oWMI = GetObject("winmgmts:root\CIMv2")

    'Taking again tools machine time on machine because if there is a significant time difference WMI may not work
    Set colItems = oWMI.ExecQuery("Select * from Win32_OperatingSystem",,48)
    If Err.number = 0 Then
        For Each objItem in colItems
            Wscript.Echo "Local machine CurrentTimeZone is: " & objItem.CurrentTimeZone
            Wscript.Echo "Local machine LocalDateTime is: " & objItem.LocalDateTime
            objFile.Writeline "Local machine CurrentTimeZone is: " & objItem.CurrentTimeZone
            objFile.Writeline "Local machine LocalDateTime is: " & objItem.LocalDateTime
            strLocalMachineTimeZone = objItem.CurrentTimeZone
            strLocalMachineDateTime = objItem.LocalDateTime
        Next

    Else

        CheckFailureCount = CheckFailureCount + 1
    	Wscript.Echo "An error occured when retriving time tools machine (step2):<" &  sServerName  & "> using WMI - FAILED."
	    Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
	    objFile.Writeline "An error occured when retriving time tools machine (step2):<" &  sServerName  & "> using WMI - FAILED."
        objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
	    objFileFail.Writeline "An error occured when retriving time tools machine (step2):<" &  sServerName  & "> using WMI - FAILED."
        objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
        Err.Clear

    End If


    Set oWMI = Nothing


    Set oWMI = GetObject("winmgmts:\\" + sServerName + "\root\CIMv2")

    If Err.Number <> 0 Then
	    bSuccessfulOutput = false
        CheckFailureCount = CheckFailureCount + 1
	    Wscript.Echo "Can't access WMI CIMv2' on server:<" &  sServerName  & "> - FAILED."
	    Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
	    objFile.Writeline "Can't access WMI CIMv2' on server:<" &  sServerName  & "> - FAILED."
        objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
	    objFileFail.Writeline "Can't access WMI CIMv2' on server:<" &  sServerName  & "> - FAILED."
        objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
        Err.Clear
		
		bWMITestFailed = true

    Else
	    Wscript.Echo "Remote access of CIMv2 on server:<" &  sServerName  & "> - OK."
	    objFile.Writeline "Remote access of CIMv2 on server:<" &  sServerName  & "> - OK."
    End If
	
	If bWMITestFailed = false then

		'Get remote time on machine because if there is a significant time difference WMI may not work
		Set colItems = oWMI.ExecQuery("Select * from Win32_OperatingSystem",,48)
		If Err.number = 0 Then
			For Each objItem in colItems
				Wscript.Echo "On server <"&sServerName&"> CurrentTimeZone is: " & objItem.CurrentTimeZone
				Wscript.Echo "On server <"&sServerName&"> LocalDateTime is: " & objItem.LocalDateTime
				objFile.Writeline "On server <"&sServerName&"> CurrentTimeZone is: " & objItem.CurrentTimeZone
				objFile.Writeline "On server <"&sServerName&"> LocalDateTime is: " & objItem.LocalDateTime

				strRemoteMachineTimeZone = objItem.CurrentTimeZone
				strRemoteMachineDateTime = objItem.LocalDateTime
			Next

		Else

			CheckFailureCount = CheckFailureCount + 1
			Wscript.Echo "An error occured when retriving time on server:<" &  sServerName  & "> using WMI - FAILED."
			Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
			objFile.Writeline "An error occured when retriving time on server:<" &  sServerName  & "> using WMI - FAILED."
			objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
			objFileFail.Writeline "An error occured when retriving time on server:<" &  sServerName  & "> using WMI - FAILED."
			objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
			Err.Clear

		End If

		If strLocalMachineTimeZone <> strRemoteMachineTimeZone Then
			iToolsMachinePerformanceWarnings = iToolsMachinePerformanceWarnings + 1
			WScript.Echo "Tools machine and remote server <" &  sServerName  & "> are in different time zones (time synchronization could be an issue and may cause WMI data collection errors) - WARNING."
			objFile.Writeline "Tools machine and remote server <" &  sServerName  & "> are in different time zones (time synchronization could be an issue and may cause WMI data collection errors) - WARNING."
			objFileFail.Writeline "Tools machine and remote server <" &  sServerName  & "> are in different time zones (time synchronization could be an issue and may cause WMI data collection errors) - WARNING."
		Else
			WScript.Echo "Tools machine and remote server <" &  sServerName  & "> are in same time zones - OK."
			objFile.Writeline "Tools machine and remote server <" &  sServerName  & "> are in same time zones - OK."
		End If

		iTimeDiffRemoteServer = ABS (Left(strLocalMachineDateTime, 14) - Left(strRemoteMachineDateTime, 14))

		If ( iTimeDiffRemoteServer > 10 ) Then

			If ( iTimeDiffRemoteServer < 20 ) Then
				iToolsMachinePerformanceWarnings = iToolsMachinePerformanceWarnings + 1
				WScript.Echo "Tools machine and remote server <" &  sServerName  & "> time difference is ("&iTimeDiffRemoteServer&") high this may cause WMI issues, please sync machine times - WARNING."
				objFile.Writeline "Tools machine and remote server <" &  sServerName  & "> time difference is ("&iTimeDiffRemoteServer&") high this may cause WMI issues, please sync machine times - WARNING."
				objFileFail.Writeline "Tools machine and remote server <" &  sServerName  & "> time difference is ("&iTimeDiffRemoteServer&") high this may cause WMI issues, please sync machine times - WARNING."
			Else
				'Anything above 20 second is at high risk 
				CheckFailureCount = CheckFailureCount + 1
				WScript.Echo "Tools machine and remote server <" &  sServerName  & "> time difference is ("&iTimeDiffRemoteServer&") very high this will likely cause WMI issues, please sync machine times - FAILED."
				objFile.Writeline "Tools machine and remote server <" &  sServerName  & "> time difference is ("&iTimeDiffRemoteServer&") very high this will likely cause WMI issues, please sync machine times - FAILED."
				objFileFail.Writeline "Tools machine and remote server <" &  sServerName  & "> time difference is ("&iTimeDiffRemoteServer&") very high this will likely cause WMI issues, please sync machine times - FAILED."
			
			End If

		Else

			WScript.Echo "Tools machine and remote server <" &  sServerName  & "> time difference is ("&iTimeDiffRemoteServer&") - OK."
			objFile.Writeline "Tools machine and remote server <" &  sServerName  & "> time difference is ("&iTimeDiffRemoteServer&") - OK."

		End If
		Err.Clear


		'Create remote WMI CIMv2 Win32_Service call to check the presence of IIS W3CSVC Service
		Set colServices = oWMI.InstancesOf("Win32_Service")

		If Err.Number <> 0 Then
			bSuccessfulOutput=false
			CheckFailureCount = CheckFailureCount + 1
			Wscript.Echo "'Cant access WMI Win32_Service on server:<" &  sServerName  & "> - FAILED."
			Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
			objFile.Writeline "'Cant access WMI Win32_Service on server:<" &  sServerName  & "> - FAILED."
			objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
			objFileFail.Writeline "'Cant access WMI Win32_Service on server:<" &  sServerName  & "> - FAILED."
			objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
			Err.Clear
		Else
			Wscript.Echo "Remote access of Win32_Service on server:<" &  sServerName  & "> - OK."
			objFile.Writeline "Remote access of Win32_Service on server:<" &  sServerName  & "> - OK."
		End If






		'Call procedure to check If Remote Registry service is started on remote server
		'CheckRemoteService colServices,"RemoteRegistry",sServerName

		'Call to Check Server Service started on remote server
		CheckRemoteService colServices,"LanmanServer",sServerName

		'Call to Check Server Service started on remote server
		CheckRemoteService colServices,"LanmanWorkstation",sServerName

		'Call to check performance logs and alerts service started on remote server
		'CheckRemoteService colServices,"pla",sServerName

		'Call to check RPC Service Started on remote server
		CheckRemoteService colServices,"RpcSs",sServerName

		'Call to check WMI Service Started on remote server
		CheckRemoteService colServices,"Winmgmt",sServerName

		'Call to check SQL Service Started on remote server
		bSQLServiceFound = CheckRemoteService (colServices,"MSSQLSERVER",sServerName)





		If bSQLServiceFound = true Then

			Err.Clear
			Wscript.Echo ""
			Wscript.Echo "*** Your input is required *** SQL Server was found on machine <" &  sServerName & "> we will now be performing sysadmin role check for this machine. Simply press the <enter> key to test against this machine or enter its name now:" 
			sSQLInstanceName = WScript.StdIn.ReadLine
			If sSQLInstanceName =  "" Then
				sSQLInstanceName = sServerName
			End If
			bSQLCheckSysAdmin = true

			' perform sysadmin role check
			CheckSysAdminRole ( sSQLInstanceName )
	

			Err.Clear

		End If

		bPassedEvtLogApp = false
		bPassedEvtLogSys = false

		Err.Clear

		Set colLoggedEvents = oWMI.ExecQuery ("Select * from Win32_NTLogEvent  Where Logfile = 'Application'")

		If Err.number <> 0 Then
			CheckFailureCount = CheckFailureCount + 1
			Wscript.Echo "Cant Application Event Log on server:<" &  sServerName  & "> - FAILED."
			Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
			objFile.Writeline "Cant Application Event Log on server:<" &  sServerName  & "> - FAILED."
			objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
			objFileFail.Writeline "Cant Application Event Log on server:<" &  sServerName  & "> - FAILED."
			objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
			Err.Clear

		Else
			bPassedEvtLogApp = true
			Wscript.Echo "Event Log Application check on server:<" &  sServerName  & "> - OK."
			objFile.Writeline "Event Log Application check on server:<" &  sServerName  & "> - OK."
			Err.Clear
		End If

		Set colLoggedEvents = Nothing

		'Perform Event Log Robustness Checks by retreiveing event mutiple times Windows Event logs 
		bPassedEvtLogMultiple = true

		If bPassedEvtLogApp = true Then
			
			Wscript.Echo "Starting WMI remote robustness check and loop ("& RemoteWMIRobustnessMaxCount &") times... Started at:" & Now & ". This may take a while, please wait..."
			iCptRobustness = 0

			If bSkipWMIRobustness = false Then

				Do While iCptRobustness < RemoteWMIRobustnessMaxCount and bPassedEvtLogMultiple = true

					Set oWMI = Nothing
					Set oWMI = GetObject("winmgmts:\\" + sServerName + "\root\CIMv2")

					If Err.Number <> 0 Then            
						bPassedEvtLogMultiple = false
						Wscript.Echo "WMI robustness check failed for machine:<" &  sServerName  & "> at : "& iCptRobustness &" iteration - GetObject - FAILED."
						Wscript.Echo "WMI robustness check failed for machine:<" &  sServerName  & "> at : "& iCptRobustness &" iteration - GetObject - FAILED."
						objFile.Writeline "WMI robustness check failed for machine:<" &  sServerName  & "> at : "& iCptRobustness &" iteration - GetObject - FAILED."
						objFileFail.Writeline "WMI robustness check failed for machine:<" &  sServerName  & "> at : "& iCptRobustness &" iteration - GetObject - FAILED."
						Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
						objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
						objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description

					  Else
						  Set colLoggedEvents = oWMI.ExecQuery ("Select top 1 * from Win32_NTLogEvent  Where Logfile ='Application'")

		'                    For Each objEvent in colLoggedEvents
		'                        Wscript.Echo "Category: " & objEvent.Category & VBNewLine _
		'                        & "Computer Name: " & objEvent.ComputerName & VBNewLine _
		'                        & "Event Code: " & objEvent.EventCode & VBNewLine _
		'                        & "Message: " & objEvent.Message & VBNewLine _
		'                        & "Record Number: " & objEvent.RecordNumber & VBNewLine _
		'                        & "Source Name: " & objEvent.SourceName & VBNewLine _
		'                        & "Time Written: " & objEvent.TimeWritten & VBNewLine _
		'                        & "Event Type: " & objEvent.Type & VBNewLine _
		'                        & "User: " & objEvent.User
		'                    Next

						  Set colLoggedEvents = Nothing

						  If Err.Number <> 0 Then
								bPassedEvtLogMultiple = false
								Wscript.Echo "WMI robustness check failed for machine:<" &  sServerName  & "> at : "& iCptRobustness &" iteration - ExecQuery - FAILED."
								objFile.Writeline "WMI robustness check failed for machine:<" &  sServerName  & "> at : "& iCptRobustness &" iteration - ExecQuery - FAILED."
								objFileFail.Writeline "WMI robustness check failed for machine:<" &  sServerName  & "> at : "& iCptRobustness &" iteration - ExecQuery - FAILED."
								Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
								objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
								objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description

						  Else
								If iCptRobustness Mod 500 = 1 Then
									Wscript.Echo "WMI remote robustness check loop at: " & iCptRobustness & " of " & RemoteWMIRobustnessMaxCount
								End If
						  End If
					  

					End If
				
					iCptRobustness = iCptRobustness +1
							
				loop
			Else
				Wscript.Echo "WMI remote robustness check skipped - WARNING."
				objFile.Writeline "WMI remote robustness check skipped - WARNING."
				objFileFail.Writeline "WMI remote robustness check skipped - WARNING."

				iToolsMachinePerformanceWarnings = iToolsMachinePerformanceWarnings + 1

			End If

			Wscript.Echo "WMI remote robustness check ended at:" & Now

			If bPassedEvtLogMultiple = false Then
				CheckFailureCount = CheckFailureCount + 1
				Wscript.Echo "WMI robustness check failed for machine:<" &  sServerName  & "> at : "& iCptRobustness &" iteration - FAILED."
				objFile.Writeline "WMI robustness check failed for machine:<" &  sServerName  & "> at : "& iCptRobustness &" iteration - FAILED."
				objFileFail.Writeline "WMI robustness check failed for machine:<" &  sServerName  & "> at : "& iCptRobustness &" iteration - FAILED."
				Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
				objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
				objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
				Err.Clear
			Else
				Wscript.Echo "WMI remote robustness check - OK."
				objFile.Writeline "WMI remote robustness check - OK."
			End If

		Else
			CheckFailureCount = CheckFailureCount + 1
			Wscript.Echo "Didn't perform WMI robustness check because single Win32_Service check failed:<" &  sServerName  & "> - FAILED."
			objFile.Writeline "Didn't perform WMI robustness check because single Win32_Service check failed:<" &  sServerName  & "> - FAILED."
			objFileFail.Writeline "Didn't perform WMI robustness check because single Win32_Service check failed:<" &  sServerName  & "> - FAILED."
			Err.Clear        
		End If


		'Detect presence for IIS W3CSVC Service
		For Each objService In colServices
			If objService.Name = "W3SVC" Then
				bIISFound = true
				Wscript.Echo "Detected that IIS is installed on server:<" &  sServerName & "> - OK."
				objFile.Writeline "Detected that IIS is installed on server:<" &  sServerName & "> - OK."
			End If
		Next

		Set oWMI=Nothing

		'If IIS service has been detected on remote host try connecting to it
		If bIISFound = true Then

			'Check if CRM is installed on that server
			sRegValue = oShell.RegRead(REG_MSCRM)
			
			If Err.Number = 0 Then
				bCRMFound = true
				
				' Check if Web Administration components are installed.
				Set oWMI = Nothing
				Set oWMI = GetObject("winmgmts:\\" + sServerName + "\root\WebAdministration")

				If Err.Number <> 0 Then
					bSuccessfulOutput=false
					CheckFailureCount = CheckFailureCount + 1
					Wscript.Echo "The server <" &  sServerName & "> has IIS installed (W3SVC service present) but can't be reached via Web Administration - please verify If IIS Scripts and Tools Feature is installed on this machine. - FAILED."
					Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
					objFile.Writeline "The server <" &  sServerName & "> has IIS installed (W3SVC service present) but can't be reached via Web Administration - please verify If IIS Scripts and Tools Feature is installed on this machine. - FAILED."
					objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
					objFileFail.Writeline "The server <" &  sServerName & "> has IIS installed (W3SVC service present) but can't be reached via Web Administration - please verify If IIS Scripts and Tools Feature is installed on this machine. - FAILED."
					objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
					Err.Clear
				Else
					v_remote_webAdmin=true
					Wscript.Echo "IIS WMI WebAdminstration check completed with success on server:<" &  sServerName & "> - OK."
					objFile.Writeline "IIS WMI WebAdminstration check completed with success on server:<" &  sServerName & "> - OK."
				End If

			Else
				Wscript.Echo "CRM was not found on that server so skipping check to verify that Web Administration components are installed - OK"
				objFile.Writeline "CRM was not found on that server so skipping check to verify that Web Administration components are installed - OK"
				Err.Clear
			End If

	 

		Else
			Wscript.Echo "Not an IIS role server so skipping IIS tests for server :<" &  sServerName & ">"
			objFile.Writeline "Not an IIS role server so skipping IIS tests for server :<" &  sServerName & ">"
		End If

		'Checking If IIS local version is higher or above than remote version
		If bIISFound = true Then
		
			strRegValueIISLocal = oShell.RegRead("HKEY_LOCAL_MACHINE\" & REG_IIS_VersionString)
		
			If Err.Number <> 0 Then
				bSuccessfulOutput=false
				CheckFailureCount = CheckFailureCount + 1
				WScript.Echo "Checking for local IIS server version - FAILED."
				Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description
				objFile.Writeline "Checking for local IIS server version - FAILED."
				objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description
				objFileFail.Writeline "Checking for local IIS server version - FAILED."
				objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description
				Err.Clear
			End If
		
			Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & sServerName & "\root\default:StdRegProv")

			oReg.GetStringValue HKEY_LOCAL_MACHINE, REG_IIS_VersionString, strRegValueIISRemote

			If Err.Number <> 0 Then
				bSuccessfulOutput=false
				CheckFailureCount = CheckFailureCount + 1
				WScript.Echo "Checking for remote IIS server version - FAILED."
				Wscript.Echo "Error number:" & Err.number & " Error description:" & Err.description	
				objFile.Writeline "Checking for remote IIS server version - FAILED."
				objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description	
				objFileFail.Writeline "Checking for remote IIS server version - FAILED."
				objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description 
				Err.Clear
			Else
				If strRegValueIISLocal >= strRegValueIISRemote Then
					WScript.Echo "IIS - Checking If local version of IIS is higher or equal than remote server version ("& strRegValueIISLocal & ">=" & strRegValueIISLocal &") - OK."
					objFile.Writeline "IIS - Checking If local version of IIS is higher or equal than remote server version ("& strRegValueIISLocal & ">=" & strRegValueIISLocal &") - OK."
				Else
					WScript.Echo "IIS - Checking If local version of IIS is higher or equal than remote server version ("& strRegValueIISLocal & ">=" & strRegValueIISLocal &") - FAILED."
					objFile.Writeline "IIS - Checking If local version of IIS is higher or equal than remote server version ("& strRegValueIISLocal & ">=" & strRegValueIISLocal &") - FAILED."
					objFileFail.Writeline "IIS - Checking If local version of IIS is higher or equal than remote server version ("& strRegValueIISLocal & ">=" & strRegValueIISLocal &") - FAILED."
					CheckFailureCount = CheckFailureCount + 1
					Err.Clear
				End if
				  
			End If
		End If

		Set oWMI = Nothing

		Err.Clear

		' Checking if remote server has IIS Script and tools installed

		If bIISFound = true Then

			dwScriptAndTools = -1

			oReg.GetDWORDValue HKEY_LOCAL_MACHINE, REG_IIS_ScriptAndTools, "ManagementScriptingTools", dwScriptAndTools

			If Err.number = 0 and dwScriptAndTools = 1 Then

					WScript.Echo "IIS - Script and Tools are installed on server:<"& sServerName & "> - OK."
					objFile.Writeline "IIS - Script and Tools are installed on server:<"& sServerName & "> - OK."           

			Else
					WScript.Echo "IIS - Script and Tools are not installed or could not be verfied on server:<"& sServerName & "> InetStp Registry value: '"&dwScriptAndTools&"' - FAILED."
					WScript.Echo "Error number:" & Err.number & " Error description:" & Err.description 
					WScript.Echo "Note: Please verify that " & HKEY_LOCAL_MACHINE & "\" & REG_IIS_ScriptAndTools & "\ManagementScriptingTools is set to 1, if that not the case, consider installing the 'IIS - Script and Tools' durning non peak-load hours and performing 'appcmd add backup' before and after installation."
					objFile.Writeline "IIS - Script and Tools are not installed or could not be verfied on server:<"& sServerName & "> InetStp Registry value: '"&dwScriptAndTools&"' - FAILED."
					objFile.Writeline "Error number:" & Err.number & " Error description:" & Err.description 
					objFile.Writeline"Note: Please verify that " & HKEY_LOCAL_MACHINE & "\" & REG_IIS_ScriptAndTools & "\ManagementScriptingTools is set to 1, if that not the case, consider installing the 'IIS - Script and Tools' durning non peak-load hours and performing 'appcmd add backup' before and after installation."
					objFileFail.Writeline "IIS - Script and Tools are not installed or could not be verfied on server:<"& sServerName & "> InetStp Registry value: '"&dwScriptAndTools&"' - FAILED."
					objFileFail.Writeline"Note: Please verify that " & HKEY_LOCAL_MACHINE & "\" & REG_IIS_ScriptAndTools & "\ManagementScriptingTools is set to 1, if that not the case, consider installing the 'IIS - Script and Tools' durning non peak-load hours and performing 'appcmd add backup' before and after installation."
					objFileFail.Writeline "Error number:" & Err.number & " Error description:" & Err.description 

					CheckFailureCount = CheckFailureCount + 1
			End If
			
			Err.Clear

		End If

		'Call procedure to check access to C$ admin share on remote server
		CheckAdminShare_Admin sServerName,"C"

		'Call procedure to check access to Admin$ admin share on remote server
		CheckAdminShare_Admin sServerName,"Admin"

		'Call procedure to check access to IPC$ admin share on remote server
		'CheckAdminShare_Admin sServerName,"IPC"

		If v_remote_webAdmin=true then
			CheckIISLoggingInstalled oReg,sServerName
			CheckIISLogsEnabledandFormat sServerName
		End if
	End If

    'If all successful output results then print success otherwise failure
    If  CheckFailureCount = 0 then 
	    Wscript.Echo ""
	    Wscript.Echo "Summary for remote tests checks - CRM RaaS prerequisites completed on:<" &  sServerName & "> with - SUCCESS."
	    Wscript.Echo ""
	    Wscript.Echo ""
	    objFile.Writeline ""
	    objFile.Writeline "Summary for remote tests checks - CRM RaaS prerequisites completed on:<" &  sServerName & "> with - SUCCESS."
	    objFile.Writeline ""

        If iIISLogsFailures <> 0 Then
	        Wscript.Echo ""
	        Wscript.Echo "IIS logs warnings: " &  iIISLogsFailures & "."
	        Wscript.Echo ""
	        Wscript.Echo ""
	        objFile.Writeline ""
	        objFile.Writeline "IIS logs warnings: " &  iIISLogsFailures & "."
	        objFile.Writeline ""
	        objFileFail.Writeline ""
	        objFileFail.Writeline "IIS logs warnings: " &  iIISLogsFailures & "."
	        objFileFail.Writeline ""
        End If

    Else
	    Wscript.Echo ""
	    Wscript.Echo "Summary for remote tests checks - CRM RaaS prerequisites completed with " & CheckFailureCount & " failure(s) on server:<" &  sServerName & "> - FAILED."
	    Wscript.Echo ""
	    objFile.Writeline ""
	    objFile.Writeline "Summary for remote tests checks - CRM RaaS prerequisites completed with " & CheckFailureCount & " failure(s) on server:<" &  sServerName & "> - FAILED."
	    objFile.Writeline ""
	
	    objFileFail.Writeline "Summary for remote tests checks - CRM RaaS prerequisites completed with " & CheckFailureCount & " failure(s) on server:<" &  sServerName & "> - FAILED."
	    objFileFail.Writeline ""
		
		If bWMITestFailed = true Then
			Wscript.Echo ""
			Wscript.Echo "CRITICAL: The WMI remote test could not be performed on server:<" &  sServerName & "> this has stopped the script - verify that firewall ports are open - FAILED."
			Wscript.Echo ""
			objFile.Writeline ""
			objFile.Writeline "CRITICAL: The WMI remote test could not be performed on server:<" &  sServerName & "> this has stopped the script - verify that firewall ports are open - FAILED."
			objFile.Writeline ""
			bjFileFail.Writeline ""
			bjFileFail.Writeline "CRITICAL: The WMI remote test could not be performed on server:<" &  sServerName & "> this has stopped the script - verify that firewall ports are open - FAILED."
			bjFileFail.Writeline ""
		End If

        If iIISLogsFailures <> 0 Then
	        Wscript.Echo ""
	        Wscript.Echo "IIS logs warnings: " &  iIISLogsFailures & "."
	        Wscript.Echo ""
	        Wscript.Echo ""
	        objFile.Writeline ""
	        objFile.Writeline "IIS logs warnings: " &  iIISLogsFailures & "."
	        objFile.Writeline ""
	        objFileFail.Writeline ""
	        objFileFail.Writeline "IIS logs warnings: " &  iIISLogsFailures & "."
	        objFileFail.Writeline ""
        End If

    End If

    iTotalFailures = iTotalFailures + CheckFailureCount

    Wscript.Echo "*****************************************************************************************************"
	Wscript.Echo ""

    If iTotalFailures = 0 Then
        Wscript.Echo "Required Dynamics CRM RaaS prerequisites PASSED with " & (iToolsMachinePerformanceWarnings + iIISLogsFailures) & " warning(s)."
        objFile.Writeline "Required Dynamics CRM RaaS prerequisites PASSED with " & (iToolsMachinePerformanceWarnings + iIISLogsFailures) & " warning(s)."
        objFileFail.Writeline "Required Dynamics CRM RaaS Prerequsitites PASSED with " & (iToolsMachinePerformanceWarnings + iIISLogsFailures) & " warning(s)."
    Else
        Wscript.Echo "Required Dynamics CRM RaaS prerequisites FAILED with " & iTotalFailures & " error(s) and " &( iToolsMachinePerformanceWarnings + iIISLogsFailures) & " warning(s)."
        objFile.Writeline "Required Dynamics CRM RaaS prerequisites FAILED with " & iTotalFailures & " error(s) and " &( iToolsMachinePerformanceWarnings + iIISLogsFailures) & " warning(s)."
        objFileFail.Writeline "Required Dynamics CRM RaaS prerequisites FAILED with " & iTotalFailures & " error(s) and " &( iToolsMachinePerformanceWarnings + iIISLogsFailures) & " warning(s)."
    End If

	Wscript.Echo ""

    If iToolsMachinePerformanceWarnings <> 0 Then
	    Wscript.Echo "Tools machine performance warnings: " &  iToolsMachinePerformanceWarnings & "."
	    objFile.Writeline "Tools machine performance warnings: " &  iToolsMachinePerformanceWarnings & "."
	    objFile.Writeline ""
	    objFileFail.Writeline ""
	    objFileFail.Writeline "Tools machine performance warnings: " &  iToolsMachinePerformanceWarnings & "."
	    objFileFail.Writeline ""
	    Wscript.Echo "Prerequisites script log files: CRMRaasPreReqsFailuresLog.txt and CRMRaasPreReqsScriptLog.txt have been written down to disk in that same directory, please share these files with your PFE and TAM."
	  
    End If

    If iIISLogsFailures <> 0 Then
	    Wscript.Echo "IIS logs warnings: " &  iIISLogsFailures & "."
	    objFile.Writeline ""
	    objFile.Writeline "IIS logs warnings: " &  iIISLogsFailures & "."
	    objFile.Writeline ""
	    objFileFail.Writeline ""
	    objFileFail.Writeline "IIS logs warnings: " &  iIISLogsFailures & "."
	    objFileFail.Writeline ""
    End If

    If iToolsMachinePerformanceWarnings <> 0 or iIISLogsFailures <> 0 Then
	    Wscript.Echo ""
    End If

    If bSQLCheckSysAdmin = true Then
        Wscript.Echo "Completed successfully user sysadmin role check on <" & sSQLInstanceName & ">"
	    objFile.Writeline "Completed successfully user sysadmin role check on <" & sSQLInstanceName & ">"
    Else
        If sSQLInstanceName = "" Then
            Wscript.Echo "IMPORTANT - The SQL SysAdmin role check was not performed, please make sure you run this script and complete this check with your CRM SQL instance."
	        objFile.Writeline "IMPORTANT - The SQL SysAdmin role check was not performed, please make sure you run this script and complete this check with your CRM SQL instance."
	        objFileFail.Writeline "IMPORTANT - The SQL SysAdmin role check was not performed, please make sure you run this script and complete this check with your CRM SQL instance."
        Else
            Wscript.Echo "CRITICAL - Failed executing sysadmin user role check on <" & sSQLInstanceName & ">. Please contact your database system administrator to get SQL sysadmin privileges as this is an must have prerequisite for a sucessful RaaS data collection. Once granted please rerun the this script."
	        objFile.Writeline "CRITICAL - Failed executing sysadmin user role check on <" & sSQLInstanceName & ">. Please contact your database system administrator to get SQL sysadmin privileges as this is an must have prerequisite for a sucessful RaaS data collection. Once granted please rerun the this script."
	        objFileFail.Writeline "CRITICAL - Failed executing sysadmin user role check on <" & sSQLInstanceName & ">. Please contact your database system administrator to get SQL sysadmin privileges as this is an must have prerequisite for a sucessful RaaS data collection. Once granted please rerun the this script."
        End If
    End If
    Wscript.Echo ""

    objFileFail.Writeline "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
    Wscript.Echo "*****************************************************************************************************"

    Set oWMI=Nothing
    Set oReg=Nothing

Next
End If

If args.Count=0 then
	Wscript.Echo "<!> Note:"
	Wscript.Echo "Please make sure you run this script for each CRM and SQL server machine of your environment."
	Wscript.Echo ""
	Wscript.Echo "In order to just perform the SysAdmin and connectivity check pass SYSADMIN as the only parameter to the script for example: cscript VerifyCRMRaaSPreReq-V.vbs SYSADMIN"
	Wscript.Echo ""

'	Wscript.Echo "<Press enter key to quit>"
'	WScript.StdIn.ReadLine
End If

objFile.Close
objFileFail.Close

Set objFileFail=Nothing
Set objFile=Nothing
Set objFileSys=Nothing
Set args=Nothing
Set oShell=Nothing
Wscript.Quit

