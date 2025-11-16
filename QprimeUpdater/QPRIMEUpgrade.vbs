'===========================================================================================================
'*	Script:		QPrimeUpdate.vbs
'*	Version:	v1.0.5   
'*	Author:		Mikael Almond (ISB DDI)
'*	Date:		20/1/2006
'*	Copyright: 	Queensland Police Service, 2006
'*
'*	Purpose:	Uninstall old clients and install latest client
'*
'-------------------------------------------------------------------------------------
'*  Change history:
'*   	Date		Author		Reference		Description
'*   	----		------		---------   	-----------
'*		20/01/06		MAl			v1.00			Initial Version
'*		23/01/06		ARu			v1.01			Fixed uninstallation process
'*		27/01/06		MAl			v1.02			Release Version
'*		03/02/06		MAl/ARu		v1.03			Corrected mistakes with error handling (CORRECTED x2)
'*		30/03/06		ARu			v1.04			Log MSI installer error codes
'*		10/04/06		MAl			v1.05			Delete leftover files from QPRIME installation/uninstallation
'*		15/06/06		MAl/ARu		v1.06			Improved error handling
'*											Installation performed via passed command line
'*		11/08/06		MAl			v1.07			Code Cleanup
'*		22/08/06		ARu			v1.08			Changes to log entries for uninstall and install commands.
'*      28/08/17		Bill W		v1.09			Support 64 bit Windows.
'===========================================================================================================
'*  Comments:
'*			Log file written to C:\Winnt\logs\QPrimeClientUpdate.log
'*			Uninstallation method change from win32_products' uninstall method to msiexec due to performance problem with win32_products
'*				problems were encountered accessing win32_product class when the script is run
'*			Installation method changed to simply run the passed command line
'===========================================================================================================

Option Explicit
On Error Resume Next

Dim strLogFileName, strLogFilePath, strClientFolder, strOsArch
Dim strLine, strSoftware, strDIR1, strDIR2, strInstCommand, strErrorMessage
Dim objShell, objFileSys, objlogfile, objTempFile
Dim intErr

Const ERROR = True 'for logging errors
Const SUCCESS = False 'for logging success

'open text file parameters
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Set objLogFile = Nothing

'CREATE THE SHELL OBJECT. Is needed to set variables below
Set objShell = CreateObject("Wscript.Shell")
	If Err.number <> 0 Then Call RunQuit(3, "")

'-----------------------------------------------------------------------------------------------------------------------
'log file path and name
strLogFileName = "QPRIMEClientUpdate.log"
strLogFilePath = objShell.ExpandEnvironmentStrings("%SystemRoot%") & "\Logs\"

'-------------------------
'Check Windows architecture (32 or 64 bit)
'-------------------------
strOsArch = GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth

'directories containing files leftover from QPRIME installation/uninstallation
If strOsArch = "32" Then
    strDIR1 = objShell.ExpandEnvironmentStrings("%PROGRAMFILES%") & "\Niche\NicheRMS\BLOBCache\"
    strDIR2 = objShell.ExpandEnvironmentStrings("%PROGRAMFILES%") & "\Niche\NicheRMS\nicheappcache\"
ElseIf strOsArch = "64" Then
    strDIR1 = objShell.ExpandEnvironmentStrings("%PROGRAMFILES(x86)%") & "\Niche\NicheRMS\BLOBCache\"
    strDIR2 = objShell.ExpandEnvironmentStrings("%PROGRAMFILES(x86)%") & "\Niche\NicheRMS\nicheappcache\"
End If

'temporary file path and name
Const TEMPFILENAME = "QPRIMEClient.tmp"
Const TEMPFILEPATH = "C:\QPS\Logs\"
'-----------------------------------------------------------------------------------------------------------------------

'CREATE THE FILE SYSTEM OBJECT. Required to access the hard drive
Set objFileSys = CreateObject("Scripting.FileSystemObject")
	If Err.Number <> 0 Then Call RunQuit(3, "")

'Check if the log location exists
If Not objFileSys.FolderExists(strLogFilePath) Then
	objFileSys.CreateFolder(strLogFilePath)
		If Err.Number <> 0 Then Call RunQuit(3, "Error creating log path")
End If

'CREATE THE LOG FILE
Set objlogfile = objFileSys.OpenTextFile(strLogFilePath & strLogFileName, ForWriting, True)
	If Err.Number <> 0 Then Call RunQuit(3, "")
Call LogHandler("------------------------------------------------------------", "", SUCCESS)
Call LogHandler("LOG FILE: QPRIMEUpgrade", "", SUCCESS)
Call LogHandler("------------------------------------------------------------", "", SUCCESS)
	
'initialise variables to zero length strings for error handling
strSoftware = ""
Call LogHandler("strSoftware initialised to zero length string.", "", SUCCESS)
strClientFolder = ""
Call LogHandler("strClientFolder initialised to zero length string.", "", SUCCESS)
strInstCommand = ""
Call LogHandler("strInstCommand initialised to zero length string.", "", SUCCESS)

'CHECK IF TMP FILE ALREADY EXISTS
Call LogHandler("Checking for Temporary File: " & TEMPFILEPATH & TEMPFILENAME, "", SUCCESS)
If objFileSys.FileExists(TEMPFILEPATH & TEMPFILENAME) Then
	Call LogHandler("Temporary File found.", "", SUCCESS)
	Set objTempFile = objFileSys.GetFile(TEMPFILEPATH & TEMPFILENAME)
	If Err.number <> 0 Then
		Call LogHandler(space(3) & "Temporary file could not be accessed.", "Error accessing temporary file", ERROR)
	Else
		Call LogHandler(space(3) & "Temporary file accessed. Checking creation date...", "", SUCCESS)
	End If
	
	'CHECK TEMPORARY FILE CREATION DATE
	If objTempFile.DateCreated > (Now - 0.004) Then 'younger than 5 minutes
		Call LogHandler(space(3) & "Creation date younger than 5 minutes old. Opening File...", "", SUCCESS)
		Set objTempfile = objFileSys.OpenTextFile(TEMPFILEPATH & TEMPFILENAME, ForReading, False)
			If Err.number <> 0 Then
				Call LogHandler(space(3) & "Temporary file could not be opened for reading.", "Error opening temporary file", ERROR)
			Else
				Call LogHandler(space(3) & "Temporary file opened for reading.", "", SUCCESS)
			End If
			
		'Get values from temporary file
		Call LogHandler(space(3) & "Retrieving Session Values.", "", SUCCESS)
		Do Until objTempFile.AtEndOfStream
			strLine = objTempFile.ReadLine
			If UCASE(Left(strLine, 8)) = "SOFTWARE" Then
				strSoftware = Mid(strLine, 10)
				Call LogHandler(space(6) & "Software name found. Value is " & strSoftware, "", SUCCESS)
			ElseIf UCASE(Left(strLine, 13)) = "CLIENT FOLDER" Then
				strClientFolder = Mid(strLine, 15)
				Call LogHandler(space(6) & "Client Folder found. Value is " & strClientFolder, "", SUCCESS)
			ElseIf UCASE(Left(strLine, 15)) = "INSTALL COMMAND" Then
				strInstCommand = Mid(strLine, 17)
				Call LogHandler(space(6) & "Install Command found. Value is " & strInstCommand, "", SUCCESS)
			End If
		Loop
		
		objTempfile.Close
		Call LogHandler(space(3) & "Temporary file closed.", "", SUCCESS)
		
		'Verify values have been successfully extracted from temporary file
		If strSoftware = "" Then Call LogHandler(space(6) & "Software Name could not be extracted from temporary file.",_
		"Information not found in Temp File", ERROR)
		If strClientFolder = "" Then Call LogHandler(space(6) & "Client Folder Name could not be extracted from temporary file.",_
		"Information not found in Temp File", ERROR)
		If strInstCommand = "" Then Call LogHandler(space(6) & "Install Command could not be extracted from temporary file.",_
		"Information not found in Temp File", ERROR)
	Else 'old temporary file
		Call LogHandler(space(3) & "Creation date indicates old temporary file.", "", SUCCESS)
		Call RunQuit(2, "")
	End If
Else 'file cannot be found
	Call LogHandler("Temporary file could not be found.", "", SUCCESS)
	Call RunQuit(2, "")
End If

Call UnInstallOld
Call DeleteFiles(strDIR1)
Call DeleteFiles(strDIR2)
Call InstallNew

Call RunQuit(0, "")
		
'=======================================================================================================================
'UNINSTALL OLD VERSION IF IT EXISTS
'-----------------------------------------------------------------------------------------------------------------------
Sub UninstallOld
	On Error Resume Next
	
	Dim objSoftware
	Dim strQuery, strUnInstCommand
	Dim colSoftware
	
	Const wbemFlagReturnWhenComplete = 0
	
	strQuery = "SELECT DisplayName, ProdID FROM Win32Reg_AddRemovePrograms WHERE DisplayName = '" & strSoftware & "'"
	
	Call LogHandler("Preparing to uninstall old clients if installed.", "", SUCCESS)
	Set colSoftware = GetObject("winmgmts:\root\cimv2").ExecQuery(strQuery,,wbemFlagReturnWhenComplete)
	If Err.Number <> 0 Then
		Call LogHandler(space(3) & "Unable to query installed software.", "Unable to query installed software", ERROR)
	Else
		Call LogHandler(space(3) & "Queried installed software.", "", SUCCESS)
	End If
	If colSoftware.Count > 0 Then
		For Each objSoftware In colSoftware
			Call LogHandler(space(3) & "'" & objSoftware.DisplayName & "' Found and will be uninstalled.", "", SUCCESS)
			strUnInstCommand = "MSIEXEC.EXE /qn /x " & objSoftware.ProdID
			Call LogHandler(space(6) & "Running Command. '" & strUnInstCommand & "'", "", SUCCESS)
			intErr = objShell.Run(strUnInstCommand,,True)
			If Err.Number <> 0 Then
				Call LogHandler(space(6) & "Failed to Uninstall '" & objSoftware.DisplayName & "'.",_
				"Uninstall Error Code " & Err.Number, ERROR)
			ElseIf intErr <> 0 Then
				Call LogHandler(space(6) & "Failed to Uninstall '" & objSoftware.DisplayName & "'. Exit Code: " & intErr,_
				"Uninstall Exit Code " & intErr, ERROR)
			Else
				Call LogHandler(space(6) & "Sucessfully uninstalled '" & objSoftware.DisplayName & "'", "", SUCCESS)
			End If
		Next
	Else
		Call LogHandler(space(3) & "No Software Found to uninstall.", "", SUCCESS)
	End If
End Sub
'=======================================================================================================================

'=======================================================================================================================
'INSTALL NEW VERSION
'-----------------------------------------------------------------------------------------------------------------------
Sub InstallNew
	On Error Resume Next
	
	Call LogHandler("Preparing to install new client.", "", SUCCESS)
	'Set working directory to client folder
	objShell.CurrentDirectory = strClientFolder
	Call LogHandler("Set Current Working Directory to: " & objShell.CurrentDirectory, "", SUCCESS)
	Call LogHandler("Running Install Command: " & strInstCommand, "", SUCCESS)
	intErr = objShell.Run(strInstCommand,,True)
	If Err.Number <> 0 Then
		Call LogHandler(space(3) & "Failed to install Client.",_ 
		"Install Error Code " & Err.Number, ERROR)
	ElseIf intErr <> 0 Then
		Call LogHandler(space(3) & "Failed to install Client. Exit Code: " & intErr,_ 
		"Install Exit Code " & intErr, ERROR)
	Else
		Call LogHandler(space(3) & "Sucessfully installed Client", "", SUCCESS)
	End If
End Sub
'=======================================================================================================================

'=======================================================================================================================
'LOG HANDLER
'-----------------------------------------------------------------------------------------------------------------------
Sub LogHandler(strMsg, strReturnedError, blnIsError)

	'Save details from Err object because it will be cleared after On Error Resume Next is turned on
	Dim strErrNumber, strErrDescription, strErrSource
	strErrNumber = Err.Number
	strErrDescription = Err.Description
	strErrSource = Err.Source

	On Error Resume Next

	If blnIsError Then
		If strErrNumber <> 0 Then
			objLogFile.WriteLine("FAIL: " & Date & " - " & Time & " - " & strMsg & " - Error Number: " &_
			strErrNumber & " - Error Description: " & strErrDescription)
		Else
			objLogFile.WriteLine("FAIL: " & Date & " - " & Time & " - " & strMsg)
		End If
		Call RunQuit(1, strReturnedError)
	Else
		objLogFile.WriteLine("PASS: " & Date & " - " & Time & " - " & strMsg)
	End If
End Sub
'=======================================================================================================================

'=======================================================================================================================
'QUIT THE SCRIPT
'-----------------------------------------------------------------------------------------------------------------------
Sub RunQuit(intCode, strErrorMessage)
	On Error Resume Next

	' #0 - Update completed without error
	' #1 - Post Log file (logable) fatal errors
	' #2 - Old or no temp file
	' #3 - Pre-Log file (not logable) fatal errors
	
	Select Case intCode
		Case 0
			Call WriteTemp("SUCCESS")
			Call LogHandler("------------------------------------------------------------", "", SUCCESS)
			Call LogHandler("Script finished Execution. " & intCode, "", SUCCESS)
			Call LogHandler("------------------------------------------------------------", "", SUCCESS)
			
		Case 1
			Call WriteTemp("ERROR " & strErrorMessage)
			Call LogHandler("------------------------------------------------------------", "", SUCCESS)
			Call LogHandler("Script Halted Execution due to Error. " & intCode, "", SUCCESS)
			Call LogHandler("------------------------------------------------------------", "", SUCCESS)
			
		Case 2
			Call LogHandler("------------------------------------------------------------", "", SUCCESS)
			Call LogHandler("Script finished execution. " & intCode, "", SUCCESS)
			Call LogHandler("------------------------------------------------------------", "", SUCCESS)
			
		Case 3
			'nothing extra required
	End Select
	If Not objLogFile is Nothing Then objLogFile.Close
	WScript.Quit(intCode)
End Sub
'=======================================================================================================================

'=======================================================================================================================
'WRITE TO THE TEMP FILE
'-----------------------------------------------------------------------------------------------------------------------
Sub WriteTemp(strTemp)
	On Error Resume Next
	
	Call LogHandler("Opening Temporary file to write Exit state.", "", SUCCESS)
	Set objTempfile = objFileSys.OpenTextFile(TEMPFILEPATH & TEMPFILENAME, ForWriting, False)
	If Err.Number <> 0 then
		Call LogHandler("ERROR: Failed to open temporary file for writing.", "", SUCCESS)
	Else
		Call LogHandler("Temporary file successfully opened for writing.", "", SUCCESS)
		objTempfile.WriteLine(strTemp)
		Call LogHandler(space(3) & "Temporary file updated with exit status of " & strTemp, "", SUCCESS)
		objTempfile.Close
		Call LogHandler(space(3) & "Temporary file closed.", "", SUCCESS)
	End If
End Sub
'=======================================================================================================================

'=======================================================================================================================
'DELETE LEFTOVER FILES FROM QPRIME INSTALL DIRECTORY
'-----------------------------------------------------------------------------------------------------------------------
Sub DeleteFiles(strDirectory)
	On Error Resume Next

	Const FORCE = True

	Call LogHandler("Deleting files leftover from QPRIME.", "", SUCCESS)
	If objFileSys.FolderExists(strDirectory) Then
		objFileSys.DeleteFile strDirectory & "*.*", FORCE
		If Err.Number <> 0 Then
			Call LogHandler("Failed to delete all leftover files in directory: " & strDirectory, "Error deleting cache files", ERROR)
		Else
			Call LogHandler("All files deleted from directory: " & strDirectory, "", SUCCESS)
		End If
	Else
		Call LogHandler("Directory: " & strDirectory & " not found.", "", SUCCESS)
	End If
End Sub
'=======================================================================================================================
