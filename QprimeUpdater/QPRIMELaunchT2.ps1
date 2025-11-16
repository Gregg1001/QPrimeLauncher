# ===========================================================================================================
# Script:      QPrimeLaunch.ps1
# Version:     v1.30
# Author:      Gregg Steffensen
# Date:        09/04/2025
# Copyright:   QPS, 2025
# Purpose:     Ensure latest version of QPRIME client is installed
# -----------------------------------------------------------------------------------------------------------
# Change history:
#   Date        Author      Reference       Description
#   ----        ------      ---------       -----------
#   Full change history preserved from original VBS file...
#  Change history:
#   	Date		Author		Reference		Description
#   	----		------		---------   	-----------
#		20/01/06	MAl			v1.00			Initial Version
#		25/01/06	MAl			v1.01			User messages added to error handling
#		31/01/06	MAl			v1.02			Updated user error messages
#		01/02/06	MAl			v1.03			Updated hard coded server INI file location
#		03/02/06	MAl/ARu		v1.04			Timeout time reduced to 5 minutes (CORRECTED x2)
#		31/03/06	MAl/ARu		v1.05			Improved logging incorporation of user status message window
#												handle occurrence links
#		11/05/06	MAl			v1.06			Moved IE progress to display only if update IS required
#		12/05/06	MAl			v1.07			Version verification does not take place when launched via an
#												occurrence link AND if the client is already running
#		14/06/06	MAl			v1.08			Updated install process to accomodate different environments
#		15/06/06	ARu			v1.09			All error messages reviewed. Issues resolved with Err object not returning info.
#		07/07/06	ARu			v1.10			Updated 2 error messages content.
#		11/08/06	MAl			v1.11			Code Cleanup
#		14/08/06	MAl/ARu		v1.12			Changed update method to use SMS instead of ASP page
#		18/08/06	ARu			v1.13			Modified the content of numerous error messages.
#		18/08/06	MAl			v1.14			Incorporation of logging fatal errors to a central database
#		22/08/06	ARu			v1.15			Added Error msg when SMS is already executing program. Modification to LogHandler.
#		22/08/06	ARu			v1.16			
#		03/06/08	RNi			v1.17			Added environment checks to the version check. Log file now displays updater version correctly.
#		19/04/10	RNi			v1.18			Prompts user to enter a valid TRAINING property code before performing its version check.
#												The user entered code is then inserted into the install command.
#												This version is ONLY to be used at the OXLEY Academy.
#		09/03/12	RNi			v1.19			Moved QPRIME Updater log details from local ini to server ini.
#												Added extra QPRIME Update Service info on server ini. If 1st doesn't exist then checks for 2nd info.
#		10/03/12	RNi			v1.20			Fatal error database connection string now sourced from server ini file.
#		23/05/12	RNi			v1.21			Removed hard-coded text when combining install commands A & B with user input.
#		08/11/13	RNi			v1.22			Added QPRIMEINABOX to the list of valid training environment codes and added a space between user input and install command B.
#		16/07/14	RNi			v1.23			Updated Academy selection list & replaced QPRIMELogo.bmp with QPRIMELogo.jpg.
#      23/08/17    BWe         V1.24           Support 64 bit Windows
#      30/05/22    NBo         V1.25           Use NextGen Client NicheRMS.exe
#		28/02/24    CGr			V1.26			Support for new NicheRMS protocol handling (case 7/8 in Sub RunQuit)
#		18/03/24    CGr			V1.27			Updating the argument string to handle Niche 5 and 6 protocol handling
#		15/05/24    CGr			V1.28			Updating the IE logo to use new NCU v6 logo
#		12/06/24    AHe			V1.29			Updated NID link handling functionality; org unit styling and opening in new window
#		19/06/24    AHe/CGr		V1.30			Added SetBarcodePrinter subroutine to update the barcodeprinter value in NicheLocal.INI to the current default printer
# ===========================================================================================================
# Comments:
#   Log file written to D:\Logs\QPRIMELaunch.log
#   Writes a temporary file to D:\Logs\QPrimeClient.tmp
#
#   LogHandler structure:
#       LogHandler("log file", "msg box text", "msg box title", "IE status text", TYPE)
#       log file - message written to the log file cannot occur until the log file has been created
#       IE status text - message displayed in the IE status window cannot occur until the status window has been created
# ===========================================================================================================


# Force script to run
Set-ExecutionPolicy Bypass -Scope Process -Force

# Change to script's directory
Set-Location -Path $PSScriptRoot


# === Variable Declarations ===
$global:strClientApp = ""
$global:strDDEApp = ""
$global:strLocalINIFile = ""
$global:strClientFolder = ""
$global:strArgument = ""
$global:strOsArch = ""
$global:strServerINIFile = ""
$global:strInstalledVersion = ""
$global:strLatestVersion = ""
$global:strInstallerName = ""
$global:strInstalledEnvironment = ""
$global:strLatestEnvironment = ""
$global:strInstalledTraining = ""
$global:strLatestTraining = ""
$global:strInputTraining = ""
$global:strValidTraining = ""
$global:strStatus = ""
$global:strInstallA = ""
$global:strInstallB = ""
$global:strInstCommand = ""
$global:strPackageID1 = ""
$global:strProgramName1 = ""
$global:strPackageID2 = ""
$global:strProgramName2 = ""
$global:strLastRunTime = ""
$global:strLine = ""

$global:strConnString = ""
$global:strDBServer = ""
$global:strDBName = ""
$global:strDBTable = ""
$global:strUpdaterVersion = ""
$global:strUpdaterEnvironment = ""
$global:strUpdaterTraining = ""

$global:intWait = 0
$global:intReturn = 0
$global:intValidTraining = 0
$global:blnLinked = $false
$global:blnProgramComplete = $false

$global:aryValidTraining = @()

# === Constants ===
$global:ForReading = 1
$global:ForWriting = 2
$global:ForAppending = 8

$global:LOGTYPE_ERROR = " FAIL"
$global:SUCCESS = " PASS"
$global:FATAL = "FATAL"

# === Objects ===
$global:objFileSys = $null
$global:objFile = $null
$global:objLogFile = $null
$global:objShell = New-Object -ComObject "WScript.Shell"
$global:objIEProgress = $null
$global:objUIResource = $null
$global:objProgram = $null

#GS
$global:strProgName = ""
$global:strPackageId = ""

$global:strLastRunTime

#---
#---------------------------------------
# 1. CHECK IF THE QPRIME CLIENT IS RUNNING
#---------------------------------------

# === Check if QPRIME Client is running ===

# [IS QPRIME CLIENT RUNNING]
function QPRIMEClientExecuting {
    try {
        $processes = Get-CimInstance Win32_Process -Filter "Name = '$CLIENTEXE'"
        return $processes.Count -gt 0
    } catch {
        LogHandler -LogFile "Unable to enumerate processes." -MsgBoxText "Error checking if QPRIME Client is running." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
        return $false
    }
}


#-------------------------------------------------------------
# 2. GET THE SERVER INI FILE LOCATION FROM QPRIMELAUNCH  INI FILE
#-------------------------------------------------------------

function GetAddresses {

   #GS

    # $qprimeLaunchIni = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) "QPRIMELaunch.ini"

    $qprimeLaunchIni = Join-Path $PSScriptRoot "QPRIMELaunch.ini"


    LogHandler -LogFile "Retrieving addresses from QPRIMELaunch INI file." -Type " PASS"

    if ($objFileSys.FileExists($qprimeLaunchIni)) {
        LogHandler -LogFile "   QPRIMELaunch INI file found at $qprimeLaunchIni" -Type " PASS"
        try {
            $file = $objFileSys.OpenTextFile($qprimeLaunchIni, $ForReading, $false)
        } catch {
            LogHandler -LogFile "   Error opening QPRIMELaunch INI file for reading." -MsgBoxText "An error has occurred checking for updates." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
            return
        }
        LogHandler -LogFile "   QPRIMELaunch INI file opened for reading." -Type " PASS"


        while (-not $file.AtEndOfStream) {
            $line = $file.ReadLine()
            switch -Regex ($line.ToUpper()) {
                "^SERVER INI FILE" {
                    $global:strServerINIFile = $line.Substring(16).Trim()
                    LogHandler -LogFile "      Server INI file location is: $($line.Substring(16).Trim())" -Type " PASS"
                }
                "^VERSION" {
                    $global:strUpdaterVersion = $line.Substring(8).Trim()
                    LogHandler -LogFile "      Updater Version is: $($line.Substring(8).Trim())" -Type " PASS"
                }
                "^INSTALLED ENVIRONMENT" {
                    $global:strUpdaterEnvironment = $line.Substring(22).Trim()
                    LogHandler -LogFile "      Updater environment is: $($line.Substring(22).Trim())" -Type " PASS"
                }
                "^INSTALLED TRAINING ENVIRONMENT" {
                    $global:strUpdaterTraining = $line.Substring(31).Trim()
                    LogHandler -LogFile "      Updater training environment is: $($line.Substring(31).Trim())" -Type " PASS"
                }
            }
        }

        $file.Close()

        #GS
        
        Write-Output "Here $strServerINIFile"

        
        if (-not $strServerINIFile) {
            LogHandler -LogFile "   Error extracting Server INI file location from QPRIMELaunch INI." -MsgBoxText "An error has occurred checking for updates." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
        }
        if (-not $strUpdaterVersion) {
            LogHandler -LogFile "   Error extracting Updater Version from QPRIMELaunch INI." -MsgBoxText "An error has occurred checking for updates." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
        }
        if (-not $strUpdaterEnvironment) {
            LogHandler -LogFile "   Error extracting Updater environment from QPRIMELaunch INI." -MsgBoxText "Unable to extract Updater environment code." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
        }
        if (-not $strUpdaterTraining) {
            LogHandler -LogFile "   Error extracting Updater training environment from QPRIMELaunch INI." -MsgBoxText "Unable to extract Updater training code." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
        }
        LogHandler -LogFile "   QPRIMELaunch INI file closed." -Type " PASS"
    } else {
        LogHandler -LogFile "   QPRIMELaunch INI file not found at $qprimeLaunchIni" -MsgBoxText "An error has occurred checking for updates." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
    }
}


#-------------------------------------------------------------
# 3.GET INSTALLED VERSION NUMBER FROM INI FILE ON LOCAL COMPUTER
# ------------------------------------------------------------

# === Get Installed Version from local INI file ===
function Test-FileExists {
    param([string]$Path)
    return $global:objFileSys.FileExists($Path)
}

function Open-INIFile {
    param([string]$Path)
    return $global:objFileSys.OpenTextFile($Path, $global:ForReading, $false)
}

function GetInstalledVersion {
    #param(
    #    [string]$LocalIniPath
    #)

    $LocalIniPath = $strLocalINIFile

    LogHandler -LogFile "Retrieving INSTALLED version number from local INI file." -Type "PASS"

    if (Test-FileExists -Path $LocalIniPath) {
        LogHandler -LogFile "Local INI file found at $LocalIniPath" -Type "PASS"

        try {
            $file = Open-INIFile -Path $LocalIniPath
        } catch {
            LogHandler -LogFile "Error opening local INI file for reading. Assumed to have old or no version installed." -Type "FAIL"
            return
        }

        LogHandler -LogFile "Local INI file opened for reading." -Type "PASS"

        try {
            while (-not $file.AtEndOfStream) {
                $line = $file.ReadLine()
                switch -Regex ($line.ToUpper()) {
                    "^INSTALLED VERSION" {
                        $global:strInstalledVersion = $line.Split('=')[1].Trim()
                        LogHandler -LogFile "Installed version is: $global:strInstalledVersion" -Type "PASS"
                    }
                    "^INSTALLED ENVIRONMENT" {
                        $global:strInstalledEnvironment = $line.Split('=')[1].Trim()
                        LogHandler -LogFile "Installed environment is: $global:strInstalledEnvironment" -Type "PASS"
                    }
                    "^INSTALLED TRAINING ENVIRONMENT" {
                        $global:strInstalledTraining = $line.Split('=')[1].Trim()
                        LogHandler -LogFile "Installed training environment is: $global:strInstalledTraining" -Type "PASS"
                    }
                }
            }


switch -Regex ($line.ToUpper()) {
    "^INSTALLED VERSION" {
        $global:strInstalledVersion = $line.Split('=')[1].Trim()
        LogHandler -LogFile "Installed version is: $global:strInstalledVersion" -Type "PASS"
    }
    "^INSTALLED ENVIRONMENT" {
        $global:strInstalledEnvironment = $line.Split('=')[1].Trim()
        LogHandler -LogFile "Installed environment is: $global:strInstalledEnvironment" -Type "PASS"
    }
    "^INSTALLED TRAINING ENVIRONMENT" {
        $global:strInstalledTraining = $line.Split('=')[1].Trim()
        LogHandler -LogFile "Installed training environment is: $global:strInstalledTraining" -Type "PASS"
    }
}


        } catch {
            LogHandler -LogFile "Error reading INI file content." -Type "FAIL"
        } finally {
            $file.Close()
            LogHandler -LogFile "Local INI file closed." -Type "PASS"
        }

        if (-not $global:strInstalledVersion) {
            LogHandler -LogFile "Error extracting installed version from local INI file." -Type "FAIL"
        }
        if (-not $global:strInstalledEnvironment) {
            LogHandler -LogFile "Error extracting installed environment from local INI file." -Type "FAIL"
        }
        if (-not $global:strInstalledTraining) {
            LogHandler -LogFile "Error extracting installed training environment from local INI file." -Type "FAIL"
        }

    } else {
        LogHandler -LogFile "Local INI file not found at $LocalIniPath. Assumed to have old or no version installed." -Type "FAIL"
    }
}




# ----------------------------------------------------------
# 4.SET NICHELOCAL.INI BARCODEPRINTER VALUE TO DEFAULT PRINTER 
# ----------------------------------------------------------

# === Set BarcodePrinter in NicheLocal.INI to Default Printer ===
function SetBarcodePrinter {
    try {
        LogHandler -LogFile "Retrieving NicheLocal.INI" -Type " PASS"
        $defaultPrinter = Get-CimInstance Win32_Printer | Where-Object { $_.Default -eq $true } | Select-Object -First 1

        if ($null -eq $defaultPrinter) {
            LogHandler -LogFile "Could not find default printer." -Type "ERROR"
            return
        }

        $barcodePrinterName = $defaultPrinter.Name
        LogHandler -LogFile "Default Printer found: $barcodePrinterName" -Type " PASS"

        $nicheIniPath = "$env:ProgramFiles (x86)\Niche\NicheRMS\ClientBin\NicheLocal.INI"
        if (-Not (Test-Path $nicheIniPath)) {
            LogHandler -LogFile "Couldn't read NicheLocal.ini, barcode printer not set." -Type "ERROR"
            return
        }

        $content = Get-Content -Path $nicheIniPath
        $updated = $false
        $printerCount = 0

        $newContent = $content | ForEach-Object {
            if ($_ -like "BarcodePrinter=*") {
                $printerCount++
                if ($printerCount -eq 1) {
                    LogHandler -LogFile "Barcode Printer set to $barcodePrinterName" -Type " PASS"
                    return "BarcodePrinter=$barcodePrinterName"
                } else {
                    return ""
                }
            } else {
                return $_
            }
        }

        if ($printerCount -ne 1) {
            LogHandler -LogFile "$printerCount instances of Barcode Printer found, barcode printing may not work" -Type "ERROR"
        }

        Set-Content -Path $nicheIniPath -Value $newContent -Force
    } catch {
        LogHandler -LogFile "Couldn't update NicheLocal.ini" -Type "ERROR"
    }
}

#-------------------------------------------------
# 5.GET LATEST TRAINING ENVIRONMENT FROM USER INPUT
#-------------------------------------------------


# === Get latest training environment from user input ===
function GetInputTraining {
    try {
        LogHandler -LogFile "Requesting LATEST Training Environment code from user input." -Type " PASS"
        $global:strInputTraining = [System.Windows.Forms.Interaction]::InputBox("Enter a database name:", "Training Database")
        $global:intValidTraining = 1
        foreach ($strValidTraining in $global:aryValidTraining) {
            if ($strValidTraining.ToUpper() -eq $strInputTraining.ToUpper()) {
                $global:intValidTraining = 0
                LogHandler -LogFile "User input: $strInputTraining is a valid Training Environment Code." -Type " PASS"
                break
            }
        }

        if ($intValidTraining -eq 1) {
            LogHandler -LogFile "User input: $strInputTraining does not match a valid Training Environment Code." -Type " PASS"
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.MessageBox]::Show("The name entered does not exist. Click OK and re-enter database name.", "Input Error", 'OK', 'Question') | Out-Null
        }
    } catch {
        LogHandler -LogFile "Error processing training environment input." -Type "ERROR"
    }
}

#--------------------------------------------------------
# 6.GET LATEST VERSION NUMBER FROM INI FILE ON NICHE SERVER
#--------------------------------------------------------
#[CHECK VERSION AND ENVIRONMENT]

# === Get latest version from INI on Niche server ===
function GetLatestVersion {
    try {
        LogHandler -LogFile "Retrieving LATEST version number from server INI file." -Type "PASS"

        #[IS SERVER INI AVAILABLE AND ON?]
        if (Test-Path $global:strServerINIFile) {
            LogHandler -LogFile "Server INI file found at $global:strServerINIFile" -Type "PASS"
            $lines = Get-Content $global:strServerINIFile

            foreach ($line in $lines) {
                $split = $line.Split('=', 2)
                if ($split.Length -ne 2) { continue }  # skip invalid lines
                $key = $split[0].Trim().ToUpper()
                $value = $split[1].Trim()

                switch -Wildcard ($key) {
                    "RELEASED VERSION"               { $global:strLatestVersion     = $value; LogHandler -LogFile "Latest Version: $value" -Type "PASS" }
                    "RELEASED ENVIRONMENT"            { $global:strLatestEnvironment = $value; LogHandler -LogFile "Latest Environment: $value" -Type "PASS" }
                    "RELEASED TRAINING ENVIRONMENT"   { $global:strLatestTraining    = $value; LogHandler -LogFile "Latest Training Environment: $value" -Type "PASS" }
                    "RELEASED INSTALLER"              { $global:strInstallerName     = $value; LogHandler -LogFile "Installer Name: $value" -Type "PASS" }
                    "STATUS"                          { $global:strStatus            = $value; LogHandler -LogFile "Status: $value" -Type "PASS" }
                    "INSTALL COMMAND"                 { $global:strInstCommand       = $value; LogHandler -LogFile "Install Command: $value" -Type "PASS" }
                    "INSTALL A"                       { $global:strInstallA          = $value; LogHandler -LogFile "Install A: $value" -Type "PASS" }
                    "INSTALL B"                       { $global:strInstallB          = $value; LogHandler -LogFile "Install B: $value" -Type "PASS" }
                    "CONNECTION STRING"               { $global:strConnString        = $value; LogHandler -LogFile "Connection String: $value" -Type "PASS" }
                    "DATABASE SERVER"                 { $global:strDBServer          = $value; LogHandler -LogFile "Database Server: $value" -Type "PASS" }
                    "DATABASE NAME"                   { $global:strDBName            = $value; LogHandler -LogFile "Database Name: $value" -Type "PASS" }
                    "DATABASE TABLE"                  { $global:strDBTable           = $value; LogHandler -LogFile "Database Table: $value" -Type "PASS" }
                    "PACKAGE ID1"                     { $global:strPackageID1        = $value; LogHandler -LogFile "Package ID1: $value" -Type "PASS" }
                    "PROGRAM NAME1"                   { $global:strProgramName1      = $value; LogHandler -LogFile "Program Name1: $value" -Type "PASS" }
                    "PACKAGE ID2"                     { $global:strPackageID2        = $value; LogHandler -LogFile "Package ID2: $value" -Type "PASS" }
                    "PROGRAM NAME2"                   { $global:strProgramName2      = $value; LogHandler -LogFile "Program Name2: $value" -Type "PASS" }
                }
            }

            $requiredVars = @(
                'strLatestVersion', 'strLatestEnvironment', 'strLatestTraining', 'strInstallerName', 'strStatus', 'strInstCommand',
                'strInstallA', 'strInstallB', 'strConnString', 'strDBServer', 'strDBName', 'strDBTable',
                'strPackageID1', 'strProgramName1', 'strPackageID2', 'strProgramName2'
            )

            foreach ($var in $requiredVars) {
                if (-not (Get-Variable -Name $var -Scope Global).Value) {
                    LogHandler -LogFile "Error extracting $var from server INI file." -Type "FATAL"
                }
            }

            LogHandler -LogFile "Server INI file closed." -Type "PASS"

            if ($global:strStatus.ToUpper() -eq "ON") {
                LogHandler -LogFile "QPRIME is ONline. Status: ON." -Type "PASS"
            } else {
                LogHandler -LogFile "QPRIME is OFFline. Status: $($global:strStatus)" -Type "ERROR"
                RunQuit 2
            }
        } else {
            #LogHandler -LogFile "Server INI file not found at $global:strServerINIFile" -Type "FATAL"

            LogHandler -LogFile "Server INI file not found at $global:strServerINIFile"-MsgBoxText "Unable to contact the QPRIME server. If this is a`nplanned outage, refer to your notification message." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"

            RunQuit 2
        }
    } catch {

    #[IS SERVER INI AVAILABLE - AND ON]
        LogHandler -LogFile "Exception encountered during server INI parsing." -Type "FATAL"
        RunQuit 2
    }
}


#---------------------------------
# 7. HANDLE SUCCESS AND ERROR LOGGING
#---------------------------------

#GBS fix at homeoffice

function LogHandler {
    param (
        [string]$LogFile = "",
        [string]$MsgBoxText = "",
        [string]$MsgBoxTitle = "",
        [string]$IEStatusText  = "",
        [string]$Type = "INFO"
    )

    $errorInfo = 0

    if ($LogFile -and $global:objLogFile) {
        if ($errorInfo -ne 0) {
            $global:objLogFile.WriteLine("{$Type}: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $LogFile - Error Code: $errorInfo")
        } else {
            $global:objLogFile.WriteLine("{$Type}: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $LogFile")
        }
    }

    if ($Type -eq 'FATAL') {
        if (-not $global:objLogFile) {
            DBLogFatal -LogType $Type -LogMessage $MsgBoxText
        } else {
            DBLogFatal -LogType $Type -LogMessage $LogFile
        }
    }

    if ($MsgBoxText) {
        Add-Type -AssemblyName System.Windows.Forms
        if ($errorInfo -ne 0) {
            [System.Windows.Forms.MessageBox]::Show("$MsgBoxText`n`n$global:MSGSUFFIX`n`nError Details:`n  Code: $errorInfo", $MsgBoxTitle, 'OK', 'Error')
        } else {
            [System.Windows.Forms.MessageBox]::Show("$MsgBoxText`n`n$global:MSGSUFFIX", $MsgBoxTitle, 'OK', 'Warning')
        }
    }

    if ($Type -eq 'FATAL') {
        if (-not $global:objLogFile) {
            RunQuit 1
        } else {
            RunQuit 2
        }
    }

    if ($IEStatusText -and $global:progressForm -and !$global:progressForm.IsDisposed -and $global:progressLabel -and !$global:progressLabel.IsDisposed) {
        if ($global:progressLabel.InvokeRequired) {
            try {
                $global:progressLabel.Invoke([Action]{ $global:progressLabel.Text = $using:IEStatusText })
            } catch {}
        } else {
            $global:progressLabel.Text = $IEStatusText 
        }
    }
}

# -----------------------------------------------------------
# 8. CONNECT TO THE CENTRAL LOGGING DATABASE AND LOG FATAL ERROR
# -----------------------------------------------------------
function DBLogFatal {
    param (
        [string]$LogType,
        [string]$LogMessage
    )

    if (-not $global:objShell) { return }

    $PCName = $env:COMPUTERNAME
    $UserDomain = $env:USERDOMAIN
    $UserName = $env:USERNAME

    if (-not $PCName) {
        LogHandler -LogFile "Unable to expand COMPUTERNAME environment variable." -Type "ERROR"
        return
    }
    if (-not $UserDomain) {
        LogHandler -LogFile "Unable to expand USERDOMAIN environment variable." -Type "ERROR"
        return
    }
    if (-not $UserName) {
        LogHandler -LogFile "Unable to expand USERNAME environment variable." -Type "ERROR"
        return
    }

    LogHandler -LogFile "Connection string: $global:strConnString" -Type " PASS"

    $InsertSQL = @"
    INSERT INTO $($global:strDBTable)
        (PCName, Application, Version, UserDomain, UserName, LogType, LogText, LogTime)
    VALUES
        ('$PCName', 'QPRIME Updater', '$($global:strUpdaterVersion)', '$UserDomain', '$UserName', '$LogType', '$($LogMessage.Trim())', GETDATE())
"@

    LogHandler -LogFile "SQL Insert Command: $InsertSQL" -Type " PASS"

    
$connString = "Server=CIT-SMS-PR-01;Database=QPS;Integrated Security=True;"
$conn = New-Object System.Data.SqlClient.SqlConnection $connString


#    $conn = New-Object -ComObject ADODB.Connection
    if (-not $conn) {
        LogHandler -LogFile "Error creating ADO database object." -Type "ERROR"
        return
    }

    #GBS

     Write-Output TestString2 $connString

    # $conn.Open("Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=QPS;Data Source=CIT-SMS-PR-01")
     # $conn.Open("Provider=SQLOLEDB;Data Source=TCP:CIT-SMS-PR-01,1433;Initial Catalog=QPS;Integrated Security=SSPI;")

    if ($conn.State -ne 1) {
        LogHandler -LogFile "Error opening connection to database." -Type "ERROR"
        return
    }
    LogHandler -LogFile "Opened connection to database." -Type " PASS"

    try {
        $conn.Execute($InsertSQL)
        LogHandler -LogFile "SQL insert command executed." -Type " PASS"
    } catch {
        LogHandler -LogFile "Error executing SQL insert command." -Type "ERROR"
    } finally {
        $conn.Close()
        LogHandler -LogFile "Closed connection to database." -Type " PASS"
    }
}


#-------------------------------------------------------------------------
#9.REFRESH THE SMS/SCCM CLIENT TO ENSURE THE QPRIME INSTALLER IS DOWNLOADING
#-------------------------------------------------------------------------

# === Refresh SMS/SCCM Client Policies ===

#[HAS THE INSTALLER BEEN DOWNLOADED?]

function RefreshSMS {
    try {
        LogHandler -LogFile "Preparing to refresh SMS/SCCM client." -Progress "   Refreshing SMS/SCCM Client Policy" -Type " PASS"
        $smsClient = New-Object -ComObject CPApplet.CPAppletMgr
        $actions = $smsClient.GetClientActions()

        foreach ($action in $actions) {
            if ($action.Name -like '*Machine Policy*') {
                $action.PerformAction()
                LogHandler -LogFile "SMS/SCCM client action successfully performed." -Type " PASS"
            }
        }
        LogHandler -LogFile "SMS/SCCM client policy refresh finished." -Type " PASS"
    } catch {
        LogHandler -LogFile "Error connecting or performing action on SMS/SCCM client." -MsgBoxText "An error has been encountered with SMS/SCCM.`nThe SMS/SCCM Client may not be working on your machine." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
    }
}


#------------------------------------
# '10.RERUN THE SMS/SCCM ADVERTISMENT
#------------------------------------

# === Re-run SMS/SCCM Advertisement ===

#[SMS PACKAGE INITIATED]
 


 function RunSMSAdvert {
    try {
        $programs = $objUIResource.GetAvailableApplications()
        if ($null -eq $programs) {
            LogHandler -LogFile "Failed to get SMS/SCCM programs collection." -MsgBoxText "An error has been encountered launching upgrade with SMS/SCCM." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
            return
        }

        LogHandler -LogFile "$($programs.Count) SMS/SCCM programs found." -Type "PASS"

        $foundProgram1 = $false
        $foundProgram2 = $false

        function TryExecuteProgram($pkgId, $progName) {
            $wait = 0
            $program = $programs | Where-Object {
                $_.PackageID -eq $pkgId -and $_.Name -eq $progName
            }

            if ($null -eq $program) {
                return $false
            }

            LogHandler -LogFile "Program found: $($program.Id). Preparing to run." -Type "PASS"

            $lastRun = $program.LastRunTime
            LogHandler -LogFile "Last run time is: $lastRun" -Type "PASS"

            while ($true) {
                $wait++
                LogHandler -LogFile "Executing program $($progName) $($pkgId)" -Type "PASS"
                try {
                    $global:objUIResource.ExecuteProgram($progName, $pkgId, $true)

                    $global:strProgName = $progName
                    $global:strPackageId = $pkgId
                    $global:strLastRunTime = $lastRun

                    LogHandler -LogFile "Program initiated. Waiting for response." -Type "PASS"
                    return $true
                } catch {
                    $hresult = $_.Exception.HResult
                    if ($hresult -eq -2147450879) {
                        if ($wait -ge 60) {
                            LogHandler -LogFile "Timed out waiting for another SMS/SCCM program to complete." -MsgBoxText "This computer is currently installing other software.`nPlease try again later." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
                            return $false
                        } else {
                            LogHandler -LogFile "Waiting for another SMS/SCCM program to finish." -Progress "  Waiting to begin update $('.' * ($wait % 7))" -Type "PASS"
                            Start-Sleep -Seconds 1
                        }
                    } elseif ($hresult -ne 0) {
                        LogHandler -LogFile "Failed to execute SMS/SCCM program. HRESULT=$hresult" -MsgBoxText "An error has been encountered launching upgrade with SMS/SCCM." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
                        return $false
                    }
                }
            }
        }

        $foundProgram1 = TryExecuteProgram $strPackageID1 $strProgramName1

        if (-not $foundProgram1) {
            LogHandler -LogFile "Searching for package with ID: $strPackageID2, with name: $strProgramName2" -Type "PASS"
            $foundProgram2 = TryExecuteProgram $strPackageID2 $strProgramName2
        }

        if (-not $foundProgram1) {
            LogHandler -LogFile "Failed to find SMS/SCCM package with ID: $strPackageID1, with name: $strProgramName1" -Type "ERROR"
        }
        if (-not $foundProgram2) {
            LogHandler -LogFile "Failed to find SMS/SCCM package with ID: $strPackageID2, with name: $strProgramName2" -Type "ERROR"
        }

        if (-not $foundProgram1 -and -not $foundProgram2) {
            Refresh-SMS
            LogHandler -LogFile "Failed to find the QPRIME Update Service SMS/SCCM package. SMS/SCCM policy refreshed." `
                       -MsgBoxText "The QPRIME Update Service could not be found by SMS/SCCM.`nIf this is a new computer, it may not have been downloaded to your computer yet.`nPlease try again later.`n`nThis computer may also not be setup to receive QPRIME Updates." `
                       -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
        }

        $programs = $null
    } catch {
        LogHandler -LogFile "Failed to get SMS/SCCM programs collection." -MsgBoxText "An error has been encountered launching upgrade with SMS/SCCM." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
    }
}



#--------------------
# 11.IE PROGRESS MESSAGES
#--------------------

# === Initialize IE Progress Window ===

function InitIEProgress {
    param ([string]$strProgress)

    LogHandler -LogFile "Preparing to initialise Internet Explorer progress window." -Type "PASS"
    LogHandler -LogFile "Creating Internet Explorer Object." -Type "PASS"
    LogHandler -LogFile "Internet Explorer object created." -Type "PASS"

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object Windows.Forms.Form
    $form.SuspendLayout()
    $form.Text = "QPRIME Update Status"
    $form.Size = New-Object Drawing.Size(500, 220)
    $form.StartPosition = "Manual"
    $form.Location = New-Object Drawing.Point(-2000, -2000)  # Offscreen init
    $form.BackColor = [Drawing.Color]::FromArgb(0, 90, 163)
    $form.FormBorderStyle = 'FixedDialog'
    $form.ControlBox = $false
    $form.ShowInTaskbar = $false

    $scriptPath = $PSCommandPath
    $splitPath = Split-Path -Parent $scriptPath
    $logoPath = Join-Path -Path $splitPath -ChildPath "QPRIMELogo.png"

    $img = New-Object Windows.Forms.PictureBox
    $img.Image = [System.Drawing.Image]::FromFile($logoPath)
    $img.Size = New-Object Drawing.Size(100, 100)
    $img.Location = New-Object Drawing.Point(10, 10)
    $img.SizeMode = 'StretchImage'
    $form.Controls.Add($img)

    $label1 = New-Object Windows.Forms.Label
    $label1.Text = $strProgress
    $label1.ForeColor = 'White'
    $label1.BackColor = [System.Drawing.Color]::Transparent
    $label1.Font = New-Object Drawing.Font("Arial", 14, [Drawing.FontStyle]::Regular)
    $label1.Location = New-Object Drawing.Point(120, 20)
    $label1.Size = New-Object Drawing.Size(350, 35)
    $form.Controls.Add($label1)

    $label2 = New-Object Windows.Forms.Label
    $label2.Text = "Please Wait"
    $label2.ForeColor = 'White'
    $label2.BackColor = [System.Drawing.Color]::Transparent
    $label2.Font = New-Object Drawing.Font("Arial", 10, [Drawing.FontStyle]::Italic)
    $label2.Location = New-Object Drawing.Point(120, 60)
    $label2.Size = New-Object Drawing.Size(350, 35)
    $form.Controls.Add($label2)

    $global:progressForm = $form
    $global:progressLabel = $label1

    $form.ResumeLayout($false)
    $form.PerformLayout()
    $form.Show()
    $form.Refresh()

    Start-Sleep -Milliseconds 100  # Give time to finish drawing
    $form.Location = New-Object Drawing.Point(0, 0)  # Show onscreen at top-left

    LogHandler -LogFile "Internet Explorer page loaded." -Type "PASS"
}



#--------------------------------------
# 12.LAUNCH THE CLIENT AND QUIT THE SCRIPT
#--------------------------------------

# === Run Client and Quit Script ===
function RunQuit {
    param([int]$ExitCode)


    try {
        if ($global:objIEProgress -ne $null) {
            $global:objIEProgress.Quit()
            $global:objIEProgress = $null
        }



        if ($global:progressForm -and !$global:progressForm.IsDisposed) {
            try {
            $global:progressForm.Invoke([Action]{ $global:progressForm.Close() })
        } catch {
            $global:progressForm.Close()
        }
            $global:progressForm = $null
            $global:progressLabel = $null
        }


        switch ($ExitCode) {
            1 {}
            2 {
                RemoveTempFile
                LogHandler -LogFile ("-"*60) -Type " PASS"
                LogHandler -LogFile "Script Halted Execution due to Fatal Error. $ExitCode" -Type "ERROR"
                LogHandler -LogFile ("-"*60) -Type " PASS"
            }
            3 {}
            5 {
                SetBarcodePrinter
                LaunchClient -Path $global:strClientApp
                LogHandler -LogFile ("-"*60) -Type " PASS"
                LogHandler -LogFile "Script Finished Executing. $ExitCode" -Type " PASS"
                LogHandler -LogFile ("-"*60) -Type " PASS"
            }
            6 {
                RemoveTempFile
                SetBarcodePrinter
                LaunchClient -Path $global:strClientApp
                LogHandler -LogFile ("-"*60) -Type " PASS"
                LogHandler -LogFile "Script Finished Executing. $ExitCode" -Type " PASS"
                LogHandler -LogFile ("-"*60) -Type " PASS"
            }
            7 {
                LaunchClient -Path "$global:strClientApp -D urlcommand $global:strArgument"
                LogHandler -LogFile ("-"*60) -Type " PASS"
                LogHandler -LogFile "Script Finished Executing. $ExitCode" -Type " PASS"
                LogHandler -LogFile ("-"*60) -Type " PASS"
            }
            8 {
                RemoveTempFile
                LaunchClient -Path "$global:strClientApp -D urlcommand $global:strArgument"
                LogHandler -LogFile ("-"*60) -Type " PASS"
                LogHandler -LogFile "Script Finished Executing. $ExitCode" -Type " PASS"
                LogHandler -LogFile ("-"*60) -Type " PASS"
            }
        }

        if ($global:objLogFile -ne $null) {
            $global:objLogFile.Close()
            $global:objLogFile = $null
        }

        $global:objFileSys = $null
        $global:objFile = $null
        $global:objShell = $null
        $global:objProgram = $null
        $global:objTempProgram = $null
        $global:objUIResource = $null

        Exit $ExitCode

    } catch {
        Exit 99
    }
}

#----------------------------------------------------
#13.CHECK FOR AND DELETE THE TEMPORARY FILE IF IT EXISTS
#----------------------------------------------------

# === Delete Temporary File If It Exists ===
function RemoveTempFile {
    try {
        LogHandler -LogFile "Checking for temporary file." -Type " PASS"
        $tempFile = Join-Path $global:TEMPFILEPATH $global:TEMPFILENAME
        if (Test-Path $tempFile) {
            LogHandler -LogFile "Temporary file found." -Type " PASS"
            Remove-Item $tempFile -Force
            LogHandler -LogFile "Temporary file deleted." -Type " PASS"
        } else {
            LogHandler -LogFile "Temporary file not found." -Type " PASS"
        }
    } catch {
        LogHandler -LogFile "Error deleting temporary file. Need to delete manually." -Type "ERROR"
    }
}

#------------------------
#14.LAUNCH THE QPRIME CLIENT
#------------------------

# === Launch the QPRIME Client ===

#[ENTER QPRIME]
function LaunchClient {
    param([string]$Path)
    try {
        LogHandler -LogFile "Launching QPRIME Client. Running: $Path" -Type " PASS"

        #gs
        #$global:objShell.Run("C:\Program Files (x86)\Niche\\NicheRMS\ClientBin\NicheRMS.exe")
        
        #Set objShell = CreateObject("Wscript.Shell")

        $execPath = "C:\Program Files (x86)\Niche\NicheRMS\ClientBin\NicheRMS.exe"
        Start-Process -FilePath $execPath -WorkingDirectory (Split-Path $execPath)

        LogHandler -LogFile "QPRIME Client Launched." -Type " PASS"
    } catch {
        LogHandler -LogFile "Error launching QPRIME client." -MsgBoxText "QPRIME Client failed to Launch." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
    }
}


# === Start of Main Section ========


#For the Status Window
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#Temp - remove
#RefreshSMS

#GS
#$TEMPFILENAME = "QPRIMEClient.tmp"
#$TEMPFILEPATH = "C:\QPS\Logs\"
#RemoveTempFile

# === Create the Shell Object ===
try {
    $global:objShell = New-Object -ComObject "WScript.Shell"
} catch {
    LogHandler -LogFile "" -MsgBoxText "Error creating the Shell Object." `
               -MsgBoxTitle "QPRIME Updater: Error" -IEStatusText "" -Type "FATAL"
    return
}

# === Check Windows architecture (32 or 64 bit) ===
$cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
$strOsArch = $cpu.AddressWidth.ToString()

# === Define Constants ===
$CLIENTEXE = "NicheRMS.exe"
$SOFTWARE = "QPrime Client"
$LOGFILENAME = "QPRIMEClientLaunchPS.log"
$LOGFILEPATH = "C:\QPS\Logs\"
$TEMPFILENAME = "QPRIMEClient.tmp"
$TEMPFILEPATH = "C:\QPS\Logs\"
$MSGSUFFIX = "For further assistance contact the Service Desk."

# === Build paths depending on architecture ===
if ($strOsArch -eq "32") {
    $strClientApp = $objShell.ExpandEnvironmentStrings("%ProgramFiles%") + "\Niche\NicheRMS\ClientBin\" + $CLIENTEXE
    $strLocalINIFile = $objShell.ExpandEnvironmentStrings("%ProgramFiles%") + "\Niche\NicheRMS\ClientBin\QPrimeInstalledVersion.ini"
} elseif ($strOsArch -eq "64") {
    $strClientApp = $objShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%") + "\Niche\NicheRMS\ClientBin\" + $CLIENTEXE
    $strLocalINIFile = $objShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%") + "\Niche\NicheRMS\ClientBin\QPrimeInstalledVersion.ini"
}

# === Define DDE Launcher App ===
$scriptFullName = $MyInvocation.MyCommand.Path
$strDDEApp = $scriptFullName.Substring(0, $scriptFullName.LastIndexOf('\') + 1) + "LAUNCHER.EXE"

# === Local Client Installer Folder ===
$strClientFolder = $objShell.ExpandEnvironmentStrings("%SystemDrive%") + "\QPS\QPrimeClients\"

# === Training Code Handling ===
$global:intValidTraining = 1
$global:aryValidTraining = @(
    "ACA000", "ACA001", "ACA002", "ACA003", "ACA004",
    "ACA005", "ACA006", "ACA007", "ACA008"
)

# === Argument Handling ===
$global:blnLinked = $false
$global:strArgument = ""

# to test GBS
# === Refresh SMS/SCCM Client Policies (if installer not found) ===
#LogHandler -LogFile "QPRIME installer: $strInstallerName not found locally in $strClientFolder" -Type " FAIL"
#RefreshSMS
#LogHandler -LogFile "QPRIME installer: $strInstallerName not found locally in $strClientFolder. SMS/SCCM policy refreshed." `
#            -MsgBoxText "A new version of QPRIME has been released however the installer`nhas not been downloaded to your computer yet, or the download has failed.`nPlease try again later.`n`nYour computer may also not be setup to receive QPRIME Updates." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"

#Write-Output TheEnd

if ($args.Count -eq 1) {
    $global:blnLinked = $true
    $strArgument = $args[0]
    if ($strArgument.StartsWith("nds://")) {
        $strArgument = $strArgument -replace "^nds://", "nicherms:open:"
        $strArgument = $strArgument.TrimEnd("/")
    }

    if ($strArgument.StartsWith("nicherms:open:1410")) {
        $strArgument += ":OpenStyle:OpPoliceEmployee"
    } else {
        $strArgument += ":OpenType:OpenWindow"
    }

    $global:strArgument = $strArgument
}
elseif ($args.Count -eq 0) {
    $global:blnLinked = $false
}
else {
    LogHandler -LogFile "" -MsgBoxText "Too many arguments passed." `
               -MsgBoxTitle "QPRIME Updater: Error" -IEStatusText "" -Type "FATAL"
    return
}

# === Check if QPRIME Client is already running ===
if (-not $blnLinked -and (QPRIMEClientExecuting)) {
    LogHandler -LogFile "" -MsgBoxText "QPRIME Client is already running." `
               -MsgBoxTitle "QPRIME Updater: Information" -IEStatusText "" -Type " FAIL"
    RunQuit 1
}

#---

# === Create the File System Object ===
try {
    $global:objFileSys = New-Object -ComObject "Scripting.FileSystemObject"
} catch {
    LogHandler -LogFile "" -MsgBoxText "Error creating the File System Object." -MsgBoxTitle "QPRIME Updater: Error" -IEStatusText "" -Type "FATAL"
    return
}

# === Check if TMP file already exists ===
$tmpFile = Join-Path $TEMPFILEPATH $TEMPFILENAME
if ($objFileSys.FileExists($tmpFile)) {
    try {
        $global:objFile = $objFileSys.GetFile($tmpFile)
        if ((Get-Date $objFile.DateCreated) -gt (Get-Date).AddMinutes(-5)) {
            LogHandler -LogFile "" -MsgBoxText "An update to QPRIME is in progress. Please wait`nfor this to complete before trying again." `
                       -MsgBoxTitle "QPRIME Updater: Information" -IEStatusText "" -Type " FAIL"
            RunQuit 1
        } else {
            $objFile.Delete()
        }
    } catch {
        LogHandler -LogFile "" -MsgBoxText "Temporary file cannot be accessed or deleted." -MsgBoxTitle "QPRIME Updater: Error" -IEStatusText "" -Type "FATAL"
        return
    }
    $global:objFile = $null
}

# === Check if log file path exists ===
if (-not $objFileSys.FolderExists($LOGFILEPATH)) {
    try {
        $objFileSys.CreateFolder($LOGFILEPATH) | Out-Null
    } catch {
        LogHandler -LogFile "" -MsgBoxText "Error creating log file path." -MsgBoxTitle "QPRIME Updater: Error" -IEStatusText "" -Type "FATAL"
        return
    }
}

# === Create the log file ===
$logFileFullPath = Join-Path $LOGFILEPATH $LOGFILENAME
try {
    $global:objLogFile = $objFileSys.OpenTextFile($logFileFullPath, $ForWriting, $true)
} catch {
    LogHandler -LogFile "" -MsgBoxText "Error creating the Log File." -MsgBoxTitle "QPRIME Updater: Error" -IEStatusText "" -Type "FATAL"
    return
}
LogHandler -LogFile "------------------------------------------------------------" -MsgBoxText "" -MsgBoxTitle "" -IEStatusText "" -Type " PASS"
LogHandler -LogFile ("LOG FILE: " + (Split-Path -Leaf $MyInvocation.MyCommand.Path)) -MsgBoxText "" -MsgBoxTitle "" -IEStatusText "" -Type " PASS"
LogHandler -LogFile "------------------------------------------------------------" -MsgBoxText "" -MsgBoxTitle "" -IEStatusText "" -Type " PASS"

if ($blnLinked) {
    LogHandler -LogFile "Script Started via an occurrence link." -MsgBoxText "" -MsgBoxTitle "" -IEStatusText "" -Type " PASS"
    LogHandler -LogFile "Argument: $strArgument" -MsgBoxText "" -MsgBoxTitle "" -IEStatusText "" -Type " PASS"
    LogHandler -LogFile "Checking for running QPRIME client." -MsgBoxText "" -MsgBoxTitle "" -IEStatusText "" -Type " PASS"

    if (QPRIMEClientExecuting) {
        LogHandler -LogFile "   QPRIME client running. Link will be executed immediately." -MsgBoxText "" -MsgBoxTitle "" -IEStatusText "" -Type " PASS"
        RunQuit 7
    } else {
        LogHandler -LogFile "   QPRIME client not running. Version verification will be performed." -MsgBoxText "" -MsgBoxTitle "" -IEStatusText "" -Type " PASS"
    }
} else {
    LogHandler -LogFile "Script Started directly. No argument passed." -MsgBoxText "" -MsgBoxTitle "" -IEStatusText "" -Type " PASS"
}

InitIEProgress -IEText "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Checking For Updates"

# === Clear state variables ===
$strInstalledVersion = ""; LogHandler -LogFile "strInstalledVersion initialised to zero length string." -Type " PASS"
$strLatestVersion = ""; LogHandler -LogFile "strLatestVersion initialised to zero length string." -Type " PASS"
$strInstallerName = ""; LogHandler -LogFile "strInstallerName initialised to zero length string." -Type " PASS"
$strStatus = ""; LogHandler -LogFile "strStatus initialised to zero length string." -Type " PASS"
$strServerINIFile = ""; LogHandler -LogFile "strServerINIFile initialised to zero length string." -Type " PASS"
$strInstallA = ""; LogHandler -LogFile "strInstallA initialised to zero length string." -Type " PASS"
$strInstallB = ""; LogHandler -LogFile "strInstallB initialised to zero length string." -Type " PASS"
$strInstCommand = ""; LogHandler -LogFile "strInstCommand initialised to zero length string." -Type " PASS"
$strConnString = ""; LogHandler -LogFile "strConnString initialised to zero length string." -Type " PASS"
$strPackageID1 = ""; LogHandler -LogFile "strPackageID1 initialised to zero length string." -Type " PASS"
$strProgramName1 = ""; LogHandler -LogFile "strProgramName1 initialised to zero length string." -Type " PASS"
$strPackageID2 = ""; LogHandler -LogFile "strPackageID2 initialised to zero length string." -Type " PASS"
$strProgramName2 = ""; LogHandler -LogFile "strProgramName2 initialised to zero length string." -Type " PASS"
$strDBServer = ""; LogHandler -LogFile "strDBServer initialised to zero length string." -Type " PASS"
$strDBName = ""; LogHandler -LogFile "strDBName initialised to zero length string." -Type " PASS"
$strDBTable = ""; LogHandler -LogFile "strDBTable initialised to zero length string." -Type " PASS"
$strUpdaterVersion = ""; LogHandler -LogFile "strUpdaterVersion initialised to zero length string." -Type " PASS"
$strInstalledEnvironment = ""; LogHandler -LogFile "strInstalledEnvironment initialised to zero length string." -Type " PASS"
$strLatestEnvironment = ""; LogHandler -LogFile "strLatestEnvironment initialised to zero length string." -Type " PASS"
$strInstalledTraining = ""; LogHandler -LogFile "strInstalledTraining initialised to zero length string." -Type " PASS"
$strLatestTraining = ""; LogHandler -LogFile "strLatestTraining initialised to zero length string." -Type " PASS"
$strInputTraining = ""; LogHandler -LogFile "strInputTraining initialised to zero length string." -Type " PASS"

# === Main Logic Placeholder ===
GetAddresses
GetInstalledVersion


if ($strUpdaterTraining.ToUpper() -eq "OXLEY") {
    LogHandler -LogFile "Updater configured for OXLEY Academy" -Type " PASS"
    GetInputTraining
    if ($intValidTraining -eq 1) {
        LogHandler -LogFile "No valid Training Code Received. Prompting user to re-enter code." -Type " PASS"
        GetInputTraining
    }
    GetLatestVersion
    $strInstCommand = $strInstallA + $strInputTraining.ToUpper() + $strInstallB
    LogHandler -LogFile "   Install Command A, User Input & Install Command B Joined: " -Type " PASS"
    LogHandler -LogFile "       $strInstCommand" -Type " PASS"
} else {
    LogHandler -LogFile "Updater is not configured for OXLEY Academy." -Type " PASS"
    GetLatestVersion
}

# === Version Comparison ===
#[CHECK VERSION AND ENVIRONMENT]

LogHandler -LogFile "Checking installed version: $strInstalledVersion against latest version: $strLatestVersion" -Type " PASS"
LogHandler -LogFile "Checking installed environment: $strInstalledEnvironment against latest environment: $strLatestEnvironment" -Type " PASS"

if ($strUpdaterTraining.ToUpper() -eq "OXLEY") {
    LogHandler -LogFile "Checking installed training environment: $strInstalledTraining against user selected training environment: $($strInputTraining.ToUpper())" -Type " PASS"
    if (($strInstalledVersion.ToUpper() -eq $strLatestVersion.ToUpper()) -and
        ($strInstalledEnvironment.ToUpper() -eq $strLatestEnvironment.ToUpper()) -and
        ($strInstalledTraining.ToUpper() -eq $strInputTraining.ToUpper())) {
        LogHandler -LogFile "Latest version already installed with the latest environment settings. NO update required." -Type " PASS"
        RunQuit ($(if ($blnLinked) { 7 } else { 5 }))
    } else {
        LogHandler -LogFile "Latest version is not installed with the latest environment settings. Update required." -IEStatusText "Preparing to Update" -Type " PASS"
    }
} else {
    LogHandler -LogFile "Checking installed training environment: $strInstalledTraining against latest training environment: $strLatestTraining" -Type " PASS"
    if (($strInstalledVersion.ToUpper() -eq $strLatestVersion.ToUpper()) -and
        ($strInstalledEnvironment.ToUpper() -eq $strLatestEnvironment.ToUpper()) -and
        ($strInstalledTraining.ToUpper() -eq $strLatestTraining.ToUpper())) {
        LogHandler -LogFile "Latest version already installed with the latest environment settings. NO update required." -Type " PASS"
        #[]-->
        #[RUN QPRIME CLIENT]
        RunQuit ($(if ($blnLinked) { 7 } else { 5 }))
    } else {
        LogHandler -LogFile "Latest version is not installed with the latest environment settings. Update required." -IEStatusText "Preparing to Update" -Type " PASS"
    }
}

#--- 
#[HERE]

# === Check for latest QPRIME installer ===
LogHandler -LogFile "Checking for latest QPRIME installer" -Type " PASS"
$installerPath = Join-Path $strClientFolder $strInstallerName

if ($objFileSys.FileExists($installerPath)) {
    LogHandler -LogFile "Latest QPRIME installer found locally in $strClientFolder" -Type " PASS"

    # === Write values into temp file ===
    LogHandler -LogFile "Preparing to write variables to temporary file." -Type " PASS"
    try {
        $global:objFile = $objFileSys.OpenTextFile((Join-Path $TEMPFILEPATH $TEMPFILENAME), $ForWriting, $true)
    } catch {
        LogHandler -LogFile "   Error opening/creating temporary file for writing." -MsgBoxText "Unable to create temporary file." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
        return
    }

    $now = Get-Date
    $objFile.WriteLine("--- TEMP FILE: $TEMPFILEPATH$TEMPFILENAME. DATE/TIME: $now ---")
    $objFile.WriteLine("------------------------------------------------------------")
    $objFile.WriteLine("Software=$SOFTWARE")
    LogHandler -LogFile "      SOFTWARE ($SOFTWARE) written to temporary file." -Type " PASS"
    $objFile.WriteLine("Client Folder=$strClientFolder")
    LogHandler -LogFile "      strClientFolder ($strClientFolder) written to temporary file." -Type " PASS"
    $objFile.WriteLine("Install Command=$strInstCommand")
    LogHandler -LogFile "      strInstCommand ($strInstCommand) written to temporary file." -Type " PASS"
    $objFile.WriteLine("------------------------------------------------------------")
    $objFile.Close()
    LogHandler -LogFile "      Temporary file closed." -Type " PASS"
    $global:objFile = $null

    # === Re-run SMS/SCCM advertisement ===
    LogHandler -LogFile "Preparing to re-run SMS/SCCM program." -Type " PASS"
    try {
        $global:objUIResource = New-Object -ComObject "UIResource.UIResourceMgr"
    } catch {
        LogHandler -LogFile "   Could not create SMS/SCCM Client Resource Object." -MsgBoxText "An error has been encountered with SMS/SCCM.`nThe SMS/SCCM Client may not be working on your machine." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
        return
    }
    LogHandler -LogFile "   Resource Object created." -Type " PASS"
    RunSMSAdvert

    # === Wait for status change ===
    $intWait = 0
    $blnProgramComplete = $false
    while (-not $blnProgramComplete) {
        Start-Sleep -Seconds 1
        $intWait++
        LogHandler -IEStatusText ("Updating QPRIME Client" + ("." * ($intWait % 7))) -Type " PASS"

        if ($intWait -eq 300) {
            LogHandler -LogFile "         Timed out waiting for status return from the upgrade script." -MsgBoxText "Timed out waiting for upgrade to complete." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
            return
        } elseif ($intWait % 10 -eq 0) {
            LogHandler -LogFile "         Checking installation status." -Type " PASS"
            try {
                #GS
                #$global:objTempProgram = $objUIResource.GetProgram($objProgram.Id, $objProgram.PackageId)

                $global:objTempProgram = $objUIResource.GetProgram( $global:strProgName, $global:strPackageId )


            } catch {
                LogHandler -LogFile "            Could not get a refreshed instance of program." -MsgBoxText "An error has been encountered with SMS/SCCM.`nThe SMS/SCCM Client may not be working on your machine." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
                return
            }
            LogHandler -LogFile "            Got a refreshed instance of program." -Type " PASS"

            if ($objTempProgram.LastRunTime -ne $global:strLastRunTime) {
                $intReturn = $objTempProgram.LastExitCode
                LogHandler -LogFile "            Last run time has changed to: $($objTempProgram.LastRunTime). Program has completed." -Type " PASS"
                LogHandler -LogFile "            Last Run exit code: $intReturn" -Type " PASS"
                $blnProgramComplete = $true
            } else {
                LogHandler -LogFile "            Program still executing." -Type " PASS"
            }
            $global:objTempProgram = $null
        }
    }


    # === Read response from TEMP file ===
    LogHandler -LogFile "Reading response from upgrade script from temporary file." -Type " PASS"
    if (-not $objFileSys.FileExists($tmpFile)) {
        LogHandler -LogFile "   Error, temporary file deleted by another program." -MsgBoxText "Temporary file incorrectly removed." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
        return
    }
    LogHandler -LogFile "   Temporary file found." -Type " PASS"

    try {
        $global:objFile = $objFileSys.OpenTextFile($tmpFile, $ForReading)
    } catch {
        LogHandler -LogFile "   Error opening temporary file for reading." -MsgBoxText "An error has been encountered while retrieving installation status." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
        return
    }


    LogHandler -LogFile "   Temporary file opened for reading." -Type " PASS"


    #[QPRIME CLIENT UPDATED - SUCCEED]


    while (-not $objFile.AtEndOfStream) {
        $strLine = $objFile.ReadLine()
        LogHandler -LogFile "      Temporary file contains: $strLine" -Type " PASS"

        #gs by-pass 12/5/25
        #$objFile.Close(); $global:objFile = $null
        #    LogHandler -LogFile "      Temporary file Closed." -Type " PASS"
        #    LogHandler -LogFile "   Upgrade completed successfully. SMS/SCCM Returned: $intReturn" -Type " PASS"
        #    RunQuit(6)
        #    return

        if ($strLine.ToUpper() -eq "SUCCESS") {
            $objFile.Close(); $global:objFile = $null
            LogHandler -LogFile "      Temporary file Closed." -Type " PASS"
            LogHandler -LogFile "   Upgrade completed successfully. SMS/SCCM Returned: $intReturn" -Type " PASS"
            
            RunQuit ($(if ($blnLinked) { 8 } else { 6 }))
            return

        } elseif ($strLine.ToUpper().StartsWith("ERROR")) {
            $objFile.Close(); $global:objFile = $null
            LogHandler -LogFile "      Temporary file Closed." -Type " PASS"
            LogHandler -LogFile "   Upgrade failed to complete due to errors. QPRIME Upgrade Returned: $($strLine.Substring(6)) SMS/SCCM Returned: $intReturn" `
                       -MsgBoxText "Upgrade failed to complete due to errors.`nQPRIME Upgrade Returned: $($strLine.Substring(6))" -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
            return
        } else {
            $objFile.Close(); $global:objFile = $null
            LogHandler -LogFile "      Temporary file Closed." -Type " PASS"
            LogHandler -LogFile "   Temporary file did not contain a response. SMS/SCCM Returned: $intReturn" `
                       -MsgBoxText "Upgrade failed to complete.`nQPRIME Upgrade: No Response`nSMS/SCCM Returned: $intReturn" -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
            return
        }
    }
} else {
    LogHandler -LogFile "QPRIME installer: $strInstallerName not found locally in $strClientFolder" -Type " FAIL"
    RefreshSMS

   
    LogHandler -LogFile "QPRIME installer: $strInstallerName not found locally in $strClientFolder. SMS/SCCM policy refreshed." `
               -MsgBoxText "A new version of QPRIME has been released however the installer`nhas not been downloaded to your computer yet, or the download has failed.`nPlease try again later.`n`nYour computer may also not be setup to receive QPRIME Updates." -MsgBoxTitle "QPRIME Updater: Error" -Type "FATAL"
}

#---

# === Check if QPRIME Client is running ===
# === Get Server INI File Location from QPRIMELaunch.ini ===

















