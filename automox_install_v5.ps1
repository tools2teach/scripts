<#
.SYNOPSIS
    This Test script checks to see if an application is installed
.DESCRIPTION
    This script queries the installed files for 64bit software
    and returns a 0 if product is installed or 1 if not
.Notes
    File Name       :App_Install_Test.ps1
    Author          :Automox
    Prerequisite    :PowerShell V2 over Win7 and later
#>
#Handle Exit Codes:
trap { $host.ui.WriteErrorLine($_.Exception); exit 90 }

function App_Install_Test {

	<#
    .SYNOPSIS
        This function Checks to see if app is installed on system or not .
    .DESCRIPTION
        After checking the app availability, based on the exit code tool will decide to call remediation code or not.
    #>

  ## Name of the desired application ##
  $appName = 'Notepad'
  ## You must also hard-code the app name in the $scriptblock below 

	# Finding the Systen Directory (system32 Directory).
	$sysDir = [Environment]::SystemDirectory

	# Making the 64 bit path name blank.
	$64BIT = ""

	# Script block to execute with powershell
	# Match name in the block should be hard coded.
	$scriptBlock = {$key = Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' |
					Where-Object { $_.DisplayName -match 'TigerConnect'}
					return $key
					}

	# Get the Registry value for the 64 bit software installed on the 64 bit machine. as the automox process is 32 bit
	if ((Get-WmiObject Win32_OperatingSystem).OSArchitecture -eq "64-bit")
	{
		# Call for accessing the 64 bit registry in case the 32 bit process cannot access the registry.
		Try
		{
			$installed64 = @(& "$env:SystemRoot\sysnative\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy Bypass -WindowStyle Hidden -NoProfile -NonInteractive -Command $scriptBlock)
			$64BIT = $installed64.DisplayName
		}
		Catch
		{
			$Exception = $error[0].Exception.Message + "`nAt Line " + $error[0].InvocationInfo.ScriptLineNumber;
			Write-Output $Exception
			exit 90
		}

	}
	# Check for the Availability of the software and exit with relevent exit code.
	if ($64BIT -ne $null -and $64BIT.Trim() -match $appName) {
		#Application Found, Automox can handle the updates!
		exit 0
	} else {
		# Application is not installed! Run Remediation Script to install it.
		exit 1
	}
}

App_Install_Test