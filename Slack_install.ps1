#REQUIRES -Version 2.0
<#
.SYNOPSIS
    This script allows an admin to install Slack
.DESCRIPTION
    This script downloads the system-wide-slack installer. Once installed, the program
    checks to see if slack is installed WHEN A USER LOGS IN. If slack is not installed
    the program installs it in the packground.
    If a user has uninstalled slack, this will not reinstall it.
.NOTES
    File Name       :Slack_Install.ps1
    Author          :Automox
    Prerequisite    :PowerShell V2 over win7 and upper
#>
#Handle Exit Codes:
trap {  $host.ui.WriteErrorLine($_.Exception); exit 90 }

function Slack_Install{

    <#
    .SYNOPSIS
        This function allows automox to Install the Slack on Windows .
    .DESCRIPTION
        This function download and installs the latest version of Slack on system.
    .EXAMPLE
        Slack_Install
    .NOTES
    #>

    #############Change the settings in this block#######################
    # Save the installer in the powershell current working directory.
    $saveFilePath = "$PSScriptRoot\SlackSetup.msi"
	
	# If any old installer file is existed, Delete the old installer.
    if (Test-Path -Path "$saveFilePath") {
		Remove-Item -force "$saveFilePath" -ErrorAction SilentlyContinue
    }

    # Download msi Installer based on the bit version of the Windows OS
    if ((Get-WmiObject Win32_OperatingSystem).OSArchitecture -eq "64-bit") {
        $URL = "http://slack.com/ssb/download-win64-msi"
    } else  {
        $URL = "http://slack.com/ssb/download-win-msi"
    }

    ###### Download the msi installer file for the installation ######
    $downloader = (new-object System.Net.WebClient)
    Try {
        $downloader.DownloadFile("$URL", "$saveFilePath")
        Write-Output "Slack Installer download finished..."
    }
    Catch { Write-Output "File download failed. Check URL [$URL] and try again."
            exit 90
    }

    ###### Installing  slack on system ######
    Try {
        $process=Start-Process -FilePath msiexec.exe -ArgumentList '/i',$saveFilePath,'/q' -Wait -PassThru -ErrorAction Stop
        Write-Output "Slack Installion Finished..."
    }
    Catch { $Exception = $error[0].Exception.Message + "`nAt Line " + $error[0].InvocationInfo.ScriptLineNumber
            Write-Output $Exception
            exit 90
    }

    # Avoid error while deleting the installer as it is still in the process of installation.
    do {
        if ((Get-Content $saveFilePath)) {
            Write-Output "Installer complete...cleaning up installation files"
            Remove-Item -force "$saveFilePath" -ErrorAction SilentlyContinue
            exit $process.ExitCode
        } else {
            Write-Output "Waiting for Installer File Unlock"
        }
        $fileunlocked = 1
    }	Until ($fileunlocked)

}

Slack_Install