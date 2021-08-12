<#
.SYNOPSIS
    Force M365 Apps to Update - Remediation Script
    OS Support: Windows 8.1 and above
    Powershell: 3.0 and above
    Run Type: Evaluation or OnDemand
.DESCRIPTION
    This Worklet is designed to grant an Admin the ability to force an existing installation of Microsoft 365 Apps (Office)
    to update to the latest build of its channel, either defined by existing policy\update keys, or by using the fallback
    $channel variable. The installation will run in USER context via a scheduled task. This Remediation script will detect
    orphaned tasks that have completed and remove them as needed.

    Office Deployment Tool: You will need to upload the most recent Office Deployment Tool (Setup.exe) when creating this
    Worklet policy. The ODT can be found at https://aka.ms/ODT

    NOTE: While this Worklet has the capability to change the channel of an existing installation, it should not be used as
    a replacement to a functioning channel management policy.

    Usage:
    This remediation script uses several variables to define the configuration and behavior of the M365 Apps Upgrade as well
    as allow the admin to define a fallback channel that can be forced regardless of the existing channel installed on a the
    device.

    $channel: This is the channel to be used as a fallback when GPO/Policy keys, and manual UpdateURL keys, are not present
    on a device. The following channel names will be accepted:

        Current: Provides users with new Office features as soon as they are ready, but on no set schedule.
        FirstReleaseCurrent: Preview release of Current channel.
        MonthlyEnterprise: Provides users with new Office features only once a month and on a predictable schedule.
        Deferred: For select devices in your organization, where extensive testing is needed before deploying new Office features.
        FirstReleaseDeferred: Preview release of Deferred channel.
        InsiderFast: Beta release channel. Frequent updates. Stable builds are sent to FirstReleaseCurrent.

    $excludeApps: This defines an application, or list of applications to be excluded from the upgrade installation. Each
    app should be enclosed in single quotes and separated by a comma. For a full list of App ID's that are supported visit
    https://docs.microsoft.com/en-us/deployoffice/office-deployment-tool-configuration-options

    $visibility: This defines the visibility of the M365 Apps installer.

        FULL: This is the default setting and will display the progress bar to the user. If $forceAppshutdown is set to False, the
        user will be presented with a dialog informing them to close the affected apps and click "continue"

        NONE: This will make the installation completely silent to the user. It is recommended to set $forceAppshutdown to True
        as the installer will wait for the applications to close without presenting a dialog to the user.

    $forceAppshutdown: This defines wether or not open M365 Apps should be shutdown automatically. When using False, the
    worklet will copy the ODT and xml file to %windir%\temp to allow the worklet to exit while waiting for the user. Upon re-evaluation
    or re-running the policy manually, it will then determine the state of the installation and cleanup any remaining scheduled tasks.

        True: This will force any M365 Apps to be closed automatically as the installer finds them. While M365 does have a robust
        data recovery option to restore unsaved work, there could be edge case scenarios where data loss may still occur.

        False: This will prompt the user instructing them to close an affected app that is preventing the installation from continueing.
        Clicking "Continue" will automatically close any affected apps. It is recommended to only set this setting when $visibility
        is also set to full. If the installation is completely silent, the user will not be prompted to close applications and the
        installer will continue to wait indefinitely.

    $forcefallback: Determines if the $channel should be enforced by default or rely on versioning information.
        True: By setting this to true, all versioning checks will be ignored and the installation will only verify if the device
        is currently on the channel defined by $channel.

        False: This is the default setting and will download the latest version from the associated channels CDN to verify against
        the installed version. If no policy keys are defined on the device, $channel will be used instead.
.EXAMPLE
    $channel = 'Current'
    $excludeApps = 'Groove','Lync','Teams'
    $visibility = 'FULL'
    $forceAppshutdown = 'False'
    $forcefallback = 'False'
.LINK
    https://www.microsoft.com/en-us/download/details.aspx?id=49117
.NOTES
    Author: eliles
    Date: May 21 2021
#>
######## Make changes within this block ########
$channel = 'Current'
$excludeApps = 'Groove','Lync'
$visibility = 'Silent'
$forceAppshutdown = 'False'
$forcefallback = 'False'
################################################

# Predefined Variables that could be modified if needed
$prodID = 'O365ProPlusRetail' #Office ProductID to install
$runAsaccount = 'USERS' # Account used to run scheduled task USERS or SYSTEM preferred
$fallbackLang = 'EN-US' # Fallback Language to use if "MatchInstalled" is not available on CDN

# Scheduled task check/cleanup
$taskcheck = (Get-ScheduledTask -TaskName 'Force M365 Apps Update' -ErrorAction SilentlyContinue).state
if($taskcheck)
{
    if($taskcheck -eq 'Ready')
    {
       Write-Output "Found Completed Scheduled Task - Deleting Task."
       Unregister-ScheduledTask 'Force M365 Apps Update' -Confirm:$false
    }
    else
    {
        Write-Output "Scheduled Task is running - Exiting Worklet"
        Exit 0
    }
}

# Define Script directory
$scriptDir = Split-Path $Script:MyInvocation.MyCommand.Path -Parent

# Define Channel URLs
$channels = @{
    Current = 'http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60'
    FirstReleaseCurrent = 'http://officecdn.microsoft.com/pr/64256afe-f5d9-4f86-8936-8840a6a4f5be'
    MonthlyEnterprise = 'http://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6'
    Deferred = 'http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114'
    FirstReleaseDeferred = 'http://officecdn.microsoft.com/pr/b8f9b850-328d-4355-9145-c59439a0c4cf'
    InsiderFast = 'http://officecdn.microsoft.com/pr/5440fd1f-7ecb-4221-8110-145efaa6372f'
}

# Capturing Office config
if([System.Environment]::Is64BitOperatingSystem)
{
    # Opening 64bit config
    $64Hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
    $64Config = $64Hklm.OpenSubKey("SOFTWARE\Microsoft\Office\ClickToRun\Configuration")

    # Config fail-safe
    if(!($64Config))
    {
        $64Hklm.close()
        Write-Output "Not Applicable - Microsoft 365 not installed"
        exit 0
    }

    # Capture config data
    $curVer = $64Config.getvalue("VersionToReport")
    $curPlat = $64Config.getvalue("Platform")
    $curUpdurl = $64Config.getvalue("UpdateURL")
    $curUpdchannel = $64Config.GetValue("UpdateChannel")
    $64Hklm.close()
}
else
{
    #x86 config path
    $86Config = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"

    # Config fail-safe
    if(!(Test-Path $86Config))
    {
        Write-Output "Not Applicable - Microsoft 365 not installed"
        exit 0
    }

    # Capture config data
    $curVer = (Get-ItemProperty -Path $86Config -Name "VersionToReport").VersionToReport
    $curPlat = (Get-ItemProperty -Path $86Config -Name "Platform").Platform
    $curUpdurl = (Get-ItemProperty -Path $86Config -Name "UpdateURL").UpdateURL
    $curUpdchannel = (Get-ItemProperty -Path $86Config -Name "UpdateChannel").UpdateChannel
}

# Check for UpdateBranch policy
$curPolicy = (Get-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\OfficeUpdate -Name UpdateBranch -ErrorAction SilentlyContinue).UpdateBranch

# Determine active channel and evaluate version if fallback not forced
if(!($forcefallback -eq "true"))
{
    if($curPolicy)
    {
        $sourceCDN = $channels[$curPolicy]
    }
    elseif($curUpdurl)
    {
        $sourceCDN = $curUpdurl
    }
    else
    {
        $sourceCDN = $channels[$channel]
    }

    # Download cab from CDN and extract VersionDescriptor.xml
    Invoke-WebRequest -Uri "$sourceCDN/Office/Data/v32.cab" -OutFile "$scriptDir\v32.cab"
    Start-Process -FilePath expand.exe -ArgumentList "$scriptDir\v32.cab -f:VersionDescriptor.xml $scriptDir" -Wait
    [XML]$verXML = Get-Content -Path "$scriptDir\VersionDescriptor.xml"
    $cdnVer64 = $verXML.version.Available.I640Version
    $cdnVer86 = $verXML.version.Available.I320Version

    # Preflight install check
    if($curPlat -like '*64')
    {
        if([version]$curVer -ge [version]$cdnVer64)
        {
            Write-Output "Compliant"
            Exit 0
        }
    }
    else
    {
        if([version]$curVer -ge [version]$cdnVer86)
        {
            Write-Output "Compliant"
            Exit 0
        }
    }
}

# Forced Fallback channel eval
if($curUpdchannel -eq $channels[$channel])
{
    Write-Output "Update Channel is already set to $channel - Exiting"
    Exit 0
}

# Standardize bitness format for XML
if($curPlat -like '*86')
{
    $bitness = '32'
}
else
{
    $bitness = '64'
}

# Define Channel for XML
if(!($forcefallback -eq 'True'))
{
    $convertChannel = $channels.GetEnumerator() | Where-Object {$_.Value -eq "$sourceCDN"}
    $xmlChannel = ($convertChannel).Name
}
else
{
    $xmlChannel = $channel
}

#Create XML file and Add root Configuration Element
[System.XML.XMLDocument]$ConfigFile = New-Object System.XML.XMLDocument
[System.XML.XMLElement]$ConfigurationRoot=$ConfigFile.CreateElement("Configuration")
$ConfigFile.appendChild($ConfigurationRoot) | Out-Null

# Add "Add" Element and set attributes
[System.XML.XMLElement]$addElement=$configFile.CreateElement("Add")
$configurationRoot.appendChild($addElement) | Out-Null
$addElement.SetAttribute("OfficeClientEdition","$bitness") | Out-Null
$addElement.SetAttribute("Channel","$xmlChannel") | Out-Null

#Add the Product Element under Add and set the ID
[System.XML.XMLElement]$productElement=$configFile.CreateElement("Product")
$addElement.appendChild($productElement) | Out-Null
$productElement.SetAttribute("ID","$prodID") | Out-Null

#Add the Language Element under Product and set the IDs
[System.XML.XMLElement]$languageElement=$configFile.CreateElement("Language")
$productElement.appendChild($languageElement) | Out-Null
$languageElement.SetAttribute("ID","MatchInstalled") | Out-Null
$languageElement.SetAttribute("Fallback","$fallbackLang") | Out-Null

#Add the ExcludeApp Element under Product and set the IDs
if($excludeApps)
{
    foreach($app in $excludeApps)
    {
        [System.XML.XMLElement]$exAppelement=$configFile.CreateElement("ExcludeApp")
        $productElement.appendChild($exAppelement) | Out-Null
        $exAppelement.SetAttribute("ID","$app") | Out-Null
    }
}

#Add the Updates Element under Configuration
[System.XML.XMLElement]$updElement=$configFile.CreateElement("Updates")
$configurationRoot.appendChild($updElement) | Out-Null
$updElement.SetAttribute("Enabled","TRUE") | Out-Null

#Add the Display Element under Configuration
[System.XML.XMLElement]$dispElement=$configFile.CreateElement("Display")
$configurationRoot.appendChild($dispElement) | Out-Null
$dispElement.SetAttribute("Level","$visibility") | Out-Null
$dispElement.SetAttribute("AcceptEULA","TRUE") | Out-Null

#Add the Property Element under Configuration
[System.XML.XMLElement]$propElement=$configFile.CreateElement("Property")
$configurationRoot.appendChild($propElement) | Out-Null
$propElement.SetAttribute("Name","FORCEAPPSHUTDOWN") | Out-Null
$propElement.SetAttribute("Value","$forceAppshutdown") | Out-Null

# Saves config to script directory
$configFile.Save("$scriptDir\update.xml") | Out-Null
$xmlpath = "$scriptDir\update.xml"

# Output Channel to activity log
if($curPolicy)
{
    Write-Output "Policy Found: $curPolicy"
}
elseif($curUpdurl)
{
    Write-Output "UpdateURL Key Found: $xmlChannel"
}
else
{
    Write-Output "Using Fallback Channel: $channel"
}

# Creates scheduled task to run in USER context and deletes task when done
$ShedService = New-Object -comobject 'Schedule.Service'
$ShedService.Connect()

$Task = $ShedService.NewTask(0)
$Task.RegistrationInfo.Description = "Forces Update of M365 Apps"
$Task.Settings.Enabled = $true
$Task.Settings.AllowDemandStart = $true
$task.Settings.DisallowStartIfOnBatteries = $false
$Task.Principal.RunLevel = 1

$trigger = $task.triggers.Create(7)
$trigger.Enabled = $true

$action = $Task.Actions.Create(0)
$action.Path = "$scriptDir\setup.exe"
$action.Arguments = "/configure $xmlpath"

$taskFolder = $ShedService.GetFolder("\")

# Copy ODT/Config to WINDIR\TEMP if appshutdown false
if($forceAppshutdown -eq 'False')
{
    Write-Output "ForceAppShutdown is set to FALSE. Copying Files to "$env:WINDIR\Temp" to run outside of Worklet."
    Copy-Item -Path "$scriptDir\setup.exe" -Destination "$env:WINDIR\temp" -Force
    Copy-Item -Path "$scriptDir\update.xml" -Destination "$env:WINDIR\temp" -Force
    $xmlpath = "$env:WINDIR\Temp\update.xml"
    $action.Path = "$env:WINDIR\Temp\setup.exe"
    $action.Arguments = "/configure $xmlpath"
    Write-Output "Initializing Update and exiting Worklet. Wait for Eval or run policy again for confirmation and cleanup"
    $taskFolder.RegisterTaskDefinition("Force M365 Apps Update", $Task , 6, "$runAsaccount", $null, 4) | Out-Null
    Exit 0
}

Write-Output "Initializing Update"
$taskFolder.RegisterTaskDefinition("Force M365 Apps Update", $Task , 6, "$runAsaccount", $null, 4) | Out-Null

# Check status until task has completed
DO
{
(Get-ScheduledTask -TaskName 'Force M365 Apps Update').State | Out-Null
}
Until ((Get-ScheduledTask -TaskName 'Force M365 Apps Update').State -eq "Ready")
Unregister-ScheduledTask 'Force M365 Apps Update' -Confirm:$false

# Capture Post install info
if([System.Environment]::Is64BitOperatingSystem)
{
    $64Hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
    $64Config = $64Hklm.OpenSubKey("SOFTWARE\Microsoft\Office\ClickToRun\Configuration")
    $postVer = $64Config.getvalue("VersionToReport")
    $postUpdchannel = $64Config.getvalue("UpdateChannel")
}
else
{
    $postVer = (Get-ItemProperty -Path $86Config -Name "VersionToReport").VersionToReport
    $postUpdchannel = (Get-ItemProperty -Path $86Config -Name "UpdateChannel").UpdateChannel
}

# Final version eval
if(!($forcefallback -eq 'True'))
{
    if($curPlat -like '*64')
    {
        if([version]$postVer -ge [version]$cdnVer64)
        {
            Write-Output "Installation Successful"
            Exit 0
        }
    }
    else
    {
        if([version]$postVer -ge [version]$cdnVer86)
        {
            Write-Output "Installation Successful"
            Exit 0
        }
    }
    Write-Output "ERROR: Installation Failed - See logs at $env:WINDIR\Temp for more information"
    Exit 1
}

# Final Fallback channel eval
if($postUpdchannel -eq $channels[$channel])
{
    Write-Output "Update Channel is now set to $channel - Exiting"
    Exit 0
}
Write-Output "ERROR: Installation Failed - See logs at $env:WINDIR\Temp for more information"
Exit 1