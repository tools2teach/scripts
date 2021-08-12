<#
.SYNOPSIS
    List all workstations in the domain.  Fields include LastLogonDate and the latest BitLocker password set date (if present)
.DESCRIPTION
    List all workstations in the domain.  Fields include LastLogonDate and the latest BitLocker password set date (if present)
.PARAMETER SearchBase
    OU where the script will begin it's search
.INPUTS
    None
.OUTPUTS
    CSV in script path
.EXAMPLE
    .\New-BitLockerReport.ps1
.NOTES
    Author:             Martin Pugh
    Date:               4/9/2015
      
    Changelog:
        4/9             MLP - Initial Release
        4/15            MLP - Added code to load ActiveDirectory tools, or error out if they aren't present
#>

#Get today's date and cleanup date
$date = get-date -Uformat %Y%m%d
$Att1="WorkstationsWithBitLocker.csv"

[CmdletBinding()]
Param (
    [string]$SearchBase = "OU=Computer Accounts,DC=ais,DC=domain"
)

Try { Import-Module ActiveDirectory -ErrorAction Stop }
Catch { Write-Warning "Unable to load Active Directory module because $($Error[0])"; Exit }


Write-Verbose "Getting Workstations..." -Verbose
$Computers = Get-ADComputer -Filter * -SearchBase $SearchBase -Properties LastLogonDate
$Count = 1
$Results = ForEach ($Computer in $Computers)
{
    Write-Progress -Id 0 -Activity "Searching Computers for BitLocker" -Status "$Count of $($Computers.Count)" -PercentComplete (($Count / $Computers.Count) * 100)
    New-Object PSObject -Property @{
        ComputerName = $Computer.Name
        LastLogonDate = $Computer.LastLogonDate 
        BitLockerPasswordSet = Get-ADObject -Filter "objectClass -eq 'msFVE-RecoveryInformation'" -SearchBase $Computer.distinguishedName -Properties msFVE-RecoveryPassword,whenCreated | Sort whenCreated -Descending | Select -First 1 | Select -ExpandProperty whenCreated
    }
    $Count ++
}
Write-Progress -Id 0 -Activity " " -Status " " -Completed

$ReportPath = Join-Path (Split-Path $MyInvocation.MyCommand.Path) -ChildPath "WorkstationsWithBitLocker.csv"
Write-Verbose "Building the report..." -Verbose
$Results | Select ComputerName,LastLogonDate,BitLockerPasswordSet | Sort ComputerName | Export-Csv $ReportPath -NoTypeInformation
Write-Verbose "Report saved at: $ReportPath" -Verbose

Send-MailMessage -To "lnguyen@aiscaregroup.com" -From "aisint@aiscaregroup.com" -Subject "WorkstationsWithBitLocker" -Attachments $Att1 -SmtpServer 10.10.10.7