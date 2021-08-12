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
    Author:             Lee Nguyen
    Date:               8/5/2022
      
   
#>


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

$date = get-date -Uformat %Y%m%d
$ReportPath = "I:\BitLocker Keys\BitLocker_Report\WorkstationsWithBitLocker_$date.csv"
Write-Verbose "Building the report..." -Verbose
$Results | Select ComputerName,LastLogonDate,BitLockerPasswordSet | Sort ComputerName | Export-Csv $ReportPath -NoTypeInformation
Write-Verbose "Report saved at: $ReportPath" -Verbose

Send-MailMessage -To "lnguyen@aiscaregroup.com","bknight@aiscaregroup.com","twallace@aiscaregroup.com","sjames@aiscaregroup.com" -From "aisint@aiscaregroup.com" -Subject "WorkstationsWithBitLocker" -Attachments $ReportPath -Body "Weekly report of all workstations with Bitlocker Keys backed up to AD" -SmtpServer aispaincare-com.mail.protection.outlook.com
