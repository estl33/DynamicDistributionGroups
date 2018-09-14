[CmdletBinding()]
Param()

Try{
$numUsers = 0
[Array] $divisionCol = @()

Get-ADUser -filter * -properties Division | foreach {
#Set-ADUser -identity $_-replace @{extensionAttribute14=$_.Division}

$dObj = New-Object PSObject -Property @{
Name               = $_.Name
Division           = $_.Division
}

$numUsers ++
$divisionCol += $dObj
}

Write-Host "Copied Division to Attribute14 for total of ${numUsers} users"

$divisionCol | Select Name, Division | Sort Name | Export-Csv "\\location\fileName\ADDivisionToAttribute12\ad_division_export$(Get-Date -f MMddyy_HHmm).csv" -NoTypeInformation

Write-Host "Exported name and divisions to csv" -ForegroundColor Green

} Catch {
Write-Warning "Error retrieving or copying division to Attribute12 in Active Directory"
}