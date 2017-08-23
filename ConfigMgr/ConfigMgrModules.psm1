#region Get-PatchTuesday
#Get Patch Tuesday date
Function Get-PatchTuesday {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$Month = (Get-Date).Month,

        [Parameter(Mandatory = $false)]
        [string]$Year = (Get-Date).Year
    )

    $FirstDayOfMonth = [datetime]($Month + "/1/" + $Year)
    (0..30 | Where-Object {$FirstDayOfMonth.AddDays($_) } | Where-Object {$_.DayofWeek -like "Tue*"})[1]
}
#endregion

#region Set-MaintenanceWindow
#Create new Maintenance Window
Function Set-MaintenanceWindow {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$CollectionID,

        [Parameter(Mandatory = $true)]
        [datetime]$MaintenanceWindowStart,

        [Parameter(Mandatory = $true)]
        [datetime]$MaintenanceWindowEnd,

        [Parameter(Mandatory = $true)]
        [string]$MaintenanceWindowName
    )

    $Schedule = New-CMSchedule -Nonrecurring -Start $MaintenanceWindowStart -End $MaintenanceWindowEnd

    #Create Maintenance Window
    New-CMMaintenanceWindow -CollectionID $CollectionID -Schedule $Schedule -Name $MaintenanceWindowName -ApplyTo SoftwareUpdatesOnly
}
#endregion

#region Remove-MaintenanceWindows
#Remove all existing Maintenance Windows for a Collection
Function Remove-MaintnanceWindows {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$CollectionID
    )

    Get-CMMaintenanceWindow -CollectionId $CollectionID | ForEach-Object {
        Remove-CMMaintenanceWindow -CollectionID $CollectionID -Name $_.Name -Force
    }
}
#endregion
