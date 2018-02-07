param (
    [Parameter(Mandatory = $true)]
    [string[]]$CollectionId,

    [Parameter(Mandatory = $true)]
    [string]$MaintenanceWindowName,

    [Parameter(Mandatory = $true)]
    [string]$StartTime,

    [Parameter(Mandatory = $true)]
    [string]$EndTime,

    [Parameter(Mandatory = $false)]
    [string]$StartDate,

    [Parameter(Mandatory = $false)]
    [string]$EndDate,

    [Parameter(Mandatory = $false)]
    [int]$OffSetDays,

    [Parameter(Mandatory = $false)]
    [int]$OffSetWeeks,

    [Parameter(Mandatory = $false)]
    [switch]$RemoveExisting,

    [Parameter(Mandatory = $true)]
    [string[]]$MailTo,

    [Parameter(Mandatory = $false)]
    [string[]]$MailToFailure,

    [Parameter(Mandatory = $true)]
    [string]$MailFrom,

    [Parameter(Mandatory = $false)]
    [string]$MailSubject = "ConfigMgr Maintenance Window Creation",

    [Parameter(Mandatory = $false)]
    [string]$MailSmtpServer = "smtp.office365.com",

    [Parameter(Mandatory = $false)]
    [int]$MailPort = 587,

    [Parameter(Mandatory = $false)]
    [switch]$MailSSL,

    [Parameter(Mandatory = $false)]
    [switch]$ScomMaintenanceMode,

    [Parameter(Mandatory = $false)]
    [string]$ScomServerFQDN
)

#Import ConfigMgr helper Cmdlets
Import-Module "$env:ProgramFiles\Inframon\Scripts\SCCM\ConfigMgrModules.psm1"

#Import ConfigMgr Cmdlets
Import-Module $env:SMS_ADMIN_UI_PATH.Replace("\bin\i386", "\bin\configurationmanager.psd1")
#endregion

#Connect to ConfigMgr Site Code
$SiteCode = Get-PSDrive -PSProvider CMSITE
Set-Location "$($SiteCode.Name):\"

#If mail SSL specified, get secure string password value from local text file and create credential object from it
if ($MailSSL) {
    $MailPassword = Get-Content "$env:ProgramFiles\Inframon\Scripts\SCCM\MailConfig.txt" | ConvertTo-SecureString
    $MailCredential = New-Object -TypeName System.Management.Automation.PSCredential($MailFrom, $MailPassword)
}

#Split Mail "To" parameters on comma seperated input to allow proper string wrapping when run via command line/scheduled task
$MailTo = $MailTo -split ","
$MailToFailure = $MailToFailure -split ","

#Mail message common paramaters
$MailParams = @{
    To         = $MailTo;
    From       = $MailFrom;
    Subject    = $MailSubject;
    SmtpServer = $MailSmtpServer;
    Port       = $MailPort;
    BodyAsHtml = $true
}

#Get Patch Tuesday date for the current month
$PatchTuesday = Get-PatchTuesday

#Set the default value of start date, if not otherwise specified in parameters
if (!$StartDate) {
    $StartDate = $PatchTuesday.AddDays($OffSetDays + ($OffSetWeeks * 7)).ToShortDateString()
}

#Set the default value of end date, if not otherwise specified in parameters
if (!$EndDate) {
    $EndDate = $StartDate
}

#Create Maintenance Window Start and End from Date and Time Strings
$MaintenanceWindowStart = [DateTime]::Parse("$StartDate $StartTime")
$MaintenanceWindowEnd = [DateTime]::Parse("$EndDate $EndTime")

#Get Maintenance Window duartion for email notification
$Duration = New-TimeSpan -Start $MaintenanceWindowStart -End $MaintenanceWindowEnd

#Create mail body variables, including start of message to be used in for-each loop to add each collection name
$MailBodySuccess = "
SUCCESS: Configuration Manager Maintenance Window `"$MaintenanceWindowName`", with a duration of `"$Duration`" hours, has been created for the following collections; 
</br>
</br>
"
$MailBodyFailure = "
FAILURE: Configuration Manager Maintenance Window `"$MaintenanceWindowName`" failed to create for the following collections;
</br>
</br>
"

try {
    #Loop through each collection specified in input parameters, and set a new Maintenance Window
    foreach ($Collection in $CollectionID) {
        if ($RemoveExisting) {
            #Remove all existing Maintenance Windows
            Remove-MaintnanceWindows -CollectionID $Collection       
        }

        #Set new Maintenance Window
        Set-MaintenanceWindow -CollectionID $Collection -MaintenanceWindowStart $MaintenanceWindowStart -MaintenanceWindowEnd $MaintenanceWindowEnd -MaintenanceWindowName $MaintenanceWindowName       
    
        #Get Collection name for email notification
        $CollectionName = (Get-CMDeviceCollection -CollectionId $Collection).Name

        #Add collection name to Mail Body parameter
        $MailBodySuccess += "<i>$CollectionName</i></br>"

        #If specified, create scheduled tasks to start SCOM Maintenance Mode
        if ($ScomMaintenanceMode) {
            #Create the here-string for the 'Start-ScomMaintenanceMode' script
            $StartScomMaintenanceMode = @"
###SCRIPT AUTOMATICALLY GENERATED FROM "New-MaintenanceWindow.ps1" SCRIPT###

#Get the list of devices in the collection
try {
    `$CollectionMembers = (Get-WmiObject -ComputerName "$env:COMPUTERNAME" -Namespace "ROOT\SMS\site_$SiteCode" -Query "SELECT * FROM SMS_FullCollectionMembership WHERE CollectionID='$($Collection)'").Name
}
catch {
    throw "Error getting the members from the collection $Collection"
}

#Create a PSRemoting session to the SCOM server
try {
    `$Session = New-PSSession -ComputerName "$ScomServerFQDN" -UseSSL -ErrorAction Stop
}
catch {
    throw "Unable to establish PSRemoting session to $ScomServerFQDN"
}

#Loop through each server in the collection and start SCOM Maintenance Mode for the duration of the ConfigMgr Maintenance Window
foreach (`$Server in `$CollectionMembers) {
    Invoke-Command -Session `$Session -ScriptBlock {
        Import-Module OperationsManager
        Start-SCOMMaintenanceMode -Instance (Get-SCOMClassInstance -Name "`$using:Server.*") -EndTime "$($MaintenanceWindowEnd.AddMinutes(10).ToString())" -Comment "Scheduled Maintenance Mode for Configuration Manager Maintenance Window" -Reason "PlannedOther"
    }    
}

#Close PSRemoting session
try {
    Get-PSSession | Remove-PSSession
}
catch {
    throw "Error removing PSSession"
}
"@

            $TaskName = "SCOM Maintenance Mode - $Collection - $MaintenanceWindowName"
            $TaskPath = "Inframon\Operations Manager\Maintenance Mode"
            $TaskUser = "SVC-CM-Automation"
            $SecurePassword = Get-Content "$env:ProgramFiles\Inframon\Scripts\SCCM\TaskConfig.txt" | ConvertTo-SecureString
            $ScriptName = Start-ScomMaintenanceMode-$Collection-$($MaintenanceWindowName.Replace(' ', '')).ps1
            $ScriptPath = "$env:ProgramFiles\Inframon\Scripts\SCOM\Maintenance Mode"

            #Due to the 'Register-ScheduledTask' cmdlet not accepting 'SecureString' as an input for password, we need to create a new credential object and use this to lookup the plaintext password.
            #This is still better than having the plaintext password directly in the script, however, this is obviously less than ideal but sadly we have little choice.
            $TaskCredentials = New-Object System.Management.Automation.PSCredential -ArgumentList $TaskUser, $SecurePassword
            $TaskPassword = $TaskCredentials.GetNetworkCredential().Password

            #Define the action, trigger and settings of the Scheduled Task
            $TaskAction = New-ScheduledTaskAction -Execute "$env:windir\System32\WindowsPowerShell\v1.0\PowerShell.exe" -Argument "-File `"$ScriptPath\$ScriptName"
            $TaskTrigger = New-ScheduledTaskTrigger -Once -At ([DateTime]::SpecifyKind("$MaintenanceWindowStart", [DateTimeKind]::Local).AddMinutes(-5))
            $TaskSettings = New-ScheduledTaskSettingsSet -Compatibility Win8 -ExecutionTimeLimit "01:00:00"

            #Create 'Start-ScomMaintenanceMode' PowerShell Script for the Scheduled Task from the here-string created above
            Set-Content -Path "$ScriptPath\$ScriptName" -Value $StartScomMaintenanceMode

            #Check for an existing Scheduled Task, if found, delete it
            Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue | Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false | Out-Null

            #Create and register the Scheduled Task using the action, trigger and settings defined above
            Register-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Action $TaskAction -Settings $TaskSettings -Trigger $TaskTrigger -User $TaskUser -Password $TaskPassword | Out-Null
        }
    }

    #Send completion email
    if ($MailSSL) {
        #Use secure SSL authentication
        Send-MailMessage @MailParams -Body $MailBodySuccess -Credential $MailCredential -UseSsl
    }
    else {
        #Use anonymous authentication
        Send-MailMessage @MailParams -Body $MailBodySuccess
    }
}
catch {
    #Send failure email
    if ($MailSSL) {
        #Use secure SSL authentication
        if ($MailToFailure) {
            #Include Cc to '$MailToFailure'
            Send-MailMessage @MailParams -Cc $MailToFailure -Body $MailBodyFailure -Priority "High" -Credential $MailCredential -UseSsl    
        }
        else {
            Send-MailMessage @MailParams -Body $MailBodyFailure -Priority "High" -Credential $MailCredential -UseSsl
        }
    }
    else {
        #Use anonymous authentication
        if ($MailToFailure) {
            #Include Cc to '$MailToFailure'
            Send-MailMessage @MailParams -Cc $MailToFailure -Body $MailBodyFailure -Priority "High"    
        }
        else {
            Send-MailMessage @MailParams -Body $MailBodyFailure -Priority "High"
        }
    }
}
