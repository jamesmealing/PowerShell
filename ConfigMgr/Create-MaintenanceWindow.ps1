param
(
    [Parameter(Mandatory = $true)]
    [string[]]$CollectionID,

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
    [switch]$MailSSL
)

#Import ConfigMgr helper Cmdlets
Import-Module "$env:ProgramFiles\ConfigMgr\Scripts\ConfigMgrModules.psm1"

#Import ConfigMgr Cmdlets
Import-Module $Env:SMS_ADMIN_UI_PATH.Replace("\bin\i386", "\bin\configurationmanager.psd1")
#endregion

#Connect to ConfigMgr Site Code
$SiteCode = Get-PSDrive -PSProvider CMSITE
Set-Location "$($SiteCode.Name):\"

#If mail SSL specified, get secure string password value from local text file and create credential object from it
if($MailSSL) {
    $MailPassword = Get-Content "$env:ProgramFiles\Inframon\Scripts\Config.txt" | ConvertTo-SecureString
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

#Loop through each collection specified in input parameters, and remove all previous Maintenance Windows
foreach ($Collection in $CollectionID) {
    Remove-MaintnanceWindows -CollectionID $Collection       
}

#Get Patch Tuesday date for the current month
$PatchTuesday = Get-PatchTuesday

#Set the default value of start date, if not otherwise specified in parameters
if (!$StartDate) {
    $StartDate = $PatchTuesday.AddDays($OffSetDays + $OffSetWeeks * 7).ToShortDateString()
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
        #Set new Maintenance Window
        Set-MaintenanceWindow -CollectionID $Collection -MaintenanceWindowStart $MaintenanceWindowStart -MaintenanceWindowEnd $MaintenanceWindowEnd -MaintenanceWindowName $MaintenanceWindowName       
    
        #Get Collection name for email notification
        $CollectionName = (Get-CMDeviceCollection -CollectionId $Collection).Name

        #Add collection name to Mail Body parameter
        $MailBodySuccess += "<i>$CollectionName</i></br>"
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
        Send-MailMessage @MailParams -Cc $MailToFailure -Body $MailBodyFailure -Priority "High" -Credential $MailCredential -UseSsl
    }
    else {
        #Use anonymous authentication
        Send-MailMessage @MailParams -Cc $MailToFailure -Body $MailBodyFailure -Priority "High"
    }
}
