<#
.SYNOPSIS
   Script to check for empty groups in Entra ID/Azure AD.

.NOTES
   Created by Kiki Biancatti for SSW.
   This script should be run once per week.
#>

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Let's time this!
$Script:Stopwatch = [system.diagnostics.stopwatch]::StartNew()

try {
    # Importing the configuration file
    $config = Import-PowerShellDataFile $PSScriptRoot\Config.PSD1
}
catch {
    $RecentError = $Error[0]
    Write-Log -File $LogFile -Message "ERROR: Failed to import configuration file. $RecentError"
    exit
}

# Building variables
$LogFile = $config.LogFile
$LogModuleLocation = $config.LogModuleLocation
$TargetEmail = $config.TargetEmail
$OriginEmail = $config.OriginEmail
$TenantName = $config.TenantName
$ApplicationId = $config.ApplicationId
$Thumbprint = $config.Thumbprint
$Port  = $config.Port
$SMTPServer= $config.SMTPServer

try {
    # Importing the SSW Write-Log module
    Import-Module -Name $LogModuleLocation
}
catch {
    $RecentError = $Error[0]
    Write-Log -File $LogFile -Message "ERROR: Failed to import SSW Write-Log module. $RecentError"
    exit
}

try {
    # Install Send-MailKitMessage, because SendMailMessage is obsolete
    Import-Module -Name "Send-MailKitMessage"
}
catch {
    $RecentError = $Error[0]
    Write-Log -File $LogFile -Message "ERROR: Failed to install Send-MailKitMessage module. $RecentError"
    exit
}

try {
    # Import the Exchange module
    Import-Module ExchangeOnlineManagement
}
catch {
    $RecentError = $Error[0]
    Write-Log -File $LogFile -Message "ERROR: Failed to import ExchangeOnlineManagement module. $RecentError"
    exit
}

try {
    # Connect to Exchange Online via App Registration
    Connect-ExchangeOnline -CertificateThumbPrint $Thumbprint -AppID $ApplicationId -Organization $TenantName
}
catch {
    $RecentError = $Error[0]
    Write-Log -File $LogFile -Message "ERROR: Failed to connect to Exchange Online. $RecentError"
    exit
}

$allDistributionGroups = @()
$allDistributionGroups = Get-DistributionGroup -ResultSize Unlimited

# Create an array to hold empty groups
$emptyGroups = @()

# Loop through each Distribution group
foreach ($group in $allDistributionGroups) {
    # Get members of the group
    $groupMembers = Get-DistributionGroupMember -Identity $group.GUID

    # If the group has no members, add it to the emptyGroups array
    if ($groupMembers.Count -eq 0) {
        $emptyGroups += $group
    }
}

# List of GUIDs to exclude
$excludedGUIDs = @(
    "6037865c-f250-48c2-ad84-60bb7a8f81f6", # All Company Yammer group, default
    "2056fda0-7d3e-419e-961e-a9e4255e219d" # Viva Engage group, default
)

# Retrieve all Microsoft 365 groups
$allUnifiedGroups = Get-UnifiedGroup -ResultSize Unlimited

# Loop through each Microsoft 365 group
foreach ($group in $allUnifiedGroups) {
    # Skip groups with excluded GUIDs
    if ($excludedGUIDs -contains $group.GUID) {
        continue
    }
    
    # Get members of the group
    $groupMembers = Get-UnifiedGroupLinks -Identity $group.GUID -LinkType Members

    # If the group has no members, add it to the emptyGroups array
    if ($groupMembers.Count -eq 0) {
        $emptyGroups += $group
    }
}

# Count the empty groups
$emptyGroupCount = $emptyGroups.Count

# Output the number of empty groups
Write-Output "Number of Empty Distribution Groups: $emptyGroupCount"

# Output the list of empty groups
$emptyGroups | Format-Table Name, PrimarySmtpAddress, OrganizationalUnit

# Create a HTML table for the list of empty groups excluding OrganizationalUnit
$emptyGroupsTable = $emptyGroups | ConvertTo-Html -Property Name, PrimarySmtpAddress | Out-String

# Get the computer name
$computerName = $env:COMPUTERNAME

# Calculate elapsed time
$elapsedTime = $Script:Stopwatch.Elapsed.ToString("hh\:mm\:ss")

$emailBody = @"
<html>
<head>
    <style>
        body {
            font-family: Arial, sans-serif;
        }
        table {
            border-collapse: collapse;
            width: 100%;
        }
        th, td {
            border: 1px solid #dddddd;
            padding: 8px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
        }
    </style>
</head>
<body>
    <p>Hey SysAdmins,</p>
    
    <p>There's some empty groups in Azure AD. Do the following:</p>
    <ol>
        <li>Log in to the Azure AD portal.</li>
        <li>Navigate to 'Groups'.</li>
        <li>Search for the listed empty groups.</li>
        <li>For each group:
            <ol>
                <li>Check if there are any members that should be added.</li>
                <li>If the group is no longer needed, consider deleting it to maintain a clean directory.</li>
            </ol>
        </li>
    </ol>

    <p>Number of Empty Distribution Groups: $emptyGroupCount</p>
    $emptyGroupsTable

    <p>--Powered by PowerShell script on $computerName, scheduled task: $($MyInvocation.MyCommand.Name). Time taken: $elapsedTime</p>
</body>
</html>
"@

$emailParams = @{
    "RecipientList" = $TargetEmail
    "From"          = $OriginEmail
    "Subject"       = "Azure AD - Empty Groups Found"
    "SMTPServer"    = $SMTPServer
    "HTMLBody"      = $emailBody
    "Port"          = $Port
}

# Only send the email if there are empty groups
if ($emptyGroupCount -gt 0) {
    try {
        # Send the email
        Send-MailKitMessage @emailParams
    }
    catch {
        $RecentError = $Error[0]
        Write-Log -File $LogFile -Message "ERROR: Failed to send email. $RecentError"
    }
}

try {
    # Disconnect from Exchange Online
    Disconnect-ExchangeOnline -Confirm:$false
}
catch {
    $RecentError = $Error[0]
    Write-Log -File $LogFile -Message "ERROR: Failed to disconnect from Exchange Online. $RecentError"
}

# Stop the stopwatch
$Script:Stopwatch.Stop()