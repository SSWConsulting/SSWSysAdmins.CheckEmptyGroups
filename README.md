# Overview 
The primary objective of this script is to identify distribution groups in Azure AD that have no members. It's part of routine maintenance to ensure the Azure AD directory remains clean and free of clutter. This script is scheduled to run once per week.

    PowerShell Modules:
        ExchangeOnlineManagement
        Send-MailKitMessage (This module is preferred as SendMailMessage is obsolete)
        SSW Write-Log module

    Configuration File (Config.PSD1):
    This should be placed in the same directory as the script. It should include:
        LogFile - Path to the log file.
        LogModuleLocation - Path to the SSW Write-Log module.
        TargetEmail - Email address where notifications should be sent.
        OriginEmail - Email address that should appear in the "From" field of the notification.
        TenantName - Azure AD tenant name.
        ApplicationId - ID of the Azure application.
        Thumbprint - Certificate thumbprint for connecting to Exchange Online.

# Usage

    The script is a scheduled task that runs once a week.
    If any empty groups are detected, a notification will be sent to the provided email address.
