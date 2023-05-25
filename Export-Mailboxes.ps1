<#
Exports mailboxes and displays info in a CSV. Allows you to choose between shared mailboxes, user mailboxes, or all mailboxes. Displays the following info:
    UserPrincipalName
    Type
    IsLicensed
    Licenses
    StorageConsumed
    StorageLimit
    ArchiveStatus
    AutoExpandingArchiveEnabled
    ArchiveStorageConsumed
    ArchiveStorageQuota
    RetentionPolicy
    ForwardingSMTPAddress
    ForwardingAddress
#>

# functions
function Initialize-ColorScheme
{
    $script:successColor = "Green"
    $script:infoColor = "DarkCyan"
    $script:failColor = "Red"
    # warning color is yellow, but that is built into Write-Warning
}

function Show-Introduction
{
    Write-Host ("This script exports mailboxes and displays useful info in a CSV. `n" +
    "Allows you to choose between shared mailboxes, user mailboxes, or all mailboxes `n") -ForegroundColor $infoColor
    Read-Host "Press Enter to continue"
}

function Use-Module($moduleName)
{    
    $keepGoing = -not(Test-ModuleInstalled $moduleName)
    while ($keepGoing)
    {
        Prompt-InstallModule($moduleName)
        Test-SessionPrivileges
        Install-Module $moduleName

        if ((Test-ModuleInstalled $moduleName) -eq $true)
        {
            Write-Host "Importing module..." -ForegroundColor $infoColor
            Import-Module $moduleName
            $keepGoing = $false
        }
    }
}

function Test-ModuleInstalled($moduleName)
{    
    $module = Get-Module -Name $moduleName -ListAvailable
    return ($null -ne $module)
}

function Prompt-InstallModule($moduleName)
{
    do 
    {
        Write-Host "$moduleName module is required." -ForegroundColor $infoColor
        $confirmInstall = Read-Host -Prompt "Would you like to install the module? (y/n)"
    }
    while ($confirmInstall -inotmatch "^\s*y\s*$") # regex matches a y but allows spaces
}

function Test-SessionPrivileges
{
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentSessionIsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($currentSessionIsAdmin -ne $true)
    {
        Write-Host ("Please run script with admin privileges.`n" +
        "1. Open Powershell as admin.`n" +
        "2. CD into script directory.`n" +
        "3. Run .\scriptname`n") -ForegroundColor $failColor
        Read-Host "Press Enter to exit"
        exit
    }
}

function TryConnect-ExchangeOnline
{
    $connectionStatus = Get-ConnectionInformation -ErrorAction SilentlyContinue

    while ($null -eq $connectionStatus)
    {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor $infoColor
        Connect-ExchangeOnline -ErrorAction SilentlyContinue
        $connectionStatus = Get-ConnectionInformation

        if ($null -eq $connectionStatus)
        {
            Write-Warning "Failed to connect to Exchange Online."
            Read-Host "Press Enter to try again"
        }
    }
}

function TryConnect-MsolService
{
    Get-MsolDomain -ErrorVariable errorConnecting -ErrorAction SilentlyContinue | Out-Null

    if ($errorConnecting)
    {
        Read-Host "You must also connect to MsolService, press Enter to continue"
    }

    while ($errorConnecting)
    {
        Write-Host "Connecting to MsolService..." -ForegroundColor $infoColor
        Connect-MsolService -ErrorAction SilentlyContinue
        Get-MSolDomain -ErrorVariable errorConnecting -ErrorAction SilentlyContinue | Out-Null   

        if ($errorConnecting)
        {
            Read-Host -Prompt "Failed to connect to MsolService. Press Enter to try again"
        }
    }
}

function Prompt-DesiredMailboxType
{
    Write-Host "Which mailbox types do you want to retrieve? `n"
    Write-Host "[1] Shared mailboxes"
    Write-Host "[2] User mailboxes"
    Write-Host "[3] All mailboxes"

    while ($true)
    {
        $response = Read-Host

        if ($response -imatch "^\s*[123]\s*$") # regex matches a 1, 2, or 3 but allows spaces
        {
            break
        }
        Write-Warning "Please enter a 1, 2, or 3."
    }
    
    return [int]($response.Trim())
}

function Prompt-DesiredLicenseOutput
{
    
    Write-Host "`nWhich output do you prefer? `n"
    Write-Host "[1] All mailboxes"
    Write-Host "[2] Only mailboxes with a license"

    while ($true)
    {
        $response = Read-Host

        if ($response -imatch "^\s*[12]\s*$") # regex matches a 1 or 2 but allows spaces
        {
            break
        }
        Write-Warning "Please enter a 1 or 2."
    }
    
    return [int]($response.Trim())
}

function Export-MailboxData($desiredType, $desiredLicenseOutput)
{
    $path = New-DesktopPath -fileName "Mailbox export" -fileExt "csv"

    switch ($desiredType)
    {
        1 { $getMbExpression = "Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited" }
        2 { $getMbExpression = "Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited" }
        3 { $getMbExpression = "Get-Mailbox -ResultSize Unlimited" }
    }

    if ($desiredLicenseOutput -eq 1) # export mailboxes with or without a license
    {
        Invoke-Expression $getMbExpression |
        Write-ProgressInPipeline -activity "Exporting mailbox data..." -status "mailboxes processed" |
        Get-MailboxData |
        Export-CSV -Path $path -Append -NoTypeInformation
    }
    else # export only mailboxes with license(s) assigned
    {
        Invoke-Expression $getMbExpression  |        
        Write-ProgressInPipeline -activity "Exporting mailbox data..." -status "mailboxes processed" |
        Where-Object { $($_ | Get-MsolUser).isLicensed -eq $true } |
        Get-MailboxData |
        Export-CSV -Path $path -Append -NoTypeInformation
    }

    Write-Host "Finished exporting to $path" -ForegroundColor $successColor
}

function Write-ProgressInPipeline
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Position=0, Mandatory, ValueFromPipeline)]
        [object[]] $inputObjects,
        [string] $activity = "Processing items...",
        [string] $status = "items processed"
    )

    Begin 
    { 
        $itemsProcessed = 0 
    }

    Process
    {
        Write-Progress -Activity $activity -Status "$itemsProcessed $status"
        $itemsProcessed++
        return $_
    }
}

function Get-MailboxData
{
    [CmdletBinding()]
    Param 
    (
        [Parameter(Position=0, Mandatory, ValueFromPipeline)]
        $mailbox
    )

    Process
    {
        $mailboxUserData = $mailbox | Get-MsolUser
        $mailboxStats = $mailbox | Get-EXOMailboxStatistics

        if ($mailbox.ArchiveStatus -eq "Active")
        {
            $archiveMailboxStats = $mailbox | Get-EXOMailboxStatistics -Archive
            $archiveStorageConsumed = $archiveMailboxStats.TotalItemSize
        }
        else
        {
            $archiveStorageConsumed = ""
        }

        [PSCustomObject]@{
            UserPrincipalName = $mailbox.UserPrincipalName
            DisplayName = $mailbox.DisplayName
            Type = $mailbox.RecipientTypeDetails
            IsLicensed = $mailboxUserData.isLicensed
            Licenses = Get-LicensesAsString $mailboxUserData
            HiddenFromGAL = $mailbox.HiddenFromAddressListsEnabled
            StorageConsumed = $mailboxStats.TotalItemSize
            StorageLimit = $mailbox.ProhibitSendReceiveQuota
            ArchiveStatus = $mailbox.ArchiveStatus
            AutoExpandingArchiveEnabled = $mailbox.AutoExpandingArchiveEnabled
            ArchiveStorageConsumed = $archiveStorageConsumed
            ArchiveStorageQuota = $mailbox.ArchiveQuota
            RetentionPolicy = $mailbox.RetentionPolicy
            ForwardingSMTPAddress = $mailbox.ForwardingSMTPAddress   
            ForwardingAddress = $mailbox.ForwardingAddress 
        }
    }
}

function Get-LicensesAsString($mailboxUserData)
{
    $licenses = $mailboxUserData.Licenses.AccountSkuId

    if ($null -eq $licenses)
    {
        return $null
    }

    $licensesAsString = ""

    foreach ($license in $licenses)
    {
        $licensesAsString += $license.ToString()
        $licensesAsString += ", "
    }

    return $licensesAsString
}

function New-DesktopPath($fileName, $fileExt)
{
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $timeStamp = (Get-Date -Format yyyy-MM-dd-hh-mm).ToString()
    return "$desktopPath\$fileName $timeStamp.$fileExt"
}

# main
Initialize-ColorScheme
Show-Introduction
Use-Module("ExchangeOnlineManagement")
Use-Module("MSOnline")
TryConnect-ExchangeOnline
TryConnect-MsolService
$desiredMailboxType = Prompt-DesiredMailboxType
$desiredLicenseOutput = Prompt-DesiredLicenseOutput
Export-MailboxData -desiredType $desiredMailboxType -desiredLicenseOutput $desiredLicenseOutput
Read-Host "Press Enter to exit"
