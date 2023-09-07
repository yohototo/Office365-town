param (
    [switch]$noprompt = $false,   ## if -noprompt used then user will not be asked for any input
    [switch]$noupdate = $false,   ## if -noupdate used then module will not be checked for more recent version
    [switch]$debug = $false       ## if -debug create a log file
)

## Variables
$systemmessagecolor = "cyan"
$processmessagecolor = "green"
$errormessagecolor = "red"
$warningmessagecolor = "yellow"
$logFilePath = "..\o365-connect-aad.txt"

## Function to log messages
function Log-Message {
    param (
        [string]$message,
        [string]$color = "white"
    )
    Write-Host -ForegroundColor $color $message
    if ($debug) {
        Add-Content -Path $logFilePath -Value $message
    }
}

Clear-Host
if ($debug) {
    Log-Message "Script activity logged at $logFilePath"
    Start-Transcript -Path $logFilePath | Out-Null
}

Log-Message "Azure AD Connection script started`n" -ForegroundColor $systemmessagecolor
Log-Message "Prompt = $(-not $noprompt)" -ForegroundColor $processmessagecolor

# Check if Azure AD module is installed
if (Get-Module -ListAvailable -Name AzureAD) {
    Log-Message "Azure AD PowerShell module installed" -ForegroundColor $processmessagecolor
}
else {
    Log-Message "[001] - Azure AD PowerShell module not installed`n" -ForegroundColor $warningmessagecolor -BackgroundColor $errormessagecolor
    if (-not $noprompt) {
        do {
            $response = Read-Host -Prompt "`nDo you wish to install the Azure AD PowerShell module (Y/N)?"
        } until (-not [string]::IsNullOrEmpty($response))
        if ($response -eq 'Y' -or $response -eq 'y') {
            Log-Message "Installing Azure AD PowerShell module - Administration escalation required" -ForegroundColor $processmessagecolor
            Start-Process powershell -Verb RunAs -ArgumentList "Install-Module -Name AzureAD -Force -Confirm:$false" -Wait -WindowStyle Hidden
            Log-Message "Azure AD PowerShell module installed" -ForegroundColor $processmessagecolor
        }
        else {
            Log-Message "Terminating script" -ForegroundColor $processmessagecolor
            if ($debug) {
                Stop-Transcript | Out-Null
            }
            exit 1
        }
    }
    else {
        Log-Message "Installing Azure AD PowerShell module - Administration escalation required" -ForegroundColor $processmessagecolor
        Start-Process powershell -Verb RunAs -ArgumentList "Install-Module -Name AzureAD -Force -Confirm:$false" -Wait -WindowStyle Hidden
        Log-Message "Azure AD PowerShell module installed" -ForegroundColor $processmessagecolor
    }
}

# Check for module updates
if (-not $noupdate) {
    Log-Message "Check whether a newer version of Azure AD PowerShell module is available" -ForegroundColor $processmessagecolor
    $installedModule = Get-InstalledModule -Name AzureAD | Sort-Object Version -Descending | Select-Object -First 1
    $onlineModule = Find-Module -Name AzureAD | Sort-Object Version -Descending | Select-Object -First 1
    if ([Version]$installedModule.Version -ge [Version]$onlineModule.Version) {
        Log-Message "Local module $($installedModule.Version) is up to date" -ForegroundColor $processmessagecolor
    }
    else {
        Log-Message "Local module $($installedModule.Version) is outdated. Online module version is $($onlineModule.Version)" -ForegroundColor $warningmessagecolor
        if (-not $noprompt) {
            do {
                $response = Read-Host -Prompt "`nDo you wish to update the Azure AD PowerShell module (Y/N)?"
            } until (-not [string]::IsNullOrEmpty($response))
            if ($response -eq 'Y' -or $response -eq 'y') {
                Log-Message "Updating Azure AD PowerShell module - Administration escalation required" -ForegroundColor $processmessagecolor
                Start-Process powershell -Verb RunAs -ArgumentList "Update-Module -Name AzureAD -Force -Confirm:$false" -Wait -WindowStyle Hidden
                Log-Message "Azure AD PowerShell module updated" -ForegroundColor $processmessagecolor
            }
            else {
                Log-Message "Azure AD PowerShell module not updated" -ForegroundColor $processmessagecolor
            }
        }
        else {
            Log-Message "Updating Azure AD PowerShell module - Administration escalation required" -ForegroundColor $processmessagecolor
            Start-Process powershell -Verb RunAs -ArgumentList "Update-Module -Name AzureAD -Force -Confirm:$false" -Wait -WindowStyle Hidden
            Log-Message "Azure AD PowerShell module updated" -ForegroundColor $processmessagecolor
        }
    }
}

# Load Azure AD module
Log-Message "Azure AD PowerShell module loading" -ForegroundColor $processmessagecolor
try {
    Import-Module AzureAD -ErrorAction Stop | Out-Null
    Log-Message "Azure AD PowerShell module loaded" -ForegroundColor $processmessagecolor
}
catch {
    Log-Message "[002] - Unable to load Azure AD PowerShell module`n" -ForegroundColor $errormessagecolor
    Log-Message $_.Exception.Message -ForegroundColor $errormessagecolor
    if ($debug) {
        Stop-Transcript | Out-Null
    }
    exit 2
}

# Connect to Azure AD
Log-Message
