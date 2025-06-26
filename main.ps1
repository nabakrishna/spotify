<#
.SYNOPSIS
    Installs Spicetify and optionally the Spicetify Marketplace with enhanced features.
.DESCRIPTION
    This script automates the installation of Spicetify for customizing Spotify,
    including architecture detection, version management, and optional Marketplace installation.
    Improved with better error handling, progress tracking, and user experience.
.NOTES
    File Name      : Install-Spicetify.ps1
    Author         : Your Name
    Prerequisite   : PowerShell 5.1+ (Windows)
    Version        : 2.0
#>

#region Initial Setup
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Configure console for better output
if ($Host.UI.RawUI) {
    $Host.UI.RawUI.WindowTitle = "Spicetify Installer"
}
#endregion

#region Configuration
$Config = @{
    SpicetifyPath = "$env:LOCALAPPDATA\spicetify"
    OldSpicetifyPath = "$HOME\spicetify-cli"
    GitHubAPI = "https://api.github.com/repos/spicetify/cli/releases/latest"
    MinPSVersion = [version]'5.1'
    ProgressPreference = 'SilentlyContinue' # Suppress progress bars for faster downloads
}

# Colors for consistent messaging
$Color = @{
    Success = 'Green'
    Error = 'Red'
    Warning = 'Yellow'
    Info = 'Cyan'
    Default = 'Gray'
}
#endregion

#region Helper Functions
function Write-Status {
    param(
        [string]$Message,
        [string]$Status,
        [string]$StatusColor = $Color.Default
    )
    
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] " -NoNewline -ForegroundColor $Color.Default
    Write-Host $Message -NoNewline
    if ($Status) {
        Write-Host " $Status" -ForegroundColor $StatusColor
    }
    else {
        Write-Host
    }
}

function Test-Administrator {
    try {
        $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Get-Architecture {
    switch ($env:PROCESSOR_ARCHITECTURE) {
        'AMD64' { return 'x64' }
        'ARM64' { return 'arm64' }
        default { return 'x86' }
    }
}

function Invoke-SafeWebRequest {
    param(
        [string]$Uri,
        [string]$OutFile,
        [int]$RetryCount = 3,
        [int]$RetryDelay = 2
    )
    
    for ($i = 1; $i -le $RetryCount; $i++) {
        try {
            Invoke-WebRequest -Uri $Uri -UseBasicParsing -OutFile $OutFile
            return $true
        }
        catch {
            if ($i -eq $RetryCount) {
                throw
            }
            Start-Sleep -Seconds $RetryDelay
        }
    }
}
#endregion

#region Core Functions
function Move-OldInstallation {
    if (Test-Path -Path $Config.OldSpicetifyPath) {
        Write-Status "Found old installation at $($Config.OldSpicetifyPath)" -Status "MIGRATING" -StatusColor $Color.Info
        
        try {
            if (-not (Test-Path -Path $Config.SpicetifyPath)) {
                New-Item -Path $Config.SpicetifyPath -ItemType Directory -Force | Out-Null
            }
            
            Get-ChildItem -Path $Config.OldSpicetifyPath | ForEach-Object {
                Copy-Item -Path $_.FullName -Destination $Config.SpicetifyPath -Recurse -Force
            }
            
            Remove-Item -Path $Config.OldSpicetifyPath -Recurse -Force
            Write-Status "Migration completed successfully" -Status "SUCCESS" -StatusColor $Color.Success
        }
        catch {
            Write-Status "Failed to migrate old installation" -Status "ERROR" -StatusColor $Color.Error
            Write-Host "  Error: $_" -ForegroundColor $Color.Error
        }
    }
}

function Get-LatestVersion {
    try {
        Write-Status "Fetching latest Spicetify version from GitHub"
        $response = Invoke-RestMethod -Uri $Config.GitHubAPI
        $version = $response.tag_name -replace 'v', ''
        Write-Status "Latest version found: $version" -Status "SUCCESS" -StatusColor $Color.Success
        return $version
    }
    catch {
        Write-Status "Failed to fetch latest version" -Status "ERROR" -StatusColor $Color.Error
        Write-Host "  Error: $_" -ForegroundColor $Color.Error
        throw
    }
}

function Install-Spicetify {
    param(
        [string]$Version,
        [string]$Architecture
    )
    
    $tempFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "spicetify-$Version-$Architecture.zip")
    $downloadUrl = "https://github.com/spicetify/cli/releases/download/v$Version/spicetify-$Version-windows-$Architecture.zip"
    
    try {
        Write-Status "Downloading Spicetify v$Version ($Architecture)"
        if (Invoke-SafeWebRequest -Uri $downloadUrl -OutFile $tempFile) {
            Write-Status "Download completed" -Status "SUCCESS" -StatusColor $Color.Success
        }
        
        Write-Status "Extracting to $($Config.SpicetifyPath)"
        Expand-Archive -Path $tempFile -DestinationPath $Config.SpicetifyPath -Force
        Write-Status "Extraction completed" -Status "SUCCESS" -StatusColor $Color.Success
        
        # Clean up
        Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Status "Installation failed" -Status "ERROR" -StatusColor $Color.Error
        Write-Host "  Error: $_" -ForegroundColor $Color.Error
        throw
    }
}

function Update-PathEnvironment {
    try {
        Write-Status "Updating PATH environment variable"
        
        $userPath = [Environment]::GetEnvironmentVariable('PATH', [EnvironmentVariableTarget]::User)
        
        # Remove old paths if they exist
        $userPath = ($userPath -split ';' | Where-Object {
            $_ -ne $Config.OldSpicetifyPath -and 
            $_ -ne $Config.SpicetifyPath
        }) -join ';'
        
        # Add new path if not already present
        if ($userPath -notlike "*$($Config.SpicetifyPath)*") {
            $userPath = "$userPath;$($Config.SpicetifyPath)"
            [Environment]::SetEnvironmentVariable('PATH', $userPath, [EnvironmentVariableTarget]::User)
            $env:PATH = "$env:PATH;$($Config.SpicetifyPath)"
        }
        
        Write-Status "PATH updated successfully" -Status "SUCCESS" -StatusColor $Color.Success
    }
    catch {
        Write-Status "Failed to update PATH" -Status "ERROR" -StatusColor $Color.Error
        Write-Host "  Error: $_" -ForegroundColor $Color.Error
    }
}

function Install-Marketplace {
    try {
        Write-Status "Starting Marketplace installation" -Status "INFO" -StatusColor $Color.Info
        
        $marketplaceScript = Join-Path -Path $env:TEMP -ChildPath "spicetify-marketplace-install.ps1"
        $marketplaceUrl = 'https://raw.githubusercontent.com/spicetify/spicetify-marketplace/main/resources/install.ps1'
        
        Write-Status "Downloading Marketplace installer"
        if (Invoke-SafeWebRequest -Uri $marketplaceUrl -OutFile $marketplaceScript) {
            Write-Status "Running Marketplace installer"
            & $marketplaceScript
            Remove-Item -Path $marketplaceScript -Force
            Write-Status "Marketplace installation completed" -Status "SUCCESS" -StatusColor $Color.Success
        }
    }
    catch {
        Write-Status "Marketplace installation failed" -Status "ERROR" -StatusColor $Color.Error
        Write-Host "  Error: $_" -ForegroundColor $Color.Error
    }
}
#endregion

#region Main Execution
try {
    # Clear screen and show header
    Clear-Host
    Write-Host "==============================================" -ForegroundColor $Color.Info
    Write-Host "          SPICETIFY INSTALLER v2.0           " -ForegroundColor $Color.Info
    Write-Host "==============================================" -ForegroundColor $Color.Info
    Write-Host ""

    # Check PowerShell version
    if ($PSVersionTable.PSVersion -lt $Config.MinPSVersion) {
        Write-Status "PowerShell $($Config.MinPSVersion) or higher is required" -Status "ERROR" -StatusColor $Color.Error
        Write-Host "  Current version: $($PSVersionTable.PSVersion)" -ForegroundColor $Color.Warning
        Write-Host "  Upgrade guide: https://aka.ms/install-powershell" -ForegroundColor $Color.Info
        throw "Incompatible PowerShell version"
    }

    # Check for admin rights
    if (Test-Administrator) {
        Write-Status "Running as administrator is not recommended" -Status "WARNING" -StatusColor $Color.Warning
        $choice = $Host.UI.PromptForChoice(
            "Continue?", 
            "Running as admin may cause issues. Continue anyway?", 
            [System.Management.Automation.Host.ChoiceDescription[]]@(
                New-Object System.Management.Automation.Host.ChoiceDescription("&No", "Exit installer"),
                New-Object System.Management.Automation.Host.ChoiceDescription("&Yes", "Continue anyway")
            ), 
            0
        )
        
        if ($choice -eq 0) {
            Write-Status "Installation cancelled by user" -Status "INFO" -StatusColor $Color.Info
            exit 0
        }
    }

    # Main installation process
    Move-OldInstallation
    $version = Get-LatestVersion
    $arch = Get-Architecture
    Install-Spicetify -Version $version -Architecture $arch
    Update-PathEnvironment

    Write-Host ""
    Write-Status "Spicetify v$version installed successfully!" -Status "COMPLETE" -StatusColor $Color.Success
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor $Color.Info
    Write-Host "1. Run " -NoNewline
    Write-Host "spicetify backup apply" -ForegroundColor $Color.Success -NoNewline
    Write-Host " to set up Spicetify"
    Write-Host "2. Run " -NoNewline
    Write-Host "spicetify -h" -ForegroundColor $Color.Success -NoNewline
    Write-Host " to see all available commands"
    Write-Host ""

    # Marketplace installation prompt
    $choice = $Host.UI.PromptForChoice(
        "Install Marketplace?", 
        "Would you like to install Spicetify Marketplace for themes and extensions?", 
        [System.Management.Automation.Host.ChoiceDescription[]]@(
            New-Object System.Management.Automation.Host.ChoiceDescription("&Yes", "Install Marketplace"),
            New-Object System.Management.Automation.Host.ChoiceDescription("&No", "Skip Marketplace")
        ), 
        0
    )

    if ($choice -eq 0) {
        Install-Marketplace
    }
    else {
        Write-Status "Marketplace installation skipped" -Status "INFO" -StatusColor $Color.Info
    }

    Write-Host ""
    Write-Host "Installation complete!" -ForegroundColor $Color.Success
    Write-Host "Press any key to exit..." -ForegroundColor $Color.Default
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}
catch {
    Write-Host ""
    Write-Status "Installation failed" -Status "ERROR" -StatusColor $Color.Error
    Write-Host "  Error: $_" -ForegroundColor $Color.Error
    Write-Host "  StackTrace: $($_.ScriptStackTrace)" -ForegroundColor $Color.Error
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor $Color.Default
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit 1
}
#endregion
