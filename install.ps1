# Outlook CLI 2.0 Installer for Windows
$ErrorActionPreference = "Stop"

$RepoUrl = "https://github.com/ob-cheng/outlook-cli-2.0.git"
$InstallDir = "$env:LOCALAPPDATA\outlook-cli"

Write-Host "Installing Outlook CLI 2.0..." -ForegroundColor Cyan

# Step 1: Check Python
try {
    $pythonVersion = python --version 2>&1
    Write-Host "Found $pythonVersion"
} catch {
    Write-Host "Error: Python is required but not installed." -ForegroundColor Red
    Write-Host "Download from https://www.python.org/downloads/" -ForegroundColor Yellow
    exit 1
}

# Step 2: Clone or update repo
if (Test-Path $InstallDir) {
    Write-Host "Updating existing installation..."
    Push-Location $InstallDir
    git pull --quiet
    Pop-Location
} else {
    Write-Host "Cloning repository..."
    git clone --quiet $RepoUrl $InstallDir
}

# Step 3: Install dependencies
Write-Host "Installing dependencies..."
Push-Location $InstallDir
pip install -q -r requirements.txt
Pop-Location

# Step 4: Create batch launcher
$LauncherDir = "$env:LOCALAPPDATA\Microsoft\WindowsApps"
$Launcher = "$LauncherDir\outlook-cli.cmd"

$LauncherContent = @"
@echo off
python "$InstallDir\outlook.py" %*
"@

Set-Content -Path $Launcher -Value $LauncherContent -Encoding ASCII

# Step 5: Install AI agent skills
$SkillMarker = "$InstallDir\.skills-installed"
if (-not (Test-Path $SkillMarker)) {
    Write-Host "Installing AI agent skills..."

    $Agents = @{
        "$env:USERPROFILE\.claude" = "Claude Code"
        "$env:USERPROFILE\.cursor" = "Cursor"
        "$env:USERPROFILE\.windsurf" = "Windsurf"
        "$env:USERPROFILE\.copilot" = "GitHub Copilot"
        "$env:USERPROFILE\.hermes" = "Hermes Agent"
        "$env:USERPROFILE\.openclaw" = "OpenClaw"
    }

    foreach ($entry in $Agents.GetEnumerator()) {
        $dir = $entry.Key
        $name = $entry.Value

        if (Test-Path $dir) {
            Write-Host "  Found $name"
            $skillDir = "$dir\skills\outlook-cli"
            New-Item -ItemType Directory -Path $skillDir -Force | Out-Null
            Copy-Item "$InstallDir\SKILL.md" "$skillDir\SKILL.md" -Force
            Write-Host "    Installed: $skillDir\SKILL.md"
        }
    }

    New-Item -ItemType File -Path $SkillMarker -Force | Out-Null
}

Write-Host ""
Write-Host "Outlook CLI 2.0 installed successfully!" -ForegroundColor Green
Write-Host ""
Write-Host "Usage:" -ForegroundColor Cyan
Write-Host "  outlook-cli search --unread"
Write-Host "  outlook-cli send --to user@example.com --subject 'Hello' --body 'Hi there'"
Write-Host "  outlook-cli cal list"
Write-Host ""
Write-Host "Run 'outlook-cli --help' for more commands."
