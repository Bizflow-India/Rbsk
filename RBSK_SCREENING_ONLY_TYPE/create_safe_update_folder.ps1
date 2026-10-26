$ErrorActionPreference = "Stop"

$source = Split-Path -Parent $MyInvocation.MyCommand.Path
$destination = Join-Path $source "RBSK_SAFE_UPDATE"
$resolvedSource = [System.IO.Path]::GetFullPath($source)
$resolvedDestination = [System.IO.Path]::GetFullPath($destination)

if (-not $resolvedDestination.StartsWith($resolvedSource, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Unsafe update output path: $resolvedDestination"
}

if (Test-Path -LiteralPath $destination) {
    Remove-Item -LiteralPath $destination -Recurse -Force
}
New-Item -ItemType Directory -Path $destination | Out-Null

$programFiles = @(
    "RBSK_Easy_Launcher.pyw",
    "Students.py",
    "RUN_EXCEL_SCREENING.py",
    "RUN_SCREENING_NOW.py",
    "OPEN_RBSK_SCREENING_APP.bat",
    "RUN_SCREENING_ONLY.bat",
    "ONE_CLICK_INSTALL_AND_RUN.bat",
    "SETUP_RBSK.bat",
    "INSTALL_REQUIRED_FILES.bat",
    "CREATE_DESKTOP_SHORTCUT.bat",
    "create_shortcut.ps1",
    "README_1CLICK_SETUP.txt",
    "CREATE_SAFE_UPDATE_FOLDER.bat",
    "create_safe_update_folder.ps1",
    "SAFE_UPDATE_INSTRUCTIONS.txt"
)

$copied = 0
foreach ($name in $programFiles) {
    $sourceFile = Join-Path $source $name
    if (Test-Path -LiteralPath $sourceFile -PathType Leaf) {
        Copy-Item -LiteralPath $sourceFile -Destination (Join-Path $destination $name) -Force
        $copied++
    }
}

$forbidden = @(
    "rbsk_license.json",
    "rbsk_saved_login.json",
    "logs",
    "rbsk_browser_profile",
    "rbsk_browser_profiles"
)
foreach ($name in $forbidden) {
    if (Test-Path -LiteralPath (Join-Path $destination $name)) {
        throw "Protected PC data was included unexpectedly: $name"
    }
}

Write-Host ""
Write-Host "Safe update folder created successfully:" -ForegroundColor Green
Write-Host "  $destination"
Write-Host ""
Write-Host "Program files included: $copied"
Write-Host "Licence, tokens, passwords, profiles, logs, progress, and Excel data are excluded."
Write-Host "Copy the CONTENTS of RBSK_SAFE_UPDATE into the other PC's existing RBSK folder and choose Replace."

