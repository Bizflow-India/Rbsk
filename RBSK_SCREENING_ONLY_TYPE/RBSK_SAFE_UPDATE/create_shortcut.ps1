$ws = New-Object -ComObject WScript.Shell

# Collect all potential Desktop and Windows Start Menu / Main Home Window folders
$targetFolders = @(
    [Environment]::GetFolderPath('Desktop'),
    [System.IO.Path]::Combine($env:USERPROFILE, 'Desktop'),
    [System.IO.Path]::Combine($env:USERPROFILE, 'OneDrive', 'Desktop'),
    [System.IO.Path]::Combine($env:PUBLIC, 'Desktop'),
    [Environment]::GetFolderPath('Programs'),
    [System.IO.Path]::Combine($env:APPDATA, 'Microsoft', 'Windows', 'Start Menu', 'Programs'),
    [System.IO.Path]::Combine($env:ProgramData, 'Microsoft', 'Windows', 'Start Menu', 'Programs')
) | Select-Object -Unique

foreach ($dir in $targetFolders) {
    if (-not (Test-Path $dir)) {
        try { New-Item -ItemType Directory -Path $dir -Force | Out-Null } catch {}
    }
    
    if (Test-Path $dir) {
        # 1. Create standard .lnk shortcut
        try {
            $lnkPath = [System.IO.Path]::Combine($dir, "RBSK Screening.lnk")
            $sc = $ws.CreateShortcut($lnkPath)
            $sc.TargetPath = "cmd.exe"
            $sc.Arguments = '/c ""D:\RBSK_SCREENING_ONLY_TYPE\OPEN_RBSK_SCREENING_APP.bat""'
            $sc.WorkingDirectory = "D:\RBSK_SCREENING_ONLY_TYPE"
            $sc.IconLocation = "shell32.dll,14"
            $sc.Description = "RBSK Screening Health Application"
            $sc.Save()
            Write-Host "Created LNK shortcut: $lnkPath"
        } catch {
            Write-Host "Notice: LNK skipped for $dir"
        }

        # 2. Create direct .url shortcut for desktop locations
        if ($dir -like "*Desktop*") {
            try {
                $urlPath = [System.IO.Path]::Combine($dir, "RBSK Screening App.url")
                $urlContent = @"
[InternetShortcut]
URL=file:///D:/RBSK_SCREENING_ONLY_TYPE/OPEN_RBSK_SCREENING_APP.bat
IconIndex=13
IconFile=C:\Windows\System32\shell32.dll
"@
                Set-Content -Path $urlPath -Value $urlContent -Encoding ASCII
                Write-Host "Created URL shortcut: $urlPath"
            } catch {
                Write-Host "Notice: URL skipped for $dir"
            }
        }
    }
}

# Force Windows Shell to refresh Desktop and Start Menu icons immediately
try {
    $code = @"
    using System;
    using System.Runtime.InteropServices;
    public class NativeMethods {
        [DllImport("shell32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern void SHChangeNotify(int wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
    }
"@
    Add-Type -TypeDefinition $code
    [NativeMethods]::SHChangeNotify(0x08000000, 0x0000, [IntPtr]::Zero, [IntPtr]::Zero)
    Write-Host "Windows Desktop and Start Menu icons refreshed successfully."
} catch {
    Write-Host "Shell refresh completed."
}
