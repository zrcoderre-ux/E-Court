<#
    LA Court E-Court Suite - native host installer (Windows).

    Registers the native messaging host so the extension's Export popup can
    launch the Word mail-merge template on Download.

    What it does:
      1. Rewrites com.lacourt.ecourt_host.json in place so its "path" points at
         the absolute location of ecourt_host.bat and its "allowed_origins"
         lists your extension.
      2. Adds the registry key Chrome (and optionally Edge) reads to find the
         host manifest.

    Usage (from a normal PowerShell prompt, in this folder):
        .\install.ps1 -ExtensionId <your-extension-id>

    Find <your-extension-id> at chrome://extensions (turn on Developer mode;
    it's the long id string under the extension's name).

    Add -IncludeEdge to also register for Microsoft Edge.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ExtensionId,

    [switch]$IncludeEdge
)

$ErrorActionPreference = "Stop"

$here         = Split-Path -Parent $MyInvocation.MyCommand.Path
$hostName     = "com.lacourt.ecourt_host"
$manifestPath = Join-Path $here "$hostName.json"
$batPath      = Join-Path $here "ecourt_host.bat"

if (-not (Test-Path $manifestPath)) { throw "Manifest not found: $manifestPath" }
if (-not (Test-Path $batPath))      { throw "Launcher not found: $batPath" }

# --- 1. Patch the host manifest (absolute bat path + this extension) ---------
$manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
$manifest.path = $batPath
$manifest.allowed_origins = @("chrome-extension://$ExtensionId/")
$manifest | ConvertTo-Json -Depth 5 | Set-Content $manifestPath -Encoding UTF8

# --- 2. Register the host in the current-user registry -----------------------
function Register-Host([string]$browserKey, [string]$browserLabel) {
    $regKey = "HKCU:\Software\$browserKey\NativeMessagingHosts\$hostName"
    New-Item -Path $regKey -Force | Out-Null
    # A registry key's default value is set by Set-Item on the key itself.
    Set-Item -Path $regKey -Value $manifestPath
    Write-Host "Registered for $browserLabel."
}

Register-Host "Google\Chrome" "Google Chrome"
if ($IncludeEdge) { Register-Host "Microsoft\Edge" "Microsoft Edge" }

Write-Host ""
Write-Host "Done."
Write-Host "  Host name : $hostName"
Write-Host "  Manifest  : $manifestPath"
Write-Host "  Launcher  : $batPath"
Write-Host "  Extension : $ExtensionId"
Write-Host ""
Write-Host "Reminder: set TEMPLATE_PATH in ecourt_host.py to your .dotm, and"
Write-Host "fully quit + reopen Chrome so it picks up the new host."
