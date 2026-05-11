# wpnUserService_Kill.ps1
# Fast, forceful restart of the Windows Push Notifications User Service (per-user)
# Requires no service name guessing and avoids slow GUI stop/start

# Ensure script is running elevated
if (-not ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script must be run as Administrator."
    exit 1
}

Write-Host "Killing WpnUserService svchost instance..." -ForegroundColor Yellow

# Kill the svchost hosting the per-user WpnUserService
taskkill /F /FI "SERVICES eq WpnUserService_*" | Out-Null

# Small pause to allow SCM to respawn the service cleanly
Start-Sleep -Milliseconds 5000

# Verify service is back
$svc = Get-Service -Name WpnUserService* -ErrorAction SilentlyContinue

if ($svc -and $svc.Status -eq 'Running') {
    Write-Host "WpnUserService restarted successfully." -ForegroundColor Green
} else {
    Write-Warning "WpnUserService did not report Running state yet."
}