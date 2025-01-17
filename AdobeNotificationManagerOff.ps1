# Define the AppID for Adobe Acrobat
$appId = "Adobe.Acrobat.Notification.Manager"

# Define the registry path
$regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\$appId"

# Check if the registry key exists, if not, create it
if (-not (Test-Path $regPath)) {
    New-Item -Path $regPath -Force
}

# Set the "Enabled" value to 0 to disable notifications
Set-ItemProperty -Path $regPath -Name "Enabled" -Value 0 -Type DWord

# Confirm the change
$enabledValue = Get-ItemProperty -Path $regPath -Name "Enabled"
if ($enabledValue.Enabled -eq 0) {
    Write-Host "Notifications for $appId have been successfully disabled."
} else {
    Write-Host "Failed to disable notifications for $appId."
}