param(
    [int]$IntervalMinutes = 15,
    [string]$TaskName = "Mail-HR-Digest",
    [switch]$Remove
)
$ErrorActionPreference = 'Stop'
if ($Remove) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName 'Hotmail-HR-Digest' -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host 'Scheduled task removed. Local data is preserved.'
    exit 0
}
if ($IntervalMinutes -lt 5 -or $IntervalMinutes -gt 1440) {
    throw 'IntervalMinutes must be between 5 and 1440.'
}
$python = Join-Path $PSScriptRoot '.venv\Scripts\pythonw.exe'
if (-not (Test-Path -LiteralPath $python)) {
    throw 'Create .venv and install requirements first. See README.md.'
}
if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot '.env'))) {
    throw 'Configure .env and run login first. See README.md.'
}
$script = Join-Path $PSScriptRoot 'hr_digest.py'
$action = New-ScheduledTaskAction -Execute $python -Argument ('"' + $script + '" run') -WorkingDirectory $PSScriptRoot
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes)
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 14) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
$identity = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$principal = New-ScheduledTaskPrincipal -UserId $identity -LogonType Interactive -RunLevel Limited
$task = New-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description 'Read-only QQ mailbox HR notifications and daily digest; requires logged-in Windows user.'
Unregister-ScheduledTask -TaskName 'Hotmail-HR-Digest' -Confirm:$false -ErrorAction SilentlyContinue
Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force | Out-Null
Write-Host "Installed $TaskName, every $IntervalMinutes minutes while this Windows user is logged in."
Write-Host 'Daily digest defaults to 20:00 China time. See data/app.log and data/reports.'
