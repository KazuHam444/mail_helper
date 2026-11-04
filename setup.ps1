param(
    [string]$PythonExecutable
)
$ErrorActionPreference = 'Stop'
$projectRoot = $PSScriptRoot
$venvPython = Join-Path $projectRoot '.venv\Scripts\python.exe'

if (-not (Test-Path -LiteralPath $venvPython)) {
    $pythonArgs = @()
    if ($PythonExecutable) {
        $launcher = $PythonExecutable
    } elseif (Get-Command py -ErrorAction SilentlyContinue) {
        $launcher = 'py'
        $pythonArgs = @('-3')
    } elseif (Get-Command python -ErrorAction SilentlyContinue) {
        $launcher = 'python'
    } else {
        throw 'Python was not found. Install Python 3.12 or 3.13 with Python Launcher or PATH enabled.'
    }
    & $launcher @pythonArgs -c 'import sys; sys.exit(0 if sys.version_info >= (3, 8) else 1)'
    if ($LASTEXITCODE -ne 0) { throw 'A working Python 3.8+ installation is required.' }
    & $launcher @pythonArgs -m venv (Join-Path $projectRoot '.venv')
    if ($LASTEXITCODE -ne 0) { throw 'Failed to create the virtual environment.' }
}

& $venvPython -m pip install -r (Join-Path $projectRoot 'requirements.txt')
if ($LASTEXITCODE -ne 0) { throw 'Dependency installation failed. Check Python, network and pip settings.' }
$config = Join-Path $projectRoot '.env'
if (-not (Test-Path -LiteralPath $config)) {
    Copy-Item -LiteralPath (Join-Path $projectRoot '.env.example') -Destination $config
    Write-Host 'Created .env from the configuration template.'
} else {
    Write-Host 'Existing .env preserved.'
}
Write-Host 'Setup complete. Configure QQ_EMAIL / QQ_AUTH_CODE in .env, then follow README.md for IMAP login and scheduling.'
