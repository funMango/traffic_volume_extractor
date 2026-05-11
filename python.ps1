$venvPython = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"

if (-not (Test-Path -LiteralPath $venvPython)) {
    Write-Error "Python virtual environment not found: $venvPython"
    exit 1
}

& $venvPython @args
exit $LASTEXITCODE
