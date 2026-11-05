$ErrorActionPreference = 'Stop'
Push-Location (Join-Path $PSScriptRoot '..')
try {
    python scripts/prepare-icons.py
    if ($LASTEXITCODE -ne 0) { throw 'Icon preparation failed' }
    Push-Location frontend
    try {
        npm ci
        if ($LASTEXITCODE -ne 0) { throw 'npm ci failed' }
        npm test
        if ($LASTEXITCODE -ne 0) { throw 'Frontend tests failed' }
        npm run build
        if ($LASTEXITCODE -ne 0) { throw 'Frontend build failed' }
    } finally { Pop-Location }
    python -m PyInstaller --noconfirm --distpath dist/2.8.0 packaging/desktop.spec
    if ($LASTEXITCODE -ne 0) { throw 'Desktop build failed' }
    $clientDir = Join-Path (Get-Location).Path 'dist/2.8.0/Shitu'
    # Discover the executable instead of hardcoding its Chinese name: Windows
    # PowerShell 5.1 reads BOM-less UTF-8 scripts as ANSI and would mangle it.
    $deadline = (Get-Date).AddSeconds(30)
    while (-not (Get-ChildItem $clientDir -Filter *.exe -ErrorAction SilentlyContinue) -and (Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 250
    }
    $clientPath = (Get-ChildItem $clientDir -Filter *.exe | Select-Object -First 1).FullName
    if (-not $clientPath) { throw "Packaged desktop executable not found in $clientDir" }
    foreach ($compact in @($false, $true)) {
        $testArguments = @('--gui-smoke-test', '--test-image-list')
        if ($compact) { $testArguments += '--compact-window' }
        $clientTest = Start-Process -FilePath $clientPath -ArgumentList $testArguments -PassThru -WindowStyle Hidden
        if (-not $clientTest.WaitForExit(45000)) {
            $clientTest.Kill()
            throw 'Packaged desktop GUI smoke test timed out'
        }
        if ($clientTest.ExitCode -ne 0) { throw 'Packaged desktop GUI smoke test failed' }
    }
    Compress-Archive -LiteralPath 'dist/2.8.0/Shitu' -DestinationPath 'dist/Shitu-2.8.0-Windows-x64.zip' -Force
} finally { Pop-Location }
