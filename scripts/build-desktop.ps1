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
    python -m PyInstaller --noconfirm --distpath dist/2.6.0 packaging/desktop.spec
    if ($LASTEXITCODE -ne 0) { throw 'Desktop build failed' }
    $clientPath = Join-Path (Get-Location).Path 'dist/2.6.0/Shitu/拾图 Shitu.exe'
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
    Compress-Archive -LiteralPath 'dist/2.6.0/Shitu' -DestinationPath 'dist/Shitu-2.6.0-Windows-x64.zip' -Force
} finally { Pop-Location }
