param([switch]$InstallModel)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$installDir = Join-Path $env:LOCALAPPDATA 'Programs\Ollama'
$app = Join-Path $installDir 'ollama app.exe'
$cli = Join-Path $installDir 'ollama.exe'
$appCommand = Get-Command 'ollama app.exe' -ErrorAction SilentlyContinue
$cliCommand = Get-Command 'ollama.exe' -ErrorAction SilentlyContinue
if (-not (Test-Path -LiteralPath $app) -and $appCommand) { $app = $appCommand.Source }
if (-not (Test-Path -LiteralPath $cli) -and $cliCommand) { $cli = $cliCommand.Source }
$installer = Join-Path $env:TEMP 'ChurchPresenter-OllamaSetup.exe'
$savedOrigins = [Environment]::GetEnvironmentVariable('OLLAMA_ORIGINS', 'User')
$origins = @($savedOrigins -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
if ($origins -notcontains 'http://tauri.localhost') { $origins += 'http://tauri.localhost' }
$env:OLLAMA_ORIGINS = $origins -join ','
[Environment]::SetEnvironmentVariable('OLLAMA_ORIGINS', $env:OLLAMA_ORIGINS, 'User')

if (-not (Test-Path -LiteralPath $app) -and -not (Test-Path -LiteralPath $cli)) {
    try {
        Invoke-WebRequest -Uri 'https://ollama.com/download/OllamaSetup.exe' -OutFile $installer
        $signature = Get-AuthenticodeSignature -FilePath $installer
        $publisher = $signature.SignerCertificate.Subject
        if ($signature.Status -ne 'Valid' -or $publisher -notmatch '(^|, )O=Ollama Inc\.(,|$)') {
            throw 'The downloaded Ollama installer did not have a valid Ollama signature.'
        }
        $process = Start-Process -FilePath $installer -ArgumentList @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART') -Wait -PassThru
        if ($process.ExitCode -ne 0) { throw "Ollama installer exited with code $($process.ExitCode)." }
    }
    finally {
        Remove-Item -LiteralPath $installer -Force -ErrorAction SilentlyContinue
    }
}

Get-Process -Name 'ollama app', 'ollama' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
if (Test-Path -LiteralPath $app) {
    Start-Process -FilePath $app
} elseif (Test-Path -LiteralPath $cli) {
    Start-Process -FilePath $cli -ArgumentList 'serve' -WindowStyle Hidden
} else {
    throw 'Ollama installed, but its application could not be found.'
}

$ready = $false
for ($attempt = 0; $attempt -lt 45; $attempt++) {
    try {
        Invoke-RestMethod -Uri 'http://localhost:11434/api/tags' -TimeoutSec 2 | Out-Null
        $ready = $true
        break
    } catch {
        Start-Sleep -Seconds 1
    }
}
if (-not $ready) { throw 'Ollama did not start its local service in time.' }
if ($InstallModel) {
    if (-not (Test-Path -LiteralPath $cli)) { throw 'Ollama CLI was not found after installation.' }
    $pull = Start-Process -FilePath $cli -ArgumentList @('pull', 'qwen3:8b') -Wait -PassThru -NoNewWindow
    if ($pull.ExitCode -ne 0) { throw "The Qwen3 8B model download exited with code $($pull.ExitCode)." }
}
