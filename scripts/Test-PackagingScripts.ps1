$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$scriptPath = Join-Path $repoRoot 'build-installer.ps1'

if (-not (Test-Path -LiteralPath $scriptPath)) {
    throw "Missing build-installer.ps1"
}

$tokens = $null
$parseErrors = $null
[System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref] $tokens, [ref] $parseErrors) | Out-Null
if ($parseErrors.Count -gt 0) {
    $messages = $parseErrors | ForEach-Object { $_.Message }
    throw "build-installer.ps1 has syntax errors: $($messages -join '; ')"
}

$content = Get-Content -LiteralPath $scriptPath -Raw
$requiredSnippets = @(
    'jpackage',
    '--main-class',
    'com.xzy.ExcelTools',
    '--main-jar',
    'excel-processor-1.0-SNAPSHOT.jar',
    '--win-menu',
    '--win-shortcut',
    'java.desktop',
    'candle.exe',
    'light.exe',
    'self-extracting-exe',
    'iexpress.exe',
    'Compress-Archive',
    'WScript.Shell'
)

foreach ($snippet in $requiredSnippets) {
    if ($content -notlike "*$snippet*") {
        throw "build-installer.ps1 is missing required packaging detail: $snippet"
    }
}

Write-Host 'Packaging script checks passed.'
