[CmdletBinding()]
param(
    [ValidateSet('exe', 'msi', 'app-image', 'self-extracting-exe')]
    [string] $Type = 'exe',

    [string] $AppName = 'ExcelTools',
    [string] $AppVersion = '1.0.1',
    [string] $Vendor = 'xzy',
    [string] $Description = 'Excel file processing desktop tools',
    [string] $MainClass = 'com.xzy.ExcelTools',
    [string] $MainJar = 'excel-processor-1.0-SNAPSHOT.jar',

    [string] $JavaHome,
    [string] $WixBin,
    [switch] $DownloadWix,
    [string] $WixDownloadUrl = 'https://github.com/wixtoolset/wix3/releases/download/wix314rtm/wix314-binaries.zip',

    [switch] $SkipMaven,
    [switch] $VerboseJpackage
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$RepoRoot = $PSScriptRoot
$TargetDir = Join-Path $RepoRoot 'target'
$PackageWorkDir = Join-Path $TargetDir 'jpackage'
$PackageInputDir = Join-Path $PackageWorkDir 'input'
$PackageTempDir = Join-Path $PackageWorkDir 'temp'
$DistDir = Join-Path $RepoRoot 'dist'
$BundledModules = 'java.base,java.compiler,java.desktop,java.management,java.naming,java.rmi,java.scripting,java.security.jgss,java.sql,java.xml.crypto'

function Get-FullPath {
    param([Parameter(Mandatory = $true)][string] $Path)
    return [System.IO.Path]::GetFullPath($Path)
}

function Assert-PathInsideRepo {
    param([Parameter(Mandatory = $true)][string] $Path)

    $fullPath = Get-FullPath $Path
    $repoPath = Get-FullPath $RepoRoot
    if (-not $fullPath.StartsWith($repoPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to clean path outside repository: $fullPath"
    }
    return $fullPath
}

function Reset-RepoDirectory {
    param([Parameter(Mandatory = $true)][string] $Path)

    $fullPath = Assert-PathInsideRepo $Path
    if (Test-Path -LiteralPath $fullPath) {
        Remove-Item -LiteralPath $fullPath -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path $fullPath | Out-Null
    return $fullPath
}

function Remove-RepoPathIfExists {
    param(
        [Parameter(Mandatory = $true)][string] $Path,
        [switch] $Recurse
    )

    $fullPath = Assert-PathInsideRepo $Path
    if (Test-Path -LiteralPath $fullPath) {
        if ($Recurse) {
            Remove-Item -LiteralPath $fullPath -Recurse -Force
        } else {
            Remove-Item -LiteralPath $fullPath -Force
        }
    }
}

function Test-JdkHome {
    param([Parameter(Mandatory = $true)][string] $JdkHomePath)

    return (Test-Path -LiteralPath (Join-Path $JdkHomePath 'bin\java.exe')) -and
           (Test-Path -LiteralPath (Join-Path $JdkHomePath 'bin\jpackage.exe'))
}

function Resolve-JdkHome {
    param([string] $RequestedJavaHome)

    $candidates = New-Object System.Collections.Generic.List[string]
    if ($RequestedJavaHome) {
        $candidates.Add($RequestedJavaHome)
    }
    if ($env:JAVA_HOME) {
        $candidates.Add($env:JAVA_HOME)
    }

    $jpackageCommand = Get-Command jpackage.exe -ErrorAction SilentlyContinue
    if ($jpackageCommand) {
        $candidates.Add((Split-Path -Parent (Split-Path -Parent $jpackageCommand.Source)))
    }

    $mavenCommand = Get-Command mvn -ErrorAction SilentlyContinue
    if ($mavenCommand) {
        $mavenVersion = & mvn -version 2>&1
        foreach ($line in $mavenVersion) {
            if ($line -match 'runtime:\s*(.+)$') {
                $candidates.Add($Matches[1].Trim())
            }
        }
    }

    $knownHomes = @(
        'D:\dev\JDK17',
        (Join-Path ${env:ProgramFiles} 'Java'),
        (Join-Path ${env:ProgramFiles} 'Eclipse Adoptium')
    )
    foreach ($knownHome in $knownHomes) {
        if ($knownHome -and (Test-Path -LiteralPath $knownHome)) {
            if (Test-JdkHome $knownHome) {
                $candidates.Add($knownHome)
            }
            Get-ChildItem -LiteralPath $knownHome -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -like 'jdk*' } |
                ForEach-Object { $candidates.Add($_.FullName) }
        }
    }

    foreach ($candidate in ($candidates | Where-Object { $_ } | Select-Object -Unique)) {
        $fullCandidate = Get-FullPath $candidate
        if (Test-JdkHome $fullCandidate) {
            return $fullCandidate
        }
    }

    throw 'Could not find a JDK with jpackage.exe. Install JDK 17 or newer, or pass -JavaHome C:\path\to\jdk.'
}

function Test-WixBin {
    param([Parameter(Mandatory = $true)][string] $Path)

    return (Test-Path -LiteralPath (Join-Path $Path 'candle.exe')) -and
           (Test-Path -LiteralPath (Join-Path $Path 'light.exe'))
}

function Find-WixBinUnder {
    param([Parameter(Mandatory = $true)][string] $Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }
    if (Test-WixBin $Path) {
        return (Get-FullPath $Path)
    }

    $candle = Get-ChildItem -LiteralPath $Path -Filter candle.exe -Recurse -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($candle) {
        $candidate = Split-Path -Parent $candle.FullName
        if (Test-WixBin $candidate) {
            return (Get-FullPath $candidate)
        }
    }
    return $null
}

function Resolve-WixBin {
    param(
        [string] $RequestedWixBin,
        [switch] $ShouldDownloadWix
    )

    $candidates = New-Object System.Collections.Generic.List[string]
    if ($RequestedWixBin) {
        $candidates.Add($RequestedWixBin)
    }

    $candleCommand = Get-Command candle.exe -ErrorAction SilentlyContinue
    if ($candleCommand) {
        $candidates.Add((Split-Path -Parent $candleCommand.Source))
    }

    $programFilesX86 = ${env:ProgramFiles(x86)}
    $knownHomes = @(
        (Join-Path $PackageWorkDir 'wix'),
        (Join-Path $programFilesX86 'WiX Toolset v3.14\bin'),
        (Join-Path $programFilesX86 'WiX Toolset v3.11\bin'),
        (Join-Path ${env:ProgramFiles} 'WiX Toolset v3.14\bin'),
        (Join-Path ${env:ProgramFiles} 'WiX Toolset v3.11\bin')
    )
    foreach ($knownHome in $knownHomes) {
        if ($knownHome) {
            $candidates.Add($knownHome)
        }
    }

    foreach ($candidate in ($candidates | Where-Object { $_ } | Select-Object -Unique)) {
        $resolved = Find-WixBinUnder $candidate
        if ($resolved) {
            return $resolved
        }
    }

    if ($ShouldDownloadWix) {
        $wixDir = Reset-RepoDirectory (Join-Path $PackageWorkDir 'wix')
        $zipPath = Join-Path $PackageWorkDir 'wix-binaries.zip'
        Write-Host "Downloading WiX Toolset binaries from $WixDownloadUrl"
        Invoke-WebRequest -Uri $WixDownloadUrl -OutFile $zipPath
        Expand-Archive -LiteralPath $zipPath -DestinationPath $wixDir -Force

        $resolved = Find-WixBinUnder $wixDir
        if ($resolved) {
            return $resolved
        }
        throw "Downloaded WiX archive, but candle.exe and light.exe were not found under $wixDir"
    }

    $message = @(
        'Windows exe/msi installers require WiX Toolset 3.x (candle.exe and light.exe) on PATH.',
        'Install WiX Toolset 3.14, pass -WixBin C:\path\to\wix\bin, or rerun this script with -DownloadWix.',
        'For a no-installer folder with bundled Java, run: .\build-installer.ps1 -Type app-image'
    ) -join [Environment]::NewLine
    throw $message
}

function Invoke-Checked {
    param(
        [Parameter(Mandatory = $true)][string] $Command,
        [string[]] $Arguments = @()
    )

    Write-Host "> $Command $($Arguments -join ' ')"
    & $Command @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$Command failed with exit code $LASTEXITCODE"
    }
}

function ConvertTo-PowerShellSingleQuotedString {
    param([Parameter(Mandatory = $true)][string] $Value)

    return "'" + $Value.Replace("'", "''") + "'"
}

function Write-Utf8BomFile {
    param(
        [Parameter(Mandatory = $true)][string] $Path,
        [Parameter(Mandatory = $true)][string] $Content
    )

    $utf8WithBom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8WithBom)
}

function Write-DefaultEncodedFile {
    param(
        [Parameter(Mandatory = $true)][string] $Path,
        [Parameter(Mandatory = $true)][string] $Content
    )

    [System.IO.File]::WriteAllText($Path, $Content, [System.Text.Encoding]::Default)
}

function Assert-PowerShellFileParses {
    param([Parameter(Mandatory = $true)][string] $Path)

    $tokens = $null
    $parseErrors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref] $tokens, [ref] $parseErrors) | Out-Null
    if ($parseErrors.Count -gt 0) {
        $messages = $parseErrors | ForEach-Object {
            "line $($_.Extent.StartLineNumber): $($_.Message)"
        }
        throw "PowerShell script does not parse: $Path$([Environment]::NewLine)$($messages -join [Environment]::NewLine)"
    }
}

function New-SelfExtractingInstaller {
    param(
        [Parameter(Mandatory = $true)][string] $AppImageDir,
        [Parameter(Mandatory = $true)][string] $OutputPath
    )

    $iexpressCommand = Get-Command iexpress.exe -ErrorAction SilentlyContinue
    if (-not $iexpressCommand) {
        throw 'iexpress.exe was not found. It is required for -Type self-extracting-exe.'
    }

    $sfxDir = Reset-RepoDirectory (Join-Path $PackageWorkDir 'self-extracting')
    $payloadDir = Join-Path $sfxDir 'payload'
    New-Item -ItemType Directory -Force -Path $payloadDir | Out-Null

    $zipPath = Join-Path $payloadDir 'app.zip'
    Compress-Archive -LiteralPath $AppImageDir -DestinationPath $zipPath -Force

    $installCmdPath = Join-Path $payloadDir 'install.cmd'
    $installCmd = @'
@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
exit /b %ERRORLEVEL%
'@
    [System.IO.File]::WriteAllText($installCmdPath, $installCmd, [System.Text.Encoding]::ASCII)

    $installPs1Path = Join-Path $payloadDir 'install.ps1'
    $appNameLiteral = ConvertTo-PowerShellSingleQuotedString $AppName
    $appExeLiteral = ConvertTo-PowerShellSingleQuotedString ($AppName + '.exe')
$installPs1Template = @'
param([switch] $Quiet)

$ErrorActionPreference = 'Stop'

$AppName = __APP_NAME__
$AppExe = __APP_EXE__
$ZipPath = Join-Path $PSScriptRoot 'app.zip'
$InstallRoot = Join-Path $env:LOCALAPPDATA 'Programs'
$InstallDir = Join-Path $InstallRoot $AppName
$ExePath = Join-Path $InstallDir $AppExe

function New-AppShortcut {
    param(
        [Parameter(Mandatory = $true)][string] $ShortcutPath,
        [Parameter(Mandatory = $true)][string] $TargetPath,
        [string] $Arguments = '',
        [string] $WorkingDirectory = ''
    )

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath = $TargetPath
    $shortcut.Arguments = $Arguments
    if ($WorkingDirectory) {
        $shortcut.WorkingDirectory = $WorkingDirectory
    }
    $shortcut.IconLocation = $TargetPath
    $shortcut.Save()
}

if (Test-Path -LiteralPath $InstallDir) {
    Remove-Item -LiteralPath $InstallDir -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $InstallRoot | Out-Null
Expand-Archive -LiteralPath $ZipPath -DestinationPath $InstallRoot -Force

if (-not (Test-Path -LiteralPath $ExePath)) {
    throw "Installed launcher was not found: $ExePath"
}

$desktopPath = [Environment]::GetFolderPath('Desktop')
if ($desktopPath) {
    New-AppShortcut -ShortcutPath (Join-Path $desktopPath ($AppName + '.lnk')) -TargetPath $ExePath -WorkingDirectory $InstallDir
}

$programsPath = [Environment]::GetFolderPath('Programs')
if ($programsPath) {
    $menuDir = Join-Path $programsPath $AppName
    New-Item -ItemType Directory -Force -Path $menuDir | Out-Null
    New-AppShortcut -ShortcutPath (Join-Path $menuDir ($AppName + '.lnk')) -TargetPath $ExePath -WorkingDirectory $InstallDir
}

if (-not $Quiet) {
    [System.Windows.Forms.MessageBox]::Show($AppName + ' installed successfully.', $AppName) | Out-Null
}
'@
    $installPs1 = $installPs1Template.
        Replace('__APP_NAME__', $appNameLiteral).
        Replace('__APP_EXE__', $appExeLiteral)
    Write-Utf8BomFile -Path $installPs1Path -Content $installPs1
    Assert-PowerShellFileParses -Path $installPs1Path

    $sedPath = Join-Path $sfxDir 'package.sed'
    $sourceDirForSed = $payloadDir.TrimEnd('\') + '\'
    $sedContent = @"
[Version]
Class=IEXPRESS
SEDVersion=3
[Options]
PackagePurpose=InstallApp
ShowInstallProgramWindow=1
HideExtractAnimation=1
UseLongFileName=1
InsideCompressed=0
CAB_FixedSize=0
CAB_ResvCodeSigning=0
RebootMode=N
InstallPrompt=
DisplayLicense=
FinishMessage=$AppName installed successfully.
TargetName=$OutputPath
FriendlyName=$AppName
AppLaunched=install.cmd
PostInstallCmd=<None>
AdminQuietInstCmd=install.cmd
UserQuietInstCmd=install.cmd
SourceFiles=SourceFiles
[SourceFiles]
SourceFiles0=$sourceDirForSed
[SourceFiles0]
%APPZIP%=
%INSTALLCMD%=
%INSTALLPS1%=
[Strings]
APPZIP=app.zip
INSTALLCMD=install.cmd
INSTALLPS1=install.ps1
"@
    Write-DefaultEncodedFile -Path $sedPath -Content $sedContent

    Remove-RepoPathIfExists -Path $OutputPath
    Invoke-Checked -Command $iexpressCommand.Source -Arguments @('/N', '/Q', $sedPath)

    $deadline = (Get-Date).AddSeconds(60)
    while (-not (Test-Path -LiteralPath $OutputPath)) {
        if ((Get-Date) -gt $deadline) {
            throw "IExpress finished but installer was not created: $OutputPath"
        }
        Start-Sleep -Milliseconds 500
    }

    $outputItem = Get-Item -LiteralPath $OutputPath
    if ($outputItem.Length -le 0) {
        throw "IExpress created an empty installer: $OutputPath"
    }
}

$resolvedJdkHome = Resolve-JdkHome -RequestedJavaHome $JavaHome
$jpackagePath = Join-Path $resolvedJdkHome 'bin\jpackage.exe'

$env:JAVA_HOME = $resolvedJdkHome
$env:Path = "$(Join-Path $resolvedJdkHome 'bin');$env:Path"
Write-Host "Using JDK: $resolvedJdkHome"

if (-not $SkipMaven) {
    if (-not (Get-Command mvn -ErrorAction SilentlyContinue)) {
        throw 'Maven was not found on PATH. Install Maven or run with -SkipMaven after building the JAR manually.'
    }
    Invoke-Checked -Command 'mvn' -Arguments @('clean', 'package')
}

$sourceJar = Join-Path $TargetDir $MainJar
if (-not (Test-Path -LiteralPath $sourceJar)) {
    throw "Main JAR not found: $sourceJar"
}

$inputDir = Reset-RepoDirectory $PackageInputDir
$tempDir = Reset-RepoDirectory $PackageTempDir
New-Item -ItemType Directory -Force -Path $DistDir | Out-Null
Copy-Item -LiteralPath $sourceJar -Destination (Join-Path $inputDir $MainJar) -Force

$jpackageType = $Type
$jpackageDest = $DistDir
if ($Type -eq 'self-extracting-exe') {
    $jpackageType = 'app-image'
    $jpackageDest = Reset-RepoDirectory (Join-Path $PackageWorkDir 'app-image')
    Remove-RepoPathIfExists -Path (Join-Path $DistDir ($AppName + '-' + $AppVersion + '-setup.exe'))
} elseif ($Type -eq 'app-image') {
    Remove-RepoPathIfExists -Path (Join-Path $DistDir $AppName) -Recurse
} else {
    Remove-RepoPathIfExists -Path (Join-Path $DistDir "$AppName-$AppVersion.$Type")
    $resolvedWixBin = Resolve-WixBin -RequestedWixBin $WixBin -ShouldDownloadWix:$DownloadWix
    $env:Path = "$resolvedWixBin;$env:Path"
    Write-Host "Using WiX: $resolvedWixBin"
}

$jpackageArgs = @(
    '--type', $jpackageType,
    '--name', $AppName,
    '--app-version', $AppVersion,
    '--vendor', $Vendor,
    '--description', $Description,
    '--input', $inputDir,
    '--dest', $jpackageDest,
    '--temp', $tempDir,
    '--main-jar', $MainJar,
    '--main-class', $MainClass,
    '--add-modules', $BundledModules,
    '--jlink-options', '--strip-debug --no-header-files --no-man-pages --strip-native-commands'
)

if ($Type -ne 'app-image' -and $Type -ne 'self-extracting-exe') {
    $jpackageArgs += @(
        '--win-menu',
        '--win-menu-group', $AppName,
        '--win-shortcut',
        '--win-dir-chooser',
        '--win-per-user-install'
    )
}

if ($VerboseJpackage) {
    $jpackageArgs += '--verbose'
}

Invoke-Checked -Command $jpackagePath -Arguments $jpackageArgs

if ($Type -eq 'self-extracting-exe') {
    $generatedAppImageDir = Join-Path $jpackageDest $AppName
    if (-not (Test-Path -LiteralPath $generatedAppImageDir)) {
        throw "Application image was not generated: $generatedAppImageDir"
    }
    $installerPath = Join-Path $DistDir ($AppName + '-' + $AppVersion + '-setup.exe')
    New-SelfExtractingInstaller -AppImageDir $generatedAppImageDir -OutputPath $installerPath
    Write-Host "Self-extracting installer created: $installerPath"
} elseif ($Type -eq 'app-image') {
    $outputPath = Join-Path $DistDir $AppName
    Write-Host "Application image created: $outputPath"
} else {
    $outputPath = Join-Path $DistDir ($AppName + '-' + $AppVersion + '.' + $Type)
    Write-Host "Installer created: $outputPath"
}
