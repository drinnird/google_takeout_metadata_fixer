<#
.SYNOPSIS
  Parallel Google Photos Takeout metadata fixer for Windows.

.DESCRIPTION
  Version 6.0 adds a bounded multi-worker pool on top of the verified v5.1 metadata
  engine. Each worker owns its own persistent ExifTool -stay_open process and handles
  one media file completely (copy -> metadata write -> read-back verification) before
  taking another file. The source Takeout is never modified.

  Worker presets are exposed by the GUI. Command-line default is 4 workers; use
  -WorkerCount 1 for conservative HDD/compatibility mode.

.EXAMPLE
  .\Fix-GoogleTakeoutMetadata-v6.2.3.ps1 -InputFolder "D:\Takeout\Google Photos" -WorkerCount 4

.EXAMPLE
  .\Fix-GoogleTakeoutMetadata-v6.2.3.ps1 -InputFolder "D:\Takeout\Google Photos" -DryRun -WorkerCount 4
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$InputFolder,

    [Parameter(Position = 1)]
    [string]$OutputFolder,

    [string]$ExifToolPath,

    [switch]$DryRun,
    [switch]$ForceJsonTime,
    [switch]$TopLevelOnly,
    [switch]$OverwriteOutput,
    [switch]$SkipVerification,
    [switch]$DisablePersistentExifTool,
    [switch]$AllowSystemSleep,

    [ValidateRange(1, 8)]
    [int]$WorkerCount = 4,

    [ValidateRange(1, 500)]
    [double]$GpsAgreementMeters = 30,

    [string]$ProgressFile,
    [string]$CancelFile
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:KeepAwakeEnabled = $false
$script:LastProgressWriteUtc = [DateTime]::MinValue
$script:ChildProcesses = New-Object System.Collections.Generic.List[object]

# Same extension policy as the tested v5.1 worker engine.
$script:PhotoExtensions = @('.jpg', '.jpeg', '.heic', '.heif', '.png', '.webp', '.gif', '.tif', '.tiff', '.bmp', '.dng', '.cr2', '.cr3', '.nef', '.exr', '.sr2', '.orf', '.raw', '.360', '.3fr', '.jp2', '.eps', '.exif', '.ico', '.arw', '.avif', '.jxl')
$script:QuickTimeExtensions = @('.mp4', '.mov', '.m4v', '.3gp', '.3g2')
$script:OtherVideoExtensions = @('.mpg', '.mpeg', '.wmv', '.tod', '.mts', '.mod', '.mmv', '.mkv', '.m2ts', '.m2t', '.divx', '.avi', '.asf', '.webm')
$script:MediaExtensionLookup = @{}
foreach ($extension in $script:PhotoExtensions) { $script:MediaExtensionLookup[$extension] = $true }
foreach ($extension in $script:QuickTimeExtensions) { $script:MediaExtensionLookup[$extension] = $true }
foreach ($extension in $script:OtherVideoExtensions) { $script:MediaExtensionLookup[$extension] = $true }
$script:SignatureScanSkipExtensions = @(
    '.json', '.html', '.htm', '.txt', '.csv', '.xml', '.xmp', '.aae', '.log', '.md',
    '.pdf', '.zip', '.7z', '.rar', '.gz', '.tgz', '.tar', '.bz2', '.xz',
    '.exe', '.dll', '.msi', '.ps1', '.psm1', '.bat', '.cmd', '.com', '.scr',
    '.ini', '.cfg', '.conf', '.db', '.sqlite', '.sqlite3', '.lnk', '.url',
    '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.rtf',
    '.mp3', '.m4a', '.aac', '.wav', '.flac', '.ogg', '.oga', '.opus', '.wma', '.aif', '.aiff'
)
$script:SignatureScanSkipLookup = @{}
foreach ($extension in $script:SignatureScanSkipExtensions) { $script:SignatureScanSkipLookup[$extension] = $true }

function Write-ProgressState {
    param(
        [string]$Stage,
        [int]$Index = 0,
        [int]$Total = 0,
        [string]$Name = '',
        [string]$Message = '',
        [hashtable]$Extra = $null,
        [switch]$Force
    )
    if ([string]::IsNullOrWhiteSpace($ProgressFile)) { return }
    $nowUtc = [DateTime]::UtcNow
    $urgent = $Force -or ($Stage -in @('Scanning','Preflight','Done','Cancelled','Failed'))
    if (-not $urgent -and (($nowUtc - $script:LastProgressWriteUtc).TotalMilliseconds -lt 150)) { return }
    $script:LastProgressWriteUtc = $nowUtc
    try {
        $percent = 0
        if ($Total -gt 0) { $percent = [int][Math]::Min(100, [Math]::Max(0, (($Index / [double]$Total) * 100))) }
        $state = [ordered]@{ Stage=$Stage; Index=$Index; Total=$Total; Percent=$percent; Name=$Name; Message=$Message }
        if ($Extra) { foreach ($key in $Extra.Keys) { $state[$key] = $Extra[$key] } }
        $dir = Split-Path -Parent $ProgressFile
        if ($dir -and -not (Test-Path -LiteralPath $dir -PathType Container)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        ($state | ConvertTo-Json -Compress -Depth 5) | Set-Content -LiteralPath $ProgressFile -Encoding UTF8
    }
    catch { }
}

function Ensure-Directory {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}

function Get-DefaultOutputFolder {
    param([string]$InputPath)
    $normalized = Get-NormalizedDirectoryPath -Path $InputPath
    $parent = Split-Path -Parent $normalized
    $leaf = Split-Path -Leaf $normalized
    if (-not $parent) { $parent = [IO.Path]::GetPathRoot($normalized) }
    if (-not $leaf) { $leaf = 'Google Photos' }
    return (Join-Path $parent ($leaf + ' - Fixed'))
}

function ConvertTo-PowerShellLiteral {
    param([string]$Value)
    if ($null -eq $Value) { return "''" }
    return "'" + $Value.Replace("'", "''") + "'"
}

function Get-PowerShellExe {
    $candidate = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    return 'powershell.exe'
}

function Stop-ProcessTree {
    param([System.Diagnostics.Process]$Process)
    if ($null -eq $Process) { return }
    try { if ($Process.HasExited) { return } } catch { return }
    try {
        $taskkill = Join-Path $env:SystemRoot 'System32\taskkill.exe'
        if (Test-Path -LiteralPath $taskkill -PathType Leaf) {
            & $taskkill /PID $Process.Id /T /F 2>$null | Out-Null
            return
        }
    }
    catch { }
    try { $Process.Kill() } catch { }
}

function Read-JsonFileSafe {
    param([string]$Path)
    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
        return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json)
    }
    catch { return $null }
}

function Merge-WorkerCsvLogs {
    param([string[]]$Paths, [string]$Destination)
    $columns = @(
        'Source','Destination','Sidecar','MatchMethod','DeclaredType','DetectedType','TypeMismatch','TypeHandling',
        'Status','TimestampUTC','StoredTimestampUTC','Verification','TimeSource','TimeDecision',
        'GPS','GPSDecision','GPSDistanceMeters','GoogleGeoSource','GoogleLocationEdit','VerificationDetails','Note'
    )
    $encoding = [System.Text.UTF8Encoding]::new($true)
    $writer = New-Object System.IO.StreamWriter($Destination, $false, $encoding, 65536)
    try {
        $header = ($columns | ForEach-Object { '"' + $_.Replace('"','""') + '"' }) -join ','
        $writer.WriteLine($header)
        foreach ($path in $Paths) {
            if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
            $reader = New-Object System.IO.StreamReader($path, [Text.Encoding]::UTF8, $true, 65536)
            try {
                [void]$reader.ReadLine() # worker header
                while (-not $reader.EndOfStream) {
                    $line = $reader.ReadLine()
                    if ($null -ne $line) { $writer.WriteLine($line) }
                }
            }
            finally { $reader.Dispose() }
        }
        $writer.Flush()
    }
    finally { $writer.Dispose() }
}

function New-WorkerWrapper {
    param(
        [string]$Path,
        [string]$WorkerScript,
        [int]$WorkerId,
        [string]$ManifestPath,
        [string]$LogPath,
        [string]$SummaryPath,
        [string]$ResultPath,
        [string]$WorkerProgressPath,
        [switch]$RetryDirect
    )
    $parts = New-Object System.Collections.Generic.List[string]
    $parts.Add('& ' + (ConvertTo-PowerShellLiteral $WorkerScript))
    $parts.Add('-InputFolder ' + (ConvertTo-PowerShellLiteral $InputFolder))
    $parts.Add('-OutputFolder ' + (ConvertTo-PowerShellLiteral $OutputFolder))
    $parts.Add('-ExifToolPath ' + (ConvertTo-PowerShellLiteral $ExifToolPath))
    $parts.Add('-FileListPath ' + (ConvertTo-PowerShellLiteral $ManifestPath))
    $parts.Add('-WorkerLogPath ' + (ConvertTo-PowerShellLiteral $LogPath))
    $parts.Add('-WorkerSummaryPath ' + (ConvertTo-PowerShellLiteral $SummaryPath))
    $parts.Add('-WorkerResultPath ' + (ConvertTo-PowerShellLiteral $ResultPath))
    $parts.Add('-WorkerId ' + $WorkerId)
    $parts.Add('-InternalWorker')
    $parts.Add('-AllowSystemSleep') # parent owns the keep-awake state
    $parts.Add('-ProgressFile ' + (ConvertTo-PowerShellLiteral $WorkerProgressPath))
    if (-not [string]::IsNullOrWhiteSpace($CancelFile)) { $parts.Add('-CancelFile ' + (ConvertTo-PowerShellLiteral $CancelFile)) }
    $parts.Add('-GpsAgreementMeters ' + $GpsAgreementMeters.ToString([Globalization.CultureInfo]::InvariantCulture))
    if ($DryRun) { $parts.Add('-DryRun') }
    if ($ForceJsonTime) { $parts.Add('-ForceJsonTime') }
    if ($TopLevelOnly) { $parts.Add('-TopLevelOnly') }
    if ($SkipVerification) { $parts.Add('-SkipVerification') }
    if ($OverwriteOutput -or $RetryDirect) { $parts.Add('-OverwriteOutput') }
    if ($DisablePersistentExifTool -or $RetryDirect) { $parts.Add('-DisablePersistentExifTool') }
    $content = @(
        '$ErrorActionPreference = ''Stop''',
        ($parts -join ' ')
    ) -join [Environment]::NewLine
    [IO.File]::WriteAllText($Path, $content, [System.Text.UTF8Encoding]::new($true))
}

function Start-WorkerProcess {
    param(
        [int]$WorkerId,
        [string]$WrapperPath,
        [string]$ProgressPath,
        [string]$ResultPath,
        [string]$LogPath,
        [string]$SummaryPath,
        [string]$ManifestPath,
        [Int64]$AssignedBytes,
        [int]$AssignedFiles,
        [bool]$IsRetry
    )
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = Get-PowerShellExe
    $startInfo.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $WrapperPath + '"'
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    if (-not $process.Start()) { throw ('Could not start worker ' + $WorkerId + '.') }
    $record = [PSCustomObject]@{
        WorkerId = $WorkerId
        Process = $process
        StdOutTask = $process.StandardOutput.ReadToEndAsync()
        StdErrTask = $process.StandardError.ReadToEndAsync()
        ProgressPath = $ProgressPath
        ResultPath = $ResultPath
        LogPath = $LogPath
        SummaryPath = $SummaryPath
        ManifestPath = $ManifestPath
        AssignedBytes = $AssignedBytes
        AssignedFiles = $AssignedFiles
        IsRetry = $IsRetry
    }
    $script:ChildProcesses.Add($record)
    return $record
}

function Get-WorkerResultOrFailure {
    param($Record)
    $result = Read-JsonFileSafe -Path $Record.ResultPath
    if ($null -ne $result) { return $result }
    $stderr = ''
    try { $stderr = [string]$Record.StdErrTask.Result } catch { }
    return [PSCustomObject]@{
        WorkerId = $Record.WorkerId; Success = $false; FatalMessage = $(if ([string]::IsNullOrWhiteSpace($stderr)) { 'Worker exited without a result file.' } else { $stderr.Trim() });
        Cancelled = $false; Total = $Record.AssignedFiles; Attempted = 0; Processed = 0; Verified = 0; VerificationFailed = 0; Unmatched = 0; Failed = 1;
        TypeMismatches = 0; PreservedLocalTime = 0; PreservedGpsPrecision = 0; GoogleGpsEdits = 0; ScanErrors = 0; UnknownExtensionMedia = 0;
        ExifToolCommands = 0; ExifToolMode = 'worker failed'; ElapsedSeconds = 0; LogPath = $Record.LogPath; SummaryPath = $Record.SummaryPath; DiagnosticsPath = ''
    }
}

function Get-NormalizedDirectoryPath {
    param([string]$Path)

    $full = [IO.Path]::GetFullPath($Path)
    $root = [IO.Path]::GetPathRoot($full)
    if ($root -and $full.Length -le $root.Length) { return $root }
    return $full.TrimEnd('\', '/')
}

function Test-PathContainsPath {
    param([string]$ParentPath, [string]$ChildPath)

    $parent = Get-NormalizedDirectoryPath -Path $ParentPath
    $child = Get-NormalizedDirectoryPath -Path $ChildPath

    if ($parent.Equals($child, [StringComparison]::OrdinalIgnoreCase)) { return $true }
    $prefix = $parent
    if (-not $prefix.EndsWith([string][IO.Path]::DirectorySeparatorChar)) {
        $prefix += [IO.Path]::DirectorySeparatorChar
    }
    return $child.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)
}

function Test-CancelRequested {
    if ([string]::IsNullOrWhiteSpace($CancelFile)) { return $false }
    try { return (Test-Path -LiteralPath $CancelFile -PathType Leaf) } catch { return $false }
}

function Select-FolderInteractive {
    param([string]$Description)

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = $Description
        $dialog.ShowNewFolderButton = $false
        $result = $dialog.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            return $dialog.SelectedPath
        }
    }
    catch {
        # Fall through to Read-Host on systems without Windows Forms.
    }

    return (Read-Host $Description)
}

function Normalize-ExifToolExecutable {
    param([string]$Path)

    $resolved = (Resolve-Path -LiteralPath $Path).Path
    $leaf = Split-Path -Leaf $resolved

    # The official Windows package is named exiftool(-k).exe.  The "-k" behavior
    # pauses before termination, which is undesirable when a script invokes ExifTool
    # repeatedly.  Make a sibling exiftool.exe copy, as ExifTool's Windows instructions
    # normally recommend when using it from the command line.
    if ($leaf -ieq 'exiftool(-k).exe') {
        $plain = Join-Path (Split-Path -Parent $resolved) 'exiftool.exe'
        if (-not (Test-Path -LiteralPath $plain -PathType Leaf)) {
            try {
                Copy-Item -LiteralPath $resolved -Destination $plain -Force
            }
            catch {
                throw "Found exiftool(-k).exe, but could not create exiftool.exe beside it. Rename exiftool(-k).exe to exiftool.exe, then run this script again."
            }
        }
        return (Resolve-Path -LiteralPath $plain).Path
    }

    return $resolved
}

function Resolve-ExifTool {
    param([string]$RequestedPath)

    if ($RequestedPath) {
        if (Test-Path -LiteralPath $RequestedPath -PathType Leaf) {
            return (Normalize-ExifToolExecutable -Path $RequestedPath)
        }
        throw "ExifTool was not found at: $RequestedPath"
    }

    $scriptDir = $PSScriptRoot
    if (-not $scriptDir) { $scriptDir = (Get-Location).Path }

    foreach ($candidate in @(
        (Join-Path $scriptDir 'exiftool.exe'),
        (Join-Path $scriptDir 'exiftool(-k).exe')
    )) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return (Normalize-ExifToolExecutable -Path $candidate)
        }
    }

    foreach ($folder in @(Get-ChildItem -LiteralPath $scriptDir -Directory -Filter 'exiftool-*' -ErrorAction SilentlyContinue)) {
        foreach ($exeName in @('exiftool.exe', 'exiftool(-k).exe')) {
            $candidate = Join-Path $folder.FullName $exeName
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                return (Normalize-ExifToolExecutable -Path $candidate)
            }
        }
    }

    $cmd = Get-Command exiftool.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    throw @"
ExifTool was not found.

Extract your exiftool-13.59_64.zip and keep its EXE together with the
exiftool_files folder. Put this script next to that extracted folder,
or pass -ExifToolPath "C:\path\to\exiftool.exe".
"@
}

function Enable-SystemAwake {
    if ($AllowSystemSleep) { return }
    try {
        if (-not ('GtmfPowerState' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class GtmfPowerState {
    [DllImport("kernel32.dll")]
    public static extern uint SetThreadExecutionState(uint esFlags);
}
'@
        }
        # ES_CONTINUOUS | ES_SYSTEM_REQUIRED. The display may still turn off normally.
        [void][GtmfPowerState]::SetThreadExecutionState(0x80000001)
        $script:KeepAwakeEnabled = $true
    }
    catch { }
}

function Disable-SystemAwake {
    if (-not $script:KeepAwakeEnabled) { return }
    try { [void][GtmfPowerState]::SetThreadExecutionState(0x80000000) } catch { }
    $script:KeepAwakeEnabled = $false
}

function Format-ByteSize {
    param([Int64]$Bytes)
    if ($Bytes -ge 1TB) { return ('{0:N2} TB' -f ($Bytes / 1TB)) }
    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N1} MB' -f ($Bytes / 1MB)) }
    return ('{0:N0} KB' -f ($Bytes / 1KB))
}

function Test-OutputFreeSpace {
    param([string]$Path, [Int64]$MediaBytes, [bool]$OutputWasNonEmpty)

    $result = [PSCustomObject]@{ Checked = $false; FreeBytes = [Int64]0; RequiredBytes = [Int64]0; Message = '' }
    try {
        $root = [IO.Path]::GetPathRoot([IO.Path]::GetFullPath($Path))
        if ([string]::IsNullOrWhiteSpace($root) -or $root.StartsWith('\\')) { return $result }
        $drive = New-Object System.IO.DriveInfo($root)
        $free = [Int64]$drive.AvailableFreeSpace
        $margin = [Int64][Math]::Max(256MB, [double]$MediaBytes * 0.03)
        $required = [Int64]($MediaBytes + $margin)
        $result.Checked = $true
        $result.FreeBytes = $free
        $result.RequiredBytes = $required
        $result.Message = ('Media size ' + (Format-ByteSize $MediaBytes) + '; free space ' + (Format-ByteSize $free) + '.')
        if (-not $OutputWasNonEmpty -and $free -lt $required) {
            throw ('Not enough free space for a safe full copy. Media totals about ' + (Format-ByteSize $MediaBytes) +
                ', recommended free space is at least ' + (Format-ByteSize $required) + ', but only ' + (Format-ByteSize $free) + ' is available on ' + $root + '.')
        }
    }
    catch {
        if ($_.Exception.Message -like 'Not enough free space*') { throw }
    }
    return $result
}

function Get-MediaTypeInfo {
    param(
        [string]$Path,
        [string]$DeclaredExtension
    )

    $declared = ''
    if (-not [string]::IsNullOrWhiteSpace($DeclaredExtension)) {
        $declared = $DeclaredExtension.ToLowerInvariant()
    }

    $detected = 'Unknown'
    $effective = $declared
    $confident = $false

    try {
        $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        try {
            $buffer = New-Object byte[] 64
            $read = $stream.Read($buffer, 0, $buffer.Length)
        }
        finally {
            $stream.Dispose()
        }

        if ($read -ge 3 -and $buffer[0] -eq 0xFF -and $buffer[1] -eq 0xD8 -and $buffer[2] -eq 0xFF) {
            $detected = 'JPEG'; $effective = '.jpg'; $confident = $true
        }
        elseif ($read -ge 8 -and
                $buffer[0] -eq 0x89 -and $buffer[1] -eq 0x50 -and $buffer[2] -eq 0x4E -and $buffer[3] -eq 0x47 -and
                $buffer[4] -eq 0x0D -and $buffer[5] -eq 0x0A -and $buffer[6] -eq 0x1A -and $buffer[7] -eq 0x0A) {
            $detected = 'PNG'; $effective = '.png'; $confident = $true
        }
        elseif ($read -ge 6) {
            $sig6 = [Text.Encoding]::ASCII.GetString($buffer, 0, 6)
            if ($sig6 -eq 'GIF87a' -or $sig6 -eq 'GIF89a') {
                $detected = 'GIF'; $effective = '.gif'; $confident = $true
            }
        }

        if (-not $confident -and $read -ge 12) {
            $riff = [Text.Encoding]::ASCII.GetString($buffer, 0, 4)
            $webp = [Text.Encoding]::ASCII.GetString($buffer, 8, 4)
            if ($riff -eq 'RIFF' -and $webp -eq 'WEBP') {
                $detected = 'WebP'; $effective = '.webp'; $confident = $true
            }
        }

        if (-not $confident -and $read -ge 2 -and $buffer[0] -eq 0x42 -and $buffer[1] -eq 0x4D) {
            $detected = 'BMP'; $effective = '.bmp'; $confident = $true
        }

        # HEIC/HEIF, AVIF, MP4 and MOV are ISO Base Media File Format containers.
        # The box type "ftyp" starts at byte offset 4, the major brand at offset 8,
        # and compatible brands follow after the 4-byte minor version. Read both major
        # and compatible brands so generic "mif1" files are not incorrectly forced to
        # HEIC when an AVIF compatible brand is present.
        if (-not $confident -and $read -ge 12) {
            $boxType = [Text.Encoding]::ASCII.GetString($buffer, 4, 4)
            if ($boxType -eq 'ftyp') {
                $brands = New-Object System.Collections.Generic.List[string]
                $majorBrand = [Text.Encoding]::ASCII.GetString($buffer, 8, 4).ToLowerInvariant()
                $brands.Add($majorBrand)

                $ftypSize = (($buffer[0] -shl 24) -bor ($buffer[1] -shl 16) -bor ($buffer[2] -shl 8) -bor $buffer[3])
                $brandEnd = [Math]::Min($read, $ftypSize)
                for ($offset = 16; ($offset + 3) -lt $brandEnd; $offset += 4) {
                    $brands.Add([Text.Encoding]::ASCII.GetString($buffer, $offset, 4).ToLowerInvariant())
                }

                $heicBrands = @('heic','heix','hevc','hevx','heim','heis')
                $avifBrands = @('avif','avis')
                $hasAvifBrand = (@($brands | Where-Object { $avifBrands -contains $_ }).Count -gt 0)
                $hasHeicBrand = (@($brands | Where-Object { $heicBrands -contains $_ }).Count -gt 0)

                if ($hasAvifBrand) {
                    $detected = 'AVIF'; $effective = '.avif'; $confident = $true
                }
                elseif ($hasHeicBrand) {
                    $detected = 'HEIF/HEIC'; $effective = '.heic'; $confident = $true
                }
                elseif ($majorBrand -in @('mif1','msf1')) {
                    # Generic HEIF family: keep the user's declared HEIC/HEIF/AVIF
                    # extension when it is already plausible instead of guessing.
                    if ($declared -in @('.heic','.heif','.avif')) {
                        $detected = $(if ($declared -eq '.avif') { 'AVIF' } else { 'HEIF/HEIC' })
                        $effective = $declared
                        $confident = $true
                    }
                }
                else {
                    $detected = 'QuickTime/MP4'; $effective = $(if ($declared -in @('.mov','.m4v','.3gp','.3g2')) { $declared } else { '.mp4' }); $confident = $true
                }
            }
        }
    }
    catch {
        # Detection is an additional safety feature. Fall back to the declared extension
        # rather than preventing the rest of the Takeout from being processed.
        $detected = 'Unknown'
        $effective = $declared
        $confident = $false
    }

    $declaredFamily = switch ($declared) {
        '.jpg' { 'JPEG' }
        '.jpeg' { 'JPEG' }
        '.png' { 'PNG' }
        '.gif' { 'GIF' }
        '.webp' { 'WebP' }
        '.bmp' { 'BMP' }
        '.heic' { 'HEIF/HEIC' }
        '.heif' { 'HEIF/HEIC' }
        '.avif' { 'AVIF' }
        '.mp4' { 'QuickTime/MP4' }
        '.mov' { 'QuickTime/MP4' }
        '.m4v' { 'QuickTime/MP4' }
        '.3gp' { 'QuickTime/MP4' }
        '.3g2' { 'QuickTime/MP4' }
        default { '' }
    }

    $mismatch = $false
    if ($confident) {
        # A confidently recognized media payload with an unrecognized extension is also
        # a mismatch. This is how files such as PXL_....MP (actual MP4 bytes) are brought
        # into the media pass instead of being silently skipped.
        if ([string]::IsNullOrWhiteSpace($declaredFamily)) {
            $mismatch = $true
        }
        else {
            $mismatch = ($declaredFamily -ne $detected)
        }
    }

    $handling = 'Direct processing using declared extension.'
    if ($mismatch) {
        if ([string]::IsNullOrWhiteSpace($declaredFamily)) {
            $shownDeclared = $(if ([string]::IsNullOrWhiteSpace($declared)) { '(no extension)' } else { $declared })
            $handling = ('Detected ' + $detected + ' media content despite unrecognized ' + $shownDeclared + ' filename; include it in the scan and use a temporary ' + $effective + ' alias if metadata processing is required.')
        }
        else {
            $handling = ('Detected ' + $detected + ' content despite ' + $declared + ' filename; process through temporary ' + $effective + ' alias and preserve original filename.')
        }
    }

    return [PSCustomObject]@{
        DeclaredType = $(if ([string]::IsNullOrWhiteSpace($declared)) { '(none)' } else { $declared.TrimStart([char]'.').ToUpperInvariant() })
        DetectedType = $detected
        EffectiveExtension = $effective
        TypeMismatch = $mismatch
        Confident = $confident
        Handling = $handling
    }
}

function Get-AvailableLogPath {
    param([string]$PreferredPath)

    if (-not (Test-Path -LiteralPath $PreferredPath)) { return $PreferredPath }

    $directory = Split-Path -Parent $PreferredPath
    $stem = [IO.Path]::GetFileNameWithoutExtension($PreferredPath)
    $extension = [IO.Path]::GetExtension($PreferredPath)
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $candidate = Join-Path $directory ($stem + '-' + $stamp + $extension)
    $counter = 2
    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path $directory ($stem + '-' + $stamp + '-' + $counter + $extension)
        $counter++
    }
    return $candidate
}


# ---------------- Main parallel orchestrator ----------------
$runStopwatch = [Diagnostics.Stopwatch]::StartNew()
$jobRoot = ''
$logPath = ''
$summaryPath = ''
$fatalMessage = ''
$outputWasNonEmpty = $false
$cancelled = $false
$retryCount = 0

try {
    $scriptDir = $PSScriptRoot
    if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
    $workerScript = Join-Path $scriptDir 'Fix-GoogleTakeoutMetadata-v6.2.3-Worker.ps1'
    if (-not (Test-Path -LiteralPath $workerScript -PathType Leaf)) { throw 'Fix-GoogleTakeoutMetadata-v6.2.3-Worker.ps1 is missing from the package.' }

    if (-not $InputFolder) { $InputFolder = Select-FolderInteractive -Description 'Select your extracted Google Photos Takeout folder' }
    if ([string]::IsNullOrWhiteSpace($InputFolder) -or -not (Test-Path -LiteralPath $InputFolder -PathType Container)) { throw "Input folder does not exist: $InputFolder" }
    $InputFolder = Get-NormalizedDirectoryPath -Path (Resolve-Path -LiteralPath $InputFolder).Path

    if (-not $OutputFolder) { $OutputFolder = Get-DefaultOutputFolder -InputPath $InputFolder }
    $OutputFolder = Get-NormalizedDirectoryPath -Path $OutputFolder

    if (-not $DryRun) {
        if ((Test-PathContainsPath -ParentPath $InputFolder -ChildPath $OutputFolder) -or (Test-PathContainsPath -ParentPath $OutputFolder -ChildPath $InputFolder)) {
            throw 'The input and output folders must be separate and must not contain one another.'
        }
        if (Test-Path -LiteralPath $OutputFolder) {
            if (-not (Test-Path -LiteralPath $OutputFolder -PathType Container)) { throw "Output path exists but is not a folder: $OutputFolder" }
            $existingItem = Get-ChildItem -LiteralPath $OutputFolder -Force -ErrorAction Stop | Select-Object -First 1
            $outputWasNonEmpty = ($null -ne $existingItem)
            if ($outputWasNonEmpty -and -not $OverwriteOutput) { throw "The output folder is not empty. Choose an empty/new folder, or enable overwrite/re-run: $OutputFolder" }
        }
        Ensure-Directory -Path $OutputFolder
    }

    $ExifToolPath = Resolve-ExifTool -RequestedPath $ExifToolPath

    Write-Host ''
    Write-Host 'Google Photos Takeout Metadata Fixer v6.2.3' -ForegroundColor Cyan
    Write-Host ('Input:      ' + $InputFolder)
    Write-Host ('Output:     ' + $OutputFolder)
    Write-Host ('ExifTool:   ' + $ExifToolPath)
    Write-Host ('Workers:    ' + $WorkerCount + ' requested')
    Write-Host ('Verify:     ' + $(if ($SkipVerification) { 'disabled' } else { 'enabled' }))
    if ($DryRun) { Write-Host 'Mode:       DRY RUN' -ForegroundColor Yellow }
    Write-Host ''

    Write-ProgressState -Stage 'Scanning' -Message 'Scanning for media, including unusual extensions detected by file signature.' -Force
    $scanErrors = @()
    $scanParams = @{ LiteralPath=$InputFolder; File=$true; ErrorAction='SilentlyContinue'; ErrorVariable='+scanErrors' }
    if (-not $TopLevelOnly) { $scanParams['Recurse'] = $true }

    $mediaFiles = New-Object 'System.Collections.Generic.List[System.IO.FileInfo]'
    $mediaTypeScanCache = @{}
    $unknownExtensionMediaDetected = 0
    Get-ChildItem @scanParams | ForEach-Object {
        $candidate = $_
        $candidateExt = $candidate.Extension.ToLowerInvariant()
        if ($script:MediaExtensionLookup.ContainsKey($candidateExt)) {
            $mediaFiles.Add($candidate)
        }
        elseif (-not $script:SignatureScanSkipLookup.ContainsKey($candidateExt) -and $candidate.Length -ge 3) {
            $candidateType = Get-MediaTypeInfo -Path $candidate.FullName -DeclaredExtension $candidateExt
            if ($candidateType.Confident -and $candidateType.DetectedType -ne 'Unknown') {
                $mediaFiles.Add($candidate)
                $mediaTypeScanCache[$candidate.FullName] = $candidateType
                $unknownExtensionMediaDetected++
            }
        }
    }

    if ($mediaFiles.Count -eq 0) {
        Write-ProgressState -Stage 'Done' -Message 'No supported or signature-detected photo/video files were found.' -Extra @{ Processed=0; Verified=0; VerificationFailed=0; UnknownExtensionMedia=0; Unmatched=0; Failed=0; WorkerCount=0 } -Force
        Write-Warning 'No supported or signature-detected photo/video files were found.'
        return
    }

    $total = $mediaFiles.Count
    [Int64]$totalMediaBytes = 0
    foreach ($m in $mediaFiles) { $totalMediaBytes += [Int64]$m.Length }
    $effectiveWorkers = [Math]::Max(1, [Math]::Min($WorkerCount, $total))

    $preflightMessage = ('Found ' + $total + ' media files (' + (Format-ByteSize $totalMediaBytes) + '). Using ' + $effectiveWorkers + ' worker(s).')
    if ($unknownExtensionMediaDetected -gt 0) { $preflightMessage += (' Signature-detected unusual extensions: ' + $unknownExtensionMediaDetected + '.') }
    Write-ProgressState -Stage 'Preflight' -Total $total -Message $preflightMessage -Extra @{ WorkerCount=$effectiveWorkers; UnknownExtensionMedia=$unknownExtensionMediaDetected } -Force

    if (-not $DryRun) {
        $spaceCheck = Test-OutputFreeSpace -Path $OutputFolder -MediaBytes $totalMediaBytes -OutputWasNonEmpty $outputWasNonEmpty
        if ($spaceCheck.Checked) { Write-Host ('Disk:       ' + $spaceCheck.Message) }
    }

    if ($DryRun) {
        $dryRunLogFolder = Split-Path -Parent $InputFolder
        if (-not $dryRunLogFolder) { $dryRunLogFolder = $InputFolder }
        $logPath = Get-AvailableLogPath -PreferredPath (Join-Path $dryRunLogFolder 'metadata-fix-dry-run.csv')
        $summaryPath = Get-AvailableLogPath -PreferredPath (Join-Path $dryRunLogFolder 'metadata-fix-summary.txt')
    }
    else {
        $logPath = Get-AvailableLogPath -PreferredPath (Join-Path $OutputFolder 'metadata-fix-log.csv')
        $summaryPath = Get-AvailableLogPath -PreferredPath (Join-Path $OutputFolder 'metadata-fix-summary.txt')
    }

    $jobRoot = Join-Path ([IO.Path]::GetTempPath()) ('GoogleTakeoutMetadataFixer-v6.2.3-' + [Guid]::NewGuid().ToString('N'))
    Ensure-Directory -Path $jobRoot
    if ([string]::IsNullOrWhiteSpace($CancelFile)) {
        $CancelFile = Join-Path $jobRoot 'cancel.request'
        Remove-Item -LiteralPath $CancelFile -Force -ErrorAction SilentlyContinue
    }
    elseif (Test-CancelRequested) {
        $cancelled = $true
        Write-ProgressState -Stage 'Cancelled' -Index 0 -Total $total -Message 'Cancelled before the worker pool started.' -Extra @{ WorkerCount=0; ActiveWorkers=0; Processed=0; Verified=0; VerificationFailed=0; Unmatched=0; Failed=0; OutputFolder=$(if($DryRun){''}else{$OutputFolder}) } -Force
        return
    }

    # Greedy byte-balanced partitioning. Large videos are assigned first to the
    # currently lightest worker, which reduces the common long-tail problem where
    # one worker is left copying the largest videos after the others finish.
    $buckets = @()
    $bucketBytes = New-Object 'Int64[]' $effectiveWorkers
    for ($i=0; $i -lt $effectiveWorkers; $i++) { $buckets += ,(New-Object System.Collections.ArrayList) }
    $sortedMedia = @($mediaFiles | Sort-Object -Property Length -Descending)
    foreach ($media in $sortedMedia) {
        $target = 0
        for ($i=1; $i -lt $effectiveWorkers; $i++) { if ($bucketBytes[$i] -lt $bucketBytes[$target]) { $target = $i } }
        [void]$buckets[$target].Add($media)
        $bucketBytes[$target] += [Int64]$media.Length
    }
    $sortedMedia = $null

    $workerRecords = New-Object System.Collections.Generic.List[object]
    for ($i=0; $i -lt $effectiveWorkers; $i++) {
        $id = $i + 1
        # v6.2.3: use a deliberately simple line-delimited path manifest. Windows file
        # names cannot contain CR/LF, so one absolute path per line is unambiguous and
        # avoids the Windows PowerShell 5.1 collection/JSON edge case that could make
        # workers with more than one assigned file terminate before creating a result.
        $manifestPath = Join-Path $jobRoot ('worker-' + $id + '-manifest.txt')
        $workerLog = Join-Path $jobRoot ('worker-' + $id + '.csv')
        $workerSummary = Join-Path $jobRoot ('worker-' + $id + '-summary.txt')
        $workerResult = Join-Path $jobRoot ('worker-' + $id + '-result.json')
        $workerProgress = Join-Path $jobRoot ('worker-' + $id + '-progress.json')
        $wrapperPath = Join-Path $jobRoot ('worker-' + $id + '-run.ps1')

        $manifestLines = New-Object System.Collections.Generic.List[string]
        foreach ($media in $buckets[$i]) { [void]$manifestLines.Add($media.FullName) }
        [IO.File]::WriteAllLines($manifestPath, [string[]]$manifestLines, [System.Text.UTF8Encoding]::new($true))
        New-WorkerWrapper -Path $wrapperPath -WorkerScript $workerScript -WorkerId $id -ManifestPath $manifestPath -LogPath $workerLog -SummaryPath $workerSummary -ResultPath $workerResult -WorkerProgressPath $workerProgress
        $record = Start-WorkerProcess -WorkerId $id -WrapperPath $wrapperPath -ProgressPath $workerProgress -ResultPath $workerResult -LogPath $workerLog -SummaryPath $workerSummary -ManifestPath $manifestPath -AssignedBytes $bucketBytes[$i] -AssignedFiles $buckets[$i].Count -IsRetry $false
        $workerRecords.Add($record)
    }

    Enable-SystemAwake
    Write-Host ('Pool:       ' + $effectiveWorkers + ' worker(s), one persistent ExifTool session per worker')
    Write-Host ('Media:      ' + $total + ' files, ' + (Format-ByteSize $totalMediaBytes))
    Write-Host ''

    $lastConsole = [DateTime]::MinValue
    while ($true) {
        $active = 0
        $sumIndex = 0
        $activeNames = New-Object System.Collections.Generic.List[string]
        foreach ($record in $workerRecords) {
            try { if (-not $record.Process.HasExited) { $active++ } } catch { }
            $state = Read-JsonFileSafe -Path $record.ProgressPath
            if ($null -ne $state) {
                if ($state.PSObject.Properties['Index']) { $sumIndex += [int]$state.Index }
                if ($activeNames.Count -lt 2 -and $state.PSObject.Properties['Name'] -and -not [string]::IsNullOrWhiteSpace([string]$state.Name)) { $activeNames.Add([string]$state.Name) }
            }
        }
        $sumIndex = [Math]::Min($total, $sumIndex)
        $elapsed = [Math]::Max(0.001, $runStopwatch.Elapsed.TotalSeconds)
        $liveRate = $sumIndex / $elapsed
        $message = ('{0} / {1} files; {2} / {3} workers active; {4:0.00} files/sec' -f $sumIndex, $total, $active, $effectiveWorkers, $liveRate)
        if ($activeNames.Count -gt 0) { $message += ('; ' + ($activeNames -join ' | ')) }
        Write-ProgressState -Stage 'ParallelProcessing' -Index $sumIndex -Total $total -Message $message -Extra @{ WorkerCount=$effectiveWorkers; ActiveWorkers=$active; FilesPerSecond=[Math]::Round($liveRate,2); LogPath=$logPath; OutputFolder=$(if($DryRun){''}else{$OutputFolder}) }
        $now = [DateTime]::UtcNow
        if (($now - $lastConsole).TotalSeconds -ge 1) {
            Write-Progress -Activity 'Restoring Google Photos metadata (parallel)' -Status $message -PercentComplete ([int](($sumIndex/[double]$total)*100))
            $lastConsole = $now
        }
        if ($active -eq 0) { break }
        if (Test-CancelRequested) { $cancelled = $true }
        Start-Sleep -Milliseconds 200
    }
    Write-Progress -Activity 'Restoring Google Photos metadata (parallel)' -Completed

    # Collect initial results and retry only true worker-process failures once in
    # direct ExifTool mode. Per-file metadata failures are already contained/logged
    # by the worker and do not trigger a whole-bucket replay.
    $finalRecords = New-Object System.Collections.Generic.List[object]
    foreach ($record in $workerRecords) {
        $result = Get-WorkerResultOrFailure -Record $record
        $needsRetry = (-not [bool]$result.Success)
        if ($needsRetry -and -not $cancelled -and -not (Test-CancelRequested)) {
            $retryCount++
            $retryLog = Join-Path $jobRoot ('worker-' + $record.WorkerId + '-retry.csv')
            $retrySummary = Join-Path $jobRoot ('worker-' + $record.WorkerId + '-retry-summary.txt')
            $retryResult = Join-Path $jobRoot ('worker-' + $record.WorkerId + '-retry-result.json')
            $retryProgress = Join-Path $jobRoot ('worker-' + $record.WorkerId + '-retry-progress.json')
            $retryWrapper = Join-Path $jobRoot ('worker-' + $record.WorkerId + '-retry.ps1')
            New-WorkerWrapper -Path $retryWrapper -WorkerScript $workerScript -WorkerId $record.WorkerId -ManifestPath $record.ManifestPath -LogPath $retryLog -SummaryPath $retrySummary -ResultPath $retryResult -WorkerProgressPath $retryProgress -RetryDirect
            Write-Host ('Worker ' + $record.WorkerId + ' exited unexpectedly; retrying its bucket once in direct compatibility mode.') -ForegroundColor Yellow
            $retryRecord = Start-WorkerProcess -WorkerId $record.WorkerId -WrapperPath $retryWrapper -ProgressPath $retryProgress -ResultPath $retryResult -LogPath $retryLog -SummaryPath $retrySummary -ManifestPath $record.ManifestPath -AssignedBytes $record.AssignedBytes -AssignedFiles $record.AssignedFiles -IsRetry $true
            while (-not $retryRecord.Process.HasExited) {
                if (Test-CancelRequested) { $cancelled = $true }
                $state = Read-JsonFileSafe -Path $retryProgress
                $idx = 0
                if ($state -and $state.PSObject.Properties['Index']) { $idx = [int]$state.Index }
                Write-ProgressState -Stage 'ParallelProcessing' -Index $idx -Total $total -Message ('Retrying worker ' + $record.WorkerId + ' in compatibility mode.') -Extra @{ WorkerCount=$effectiveWorkers; ActiveWorkers=1; FilesPerSecond=0 }
                Start-Sleep -Milliseconds 250
            }
            $result = Get-WorkerResultOrFailure -Record $retryRecord
            $retryRecord | Add-Member -NotePropertyName FinalResult -NotePropertyValue $result -Force
            $finalRecords.Add($retryRecord)
        }
        else {
            $record | Add-Member -NotePropertyName FinalResult -NotePropertyValue $result -Force
            $finalRecords.Add($record)
        }
    }

    if (-not $cancelled -and (Test-CancelRequested)) { $cancelled = $true }

    $finalLogs = @($finalRecords | ForEach-Object { $_.LogPath })
    Merge-WorkerCsvLogs -Paths $finalLogs -Destination $logPath

    $processed=0; $verified=0; $verificationFailed=0; $unmatched=0; $failed=0; $typeMismatches=0
    $preservedLocal=0; $preservedGps=0; $googleGpsEdits=0; $workerScanErrors=0; $exifCommands=0; $attempted=0
    $engineModes = New-Object System.Collections.Generic.List[string]
    $workerFatalMessages = New-Object System.Collections.Generic.List[string]
    foreach ($record in $finalRecords) {
        $r = $record.FinalResult
        $processed += [int]$r.Processed
        $verified += [int]$r.Verified
        $verificationFailed += [int]$r.VerificationFailed
        $unmatched += [int]$r.Unmatched
        $failed += [int]$r.Failed
        $typeMismatches += [int]$r.TypeMismatches
        $preservedLocal += [int]$r.PreservedLocalTime
        $preservedGps += [int]$r.PreservedGpsPrecision
        $googleGpsEdits += [int]$r.GoogleGpsEdits
        $workerScanErrors += [int]$r.ScanErrors
        $exifCommands += [int]$r.ExifToolCommands
        $attempted += [int]$r.Attempted
        $engineModes.Add(('W' + $record.WorkerId + ':' + [string]$r.ExifToolMode))
        if (-not [bool]$r.Success -and -not [string]::IsNullOrWhiteSpace([string]$r.FatalMessage)) { $workerFatalMessages.Add(('Worker ' + $record.WorkerId + ': ' + [string]$r.FatalMessage)) }
        if ([bool]$r.Cancelled) { $cancelled = $true }
    }
    $scanErrorCount = @($scanErrors).Count + $workerScanErrors

    # A worker-process failure must never be reported as a successful folder run.
    # The old behavior could leave most of an output folder missing while the batch
    # coordinator still labelled the folder DONE. Count every assigned file and every
    # final worker result before publishing success.
    $workerPoolIncomplete = ($workerFatalMessages.Count -gt 0 -or $attempted -lt $total)
    $missingAttempts = [Math]::Max(0, ($total - $attempted))

    $runStopwatch.Stop()
    $elapsedSeconds = [Math]::Max(0.001, $runStopwatch.Elapsed.TotalSeconds)
    $rateCount = $(if ($cancelled) { $attempted } else { $total })
    $filesPerSecond = $rateCount / $elapsedSeconds

    $summaryLines = New-Object System.Collections.Generic.List[string]
    $summaryLines.Add('Google Takeout Metadata Fixer v6.2.3 - Run Summary')
    $summaryLines.Add('Completed: ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz'))
    $summaryLines.Add('Mode: ' + $(if($DryRun){'DRY RUN'}else{'CREATE FIXED COPY'}))
    $summaryLines.Add('Cancelled by user: ' + $(if($cancelled){'YES'}else{'NO'}))
    $summaryLines.Add('Input: ' + $InputFolder)
    $summaryLines.Add('Output: ' + $(if($DryRun){'(none - dry run)'}else{$OutputFolder}))
    $summaryLines.Add('Media files scanned: ' + $total)
    $summaryLines.Add('Signature-detected unusual extensions: ' + $unknownExtensionMediaDetected)
    $summaryLines.Add('Media bytes scanned: ' + $totalMediaBytes + ' (' + (Format-ByteSize $totalMediaBytes) + ')')
    $summaryLines.Add('Processed/matched: ' + $processed)
    $summaryLines.Add('Verified timestamps: ' + $verified)
    $summaryLines.Add('Verification failures: ' + $verificationFailed)
    $summaryLines.Add('Copied without sidecar: ' + $unmatched)
    $summaryLines.Add('Type mismatches handled: ' + $typeMismatches)
    $summaryLines.Add('Local EXIF times preserved: ' + $preservedLocal)
    $summaryLines.Add('Precise embedded GPS preserved: ' + $preservedGps)
    $summaryLines.Add('Google GPS edits applied: ' + $googleGpsEdits)
    $summaryLines.Add('Failed/warned: ' + $failed)
    $summaryLines.Add('Scan errors/skips: ' + $scanErrorCount)
    $summaryLines.Add('Workers requested: ' + $WorkerCount)
    $summaryLines.Add('Workers used: ' + $effectiveWorkers)
    $summaryLines.Add('Worker fallback retries: ' + $retryCount)
    $summaryLines.Add('Files attempted by workers: ' + $attempted + ' / ' + $total)
    $summaryLines.Add('Completeness check: ' + $(if($workerPoolIncomplete){'FAIL'}else{'PASS'}))
    if ($missingAttempts -gt 0) { $summaryLines.Add('Files not attempted because of worker failure: ' + $missingAttempts) }
    $summaryLines.Add('Elapsed: ' + $runStopwatch.Elapsed.ToString())
    $summaryLines.Add('Average rate: ' + $filesPerSecond.ToString('0.00',[Globalization.CultureInfo]::InvariantCulture) + ' files/sec')
    $summaryLines.Add('ExifTool engine: parallel worker pool')
    $summaryLines.Add('Worker engines: ' + ($engineModes -join '; '))
    $summaryLines.Add('ExifTool commands: ' + $exifCommands)
    $summaryLines.Add('Verification: ' + $(if($SkipVerification){'disabled by user'}else{'enabled'}))
    $summaryLines.Add('CSV audit log: ' + $logPath)
    if ($workerFatalMessages.Count -gt 0) { $summaryLines.Add('Worker fatal messages: ' + ($workerFatalMessages -join ' | ')) }
    $summaryLines | Set-Content -LiteralPath $summaryPath -Encoding UTF8

    $stage = $(if($cancelled){'Cancelled'}elseif($workerPoolIncomplete){'Failed'}else{'Done'})
    $message = $(if($cancelled){
        'Cancelled safely after active files completed.'
    }elseif($workerPoolIncomplete){
        ('Worker pool incomplete: attempted ' + $attempted + ' of ' + $total + ' files. The folder is NOT complete; review the summary and rerun after fixing the worker error.')
    }else{
        'Finished parallel metadata pass; all scanned media files were attempted.'
    })
    Write-ProgressState -Stage $stage -Index $attempted -Total $total -Message $message -Extra @{
        Cancelled=$cancelled; Processed=$processed; Verified=$verified; VerificationFailed=$verificationFailed;
        TypeMismatches=$typeMismatches; UnknownExtensionMedia=$unknownExtensionMediaDetected; Unmatched=$unmatched;
        PreservedLocalTime=$preservedLocal; PreservedGpsPrecision=$preservedGps; GoogleGpsEdits=$googleGpsEdits;
        Failed=$failed; ScanErrors=$scanErrorCount; LogPath=$logPath; SummaryPath=$summaryPath;
        OutputFolder=$(if($DryRun){''}else{$OutputFolder}); ElapsedSeconds=[Math]::Round($elapsedSeconds,1);
        FilesPerSecond=[Math]::Round($filesPerSecond,2); WorkerCount=$effectiveWorkers; ActiveWorkers=0;
        ExifToolMode='parallel worker pool'; ExifToolCommands=$exifCommands; WorkerRetries=$retryCount;
        Attempted=$attempted; CompletenessCheck=$(if($workerPoolIncomplete){'FAIL'}else{'PASS'}); MissingAttempts=$missingAttempts
    } -Force

    Write-Host ''
    if ($cancelled) { Write-Host 'Cancelled safely.' -ForegroundColor Yellow } else { Write-Host 'Finished.' -ForegroundColor Green }
    Write-Host ('Workers used:              ' + $effectiveWorkers)
    Write-Host ('Processed/matched:         ' + $processed)
    Write-Host ('Verified timestamps:       ' + $verified)
    Write-Host ('Verification failures:     ' + $verificationFailed)
    Write-Host ('Copied without sidecar:    ' + $unmatched)
    Write-Host ('Failed/warned:              ' + $failed)
    Write-Host ('Scan errors/skips:          ' + $scanErrorCount)
    Write-Host ('Elapsed:                    ' + $runStopwatch.Elapsed.ToString())
    Write-Host ('Average rate:               ' + $filesPerSecond.ToString('0.00',[Globalization.CultureInfo]::InvariantCulture) + ' files/sec')
    Write-Host ('ExifTool commands:          ' + $exifCommands)
    Write-Host ('Worker fallback retries:    ' + $retryCount)
    Write-Host ('Files attempted:            ' + $attempted + ' / ' + $total)
    Write-Host ('Completeness check:         ' + $(if($workerPoolIncomplete){'FAIL'}else{'PASS'}))
    Write-Host ('Log:                        ' + $logPath)
    Write-Host ('Run summary:                ' + $summaryPath)
    if (-not $DryRun) { Write-Host ('Fixed library:              ' + $OutputFolder) }
    Write-Host ''

    if ($workerPoolIncomplete -and -not $cancelled) {
        # Set a fatal status only after the CSV, summary and final progress state have
        # been written. This lets the batch coordinator mark the folder FAILED while
        # preserving useful diagnostics for the user.
        $fatalMessage = ('Parallel worker pool was incomplete: attempted ' + $attempted + ' of ' + $total + ' files.')
    }
}
catch {
    $fatalMessage = $_.Exception.Message
    Write-ProgressState -Stage 'Failed' -Message $fatalMessage -Extra @{ LogPath=$logPath; OutputFolder=$(if($DryRun){''}else{$OutputFolder}) } -Force
}
finally {
    try { Disable-SystemAwake } catch { }
    # If the orchestrator itself fails, do not orphan child PowerShell/ExifTool trees.
    if (-not [string]::IsNullOrWhiteSpace($fatalMessage)) {
        foreach ($record in $script:ChildProcesses) { try { Stop-ProcessTree -Process $record.Process } catch { } }
    }
    if (-not [string]::IsNullOrWhiteSpace($jobRoot) -and (Test-Path -LiteralPath $jobRoot -PathType Container)) {
        try { Remove-Item -LiteralPath $jobRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }
    }
}

if (-not [string]::IsNullOrWhiteSpace($fatalMessage)) {
    Write-Error -Message $fatalMessage -ErrorAction Continue
    exit 1
}
