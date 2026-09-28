<#
.SYNOPSIS
  Internal v6 worker engine: restore Google Photos Takeout JSON metadata into COPIES using ExifTool.

.DESCRIPTION
  Version 6.0 worker engine - designed for Windows PowerShell 5.1+ and ExifTool 13.59+.

  Full-library optimizations in v5:
  - Keeps one ExifTool process open for the run instead of launching it repeatedly.
  - Streams the CSV audit log incrementally instead of retaining every row in memory.
  - Uses bounded sidecar-index caches and avoids retaining parsed JSON objects for the whole run.
  - Throttles GUI progress-file writes to reduce unnecessary disk I/O.
  - Preserves higher-precision embedded GPS when it agrees with Google location data.
  - Performs a disk-space preflight and keeps Windows awake during long copy runs.

  Safety / behavior:
  - Never modifies the original Takeout media.
  - Copies ALL supported media to a separate output tree, even if no JSON is found.
  - Signature-sniffs files with unrecognized extensions so real media such as Pixel .MP
    Motion Photo components are retained instead of being silently skipped.
  - Uses photoTakenTime first, creationTime only as a fallback.
  - Uses Google Photos geoData first, then geoDataExif as a fallback.
  - Restores GPS, captions/descriptions, people tags and favorite rating when possible.
  - Restores QuickTime video date tags for MP4/MOV/M4V/3GP/3G2 using raw UTC values,
    avoiding the host-computer timezone double-conversion bug fixed in v4.
  - Verifies written capture timestamps by reading the output back with ExifTool.
  - Detects common file-content/filename mismatches (for example JPEG bytes named .HEIC),
    processes them through a temporary correctly-typed alias, and preserves the original filename.
  - Preserves an existing local EXIF time when it differs from Google's UTC timestamp by
    a plausible 15-minute timezone offset; writes/infer the offset instead of flattening
    the local wall-clock time to UTC.
  - Use -ForceJsonTime if you intentionally changed dates in Google Photos and want the
    JSON timestamp to overwrite existing embedded capture times.
  - Writes a CSV audit log with expected/stored UTC timestamps and verification status.

.EXAMPLE
  .\Fix-GoogleTakeoutMetadata-v6.2.3-Worker.ps1 -InputFolder "D:\Takeout\Google Photos" -DryRun

.EXAMPLE
  .\Fix-GoogleTakeoutMetadata-v6.2.3-Worker.ps1 -InputFolder "D:\Takeout\Google Photos"

.EXAMPLE
  .\Fix-GoogleTakeoutMetadata-v6.2.3-Worker.ps1 -InputFolder "D:\Takeout\Google Photos" -ForceJsonTime
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

    [ValidateRange(1, 500)]
    [double]$GpsAgreementMeters = 30,

    [string]$ProgressFile,

    [string]$CancelFile,

    # Internal worker-pool parameters. The v6 orchestrator writes a JSON manifest
    # containing the exact media paths assigned to this worker. These switches are
    # intentionally undocumented for normal users; use Fix-GoogleTakeoutMetadata-v6.2.3.ps1.
    [string]$FileListPath,

    [string]$WorkerLogPath,

    [string]$WorkerSummaryPath,

    [string]$WorkerResultPath,

    [ValidateRange(0, 64)]
    [int]$WorkerId = 0,

    [switch]$InternalWorker
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:JsonIndexCache = @{}
$script:JsonIndexOrder = New-Object 'System.Collections.Generic.Queue[string]'
$script:MaxJsonFolderIndexes = 128
$script:OutputDirectoryCache = @{}
$script:ExifToolSession = $null
$script:ExifToolCommandCount = 0
$script:ExifToolMode = 'direct'
$script:LastProgressWriteUtc = [DateTime]::MinValue
$script:AuditWriter = $null
$script:AuditRowCount = 0
$script:KeepAwakeEnabled = $false


# Build media-extension lookup tables once. These are consulted for every file in a
# full-library run, so avoid recreating arrays and doing linear -contains scans in
# the write/verification hot path.
$script:PhotoExtensions = @('.jpg', '.jpeg', '.heic', '.heif', '.png', '.webp', '.gif', '.tif', '.tiff', '.bmp', '.dng', '.cr2', '.cr3', '.nef', '.exr', '.sr2', '.orf', '.raw', '.360', '.3fr', '.jp2', '.eps', '.exif', '.ico', '.arw', '.avif', '.jxl')
$script:QuickTimeExtensions = @('.mp4', '.mov', '.m4v', '.3gp', '.3g2')
$script:OtherVideoExtensions = @('.mpg', '.mpeg', '.wmv', '.tod', '.mts', '.mod', '.mmv', '.mkv', '.m2ts', '.m2t', '.divx', '.avi', '.asf', '.webm')
$script:PhotoExtensionLookup = @{}
$script:QuickTimeExtensionLookup = @{}
$script:MediaExtensionLookup = @{}
foreach ($extension in $script:PhotoExtensions) {
    $script:PhotoExtensionLookup[$extension] = $true
    $script:MediaExtensionLookup[$extension] = $true
}
foreach ($extension in $script:QuickTimeExtensions) {
    $script:QuickTimeExtensionLookup[$extension] = $true
    $script:MediaExtensionLookup[$extension] = $true
}
foreach ($extension in $script:OtherVideoExtensions) { $script:MediaExtensionLookup[$extension] = $true }

# Files with an unrecognized extension are signature-sniffed so Google/phone oddities
# such as Pixel Motion Photo .MP files are not silently omitted. Avoid opening common
# sidecars/documents/archives/executables that cannot be photo/video payloads. This
# keeps the full-library scan fast even when every media item has a JSON companion.
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
        [hashtable]$Extra = $null
    )

    if ([string]::IsNullOrWhiteSpace($ProgressFile)) { return }

    # The GUI polls at 250 ms. Avoid rewriting a JSON file multiple times per media item
    # on fast storage; final/failure/scanning states are always written immediately.
    $nowUtc = [DateTime]::UtcNow
    $urgent = ($Stage -in @('Scanning', 'Preflight', 'Done', 'Cancelled', 'Failed'))
    if (-not $urgent -and (($nowUtc - $script:LastProgressWriteUtc).TotalMilliseconds -lt 150)) { return }
    $script:LastProgressWriteUtc = $nowUtc

    try {
        $percent = 0
        if ($Total -gt 0) {
            $percent = [int][Math]::Min(100, [Math]::Max(0, (($Index / [double]$Total) * 100)))
        }

        $state = [ordered]@{
            Stage = $Stage
            Index = $Index
            Total = $Total
            Percent = $percent
            Name = $Name
            Message = $Message
        }
        if ($Extra) {
            foreach ($key in $Extra.Keys) { $state[$key] = $Extra[$key] }
        }

        $progressDir = Split-Path -Parent $ProgressFile
        if ($progressDir -and -not (Test-Path -LiteralPath $progressDir -PathType Container)) {
            New-Item -ItemType Directory -Path $progressDir -Force | Out-Null
        }
        ($state | ConvertTo-Json -Compress -Depth 4) | Set-Content -LiteralPath $ProgressFile -Encoding UTF8
    }
    catch {
        # Progress reporting must never stop the metadata job.
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


function Start-ExifToolSession {
    param([string]$ExifTool)

    if ($DisablePersistentExifTool) {
        $script:ExifToolMode = 'direct compatibility mode'
        return $false
    }

    try {
        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = $ExifTool
        $startInfo.Arguments = '-charset filename=UTF8 -stay_open True -@ -'
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        $startInfo.RedirectStandardInput = $true
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true
        if ($startInfo.PSObject.Properties['StandardOutputEncoding']) {
            $startInfo.StandardOutputEncoding = [System.Text.UTF8Encoding]::new($false)
        }
        if ($startInfo.PSObject.Properties['StandardErrorEncoding']) {
            $startInfo.StandardErrorEncoding = [System.Text.UTF8Encoding]::new($false)
        }

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $startInfo
        if (-not $process.Start()) { throw 'ExifTool process did not start.' }

        # Write UTF-8 arguments directly to the stdin pipe. This is important because
        # ExifTool arg files do not receive normal Windows command-line recoding.
        $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
        $writer = [System.IO.StreamWriter]::new($process.StandardInput.BaseStream, $utf8NoBom)
        $writer.AutoFlush = $true
        $writer.NewLine = "`n"

        # Drain stderr continuously so warnings can never fill the pipe and stall a run.
        $stderrTask = $process.StandardError.ReadToEndAsync()

        $script:ExifToolSession = [PSCustomObject]@{
            Process = $process
            Writer = $writer
            Counter = 0
            StderrTask = $stderrTask
            Path = $ExifTool
        }
        $script:ExifToolMode = 'persistent stay_open'

        $test = Invoke-ExifToolCommand -ExifTool $ExifTool -Arguments @('-ver') -SuppressErrors
        if ($test.ExitCode -ne 0 -or $test.OutputLines.Count -eq 0) {
            throw 'Persistent ExifTool self-test failed.'
        }
        return $true
    }
    catch {
        try { Stop-ExifToolSession | Out-Null } catch { }
        $script:ExifToolSession = $null
        $script:ExifToolMode = 'direct fallback'
        Write-Warning ('High-speed ExifTool session could not be started; using slower compatibility mode. ' + $_.Exception.Message)
        return $false
    }
}

function Stop-ExifToolSession {
    if ($null -eq $script:ExifToolSession) { return '' }

    $session = $script:ExifToolSession
    $script:ExifToolSession = $null
    $stderrText = ''
    try {
        if ($session.Process -and -not $session.Process.HasExited) {
            try {
                $session.Writer.WriteLine('-stay_open')
                $session.Writer.WriteLine('False')
                $session.Writer.Flush()
                $session.Writer.Close()
            }
            catch { }

            if (-not $session.Process.WaitForExit(5000)) {
                try { $session.Process.Kill() } catch { }
                try { $session.Process.WaitForExit(2000) | Out-Null } catch { }
            }
        }
        if ($session.StderrTask) {
            try { $stderrText = [string]$session.StderrTask.Result } catch { }
        }
    }
    finally {
        try { $session.Process.Dispose() } catch { }
    }
    return $stderrText
}

function Invoke-ExifToolCommand {
    param(
        [string]$ExifTool,
        [string[]]$Arguments,
        [switch]$SuppressErrors
    )

    $script:ExifToolCommandCount++

    # stay_open uses one argument per line. Multiline metadata values are uncommon,
    # but are safer through the normal Windows command-line path than through an
    # argfile protocol, so fall back for just that command.
    $requiresDirect = $false
    foreach ($arg in $Arguments) {
        if ($null -ne $arg -and ([string]$arg).IndexOfAny(@([char]13, [char]10)) -ge 0) {
            $requiresDirect = $true
            break
        }
    }

    $session = $script:ExifToolSession
    if (-not $requiresDirect -and $null -ne $session -and $session.Process -and -not $session.Process.HasExited) {
        try {
            $session.Counter = [int]$session.Counter + 1
            $id = [int]$session.Counter
            $statusPrefix = '__GTMF_STATUS_' + $id + '='
            $ready = '{ready' + $id + '}'

            foreach ($arg in $Arguments) {
                $session.Writer.WriteLine([string]$arg)
            }
            $session.Writer.WriteLine('-echo3')
            $session.Writer.WriteLine($statusPrefix + '${status}')
            $session.Writer.WriteLine('-execute' + $id)
            $session.Writer.Flush()

            $output = New-Object System.Collections.Generic.List[string]
            $status = $null
            while ($true) {
                $line = $session.Process.StandardOutput.ReadLine()
                if ($null -eq $line) {
                    throw 'ExifTool closed its output pipe unexpectedly.'
                }
                if ($line -eq $ready) { break }
                if ($line.StartsWith($statusPrefix, [StringComparison]::Ordinal)) {
                    $parsedStatus = 1
                    if ([int]::TryParse($line.Substring($statusPrefix.Length), [ref]$parsedStatus)) {
                        $status = $parsedStatus
                    }
                    continue
                }
                $output.Add($line)
            }
            if ($null -eq $status) { $status = 1 }
            return [PSCustomObject]@{ ExitCode = [int]$status; OutputLines = [string[]]$output; Mode = 'persistent' }
        }
        catch {
            # If the persistent process unexpectedly fails, shut it down and transparently
            # fall back to direct execution for the current and subsequent commands.
            try { Stop-ExifToolSession | Out-Null } catch { }
            $script:ExifToolMode = 'direct fallback'
        }
    }

    try {
        if ($SuppressErrors) {
            $raw = & $ExifTool @Arguments 2>$null
        }
        else {
            $raw = & $ExifTool @Arguments 2>&1
        }
        $exitCode = $LASTEXITCODE
        $lines = New-Object System.Collections.Generic.List[string]
        foreach ($item in @($raw)) {
            if ($null -ne $item) { $lines.Add([string]$item) }
        }
        return [PSCustomObject]@{ ExitCode = [int]$exitCode; OutputLines = [string[]]$lines; Mode = 'direct' }
    }
    catch {
        return [PSCustomObject]@{ ExitCode = 1; OutputLines = @($_.Exception.Message); Mode = 'direct' }
    }
}

function Ensure-OutputDirectory {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    $key = $Path.ToLowerInvariant()
    if ($script:OutputDirectoryCache.ContainsKey($key)) { return }
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
    $script:OutputDirectoryCache[$key] = $true
}

function ConvertTo-CsvCell {
    param($Value)
    if ($null -eq $Value) { return '""' }
    $text = [Convert]::ToString($Value, [Globalization.CultureInfo]::InvariantCulture)
    return '"' + $text.Replace('"', '""') + '"'
}

function Open-AuditLog {
    param([string]$Path, [string[]]$Columns)

    $dir = Split-Path -Parent $Path
    Ensure-OutputDirectory -Path $dir
    $encoding = [System.Text.UTF8Encoding]::new($true)
    $writer = [System.IO.StreamWriter]::new($Path, $false, $encoding, 65536)
    $writer.WriteLine((($Columns | ForEach-Object { ConvertTo-CsvCell $_ }) -join ','))
    $writer.Flush()
    $script:AuditWriter = $writer
    $script:AuditRowCount = 0
}

function Write-AuditRow {
    param($Row, [string[]]$Columns, [switch]$ForceFlush)

    if ($null -eq $script:AuditWriter) { return }
    $cells = New-Object System.Collections.Generic.List[string]
    foreach ($column in $Columns) {
        $value = ''
        if ($null -ne $Row -and $Row.PSObject.Properties[$column]) { $value = $Row.$column }
        $cells.Add((ConvertTo-CsvCell $value))
    }
    $script:AuditWriter.WriteLine(($cells -join ','))
    $script:AuditRowCount++
    if ($ForceFlush -or (($script:AuditRowCount % 25) -eq 0)) { $script:AuditWriter.Flush() }
}

function Close-AuditLog {
    if ($null -ne $script:AuditWriter) {
        try { $script:AuditWriter.Flush() } catch { }
        try { $script:AuditWriter.Close() } catch { }
        try { $script:AuditWriter.Dispose() } catch { }
        $script:AuditWriter = $null
    }
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

function New-TemporaryMediaAlias {
    param(
        [string]$SourcePath,
        [string]$Extension
    )

    $tempRoot = Join-Path ([IO.Path]::GetTempPath()) 'GoogleTakeoutMetadataFixer'
    if (-not (Test-Path -LiteralPath $tempRoot -PathType Container)) {
        New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    }

    if ([string]::IsNullOrWhiteSpace($Extension) -or -not $Extension.StartsWith('.')) {
        $Extension = '.bin'
    }

    $tempPath = Join-Path $tempRoot (('media-' + [Guid]::NewGuid().ToString('N')) + $Extension)
    Copy-Item -LiteralPath $SourcePath -Destination $tempPath -Force
    return $tempPath
}

function Get-OutputWorkingPath {
    param(
        [string]$Destination,
        [string]$EffectiveExtension,
        [bool]$NeedsAlias
    )

    if (-not $NeedsAlias) { return $Destination }
    $directory = Split-Path -Parent $Destination
    return (Join-Path $directory (('.gtmf-' + [Guid]::NewGuid().ToString('N')) + $EffectiveExtension))
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

function Get-DefaultOutputFolder {
    param([string]$InputPath)

    $normalized = Get-NormalizedDirectoryPath -Path $InputPath
    $parent = Split-Path -Parent $normalized
    $leaf = Split-Path -Leaf $normalized
    if (-not $parent) { $parent = [IO.Path]::GetPathRoot($normalized) }
    if (-not $leaf) { $leaf = 'Google Photos' }
    return (Join-Path $parent ($leaf + ' - Fixed'))
}

function Get-RelativePathSimple {
    param([string]$BasePath, [string]$FullPath)

    $base = Get-NormalizedDirectoryPath -Path $BasePath
    $full = [IO.Path]::GetFullPath($FullPath)

    if (-not $full.StartsWith($base, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path '$full' is not below '$base'."
    }

    return $full.Substring($base.Length).TrimStart('\', '/')
}

function Read-GoogleJson {
    param([string]$JsonPath)

    # Do not retain every parsed sidecar object for a full-library run. A large
    # Takeout can contain hundreds of thousands of JSON files; keeping all of them
    # alive causes avoidable memory growth. Sidecars are small and normally parsed
    # once, so reparsing the rare fallback match is cheaper than unbounded RAM use.
    try {
        return (Get-Content -LiteralPath $JsonPath -Raw -Encoding UTF8 | ConvertFrom-Json)
    }
    catch {
        return $null
    }
}

function Add-JsonIndexValue {
    param([hashtable]$Map, [string]$Key, [string]$Value)
    if ([string]::IsNullOrWhiteSpace($Key)) { return }
    $normalized = $Key.ToLowerInvariant()
    if (-not $Map.ContainsKey($normalized)) {
        $Map[$normalized] = New-Object System.Collections.Generic.List[string]
    }
    $Map[$normalized].Add($Value)
}

function Get-JsonFolderIndex {
    param([string]$Folder)

    $key = $Folder.ToLowerInvariant()
    if ($script:JsonIndexCache.ContainsKey($key)) { return $script:JsonIndexCache[$key] }

    # Index each folder once, but keep only a bounded number of folder indexes.
    # Store paths rather than FileInfo objects to reduce memory overhead.
    $jsonItems = @(Get-ChildItem -LiteralPath $Folder -File -Filter '*.json' -ErrorAction SilentlyContinue)
    $paths = New-Object System.Collections.Generic.List[string]
    $byName = @{}
    $filenameBaseMap = @{}
    foreach ($file in $jsonItems) {
        $paths.Add($file.FullName)
        $lowerName = $file.Name.ToLowerInvariant()
        $byName[$lowerName] = $file.FullName

        # Build a cheap basename index from the sidecar filename itself. This resolves
        # most Live Photo companion videos without parsing every JSON file in a folder.
        $stem = $file.Name.Substring(0, $file.Name.Length - 5) # remove .json
        $mediaLikeName = [regex]::Replace($stem, '\.supplemental-metadata(?:\(\d+\))?$', '', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
        $mediaLikeExtension = [IO.Path]::GetExtension($mediaLikeName).ToLowerInvariant()
        $candidateBase = [IO.Path]::GetFileNameWithoutExtension($mediaLikeName)
        if ($script:MediaExtensionLookup.ContainsKey($mediaLikeExtension) -and -not [string]::IsNullOrWhiteSpace($candidateBase)) {
            Add-JsonIndexValue -Map $filenameBaseMap -Key $candidateBase -Value $file.FullName
        }
    }
    $jsonItems = $null

    $index = [PSCustomObject]@{
        Files = $paths
        ByName = $byName
        FilenameBaseMap = $filenameBaseMap
        Stems = $null
        TitleMap = @{}
        BaseMap = @{}
        FallbackBuilt = $false
    }
    $script:JsonIndexCache[$key] = $index
    $script:JsonIndexOrder.Enqueue($key)
    while ($script:JsonIndexOrder.Count -gt $script:MaxJsonFolderIndexes) {
        $oldKey = $script:JsonIndexOrder.Dequeue()
        if ($oldKey -ne $key) { [void]$script:JsonIndexCache.Remove($oldKey) }
    }
    return $index
}

function Ensure-JsonStemIndex {
    param($Index)
    if ($null -ne $Index.Stems) { return }

    $stems = New-Object System.Collections.Generic.List[object]
    foreach ($jsonPath in $Index.Files) {
        $name = [IO.Path]::GetFileName([string]$jsonPath)
        if ($name.Length -gt 5) {
            $stems.Add([PSCustomObject]@{ Path = [string]$jsonPath; Stem = $name.Substring(0, $name.Length - 5).ToLowerInvariant() })
        }
    }
    $Index.Stems = $stems
}

function Ensure-JsonFallbackIndex {
    param($Index)
    if ($Index.FallbackBuilt) { return }

    foreach ($jsonPath in $Index.Files) {
        $data = Read-GoogleJson -JsonPath ([string]$jsonPath)
        if ($null -eq $data -or $null -eq $data.PSObject.Properties['title']) { continue }
        $title = [string]$data.title
        if ([string]::IsNullOrWhiteSpace($title)) { continue }
        Add-JsonIndexValue -Map $Index.TitleMap -Key $title -Value ([string]$jsonPath)
        $titleBase = [IO.Path]::GetFileNameWithoutExtension($title)
        Add-JsonIndexValue -Map $Index.BaseMap -Key $titleBase -Value ([string]$jsonPath)
    }
    $Index.FallbackBuilt = $true
}

function Convert-ToNullableDouble {
    param($Value)

    if ($null -eq $Value) { return $null }
    $number = 0.0
    $text = [Convert]::ToString($Value, [Globalization.CultureInfo]::InvariantCulture)
    if ([double]::TryParse($text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $number
    }
    return $null
}


function Convert-GoogleGeoPoint {
    param($GeoObject, [string]$SourceName)

    if ($null -eq $GeoObject) { return $null }
    $lat = $null; $lon = $null; $alt = $null
    if ($GeoObject.PSObject.Properties['latitude'])  { $lat = Convert-ToNullableDouble $GeoObject.latitude }
    if ($GeoObject.PSObject.Properties['longitude']) { $lon = Convert-ToNullableDouble $GeoObject.longitude }
    if ($GeoObject.PSObject.Properties['altitude'])  { $alt = Convert-ToNullableDouble $GeoObject.altitude }
    if ($null -eq $lat -or $null -eq $lon) { return $null }
    if ($lat -lt -90 -or $lat -gt 90 -or $lon -lt -180 -or $lon -gt 180) { return $null }
    if ($lat -eq 0 -and $lon -eq 0) { return $null }
    return [PSCustomObject]@{ Latitude = [double]$lat; Longitude = [double]$lon; Altitude = $alt; Source = $SourceName }
}

function Get-MediaNameVariants {
    param([string]$Name)

    $variants = New-Object System.Collections.Generic.List[string]
    $variants.Add($Name)

    $ext = [IO.Path]::GetExtension($Name)
    $base = [IO.Path]::GetFileNameWithoutExtension($Name)

    if ($base -match '^(.*?)(?:-edited|_edited|-edit|_edit)$') {
        $variants.Add($Matches[1] + $ext)
    }

    if ($base -match '^(.*?)~\d+$') {
        $variants.Add($Matches[1] + $ext)
    }

    return @($variants | Select-Object -Unique)
}

function Find-Sidecar {
    param([System.IO.FileInfo]$Media)

    $dir = $Media.DirectoryName
    $name = $Media.Name
    $baseName = [IO.Path]::GetFileNameWithoutExtension($name)
    $index = Get-JsonFolderIndex -Folder $dir
    $candidateNames = New-Object System.Collections.Generic.List[string]
    $mediaVariants = @(Get-MediaNameVariants -Name $name)

    foreach ($variant in $mediaVariants) {
        $variantBase = [IO.Path]::GetFileNameWithoutExtension($variant)
        $candidateNames.Add($variant + '.supplemental-metadata.json')
        $candidateNames.Add($variant + '.json')
        $candidateNames.Add($variantBase + '.json')

        if ($variant -match '^(.*)\((\d+)\)(\.[^.]*)$') {
            $prefix = $Matches[1]
            $num = $Matches[2]
            $extension = $Matches[3]
            $originalName = $prefix + $extension
            $candidateNames.Add($originalName + '(' + $num + ').json')
            $candidateNames.Add($originalName + '.supplemental-metadata(' + $num + ').json')
            $candidateNames.Add($originalName + '(' + $num + ').supplemental-metadata.json')
        }
    }

    foreach ($candidateName in @($candidateNames | Select-Object -Unique)) {
        $key = $candidateName.ToLowerInvariant()
        if ($index.ByName.ContainsKey($key)) {
            return [PSCustomObject]@{ Path = [string]$index.ByName[$key]; Method = 'exact filename' }
        }
    }

    # A cheap sidecar-filename basename index resolves most Live Photo companion
    # videos without the much more expensive step of parsing all JSON in this folder.
    $baseKey = $baseName.ToLowerInvariant()
    if ($index.FilenameBaseMap.ContainsKey($baseKey)) {
        $filenameBaseMatches = @($index.FilenameBaseMap[$baseKey] | Select-Object -Unique)
        if ($filenameBaseMatches.Count -eq 1) {
            return [PSCustomObject]@{ Path = $filenameBaseMatches[0]; Method = 'sidecar filename basename / Live Photo' }
        }
    }

    # Truncated supplemental metadata names cannot be indexed by an exact key because
    # Google may cut them at varying lengths. Build the stem index lazily only when
    # inexpensive exact/basename matching has failed.
    Ensure-JsonStemIndex -Index $index
    $prefixMatches = New-Object System.Collections.Generic.List[string]
    foreach ($variant in $mediaVariants) {
        $expected = ($variant + '.supplemental-metadata').ToLowerInvariant()
        foreach ($entry in $index.Stems) {
            $stemForCompare = [string]$entry.Stem
            if ($variant -match '^(.*)\((\d+)\)(\.[^.]*)$') {
                $n = $Matches[2]
                if ($stemForCompare -match ('\(' + [regex]::Escape($n) + '\)$')) {
                    $stemForCompare = $stemForCompare.Substring(0, $stemForCompare.LastIndexOf('('))
                }
                $variantNoCounter = $Matches[1] + $Matches[3]
                $expected = ($variantNoCounter + '.supplemental-metadata').ToLowerInvariant()
            }
            if ($expected.StartsWith($stemForCompare, [StringComparison]::OrdinalIgnoreCase) -and
                $stemForCompare.Length -ge [Math]::Min(12, $expected.Length)) {
                $prefixMatches.Add([string]$entry.Path)
            }
        }
    }
    $prefixMatches = @($prefixMatches | Select-Object -Unique)
    if ($prefixMatches.Count -eq 1) {
        return [PSCustomObject]@{ Path = $prefixMatches[0]; Method = 'truncated supplemental prefix' }
    }

    # Last resort: parse JSON titles. This handles unusual duplicate/export naming,
    # but is intentionally delayed because parsing every JSON in a large folder is
    # much slower and more memory intensive than filename-only indexing.
    Ensure-JsonFallbackIndex -Index $index

    $exactTitleMatches = New-Object System.Collections.Generic.List[string]
    foreach ($variant in $mediaVariants) {
        $key = $variant.ToLowerInvariant()
        if ($index.TitleMap.ContainsKey($key)) {
            foreach ($path in $index.TitleMap[$key]) { $exactTitleMatches.Add($path) }
        }
    }
    $exactTitleMatches = @($exactTitleMatches | Select-Object -Unique)
    if ($exactTitleMatches.Count -eq 1) {
        return [PSCustomObject]@{ Path = $exactTitleMatches[0]; Method = 'JSON title' }
    }

    if ($index.BaseMap.ContainsKey($baseKey)) {
        $sameBaseMatches = @($index.BaseMap[$baseKey] | Select-Object -Unique)
        if ($sameBaseMatches.Count -eq 1) {
            return [PSCustomObject]@{ Path = $sameBaseMatches[0]; Method = 'JSON title basename / Live Photo' }
        }
    }

    return $null
}

function Get-GoogleMetadata {
    param([string]$JsonPath)

    $json = Read-GoogleJson -JsonPath $JsonPath
    if ($null -eq $json) { return $null }

    $timestamp = $null
    $timeSource = $null

    if ($json.PSObject.Properties['photoTakenTime'] -and $json.photoTakenTime -and
        $json.photoTakenTime.PSObject.Properties['timestamp']) {
        $timestamp = [string]$json.photoTakenTime.timestamp
        $timeSource = 'photoTakenTime'
    }

    if ((-not $timestamp -or $timestamp -notmatch '^\d+$' -or [Int64]$timestamp -le 0) -and
        $json.PSObject.Properties['creationTime'] -and $json.creationTime -and
        $json.creationTime.PSObject.Properties['timestamp']) {
        $timestamp = [string]$json.creationTime.timestamp
        $timeSource = 'creationTime'
    }

    $utc = $null
    if ($timestamp -and $timestamp -match '^\d+$' -and [Int64]$timestamp -gt 0) {
        try {
            $utc = [DateTimeOffset]::FromUnixTimeSeconds([Int64]$timestamp).UtcDateTime
        }
        catch {
            $utc = $null
        }
    }

    # Keep both Google location representations. geoData is Google's current state
    # (including user edits); geoDataExif represents the original embedded location.
    # The later GPS decision can therefore distinguish a real Google edit from simple
    # coordinate rounding and preserve more precise embedded coordinates when safe.
    $geoDataPoint = $null
    $geoExifPoint = $null
    if ($json.PSObject.Properties['geoData'] -and $json.geoData) {
        $geoDataPoint = Convert-GoogleGeoPoint -GeoObject $json.geoData -SourceName 'geoData'
    }
    if ($json.PSObject.Properties['geoDataExif'] -and $json.geoDataExif) {
        $geoExifPoint = Convert-GoogleGeoPoint -GeoObject $json.geoDataExif -SourceName 'geoDataExif'
    }

    $preferredGeo = $geoDataPoint
    if ($null -eq $preferredGeo) { $preferredGeo = $geoExifPoint }
    $lat = $null; $lon = $null; $alt = $null; $geoSource = ''
    if ($null -ne $preferredGeo) {
        $lat = $preferredGeo.Latitude
        $lon = $preferredGeo.Longitude
        $alt = $preferredGeo.Altitude
        $geoSource = $preferredGeo.Source
    }

    $description = ''
    if ($json.PSObject.Properties['description'] -and $null -ne $json.description) {
        $description = [string]$json.description
    }

    $title = ''
    if ($json.PSObject.Properties['title'] -and $null -ne $json.title) {
        $title = [string]$json.title
    }

    $favorite = $false
    if ($json.PSObject.Properties['favorited']) {
        if ($json.favorited -is [bool]) {
            $favorite = [bool]$json.favorited
        }
        else {
            $favorite = ([string]$json.favorited -match '^true$')
        }
    }

    $people = New-Object System.Collections.Generic.List[string]
    if ($json.PSObject.Properties['people'] -and $json.people) {
        foreach ($person in @($json.people)) {
            if ($person -and $person.PSObject.Properties['name']) {
                $personName = [string]$person.name
                if (-not [string]::IsNullOrWhiteSpace($personName)) {
                    $people.Add($personName)
                }
            }
        }
    }

    return [PSCustomObject]@{
        UtcTime = $utc
        TimeSource = $timeSource
        Latitude = $lat
        Longitude = $lon
        Altitude = $alt
        GeoSource = $geoSource
        GeoDataPoint = $geoDataPoint
        GeoDataExifPoint = $geoExifPoint
        Description = $description
        Title = $title
        Favorite = $favorite
        People = @($people | Select-Object -Unique)
    }
}

function Parse-ExifDate {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $formats = @('yyyy:MM:dd HH:mm:ss', 'yyyy:MM:dd HH:mm:ss.fff', 'yyyy-MM-ddTHH:mm:ss', 'yyyy-MM-ddTHH:mm:ssZ')
    foreach ($fmt in $formats) {
        $parsed = [DateTime]::MinValue
        if ([DateTime]::TryParseExact($Text, $fmt, [Globalization.CultureInfo]::InvariantCulture,
                [Globalization.DateTimeStyles]::AllowWhiteSpaces, [ref]$parsed)) {
            return [DateTime]::SpecifyKind($parsed, [DateTimeKind]::Unspecified)
        }
    }
    return $null
}

function Parse-OffsetMinutes {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    if ($Text -match '^([+-])(\d{2}):(\d{2})$') {
        $mins = ([int]$Matches[2] * 60) + [int]$Matches[3]
        if ($Matches[1] -eq '-') { $mins = -$mins }
        return $mins
    }
    return $null
}

function Format-Offset {
    param([int]$Minutes)

    $sign = if ($Minutes -lt 0) { '-' } else { '+' }
    $abs = [Math]::Abs($Minutes)
    $hours = [Math]::Floor($abs / 60)
    $mins = $abs % 60
    return ('{0}{1:00}:{2:00}' -f $sign, $hours, $mins)
}


function Test-ValidCoordinates {
    param($Latitude, $Longitude)
    if ($null -eq $Latitude -or $null -eq $Longitude) { return $false }
    try {
        $lat = [double]$Latitude; $lon = [double]$Longitude
        if ($lat -lt -90 -or $lat -gt 90 -or $lon -lt -180 -or $lon -gt 180) { return $false }
        if ($lat -eq 0 -and $lon -eq 0) { return $false }
        return $true
    }
    catch { return $false }
}

function Get-CoordinateDistanceMeters {
    param([double]$Latitude1, [double]$Longitude1, [double]$Latitude2, [double]$Longitude2)

    $radius = 6371008.8
    $toRad = [Math]::PI / 180.0
    $lat1 = $Latitude1 * $toRad
    $lat2 = $Latitude2 * $toRad
    $dLat = ($Latitude2 - $Latitude1) * $toRad
    $dLon = ($Longitude2 - $Longitude1) * $toRad
    $a = [Math]::Sin($dLat / 2.0) * [Math]::Sin($dLat / 2.0) +
         [Math]::Cos($lat1) * [Math]::Cos($lat2) * [Math]::Sin($dLon / 2.0) * [Math]::Sin($dLon / 2.0)
    $c = 2.0 * [Math]::Atan2([Math]::Sqrt($a), [Math]::Sqrt([Math]::Max(0.0, 1.0 - $a)))
    return $radius * $c
}

function Get-GpsDecision {
    param($Metadata, $Embedded)

    if (-not (Test-ValidCoordinates -Latitude $Metadata.Latitude -Longitude $Metadata.Longitude)) {
        return [PSCustomObject]@{
            ShouldWrite = $false; Latitude = $null; Longitude = $null; Altitude = $null
            Source = 'none'; Reason = 'no usable Google GPS; existing embedded location left unchanged'
            DistanceMeters = $null; GoogleEditDetected = $false
        }
    }

    $googleLat = [double]$Metadata.Latitude
    $googleLon = [double]$Metadata.Longitude
    $googleAlt = $Metadata.Altitude
    $googleEditDetected = $false
    $googleEditDistance = $null

    if ($null -ne $Metadata.GeoDataPoint -and $null -ne $Metadata.GeoDataExifPoint) {
        $googleEditDistance = Get-CoordinateDistanceMeters `
            -Latitude1 ([double]$Metadata.GeoDataPoint.Latitude) -Longitude1 ([double]$Metadata.GeoDataPoint.Longitude) `
            -Latitude2 ([double]$Metadata.GeoDataExifPoint.Latitude) -Longitude2 ([double]$Metadata.GeoDataExifPoint.Longitude)
        if ($googleEditDistance -gt $GpsAgreementMeters) { $googleEditDetected = $true }
    }

    $embeddedLat = $null; $embeddedLon = $null; $embeddedAlt = $null
    if ($null -ne $Embedded) {
        if ($Embedded.PSObject.Properties['GPSLatitude']) { $embeddedLat = Convert-ToNullableDouble $Embedded.GPSLatitude }
        if ($Embedded.PSObject.Properties['GPSLongitude']) { $embeddedLon = Convert-ToNullableDouble $Embedded.GPSLongitude }
        if ($Embedded.PSObject.Properties['GPSAltitude']) { $embeddedAlt = Convert-ToNullableDouble $Embedded.GPSAltitude }
    }

    if ((Test-ValidCoordinates -Latitude $embeddedLat -Longitude $embeddedLon) -and -not $googleEditDetected) {
        $distance = Get-CoordinateDistanceMeters -Latitude1 ([double]$embeddedLat) -Longitude1 ([double]$embeddedLon) -Latitude2 $googleLat -Longitude2 $googleLon
        if ($distance -le $GpsAgreementMeters) {
            return [PSCustomObject]@{
                ShouldWrite = $false
                Latitude = [double]$embeddedLat
                Longitude = [double]$embeddedLon
                Altitude = $embeddedAlt
                Source = 'embedded-preserved'
                Reason = ('preserved higher-precision embedded GPS; agrees with Google within {0:N1} m' -f $distance)
                DistanceMeters = [Math]::Round($distance, 2)
                GoogleEditDetected = $false
            }
        }
    }

    $reason = 'Google GPS written because embedded GPS was missing or materially different'
    if ($googleEditDetected) {
        $reason = ('Google geoData differs from geoDataExif by {0:N1} m; treating Google location as an edit' -f $googleEditDistance)
    }
    elseif (Test-ValidCoordinates -Latitude $embeddedLat -Longitude $embeddedLon) {
        $distance = Get-CoordinateDistanceMeters -Latitude1 ([double]$embeddedLat) -Longitude1 ([double]$embeddedLon) -Latitude2 $googleLat -Longitude2 $googleLon
        $reason = ('Google GPS written; embedded location differs by {0:N1} m' -f $distance)
    }

    return [PSCustomObject]@{
        ShouldWrite = $true
        Latitude = $googleLat
        Longitude = $googleLon
        Altitude = $googleAlt
        Source = $(if ([string]::IsNullOrWhiteSpace($Metadata.GeoSource)) { 'Google' } else { $Metadata.GeoSource })
        Reason = $reason
        DistanceMeters = $(if (Test-ValidCoordinates -Latitude $embeddedLat -Longitude $embeddedLon) {
            [Math]::Round((Get-CoordinateDistanceMeters -Latitude1 ([double]$embeddedLat) -Longitude1 ([double]$embeddedLon) -Latitude2 $googleLat -Longitude2 $googleLon), 2)
        } else { $null })
        GoogleEditDetected = $googleEditDetected
    }
}

function Get-EmbeddedMetadata {
    param(
        [string]$ExifTool,
        [string]$MediaPath,
        [bool]$IsQuickTime
    )

    $exifArgs = New-Object System.Collections.Generic.List[string]
    $exifArgs.Add('-j')
    $exifArgs.Add('-n')
    if ($IsQuickTime) {
        $exifArgs.Add('-api')
        $exifArgs.Add('QuickTimeUTC=1')
    }
    foreach ($tag in @('-DateTimeOriginal', '-CreateDate', '-ModifyDate', '-OffsetTimeOriginal', '-ExifVersion', '-GPSLatitude', '-GPSLongitude', '-GPSAltitude')) {
        $exifArgs.Add($tag)
    }
    $exifArgs.Add($MediaPath)

    try {
        $command = Invoke-ExifToolCommand -ExifTool $ExifTool -Arguments ([string[]]$exifArgs) -SuppressErrors
        if ($command.ExitCode -ne 0 -or $command.OutputLines.Count -eq 0) { return $null }
        $parsed = (($command.OutputLines | ForEach-Object { $_.ToString() }) -join "`n") | ConvertFrom-Json
        if ($parsed -is [Array]) { return $parsed[0] }
        return $parsed
    }
    catch {
        return $null
    }
}

function Get-TimeDecision {
    param(
        $Metadata,
        $Embedded,
        [bool]$IsPhoto,
        [switch]$Force
    )

    if ($null -eq $Metadata.UtcTime) {
        return [PSCustomObject]@{ WriteDate = $false; ExistingLocal = $null; OffsetMinutes = $null; Reason = 'no usable JSON timestamp' }
    }

    # Google creationTime is generally the upload/creation time in Google's system,
    # not necessarily the camera capture time.  If photoTakenTime is absent but the
    # media already has a usable embedded capture date, keep that embedded date.
    # This safeguard also applies in ForceJsonTime mode; ForceJsonTime is intended
    # to make Google's photoTakenTime win, not to replace a real capture date with
    # an upload timestamp.
    if ($Metadata.TimeSource -eq 'creationTime' -and $null -ne $Embedded) {
        $fallbackDateText = $null
        if ($IsPhoto -and $Embedded.PSObject.Properties['DateTimeOriginal']) {
            $fallbackDateText = [string]$Embedded.DateTimeOriginal
        }
        if ([string]::IsNullOrWhiteSpace($fallbackDateText) -and $Embedded.PSObject.Properties['CreateDate']) {
            $fallbackDateText = [string]$Embedded.CreateDate
        }
        $fallbackEmbeddedDate = Parse-ExifDate -Text $fallbackDateText
        if ($null -ne $fallbackEmbeddedDate) {
            return [PSCustomObject]@{
                WriteDate = $false
                ExistingLocal = $fallbackEmbeddedDate
                OffsetMinutes = $null
                Reason = 'kept embedded capture time; sidecar has creationTime only'
            }
        }
    }

    if ($Force) {
        return [PSCustomObject]@{ WriteDate = $true; ExistingLocal = $null; OffsetMinutes = 0; Reason = 'ForceJsonTime/photoTakenTime' }
    }

    # QuickTime creation/track/media timestamps are defined as UTC-oriented values.
    # For videos, use Google's absolute timestamp directly instead of applying the
    # photo-specific local-wall-clock preservation heuristic below.
    if (-not $IsPhoto) {
        return [PSCustomObject]@{ WriteDate = $true; ExistingLocal = $null; OffsetMinutes = 0; Reason = 'video timestamp restored from Google JSON' }
    }

    if ($null -eq $Embedded) {
        return [PSCustomObject]@{ WriteDate = $true; ExistingLocal = $null; OffsetMinutes = 0; Reason = 'no readable embedded timestamp' }
    }

    $dateText = $null
    if ($IsPhoto -and $Embedded.PSObject.Properties['DateTimeOriginal']) {
        $dateText = [string]$Embedded.DateTimeOriginal
    }
    if ([string]::IsNullOrWhiteSpace($dateText) -and $Embedded.PSObject.Properties['CreateDate']) {
        $dateText = [string]$Embedded.CreateDate
    }

    $existing = Parse-ExifDate -Text $dateText
    if ($null -eq $existing) {
        return [PSCustomObject]@{ WriteDate = $true; ExistingLocal = $null; OffsetMinutes = 0; Reason = 'embedded timestamp missing/unparseable' }
    }

    $jsonNaive = [DateTime]::SpecifyKind([DateTime]$Metadata.UtcTime, [DateTimeKind]::Unspecified)

    # If the image already contains an explicit offset, compare the real instant.
    $explicitOffset = $null
    if ($Embedded.PSObject.Properties['OffsetTimeOriginal']) {
        $explicitOffset = Parse-OffsetMinutes -Text ([string]$Embedded.OffsetTimeOriginal)
    }
    if ($null -ne $explicitOffset) {
        try {
            $dto = [DateTimeOffset]::new($existing, [TimeSpan]::FromMinutes([int]$explicitOffset))
            $seconds = [Math]::Abs(($dto.UtcDateTime - [DateTime]$Metadata.UtcTime).TotalSeconds)
            if ($seconds -le 120) {
                return [PSCustomObject]@{ WriteDate = $false; ExistingLocal = $existing; OffsetMinutes = [int]$explicitOffset; Reason = 'embedded timestamp + offset already matches JSON instant' }
            }
        }
        catch { }
    }

    # EXIF DateTimeOriginal is often local wall-clock time while Takeout JSON is UTC.
    # If their difference is a plausible world time-zone offset in 15-minute increments,
    # preserve the local clock and infer/write that offset instead of replacing it with UTC.
    $differenceMinutes = ($existing - $jsonNaive).TotalMinutes
    $nearestQuarterHour = [int]([Math]::Round($differenceMinutes / 15.0) * 15)
    $distanceFromQuarter = [Math]::Abs($differenceMinutes - $nearestQuarterHour)

    if ([Math]::Abs($nearestQuarterHour) -le (14 * 60) -and $distanceFromQuarter -le 2) {
        return [PSCustomObject]@{
            WriteDate = $false
            ExistingLocal = $existing
            OffsetMinutes = $nearestQuarterHour
            Reason = ('preserved existing local time; inferred UTC offset ' + (Format-Offset -Minutes $nearestQuarterHour))
        }
    }

    return [PSCustomObject]@{ WriteDate = $true; ExistingLocal = $existing; OffsetMinutes = 0; Reason = 'embedded date differs from Google time by more than a timezone-like offset' }
}

function Invoke-ExifWrite {
    param(
        [string]$ExifTool,
        [string]$MediaPath,
        [string]$Extension,
        $Metadata,
        $TimeDecision,
        $Embedded,
        $GpsDecision
    )

    $isPhoto = $script:PhotoExtensionLookup.ContainsKey($Extension)
    $isQuickTime = $script:QuickTimeExtensionLookup.ContainsKey($Extension)

    if (-not $isPhoto -and -not $isQuickTime) {
        return [PSCustomObject]@{ Success = $true; Message = 'Copied; embedded metadata writing skipped for this container (filesystem timestamp restored).' }
    }

    $exifArgs = New-Object System.Collections.Generic.List[string]
    $exifArgs.Add('-m')
    $exifArgs.Add('-overwrite_original')
    $exifArgs.Add('-api')
    $exifArgs.Add('LargeFileSupport=1')
    $hasWrite = $false

    # EXIF OffsetTime/OffsetTimeOriginal/OffsetTimeDigitized were standardized in
    # EXIF 2.31. If we add these tags to an older EXIF block (for example 2.20),
    # upgrade only the version declaration needed for standards compliance. Never
    # downgrade a newer EXIF version such as 2.32.
    $needsExif231 = $false
    if ($isPhoto) {
        $embeddedExifVersion = ''
        if ($null -ne $Embedded -and $Embedded.PSObject.Properties['ExifVersion']) {
            $embeddedExifVersion = [string]$Embedded.ExifVersion
        }
        $digits = ($embeddedExifVersion -replace '[^0-9]', '')
        if ([string]::IsNullOrWhiteSpace($digits)) {
            $needsExif231 = $true
        }
        else {
            $parsedExifVersion = 0
            if (-not [int]::TryParse($digits, [ref]$parsedExifVersion) -or $parsedExifVersion -lt 231) {
                $needsExif231 = $true
            }
        }
    }

    # IMPORTANT v4/v4.1/v5.1: Do NOT use QuickTimeUTC while writing the core QuickTime dates below.
    # Google photoTakenTime is already an absolute UTC instant. With QuickTimeUTC=1,
    # a timezone-less value is interpreted as the computer's local time and converted
    # again to UTC (for example +7 hours on Pacific daylight time). Writing the known
    # UTC clock value directly stores the correct QuickTime epoch value.

    if ($TimeDecision.WriteDate -and $null -ne $Metadata.UtcTime) {
        $utc = [DateTime]$Metadata.UtcTime
        $exifDate = $utc.ToString('yyyy:MM:dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
        $isoDate = $utc.ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)

        if ($isPhoto) {
            $hasWrite = $true
            $exifArgs.Add('-AllDates=' + $exifDate)
            $exifArgs.Add('-XMP:DateTimeOriginal=' + $isoDate)
            $exifArgs.Add('-XMP:CreateDate=' + $isoDate)
            $exifArgs.Add('-XMP:ModifyDate=' + $isoDate)
            $exifArgs.Add('-OffsetTime=+00:00')
            $exifArgs.Add('-OffsetTimeOriginal=+00:00')
            $exifArgs.Add('-OffsetTimeDigitized=+00:00')
            if ($needsExif231) { $exifArgs.Add('-ExifVersion=0231') }
        }
        elseif ($isQuickTime) {
            $hasWrite = $true
            # Group-qualify movie-header dates so ExifTool doesn't also create
            # timezone-less XMP dates as a side effect of generic -CreateDate/-ModifyDate.
            $exifArgs.Add('-QuickTime:CreateDate=' + $exifDate)
            $exifArgs.Add('-QuickTime:ModifyDate=' + $exifDate)
            $exifArgs.Add('-TrackCreateDate=' + $exifDate)
            $exifArgs.Add('-TrackModifyDate=' + $exifDate)
            $exifArgs.Add('-MediaCreateDate=' + $exifDate)
            $exifArgs.Add('-MediaModifyDate=' + $exifDate)
            $exifArgs.Add('-Keys:CreationDate=' + $isoDate)
            $exifArgs.Add('-XMP-xmp:CreateDate=' + $isoDate)
            $exifArgs.Add('-XMP-xmp:ModifyDate=' + $isoDate)
        }
    }
    elseif ($isPhoto -and $null -ne $TimeDecision.OffsetMinutes) {
        # Preserve existing local EXIF time, but make the inferred/known offset explicit.
        $hasWrite = $true
        $offsetText = Format-Offset -Minutes ([int]$TimeDecision.OffsetMinutes)
        $exifArgs.Add('-OffsetTime=' + $offsetText)
        $exifArgs.Add('-OffsetTimeOriginal=' + $offsetText)
        $exifArgs.Add('-OffsetTimeDigitized=' + $offsetText)
        if ($needsExif231) { $exifArgs.Add('-ExifVersion=0231') }
    }

    if ($null -ne $GpsDecision -and $GpsDecision.ShouldWrite -and
        (Test-ValidCoordinates -Latitude $GpsDecision.Latitude -Longitude $GpsDecision.Longitude)) {
        $hasWrite = $true
        $lat = [double]$GpsDecision.Latitude
        $lon = [double]$GpsDecision.Longitude
        if ($isPhoto) {
            $latAbs = [Math]::Abs($lat).ToString('0.########', [Globalization.CultureInfo]::InvariantCulture)
            $lonAbs = [Math]::Abs($lon).ToString('0.########', [Globalization.CultureInfo]::InvariantCulture)
            $exifArgs.Add('-GPSLatitude=' + $latAbs)
            $exifArgs.Add('-GPSLatitudeRef=' + $(if ($lat -lt 0) { 'S' } else { 'N' }))
            $exifArgs.Add('-GPSLongitude=' + $lonAbs)
            $exifArgs.Add('-GPSLongitudeRef=' + $(if ($lon -lt 0) { 'W' } else { 'E' }))
            if ($null -ne $GpsDecision.Altitude) {
                $alt = [double]$GpsDecision.Altitude
                $exifArgs.Add('-GPSAltitude=' + [Math]::Abs($alt).ToString('0.########', [Globalization.CultureInfo]::InvariantCulture))
                $exifArgs.Add('-GPSAltitudeRef=' + $(if ($alt -lt 0) { '1' } else { '0' }))
            }
        }
        elseif ($isQuickTime) {
            $latText = $lat.ToString('+0.########;-0.########;+0', [Globalization.CultureInfo]::InvariantCulture)
            $lonText = $lon.ToString('+0.########;-0.########;+0', [Globalization.CultureInfo]::InvariantCulture)
            $exifArgs.Add('-Keys:GPSCoordinates=' + $latText + $lonText + '/')
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($Metadata.Description)) {
        $hasWrite = $true
        if ($isPhoto) {
            $exifArgs.Add('-XMP:Description=' + $Metadata.Description)
            $exifArgs.Add('-IPTC:Caption-Abstract=' + $Metadata.Description)
            $exifArgs.Add('-EXIF:ImageDescription=' + $Metadata.Description)
        }
        else {
            $exifArgs.Add('-Keys:Description=' + $Metadata.Description)
        }
    }

    if ($isPhoto -and $Metadata.People -and $Metadata.People.Count -gt 0) {
        $hasWrite = $true
        foreach ($person in $Metadata.People) {
            $exifArgs.Add('-XMP-iptcExt:PersonInImage+=' + $person)
            $exifArgs.Add('-XMP-dc:Subject+=' + $person)
        }
    }

    if ($Metadata.Favorite) {
        $hasWrite = $true
        if ($isPhoto) {
            $exifArgs.Add('-Rating=5')
        }
        else {
            $exifArgs.Add('-Keys:UserRating=5')
        }
    }

    # If there is nothing to write, return successfully without launching ExifTool.
    if (-not $hasWrite) {
        return [PSCustomObject]@{ Success = $true; Message = 'No embedded tag changes required.' }
    }

    $exifArgs.Add($MediaPath)

    try {
        $command = Invoke-ExifToolCommand -ExifTool $ExifTool -Arguments ([string[]]$exifArgs)
        $text = ($command.OutputLines | ForEach-Object { $_.ToString() }) -join ' | '
        if ($command.ExitCode -eq 0) {
            return [PSCustomObject]@{ Success = $true; Message = $text }
        }
        return [PSCustomObject]@{ Success = $false; Message = ('ExifTool exit ' + $command.ExitCode + ': ' + $text) }
    }
    catch {
        return [PSCustomObject]@{ Success = $false; Message = $_.Exception.Message }
    }
}


function Get-ExifToolTagValue {
    param(
        [string]$ExifTool,
        [string]$MediaPath,
        [string]$Tag
    )

    try {
        $command = Invoke-ExifToolCommand -ExifTool $ExifTool -Arguments @('-s3', $Tag, $MediaPath) -SuppressErrors
        if ($command.ExitCode -ne 0) { return '' }
        foreach ($line in @($command.OutputLines)) {
            $text = [string]$line
            if (-not [string]::IsNullOrWhiteSpace($text)) { return $text.Trim() }
        }
    }
    catch { }
    return ''
}

function Parse-OffsetDateToUtc {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $value = $Text.Trim()

    # ExifTool commonly emits 2021:08:02 15:06:52Z or ...-07:00.
    if ($value -match '^(\d{4}):(\d{2}):(\d{2})[ T](\d{2}):(\d{2}):(\d{2})(?:\.\d+)?(Z|[+-]\d{2}:\d{2})$') {
        $iso = ('{0}-{1}-{2}T{3}:{4}:{5}{6}' -f $Matches[1], $Matches[2], $Matches[3], $Matches[4], $Matches[5], $Matches[6], $Matches[7])
        $dto = [DateTimeOffset]::MinValue
        if ([DateTimeOffset]::TryParse($iso, [Globalization.CultureInfo]::InvariantCulture,
                [Globalization.DateTimeStyles]::AllowWhiteSpaces, [ref]$dto)) {
            return $dto.UtcDateTime
        }
    }

    # ISO form already using dashes/T.
    $dto2 = [DateTimeOffset]::MinValue
    if ([DateTimeOffset]::TryParse($value, [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::AllowWhiteSpaces, [ref]$dto2)) {
        if ($value -match '(Z|[+-]\d{2}:\d{2})$') { return $dto2.UtcDateTime }
    }
    return $null
}

function Compare-UtcTimes {
    param(
        [DateTime]$Expected,
        [DateTime]$Actual,
        [double]$ToleranceSeconds = 2
    )

    return ([Math]::Abs(($Actual.ToUniversalTime() - $Expected.ToUniversalTime()).TotalSeconds) -le $ToleranceSeconds)
}

function Verify-OutputMetadata {
    param(
        [string]$ExifTool,
        [string]$MediaPath,
        [string]$Extension,
        $Metadata,
        $TimeDecision,
        $GpsDecision,
        [switch]$Disabled
    )

    if ($Disabled) {
        return [PSCustomObject]@{ Status = 'DISABLED'; StoredUtc = ''; Details = 'Post-write verification disabled by user.' }
    }

    if ($null -eq $Metadata.UtcTime) {
        return [PSCustomObject]@{ Status = 'SKIPPED'; StoredUtc = ''; Details = 'No usable Google timestamp was available to verify.' }
    }

    $shouldVerifyTime = ($Metadata.TimeSource -eq 'photoTakenTime' -or $TimeDecision.WriteDate)
    if (-not $shouldVerifyTime) {
        return [PSCustomObject]@{ Status = 'SKIPPED'; StoredUtc = ''; Details = 'Capture time was intentionally preserved because the sidecar only provided creationTime.' }
    }

    $expected = [DateTime]$Metadata.UtcTime
    $isPhoto = $script:PhotoExtensionLookup.ContainsKey($Extension)
    $isQuickTime = $script:QuickTimeExtensionLookup.ContainsKey($Extension)

    if ($isQuickTime) {
        # Read core QuickTime integer timestamps WITHOUT QuickTimeUTC. Their raw clock
        # values are specified as UTC, so this is independent of the computer timezone.
        $verifyArgs = New-Object System.Collections.Generic.List[string]
        $verifyArgs.Add('-s3')
        $verifyArgs.Add('-a')
        foreach ($tag in @('-QuickTime:CreateDate', '-QuickTime:ModifyDate', '-TrackCreateDate', '-TrackModifyDate', '-MediaCreateDate', '-MediaModifyDate')) { $verifyArgs.Add($tag) }
        $verifyArgs.Add($MediaPath)

        try {
            $verifyCommand = Invoke-ExifToolCommand -ExifTool $ExifTool -Arguments ([string[]]$verifyArgs) -SuppressErrors
            if ($verifyCommand.ExitCode -ne 0) {
                return [PSCustomObject]@{ Status = 'FAIL'; StoredUtc = ''; Details = 'ExifTool could not read the QuickTime timestamps back for verification.' }
            }

            $values = New-Object System.Collections.Generic.List[string]
            foreach ($line in @($verifyCommand.OutputLines)) {
                $text = ([string]$line).Trim()
                if (-not [string]::IsNullOrWhiteSpace($text)) { $values.Add($text) }
            }
            if ($values.Count -eq 0) {
                return [PSCustomObject]@{ Status = 'FAIL'; StoredUtc = ''; Details = 'No core QuickTime timestamps were readable after writing.' }
            }

            $bad = New-Object System.Collections.Generic.List[string]
            $storedIso = New-Object System.Collections.Generic.List[string]
            foreach ($value in $values) {
                $parsed = Parse-ExifDate -Text $value
                if ($null -eq $parsed) {
                    $bad.Add($value)
                    continue
                }
                $actualUtc = [DateTime]::SpecifyKind($parsed, [DateTimeKind]::Utc)
                $storedIso.Add($actualUtc.ToString('o'))
                if (-not (Compare-UtcTimes -Expected $expected -Actual $actualUtc)) { $bad.Add($value) }
            }

            $uniqueStored = @($storedIso | Select-Object -Unique)
            $storedText = $uniqueStored -join '; '
            if ($bad.Count -gt 0) {
                return [PSCustomObject]@{
                    Status = 'FAIL'
                    StoredUtc = $storedText
                    Details = ('QuickTime UTC mismatch. Expected ' + $expected.ToString('o') + '; raw stored value(s): ' + (($values | Select-Object -Unique) -join ', '))
                }
            }

            $keysDate = Get-ExifToolTagValue -ExifTool $ExifTool -MediaPath $MediaPath -Tag '-Keys:CreationDate'
            if (-not [string]::IsNullOrWhiteSpace($keysDate)) {
                $keysUtc = Parse-OffsetDateToUtc -Text $keysDate
                if ($null -ne $keysUtc -and -not (Compare-UtcTimes -Expected $expected -Actual $keysUtc)) {
                    return [PSCustomObject]@{
                        Status = 'FAIL'
                        StoredUtc = $storedText
                        Details = ('Keys:CreationDate also mismatched: ' + $keysDate + '. Expected ' + $expected.ToString('o'))
                    }
                }
            }

            $xmpDate = Get-ExifToolTagValue -ExifTool $ExifTool -MediaPath $MediaPath -Tag '-XMP-xmp:CreateDate'
            if (-not [string]::IsNullOrWhiteSpace($xmpDate)) {
                $xmpUtc = Parse-OffsetDateToUtc -Text $xmpDate
                if ($null -eq $xmpUtc -or -not (Compare-UtcTimes -Expected $expected -Actual $xmpUtc)) {
                    return [PSCustomObject]@{
                        Status = 'FAIL'
                        StoredUtc = $storedText
                        Details = ('XMP CreateDate was missing an unambiguous UTC match: ' + $xmpDate + '. Expected ' + $expected.ToString('o'))
                    }
                }
            }

            return [PSCustomObject]@{
                Status = 'PASS'
                StoredUtc = $storedText
                Details = ('Verified ' + $values.Count + ' core QuickTime timestamp(s), Keys creation date, and UTC-aware XMP date against Google UTC.')
            }
        }
        catch {
            return [PSCustomObject]@{ Status = 'FAIL'; StoredUtc = ''; Details = ('Verification exception: ' + $_.Exception.Message) }
        }
    }

    if ($isPhoto) {
        # One structured read replaces multiple ExifTool launches/reads for the common
        # photo path and also gives us GPS for optional write verification.
        $readback = Get-EmbeddedMetadata -ExifTool $ExifTool -MediaPath $MediaPath -IsQuickTime $false
        if ($null -eq $readback) {
            return [PSCustomObject]@{ Status = 'FAIL'; StoredUtc = ''; Details = 'Could not read photo metadata back after writing.' }
        }

        $dateText = ''
        $offsetText = ''
        if ($readback.PSObject.Properties['DateTimeOriginal']) { $dateText = [string]$readback.DateTimeOriginal }
        if ($readback.PSObject.Properties['OffsetTimeOriginal']) { $offsetText = [string]$readback.OffsetTimeOriginal }

        $actualUtc = $null
        $timeDetails = ''
        if (-not [string]::IsNullOrWhiteSpace($dateText) -and -not [string]::IsNullOrWhiteSpace($offsetText)) {
            $local = Parse-ExifDate -Text $dateText
            $offsetMinutes = Parse-OffsetMinutes -Text $offsetText
            if ($null -ne $local -and $null -ne $offsetMinutes) {
                try {
                    $dto = [DateTimeOffset]::new($local, [TimeSpan]::FromMinutes([int]$offsetMinutes))
                    $actualUtc = $dto.UtcDateTime
                    $timeDetails = ('EXIF DateTimeOriginal ' + $dateText + ' ' + $offsetText)
                }
                catch { $actualUtc = $null }
            }
        }

        if ($null -eq $actualUtc) {
            # Some formats favor XMP. Only pay for this extra read when the normal EXIF
            # date+offset path is unavailable.
            $xmpText = Get-ExifToolTagValue -ExifTool $ExifTool -MediaPath $MediaPath -Tag '-XMP:DateTimeOriginal'
            if (-not [string]::IsNullOrWhiteSpace($xmpText)) {
                $actualUtc = Parse-OffsetDateToUtc -Text $xmpText
                $timeDetails = ('XMP DateTimeOriginal ' + $xmpText)
            }
        }

        if ($null -eq $actualUtc) {
            return [PSCustomObject]@{ Status = 'FAIL'; StoredUtc = ''; Details = 'Could not read a timezone-qualified EXIF or XMP capture timestamp after writing.' }
        }
        if (-not (Compare-UtcTimes -Expected $expected -Actual $actualUtc)) {
            return [PSCustomObject]@{
                Status = 'FAIL'
                StoredUtc = $actualUtc.ToString('o')
                Details = ('Photo timestamp mismatch. Expected ' + $expected.ToString('o') + '; stored ' + $timeDetails)
            }
        }

        $gpsDetails = ''
        if ($null -ne $GpsDecision -and $GpsDecision.ShouldWrite -and
            (Test-ValidCoordinates -Latitude $GpsDecision.Latitude -Longitude $GpsDecision.Longitude)) {
            $readLat = $null; $readLon = $null
            if ($readback.PSObject.Properties['GPSLatitude']) { $readLat = Convert-ToNullableDouble $readback.GPSLatitude }
            if ($readback.PSObject.Properties['GPSLongitude']) { $readLon = Convert-ToNullableDouble $readback.GPSLongitude }
            if (-not (Test-ValidCoordinates -Latitude $readLat -Longitude $readLon)) {
                return [PSCustomObject]@{ Status = 'FAIL'; StoredUtc = $actualUtc.ToString('o'); Details = 'Timestamp verified, but GPS written from Google could not be read back.' }
            }
            $gpsError = Get-CoordinateDistanceMeters -Latitude1 ([double]$readLat) -Longitude1 ([double]$readLon) `
                -Latitude2 ([double]$GpsDecision.Latitude) -Longitude2 ([double]$GpsDecision.Longitude)
            if ($gpsError -gt 5.0) {
                return [PSCustomObject]@{
                    Status = 'FAIL'; StoredUtc = $actualUtc.ToString('o')
                    Details = ('Timestamp verified, but GPS read-back differs from intended Google GPS by {0:N1} m.' -f $gpsError)
                }
            }
            $gpsDetails = (' GPS also verified within {0:N1} m.' -f $gpsError)
        }

        return [PSCustomObject]@{
            Status = 'PASS'
            StoredUtc = $actualUtc.ToString('o')
            Details = ('Verified ' + $timeDetails + ' against Google UTC.' + $gpsDetails)
        }
    }

    return [PSCustomObject]@{ Status = 'SKIPPED'; StoredUtc = ''; Details = 'This media container is copied but does not use the photo/QuickTime timestamp verification path.' }
}

# ---------------- Main ----------------

$runStopwatch = [Diagnostics.Stopwatch]::StartNew()
$logPath = ''
$diagnosticsPath = ''
$summaryPath = ''
$fatalMessage = ''
$outputWasNonEmpty = $false

try {
    if ($InternalWorker) {
        if ([string]::IsNullOrWhiteSpace($FileListPath) -or -not (Test-Path -LiteralPath $FileListPath -PathType Leaf)) {
            throw 'Internal worker manifest is missing.'
        }
        if ([string]::IsNullOrWhiteSpace($WorkerLogPath)) { throw 'Internal worker log path is missing.' }
        if ([string]::IsNullOrWhiteSpace($WorkerSummaryPath)) { throw 'Internal worker summary path is missing.' }
        if ([string]::IsNullOrWhiteSpace($WorkerResultPath)) { throw 'Internal worker result path is missing.' }
    }

    if (-not $InputFolder) {
        $InputFolder = Select-FolderInteractive -Description 'Select your extracted Google Photos Takeout folder'
    }

    if ([string]::IsNullOrWhiteSpace($InputFolder) -or -not (Test-Path -LiteralPath $InputFolder -PathType Container)) {
        throw "Input folder does not exist: $InputFolder"
    }

    $InputFolder = Get-NormalizedDirectoryPath -Path (Resolve-Path -LiteralPath $InputFolder).Path

    if (-not $OutputFolder) { $OutputFolder = Get-DefaultOutputFolder -InputPath $InputFolder }
    $OutputFolder = Get-NormalizedDirectoryPath -Path $OutputFolder

    if (-not $DryRun) {
        if ((Test-PathContainsPath -ParentPath $InputFolder -ChildPath $OutputFolder) -or
            (Test-PathContainsPath -ParentPath $OutputFolder -ChildPath $InputFolder)) {
            throw 'The input and output folders must be separate and must not contain one another.'
        }

        if (Test-Path -LiteralPath $OutputFolder) {
            if (-not (Test-Path -LiteralPath $OutputFolder -PathType Container)) {
                throw "Output path exists but is not a folder: $OutputFolder"
            }
            # The parent v6 orchestrator performs the one authoritative empty-output
            # safety check before starting the pool. Parallel workers necessarily see
            # files created by sibling workers, so they must not reject a now-nonempty
            # shared destination directory.
            if (-not $InternalWorker) {
                $existingItem = Get-ChildItem -LiteralPath $OutputFolder -Force -ErrorAction Stop | Select-Object -First 1
                $outputWasNonEmpty = ($null -ne $existingItem)
                if ($outputWasNonEmpty -and -not $OverwriteOutput) {
                    throw "The output folder is not empty. Choose an empty/new folder, or rerun with -OverwriteOutput to re-run/replace files: $OutputFolder"
                }
            }
            else {
                $outputWasNonEmpty = $true
            }
        }
    }

    $ExifToolPath = Resolve-ExifTool -RequestedPath $ExifToolPath

    Write-Host ''
    Write-Host ('Google Photos Takeout Metadata Fixer v6.2.3 - Worker ' + $WorkerId) -ForegroundColor Cyan
    Write-Host ('Input:      ' + $InputFolder)
    Write-Host ('Output:     ' + $OutputFolder)
    Write-Host ('ExifTool:   ' + $ExifToolPath)
    Write-Host ('Time mode:  ' + $(if ($ForceJsonTime) { 'FORCE Google JSON timestamps' } else { 'preserve plausible existing local EXIF time' }))
    Write-Host ('Scan:       ' + $(if ($TopLevelOnly) { 'selected folder only' } else { 'selected folder + all subfolders' }))
    Write-Host ('Verify:     ' + $(if ($SkipVerification) { 'disabled by user' } else { 'read back written timestamps (recommended)' }))
    Write-Host ('GPS policy: preserve embedded precision when within ' + $GpsAgreementMeters.ToString('0.#', [Globalization.CultureInfo]::InvariantCulture) + ' m of Google')
    if ($DryRun) { Write-Host 'Mode:       DRY RUN (nothing will be copied or changed)' -ForegroundColor Yellow }
    Write-Host ''

    $mediaExtensionLookup = $script:MediaExtensionLookup
    $signatureSkipLookup = $script:SignatureScanSkipLookup

    Write-ProgressState -Stage 'Scanning' -Message 'Scanning selected folder for photos/videos, including signature-detected files with unusual extensions.'
    $scanErrors = @()
    $scanParams = @{
        LiteralPath = $InputFolder
        File = $true
        ErrorAction = 'SilentlyContinue'
        ErrorVariable = '+scanErrors'
    }
    if (-not $TopLevelOnly) { $scanParams['Recurse'] = $true }

    # In worker-pool mode the parent already performed the expensive full-library
    # enumeration and signature scan. Load only this worker's manifest. Standalone
    # compatibility mode retains the v5.1 scanner for direct use and troubleshooting.
    $mediaFiles = New-Object 'System.Collections.Generic.List[System.IO.FileInfo]'
    $mediaTypeScanCache = @{}
    $unknownExtensionMediaDetected = 0

    if ($InternalWorker -and -not [string]::IsNullOrWhiteSpace($FileListPath)) {
        # v6.2.3 worker manifests are one absolute media path per line. This is
        # intentionally simpler than a JSON array and avoids a Windows PowerShell 5.1
        # collection conversion failure that affected workers assigned multiple files.
        foreach ($manifestLine in (Get-Content -LiteralPath $FileListPath -Encoding UTF8)) {
            $pathText = [string]$manifestLine
            if ([string]::IsNullOrWhiteSpace($pathText)) { continue }
            if (-not (Test-Path -LiteralPath $pathText -PathType Leaf)) {
                $scanErrors += ('Manifest media path is missing: ' + $pathText)
                continue
            }
            $candidate = Get-Item -LiteralPath $pathText -Force -ErrorAction SilentlyContinue
            if ($null -eq $candidate) {
                $scanErrors += ('Manifest media path could not be opened: ' + $pathText)
                continue
            }
            [void]$mediaFiles.Add($candidate)

            # The parent already sniffed unusual extensions during discovery, but the
            # worker can cheaply reconstruct that information for its own audit row.
            $candidateExt = $candidate.Extension.ToLowerInvariant()
            if (-not $mediaExtensionLookup.ContainsKey($candidateExt) -and -not $signatureSkipLookup.ContainsKey($candidateExt) -and $candidate.Length -ge 3) {
                $candidateType = Get-MediaTypeInfo -Path $candidate.FullName -DeclaredExtension $candidateExt
                if ($candidateType.Confident -and $candidateType.DetectedType -ne 'Unknown') {
                    $mediaTypeScanCache[$candidate.FullName] = $candidateType
                    $unknownExtensionMediaDetected++
                }
            }
        }
    }
    else {
        # Known media extensions stay on the zero-extra-I/O fast path. Only files whose
        # extension is not recognized (and is not an obvious sidecar/document/archive/audio
        # file) are opened for a 64-byte signature read.
        Get-ChildItem @scanParams | ForEach-Object {
            $candidate = $_
            $candidateExt = $candidate.Extension.ToLowerInvariant()

            if ($mediaExtensionLookup.ContainsKey($candidateExt)) {
                [void]$mediaFiles.Add($candidate)
            }
            elseif (-not $signatureSkipLookup.ContainsKey($candidateExt) -and $candidate.Length -ge 3) {
                $candidateType = Get-MediaTypeInfo -Path $candidate.FullName -DeclaredExtension $candidateExt
                if ($candidateType.Confident -and $candidateType.DetectedType -ne 'Unknown') {
                    [void]$mediaFiles.Add($candidate)
                    $mediaTypeScanCache[$candidate.FullName] = $candidateType
                    $unknownExtensionMediaDetected++
                }
            }
        }
    }

    if ($mediaFiles.Count -eq 0) {
        Write-ProgressState -Stage 'Done' -Message 'No supported or signature-detected photo/video files were found.' -Extra @{ Processed = 0; Verified = 0; VerificationFailed = 0; UnknownExtensionMedia = 0; Unmatched = 0; Failed = 0; LogPath = ''; OutputFolder = $OutputFolder }
        Write-Warning 'No supported or signature-detected photo/video files were found.'
        return
    }

    $total = $mediaFiles.Count
    $totalMediaBytes = [Int64]0
    foreach ($fileForSize in $mediaFiles) { $totalMediaBytes += [Int64]$fileForSize.Length }

    $preflightMessage = ('Found ' + $total + ' media files (' + (Format-ByteSize $totalMediaBytes) + ').')
    if ($unknownExtensionMediaDetected -gt 0) {
        $preflightMessage += (' Signature-detected unusual extensions: ' + $unknownExtensionMediaDetected + '.')
    }
    Write-ProgressState -Stage 'Preflight' -Total $total -Message $preflightMessage -Extra @{ UnknownExtensionMedia = $unknownExtensionMediaDetected }

    if (-not $DryRun) {
        Ensure-OutputDirectory -Path $OutputFolder
        if (-not $InternalWorker) {
            $spaceCheck = Test-OutputFreeSpace -Path $OutputFolder -MediaBytes $totalMediaBytes -OutputWasNonEmpty $outputWasNonEmpty
            if ($spaceCheck.Checked) { Write-Host ('Disk:       ' + $spaceCheck.Message) }
        }
    }

    # Open the audit log before processing and stream rows as work completes. This keeps
    # memory usage bounded and leaves a useful partial log if a long run is interrupted.
    if ($InternalWorker -and -not [string]::IsNullOrWhiteSpace($WorkerLogPath)) {
        $logPath = $WorkerLogPath
    }
    elseif ($DryRun) {
        $dryRunLogFolder = Split-Path -Parent $InputFolder
        if (-not $dryRunLogFolder) { $dryRunLogFolder = $InputFolder }
        $logPath = Get-AvailableLogPath -PreferredPath (Join-Path $dryRunLogFolder 'metadata-fix-dry-run.csv')
    }
    else {
        $logPath = Get-AvailableLogPath -PreferredPath (Join-Path $OutputFolder 'metadata-fix-log.csv')
    }

    $auditColumns = @(
        'Source','Destination','Sidecar','MatchMethod','DeclaredType','DetectedType','TypeMismatch','TypeHandling',
        'Status','TimestampUTC','StoredTimestampUTC','Verification','TimeSource','TimeDecision',
        'GPS','GPSDecision','GPSDistanceMeters','GoogleGeoSource','GoogleLocationEdit',
        'VerificationDetails','Note'
    )
    Open-AuditLog -Path $logPath -Columns $auditColumns

    Enable-SystemAwake
    [void](Start-ExifToolSession -ExifTool $ExifToolPath)
    Write-Host ('Engine:     ' + $script:ExifToolMode)
    Write-Host ('Media:      ' + $total + ' files, ' + (Format-ByteSize $totalMediaBytes))
    Write-Host ('Log:        ' + $logPath)
    Write-Host ''

    Write-ProgressState -Stage 'Processing' -Index 0 -Total $total -Message 'Starting metadata pass.' -Extra @{ LogPath = $logPath; OutputFolder = $(if ($DryRun) { '' } else { $OutputFolder }) }

    $index = 0
    $fixed = 0
    $copiedNoSidecar = 0
    $failed = 0
    $preservedLocalTime = 0
    $preservedGpsPrecision = 0
    $googleGpsEdits = 0
    $verifiedCount = 0
    $verificationFailed = 0
    $typeMismatchCount = 0
    $scanErrorCount = @($scanErrors).Count
    $lastConsoleProgress = [DateTime]::MinValue
    $cancelled = $false

    foreach ($media in $mediaFiles) {
        if (Test-CancelRequested) {
            $cancelled = $true
            break
        }
        $index++
        $nowProgress = [DateTime]::UtcNow
        if ($index -eq 1 -or $index -eq $total -or (($nowProgress - $lastConsoleProgress).TotalMilliseconds -ge 200)) {
            $percent = [int](($index / [double]$total) * 100)
            Write-Progress -Activity 'Restoring Google Photos metadata' -Status ($index.ToString() + ' / ' + $total.ToString() + '  ' + $media.Name) -PercentComplete $percent
            $lastConsoleProgress = $nowProgress
        }
        Write-ProgressState -Stage 'Processing' -Index $index -Total $total -Name $media.Name -Message ('Processing ' + $media.Name)

        $relative = Get-RelativePathSimple -BasePath $InputFolder -FullPath $media.FullName
        $destination = Join-Path $OutputFolder $relative
        $declaredExt = $media.Extension.ToLowerInvariant()
        if ($mediaTypeScanCache.ContainsKey($media.FullName)) {
            $typeInfo = $mediaTypeScanCache[$media.FullName]
        }
        else {
            $typeInfo = Get-MediaTypeInfo -Path $media.FullName -DeclaredExtension $declaredExt
        }
        $effectiveExt = $typeInfo.EffectiveExtension
        if ([string]::IsNullOrWhiteSpace($effectiveExt)) { $effectiveExt = $declaredExt }
        if ($typeInfo.TypeMismatch) { $typeMismatchCount++ }
        $sidecarMatch = Find-Sidecar -Media $media

        if ($null -eq $sidecarMatch) {
            if ($DryRun) {
                $status = 'DRY RUN - COPY UNCHANGED (NO SIDECAR)'
            }
            else {
                try {
                    $destinationDir = Split-Path -Parent $destination
                    Ensure-OutputDirectory -Path $destinationDir
                    if ((Test-Path -LiteralPath $destination -PathType Leaf) -and -not $OverwriteOutput) { throw "Destination already exists: $destination" }
                    Copy-Item -LiteralPath $media.FullName -Destination $destination -Force
                    [IO.File]::SetCreationTimeUtc($destination, $media.CreationTimeUtc)
                    [IO.File]::SetLastWriteTimeUtc($destination, $media.LastWriteTimeUtc)
                    $status = 'COPIED UNCHANGED - NO SIDECAR'
                }
                catch {
                    $failed++
                    $row = [PSCustomObject]@{
                        Source = $media.FullName; Destination = $destination; Sidecar = ''; MatchMethod = '';
                        DeclaredType = $typeInfo.DeclaredType; DetectedType = $typeInfo.DetectedType; TypeMismatch = $(if ($typeInfo.TypeMismatch) { 'YES' } else { 'NO' }); TypeHandling = $typeInfo.Handling;
                        Status = 'FAILED'; TimestampUTC = ''; StoredTimestampUTC = ''; Verification = 'NOT RUN'; TimeSource = ''; TimeDecision = ''; GPS = '';
                        GPSDecision = ''; GPSDistanceMeters = ''; GoogleGeoSource = ''; GoogleLocationEdit = '';
                        VerificationDetails = ''; Note = $_.Exception.Message
                    }
                    Write-AuditRow -Row $row -Columns $auditColumns -ForceFlush
                    continue
                }
            }
            $copiedNoSidecar++
            $row = [PSCustomObject]@{
                Source = $media.FullName; Destination = $destination; Sidecar = ''; MatchMethod = '';
                DeclaredType = $typeInfo.DeclaredType; DetectedType = $typeInfo.DetectedType; TypeMismatch = $(if ($typeInfo.TypeMismatch) { 'YES' } else { 'NO' }); TypeHandling = $typeInfo.Handling;
                Status = $status; TimestampUTC = ''; StoredTimestampUTC = ''; Verification = 'NOT APPLICABLE'; TimeSource = ''; TimeDecision = ''; GPS = '';
                GPSDecision = ''; GPSDistanceMeters = ''; GoogleGeoSource = ''; GoogleLocationEdit = '';
                VerificationDetails = ''; Note = 'No matching Google JSON sidecar was found; media is retained rather than omitted.'
            }
            Write-AuditRow -Row $row -Columns $auditColumns
            continue
        }

        $sidecar = $sidecarMatch.Path
        $metadata = Get-GoogleMetadata -JsonPath $sidecar
        if ($null -eq $metadata) {
            $failed++
            $row = [PSCustomObject]@{
                Source = $media.FullName; Destination = $destination; Sidecar = $sidecar; MatchMethod = $sidecarMatch.Method;
                DeclaredType = $typeInfo.DeclaredType; DetectedType = $typeInfo.DetectedType; TypeMismatch = $(if ($typeInfo.TypeMismatch) { 'YES' } else { 'NO' }); TypeHandling = $typeInfo.Handling;
                Status = 'BAD JSON'; TimestampUTC = ''; StoredTimestampUTC = ''; Verification = 'NOT RUN'; TimeSource = ''; TimeDecision = ''; GPS = '';
                GPSDecision = ''; GPSDistanceMeters = ''; GoogleGeoSource = ''; GoogleLocationEdit = '';
                VerificationDetails = ''; Note = 'Sidecar could not be parsed.'
            }
            Write-AuditRow -Row $row -Columns $auditColumns -ForceFlush
            continue
        }

        $ext = $effectiveExt
        $isPhoto = $script:PhotoExtensionLookup.ContainsKey($ext)
        $isQuickTime = $script:QuickTimeExtensionLookup.ContainsKey($ext)
        $workingPath = $null
        $embeddedAlias = $null

        try {
            # In a real copy run, a filename/content mismatch is copied only once: directly
            # to the correctly-typed output working alias. Dry-run still needs a temporary
            # alias because it deliberately creates no output tree.
            if ($DryRun) {
                $embeddedReadPath = $media.FullName
                if ($typeInfo.TypeMismatch) {
                    $embeddedAlias = New-TemporaryMediaAlias -SourcePath $media.FullName -Extension $ext
                    $embeddedReadPath = $embeddedAlias
                }
                $embedded = Get-EmbeddedMetadata -ExifTool $ExifToolPath -MediaPath $embeddedReadPath -IsQuickTime $isQuickTime
            }
            else {
                $destinationDir = Split-Path -Parent $destination
                Ensure-OutputDirectory -Path $destinationDir
                if ((Test-Path -LiteralPath $destination -PathType Leaf) -and -not $OverwriteOutput) { throw "Destination already exists: $destination" }
                $workingPath = Get-OutputWorkingPath -Destination $destination -EffectiveExtension $ext -NeedsAlias ([bool]$typeInfo.TypeMismatch)
                Copy-Item -LiteralPath $media.FullName -Destination $workingPath -Force
                $embedded = Get-EmbeddedMetadata -ExifTool $ExifToolPath -MediaPath $workingPath -IsQuickTime $isQuickTime
            }
        }
        finally {
            if ($embeddedAlias -and (Test-Path -LiteralPath $embeddedAlias -PathType Leaf)) {
                Remove-Item -LiteralPath $embeddedAlias -Force -ErrorAction SilentlyContinue
            }
        }

        $timeDecision = Get-TimeDecision -Metadata $metadata -Embedded $embedded -IsPhoto $isPhoto -Force:$ForceJsonTime
        $gpsDecision = Get-GpsDecision -Metadata $metadata -Embedded $embedded

        if (-not $timeDecision.WriteDate -and $timeDecision.Reason -like 'preserved existing local time*') { $preservedLocalTime++ }
        if ($gpsDecision.Source -eq 'embedded-preserved') { $preservedGpsPrecision++ }
        if ($gpsDecision.GoogleEditDetected) { $googleGpsEdits++ }

        $gpsText = ''
        if (Test-ValidCoordinates -Latitude $gpsDecision.Latitude -Longitude $gpsDecision.Longitude) {
            $gpsText = ([double]$gpsDecision.Latitude).ToString('0.########', [Globalization.CultureInfo]::InvariantCulture) + ',' +
                       ([double]$gpsDecision.Longitude).ToString('0.########', [Globalization.CultureInfo]::InvariantCulture)
        }
        $timestampText = ''
        if ($null -ne $metadata.UtcTime) { $timestampText = $metadata.UtcTime.ToString('o') }
        $gpsDistanceText = ''
        if ($null -ne $gpsDecision.DistanceMeters) { $gpsDistanceText = ([double]$gpsDecision.DistanceMeters).ToString('0.##', [Globalization.CultureInfo]::InvariantCulture) }

        if ($DryRun) {
            $fixed++
            $row = [PSCustomObject]@{
                Source = $media.FullName; Destination = $destination; Sidecar = $sidecar; MatchMethod = $sidecarMatch.Method;
                DeclaredType = $typeInfo.DeclaredType; DetectedType = $typeInfo.DetectedType; TypeMismatch = $(if ($typeInfo.TypeMismatch) { 'YES' } else { 'NO' }); TypeHandling = $typeInfo.Handling;
                Status = 'DRY RUN - WOULD PROCESS'; TimestampUTC = $timestampText; StoredTimestampUTC = ''; Verification = 'NOT RUN - DRY RUN'; TimeSource = $metadata.TimeSource;
                TimeDecision = $timeDecision.Reason; GPS = $gpsText; GPSDecision = $gpsDecision.Reason; GPSDistanceMeters = $gpsDistanceText; GoogleGeoSource = $metadata.GeoSource; GoogleLocationEdit = $(if ($gpsDecision.GoogleEditDetected) { 'YES' } else { 'NO' });
                VerificationDetails = 'No output file was written.'; Note = ('No source/output media changes made.' + $(if ($typeInfo.TypeMismatch) { ' ' + $typeInfo.Handling } else { '' }))
            }
            Write-AuditRow -Row $row -Columns $auditColumns
            continue
        }

        try {
            $writeResult = Invoke-ExifWrite -ExifTool $ExifToolPath -MediaPath $workingPath -Extension $ext -Metadata $metadata -TimeDecision $timeDecision -Embedded $embedded -GpsDecision $gpsDecision

            $useJsonForFileTime = ($null -ne $metadata.UtcTime -and ($metadata.TimeSource -eq 'photoTakenTime' -or $timeDecision.WriteDate))
            if ($useJsonForFileTime) {
                [IO.File]::SetCreationTimeUtc($workingPath, $metadata.UtcTime)
                [IO.File]::SetLastWriteTimeUtc($workingPath, $metadata.UtcTime)
            }
            else {
                [IO.File]::SetCreationTimeUtc($workingPath, $media.CreationTimeUtc)
                [IO.File]::SetLastWriteTimeUtc($workingPath, $media.LastWriteTimeUtc)
            }

            $verification = [PSCustomObject]@{ Status = 'NOT RUN'; StoredUtc = ''; Details = 'ExifTool write did not succeed.' }
            if ($writeResult.Success) {
                Write-ProgressState -Stage 'Verifying' -Index $index -Total $total -Name $media.Name -Message ('Verifying ' + $media.Name)
                $verification = Verify-OutputMetadata -ExifTool $ExifToolPath -MediaPath $workingPath -Extension $ext -Metadata $metadata -TimeDecision $timeDecision -GpsDecision $gpsDecision -Disabled:$SkipVerification

                if ($verification.Status -eq 'FAIL') {
                    $failed++; $verificationFailed++; $status = 'VERIFY FAILED'
                }
                else {
                    $fixed++
                    if ($verification.Status -eq 'PASS') { $verifiedCount++; $status = 'PROCESSED + VERIFIED' }
                    elseif ($verification.Status -eq 'DISABLED') { $status = 'PROCESSED - VERIFY DISABLED' }
                    else { $status = 'PROCESSED' }
                }
            }
            else {
                $failed++; $status = 'EXIFTOOL WARNING/FAIL'
            }

            if (-not $workingPath.Equals($destination, [StringComparison]::OrdinalIgnoreCase)) {
                if (Test-Path -LiteralPath $destination -PathType Leaf) {
                    if (-not $OverwriteOutput) { throw "Destination already exists: $destination" }
                    Remove-Item -LiteralPath $destination -Force
                }
                Move-Item -LiteralPath $workingPath -Destination $destination -Force
                $workingPath = $destination
            }

            $row = [PSCustomObject]@{
                Source = $media.FullName; Destination = $destination; Sidecar = $sidecar; MatchMethod = $sidecarMatch.Method;
                DeclaredType = $typeInfo.DeclaredType; DetectedType = $typeInfo.DetectedType; TypeMismatch = $(if ($typeInfo.TypeMismatch) { 'YES' } else { 'NO' }); TypeHandling = $typeInfo.Handling;
                Status = $status; TimestampUTC = $timestampText; StoredTimestampUTC = $verification.StoredUtc; Verification = $verification.Status; TimeSource = $metadata.TimeSource;
                TimeDecision = $timeDecision.Reason; GPS = $gpsText; GPSDecision = $gpsDecision.Reason; GPSDistanceMeters = $gpsDistanceText; GoogleGeoSource = $metadata.GeoSource; GoogleLocationEdit = $(if ($gpsDecision.GoogleEditDetected) { 'YES' } else { 'NO' });
                VerificationDetails = $verification.Details; Note = ($writeResult.Message + $(if ($typeInfo.TypeMismatch) { ' | ' + $typeInfo.Handling } else { '' }))
            }
            Write-AuditRow -Row $row -Columns $auditColumns -ForceFlush:($status -like '*FAILED*' -or $status -like '*FAIL*')
        }
        catch {
            if ($workingPath -and -not $workingPath.Equals($destination, [StringComparison]::OrdinalIgnoreCase) -and (Test-Path -LiteralPath $workingPath -PathType Leaf)) {
                try {
                    if ((Test-Path -LiteralPath $destination -PathType Leaf) -and $OverwriteOutput) { Remove-Item -LiteralPath $destination -Force }
                    if (-not (Test-Path -LiteralPath $destination -PathType Leaf)) { Move-Item -LiteralPath $workingPath -Destination $destination -Force }
                }
                catch { Remove-Item -LiteralPath $workingPath -Force -ErrorAction SilentlyContinue }
            }
            $failed++
            $row = [PSCustomObject]@{
                Source = $media.FullName; Destination = $destination; Sidecar = $sidecar; MatchMethod = $sidecarMatch.Method;
                DeclaredType = $typeInfo.DeclaredType; DetectedType = $typeInfo.DetectedType; TypeMismatch = $(if ($typeInfo.TypeMismatch) { 'YES' } else { 'NO' }); TypeHandling = $typeInfo.Handling;
                Status = 'FAILED'; TimestampUTC = $timestampText; StoredTimestampUTC = ''; Verification = 'NOT RUN'; TimeSource = $metadata.TimeSource;
                TimeDecision = $timeDecision.Reason; GPS = $gpsText; GPSDecision = $gpsDecision.Reason; GPSDistanceMeters = $gpsDistanceText; GoogleGeoSource = $metadata.GeoSource; GoogleLocationEdit = $(if ($gpsDecision.GoogleEditDetected) { 'YES' } else { 'NO' });
                VerificationDetails = ''; Note = $_.Exception.Message
            }
            Write-AuditRow -Row $row -Columns $auditColumns -ForceFlush
        }
    }

    if (-not $cancelled -and (Test-CancelRequested)) { $cancelled = $true }

    Write-Progress -Activity 'Restoring Google Photos metadata' -Completed
    Close-AuditLog

    $exifDiagnostics = Stop-ExifToolSession
    if (-not [string]::IsNullOrWhiteSpace($exifDiagnostics)) {
        $diagnosticsPath = Get-AvailableLogPath -PreferredPath (Join-Path (Split-Path -Parent $logPath) 'exiftool-session-diagnostics.txt')
        $exifDiagnostics | Set-Content -LiteralPath $diagnosticsPath -Encoding UTF8
    }
    Disable-SystemAwake

    $runStopwatch.Stop()
    $elapsedSeconds = [Math]::Max(0.001, $runStopwatch.Elapsed.TotalSeconds)
    $rateFileCount = $(if ($cancelled) { $index } else { $total })
    $filesPerSecond = $rateFileCount / $elapsedSeconds

    $summaryFolder = Split-Path -Parent $logPath
    if ([string]::IsNullOrWhiteSpace($summaryFolder)) { $summaryFolder = $InputFolder }
    if ($InternalWorker -and -not [string]::IsNullOrWhiteSpace($WorkerSummaryPath)) {
        $summaryPath = $WorkerSummaryPath
    }
    else {
        $summaryPath = Get-AvailableLogPath -PreferredPath (Join-Path $summaryFolder 'metadata-fix-summary.txt')
    }
    $summaryLines = @(
        ('Google Takeout Metadata Fixer v6.2.3 - Worker ' + $WorkerId + ' Summary'),
        ('Completed: ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz')),
        ('Mode: ' + $(if ($DryRun) { 'DRY RUN' } else { 'CREATE FIXED COPY' })),
        ('Cancelled by user: ' + $(if ($cancelled) { 'YES' } else { 'NO' })),
        ('Input: ' + $InputFolder),
        ('Output: ' + $(if ($DryRun) { '(none - dry run)' } else { $OutputFolder })),
        ('Media files scanned: ' + $total),
        ('Signature-detected unusual extensions: ' + $unknownExtensionMediaDetected),
        ('Media bytes scanned: ' + $totalMediaBytes + ' (' + (Format-ByteSize $totalMediaBytes) + ')'),
        ('Processed/matched: ' + $fixed),
        ('Verified timestamps: ' + $verifiedCount),
        ('Verification failures: ' + $verificationFailed),
        ('Copied without sidecar: ' + $copiedNoSidecar),
        ('Type mismatches handled: ' + $typeMismatchCount),
        ('Local EXIF times preserved: ' + $preservedLocalTime),
        ('Precise embedded GPS preserved: ' + $preservedGpsPrecision),
        ('Google GPS edits applied: ' + $googleGpsEdits),
        ('Failed/warned: ' + $failed),
        ('Scan errors/skips: ' + $scanErrorCount),
        ('Elapsed: ' + $runStopwatch.Elapsed.ToString()),
        ('Average rate: ' + $filesPerSecond.ToString('0.00', [Globalization.CultureInfo]::InvariantCulture) + ' files/sec'),
        ('ExifTool engine: ' + $script:ExifToolMode),
        ('ExifTool commands: ' + $script:ExifToolCommandCount),
        ('Sidecar folder-cache limit: ' + $script:MaxJsonFolderIndexes),
        ('Verification: ' + $(if ($SkipVerification) { 'disabled by user' } else { 'enabled' })),
        ('CSV audit log: ' + $logPath),
        ('ExifTool diagnostics: ' + $(if ([string]::IsNullOrWhiteSpace($diagnosticsPath)) { '(none)' } else { $diagnosticsPath }))
    )
    $summaryLines | Set-Content -LiteralPath $summaryPath -Encoding UTF8

    if ($InternalWorker -and -not [string]::IsNullOrWhiteSpace($WorkerResultPath)) {
        $workerResult = [ordered]@{
            WorkerId = $WorkerId
            Success = $true
            FatalMessage = ''
            Cancelled = $cancelled
            Total = $total
            Attempted = $index
            Processed = $fixed
            Verified = $verifiedCount
            VerificationFailed = $verificationFailed
            Unmatched = $copiedNoSidecar
            Failed = $failed
            TypeMismatches = $typeMismatchCount
            PreservedLocalTime = $preservedLocalTime
            PreservedGpsPrecision = $preservedGpsPrecision
            GoogleGpsEdits = $googleGpsEdits
            ScanErrors = $scanErrorCount
            UnknownExtensionMedia = $unknownExtensionMediaDetected
            ExifToolCommands = $script:ExifToolCommandCount
            ExifToolMode = $script:ExifToolMode
            ElapsedSeconds = [Math]::Round($elapsedSeconds, 3)
            LogPath = $logPath
            SummaryPath = $summaryPath
            DiagnosticsPath = $diagnosticsPath
        }
        ($workerResult | ConvertTo-Json -Compress -Depth 4) | Set-Content -LiteralPath $WorkerResultPath -Encoding UTF8
    }

    $finalStage = $(if ($cancelled) { 'Cancelled' } else { 'Done' })
    $finalMessage = $(if ($cancelled) { 'Cancelled safely after the current file.' } else { 'Finished.' })
    Write-ProgressState -Stage $finalStage -Index $index -Total $total -Message $finalMessage -Extra @{
        Cancelled = $cancelled
        Processed = $fixed
        Verified = $verifiedCount
        VerificationFailed = $verificationFailed
        TypeMismatches = $typeMismatchCount
        UnknownExtensionMedia = $unknownExtensionMediaDetected
        Unmatched = $copiedNoSidecar
        PreservedLocalTime = $preservedLocalTime
        PreservedGpsPrecision = $preservedGpsPrecision
        GoogleGpsEdits = $googleGpsEdits
        Failed = $failed
        ScanErrors = $scanErrorCount
        LogPath = $logPath
        SummaryPath = $summaryPath
        DiagnosticsPath = $diagnosticsPath
        OutputFolder = $(if ($DryRun) { '' } else { $OutputFolder })
        ElapsedSeconds = [Math]::Round($elapsedSeconds, 1)
        FilesPerSecond = [Math]::Round($filesPerSecond, 2)
        ExifToolMode = $script:ExifToolMode
        ExifToolCommands = $script:ExifToolCommandCount
    }

    Write-Host ''
    if ($cancelled) { Write-Host 'Cancelled safely.' -ForegroundColor Yellow } else { Write-Host 'Finished.' -ForegroundColor Green }
    Write-Host ('Processed/matched:        ' + $fixed)
    Write-Host ('Verified timestamps:      ' + $verifiedCount)
    Write-Host ('Verification failures:    ' + $verificationFailed)
    Write-Host ('Type mismatches handled:  ' + $typeMismatchCount)
    Write-Host ('Unusual extensions found: ' + $unknownExtensionMediaDetected)
    Write-Host ('Copied without sidecar:   ' + $copiedNoSidecar)
    Write-Host ('Local EXIF time kept:      ' + $preservedLocalTime)
    Write-Host ('Precise GPS preserved:    ' + $preservedGpsPrecision)
    Write-Host ('Google GPS edits applied: ' + $googleGpsEdits)
    Write-Host ('Failed/warned:             ' + $failed)
    Write-Host ('Folders/files skipped:    ' + $scanErrorCount)
    Write-Host ('Elapsed:                   ' + $runStopwatch.Elapsed.ToString())
    Write-Host ('Average rate:              ' + $filesPerSecond.ToString('0.00', [Globalization.CultureInfo]::InvariantCulture) + ' files/sec')
    Write-Host ('ExifTool commands:         ' + $script:ExifToolCommandCount + ' (' + $script:ExifToolMode + ')')
    Write-Host ('Log:                       ' + $logPath)
    Write-Host ('Run summary:               ' + $summaryPath)
    if (-not [string]::IsNullOrWhiteSpace($diagnosticsPath)) { Write-Host ('ExifTool diagnostics:       ' + $diagnosticsPath) }
    if (-not $DryRun) { Write-Host ('Fixed library:             ' + $OutputFolder) }
    Write-Host ''

    if ($verificationFailed -gt 0) {
        Write-Host 'WARNING: One or more files failed timestamp verification. Review the CSV before uploading.' -ForegroundColor Red
    }
    elseif ($cancelled) {
        Write-Host 'Cancellation was graceful; the CSV contains the files completed before the stop request.' -ForegroundColor Yellow
    }
    else {
        Write-Host 'Post-write timestamp verification found no mismatches in files that were eligible for verification.' -ForegroundColor Green
    }
    Write-Host 'Spot-check dates, travel photos, videos, GPS and a few edited dates before uploading the full library.' -ForegroundColor Yellow
    if (-not $ForceJsonTime) {
        Write-Host 'If you intentionally changed photo dates inside Google Photos, rerun a test with -ForceJsonTime and compare.' -ForegroundColor Yellow
    }
}
catch {
    $fatalMessage = $_.Exception.Message
    if ($InternalWorker -and -not [string]::IsNullOrWhiteSpace($WorkerResultPath)) {
        try {
            $failureResult = [ordered]@{ WorkerId = $WorkerId; Success = $false; FatalMessage = $fatalMessage; Cancelled = (Test-CancelRequested); Total = 0; Attempted = 0; Processed = 0; Verified = 0; VerificationFailed = 0; Unmatched = 0; Failed = 1; TypeMismatches = 0; PreservedLocalTime = 0; PreservedGpsPrecision = 0; GoogleGpsEdits = 0; ScanErrors = 0; UnknownExtensionMedia = 0; ExifToolCommands = $script:ExifToolCommandCount; ExifToolMode = $script:ExifToolMode; ElapsedSeconds = 0; LogPath = $logPath; SummaryPath = $summaryPath; DiagnosticsPath = $diagnosticsPath }
            ($failureResult | ConvertTo-Json -Compress -Depth 4) | Set-Content -LiteralPath $WorkerResultPath -Encoding UTF8
        }
        catch { }
    }
    Write-ProgressState -Stage 'Failed' -Message $fatalMessage -Extra @{ LogPath = $logPath; OutputFolder = $(if ($DryRun) { '' } else { $OutputFolder }) }
}
finally {
    try { Close-AuditLog } catch { }
    try { [void](Stop-ExifToolSession) } catch { }
    try { Disable-SystemAwake } catch { }
}

if (-not [string]::IsNullOrWhiteSpace($fatalMessage)) {
    Write-Error -Message $fatalMessage -ErrorAction Continue
    exit 1
}
