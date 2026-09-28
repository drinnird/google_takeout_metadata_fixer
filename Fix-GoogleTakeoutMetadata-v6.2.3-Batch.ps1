<#
.SYNOPSIS
  Batch coordinator for Google Takeout Metadata Fixer v6.2.3.

.DESCRIPTION
  Processes multiple input/output folder pairs sequentially. Each folder pair is
  handed to the tested v6.2.3 parallel core, so the selected worker pool is used
  inside one folder at a time without multiplying disk contention across folders.
  Every input folder keeps its own independent output folder, CSV log, and run
  summary. The source Takeout folders are never modified.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestPath,

    [string]$ExifToolPath,

    [switch]$DryRun,
    [switch]$ForceJsonTime,
    [switch]$TopLevelOnly,
    [switch]$OverwriteOutput,
    [switch]$SkipVerification,
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

$script:LastProgressWriteUtc = [DateTime]::MinValue
$script:ActiveProcess = $null
$script:ActiveChildCancelFile = ''
$script:FatalMessage = ''

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

function Read-JsonFileSafe {
    param([string]$Path)
    try {
        if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
        $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
        return ($raw | ConvertFrom-Json)
    }
    catch { return $null }
}

function Test-CancelRequested {
    return (-not [string]::IsNullOrWhiteSpace($CancelFile) -and (Test-Path -LiteralPath $CancelFile -PathType Leaf))
}

function Signal-ChildCancel {
    if ([string]::IsNullOrWhiteSpace($script:ActiveChildCancelFile)) { return }
    try { 'cancel' | Set-Content -LiteralPath $script:ActiveChildCancelFile -Encoding ASCII -Force } catch { }
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

function Write-BatchProgress {
    param(
        [string]$Stage,
        [int]$FolderIndex = 0,
        [int]$FolderTotal = 0,
        [int]$Percent = 0,
        [string]$CurrentInput = '',
        [string]$CurrentOutput = '',
        [string]$Message = '',
        [hashtable]$Extra = $null,
        [switch]$Force
    )

    if ([string]::IsNullOrWhiteSpace($ProgressFile)) { return }
    $nowUtc = [DateTime]::UtcNow
    $urgent = $Force -or ($Stage -in @('BatchStarting','Done','Cancelled','Failed'))
    if (-not $urgent -and (($nowUtc - $script:LastProgressWriteUtc).TotalMilliseconds -lt 150)) { return }
    $script:LastProgressWriteUtc = $nowUtc

    try {
        if ($Percent -lt 0) { $Percent = 0 }
        if ($Percent -gt 100) { $Percent = 100 }
        $state = [ordered]@{
            Stage = $Stage
            Index = $FolderIndex
            Total = $FolderTotal
            Percent = $Percent
            FolderIndex = $FolderIndex
            FolderTotal = $FolderTotal
            CurrentInput = $CurrentInput
            CurrentOutput = $CurrentOutput
            Message = $Message
        }
        if ($Extra) {
            foreach ($key in $Extra.Keys) { $state[$key] = $Extra[$key] }
        }
        $parent = Split-Path -Parent $ProgressFile
        if ($parent -and -not (Test-Path -LiteralPath $parent -PathType Container)) {
            New-Item -ItemType Directory -Path $parent -Force | Out-Null
        }
        ($state | ConvertTo-Json -Compress -Depth 10) | Set-Content -LiteralPath $ProgressFile -Encoding UTF8
    }
    catch { }
}

function Get-AvailablePath {
    param([string]$PreferredPath)
    if (-not (Test-Path -LiteralPath $PreferredPath)) { return $PreferredPath }
    $dir = Split-Path -Parent $PreferredPath
    $base = [IO.Path]::GetFileNameWithoutExtension($PreferredPath)
    $ext = [IO.Path]::GetExtension($PreferredPath)
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $candidate = Join-Path $dir ($base + '-' + $stamp + $ext)
    $n = 1
    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path $dir ($base + '-' + $stamp + '-' + $n + $ext)
        $n++
    }
    return $candidate
}

function Get-SummaryValue {
    param($State, [string]$Name, $DefaultValue)
    if ($null -ne $State -and $State.PSObject.Properties[$Name]) { return $State.$Name }
    return $DefaultValue
}

$batchStopwatch = [Diagnostics.Stopwatch]::StartNew()
$jobRoot = ''
$batchSummaryPath = ''
$results = New-Object System.Collections.Generic.List[object]

try {
    $scriptDir = $PSScriptRoot
    if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
    $coreScript = Join-Path $scriptDir 'Fix-GoogleTakeoutMetadata-v6.2.3.ps1'
    if (-not (Test-Path -LiteralPath $coreScript -PathType Leaf)) { throw 'Fix-GoogleTakeoutMetadata-v6.2.3.ps1 is missing from the package.' }
    if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) { throw "Batch manifest does not exist: $ManifestPath" }

    $manifestRaw = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $jobs = @($manifestRaw)
    if ($jobs.Count -eq 0) { throw 'The batch manifest does not contain any folder pairs.' }

    foreach ($job in $jobs) {
        if (-not $job.PSObject.Properties['InputFolder'] -or [string]::IsNullOrWhiteSpace([string]$job.InputFolder)) {
            throw 'A batch row is missing its input folder.'
        }
        if (-not $DryRun -and (-not $job.PSObject.Properties['OutputFolder'] -or [string]::IsNullOrWhiteSpace([string]$job.OutputFolder))) {
            throw ('A batch row is missing its output folder: ' + [string]$job.InputFolder)
        }
    }

    $jobRoot = Join-Path ([IO.Path]::GetTempPath()) ('GoogleTakeoutMetadataFixer-v6.2.3-batch-' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $jobRoot -Force | Out-Null

    Write-BatchProgress -Stage 'BatchStarting' -FolderIndex 0 -FolderTotal $jobs.Count -Percent 0 -Message ('Preparing ' + $jobs.Count + ' folder pair(s).') -Force

    $cancelled = $false
    $folderFailures = 0
    $aggregateProcessed = 0
    $aggregateVerified = 0
    $aggregateVerificationFailed = 0
    $aggregateUnmatched = 0
    $aggregateFailed = 0
    $aggregateScanErrors = 0
    $aggregateTypeMismatches = 0
    $aggregateUnknownExtensions = 0
    $aggregateExifCommands = 0

    for ($i = 0; $i -lt $jobs.Count; $i++) {
        if (Test-CancelRequested) { $cancelled = $true; break }

        $job = $jobs[$i]
        $inputFolder = [string]$job.InputFolder
        $outputFolder = if ($job.PSObject.Properties['OutputFolder']) { [string]$job.OutputFolder } else { '' }
        $folderNumber = $i + 1
        $leaf = Split-Path -Leaf $inputFolder
        if ([string]::IsNullOrWhiteSpace($leaf)) { $leaf = $inputFolder }

        $childProgress = Join-Path $jobRoot ('folder-' + $folderNumber + '-progress.json')
        $childCancel = Join-Path $jobRoot ('folder-' + $folderNumber + '.cancel')
        Remove-Item -LiteralPath $childCancel -Force -ErrorAction SilentlyContinue
        $script:ActiveChildCancelFile = $childCancel

        $parts = New-Object System.Collections.Generic.List[string]
        $parts.Add('& ' + (ConvertTo-PowerShellLiteral $coreScript))
        $parts.Add('-InputFolder ' + (ConvertTo-PowerShellLiteral $inputFolder))
        if (-not $DryRun) { $parts.Add('-OutputFolder ' + (ConvertTo-PowerShellLiteral $outputFolder)) }
        if (-not [string]::IsNullOrWhiteSpace($ExifToolPath)) { $parts.Add('-ExifToolPath ' + (ConvertTo-PowerShellLiteral $ExifToolPath)) }
        $parts.Add('-ProgressFile ' + (ConvertTo-PowerShellLiteral $childProgress))
        $parts.Add('-CancelFile ' + (ConvertTo-PowerShellLiteral $childCancel))
        $parts.Add('-WorkerCount ' + $WorkerCount)
        $parts.Add('-GpsAgreementMeters ' + $GpsAgreementMeters.ToString([Globalization.CultureInfo]::InvariantCulture))
        if ($DryRun) { $parts.Add('-DryRun') }
        if ($ForceJsonTime) { $parts.Add('-ForceJsonTime') }
        if ($TopLevelOnly) { $parts.Add('-TopLevelOnly') }
        if ($OverwriteOutput -and -not $DryRun) { $parts.Add('-OverwriteOutput') }
        if ($SkipVerification -and -not $DryRun) { $parts.Add('-SkipVerification') }
        if ($AllowSystemSleep) { $parts.Add('-AllowSystemSleep') }

        $commandText = $parts -join ' '
        $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($commandText))
        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = Get-PowerShellExe
        $startInfo.Arguments = '-NoProfile -ExecutionPolicy Bypass -EncodedCommand ' + $encoded
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $startInfo
        if (-not $process.Start()) { throw ('Windows could not start folder job: ' + $inputFolder) }
        $script:ActiveProcess = $process
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        $cancelSignalled = $false

        while (-not $process.HasExited) {
            if (Test-CancelRequested) {
                $cancelled = $true
                if (-not $cancelSignalled) {
                    Signal-ChildCancel
                    $cancelSignalled = $true
                }
            }

            $childState = Read-JsonFileSafe -Path $childProgress
            $childPercent = 0
            $childMessage = ''
            $childStage = ''
            if ($childState) {
                if ($childState.PSObject.Properties['Percent']) { $childPercent = [int]$childState.Percent }
                if ($childState.PSObject.Properties['Message']) { $childMessage = [string]$childState.Message }
                if ($childState.PSObject.Properties['Stage']) { $childStage = [string]$childState.Stage }
            }
            $overall = [int][Math]::Min(99, [Math]::Max(0, ((($i + ($childPercent / 100.0)) / [double]$jobs.Count) * 100)))
            $message = ('Folder {0}/{1}: {2}' -f $folderNumber, $jobs.Count, $leaf)
            if (-not [string]::IsNullOrWhiteSpace($childMessage)) { $message += (' - ' + $childMessage) }
            Write-BatchProgress -Stage 'BatchProcessing' -FolderIndex $folderNumber -FolderTotal $jobs.Count -Percent $overall -CurrentInput $inputFolder -CurrentOutput $outputFolder -Message $message -Extra @{ ChildStage=$childStage; ChildPercent=$childPercent; CompletedFolders=$i }
            Start-Sleep -Milliseconds 250
        }

        $exitCode = $process.ExitCode
        $stdout = ''
        $stderr = ''
        try { $stdout = [string]$stdoutTask.Result } catch { }
        try { $stderr = [string]$stderrTask.Result } catch { }
        $finalState = Read-JsonFileSafe -Path $childProgress

        $folderStatus = 'DONE'
        if ($cancelled -or ($finalState -and $finalState.PSObject.Properties['Cancelled'] -and [bool]$finalState.Cancelled)) {
            $folderStatus = 'CANCELLED'
            $cancelled = $true
        }
        elseif ($exitCode -ne 0 -or $null -eq $finalState -or
                ($finalState.PSObject.Properties['Stage'] -and [string]$finalState.Stage -eq 'Failed') -or
                ($finalState.PSObject.Properties['CompletenessCheck'] -and [string]$finalState.CompletenessCheck -eq 'FAIL')) {
            $folderStatus = 'FAILED'
            $folderFailures++
        }

        $result = [ordered]@{
            InputFolder = $inputFolder
            OutputFolder = $(if ($DryRun) { '' } else { $outputFolder })
            Status = $folderStatus
            ExitCode = $exitCode
            Processed = [int](Get-SummaryValue -State $finalState -Name 'Processed' -DefaultValue 0)
            Verified = [int](Get-SummaryValue -State $finalState -Name 'Verified' -DefaultValue 0)
            VerificationFailed = [int](Get-SummaryValue -State $finalState -Name 'VerificationFailed' -DefaultValue 0)
            Unmatched = [int](Get-SummaryValue -State $finalState -Name 'Unmatched' -DefaultValue 0)
            Failed = [int](Get-SummaryValue -State $finalState -Name 'Failed' -DefaultValue 0)
            ScanErrors = [int](Get-SummaryValue -State $finalState -Name 'ScanErrors' -DefaultValue 0)
            TypeMismatches = [int](Get-SummaryValue -State $finalState -Name 'TypeMismatches' -DefaultValue 0)
            UnknownExtensionMedia = [int](Get-SummaryValue -State $finalState -Name 'UnknownExtensionMedia' -DefaultValue 0)
            ExifToolCommands = [int](Get-SummaryValue -State $finalState -Name 'ExifToolCommands' -DefaultValue 0)
            LogPath = [string](Get-SummaryValue -State $finalState -Name 'LogPath' -DefaultValue '')
            SummaryPath = [string](Get-SummaryValue -State $finalState -Name 'SummaryPath' -DefaultValue '')
            Message = [string](Get-SummaryValue -State $finalState -Name 'Message' -DefaultValue '')
        }
        if ($folderStatus -eq 'FAILED' -and [string]::IsNullOrWhiteSpace($result.Message)) {
            $msg = $stderr.Trim()
            if ([string]::IsNullOrWhiteSpace($msg)) { $msg = $stdout.Trim() }
            $result.Message = $msg
        }
        $results.Add([PSCustomObject]$result)

        $aggregateProcessed += $result.Processed
        $aggregateVerified += $result.Verified
        $aggregateVerificationFailed += $result.VerificationFailed
        $aggregateUnmatched += $result.Unmatched
        $aggregateFailed += $result.Failed
        $aggregateScanErrors += $result.ScanErrors
        $aggregateTypeMismatches += $result.TypeMismatches
        $aggregateUnknownExtensions += $result.UnknownExtensionMedia
        $aggregateExifCommands += $result.ExifToolCommands

        try { $process.Dispose() } catch { }
        $script:ActiveProcess = $null
        $script:ActiveChildCancelFile = ''

        $completedPercent = [int][Math]::Min(100, (($folderNumber / [double]$jobs.Count) * 100))
        Write-BatchProgress -Stage 'BatchProcessing' -FolderIndex $folderNumber -FolderTotal $jobs.Count -Percent $completedPercent -CurrentInput $inputFolder -CurrentOutput $outputFolder -Message ('Completed folder {0}/{1}: {2} [{3}]' -f $folderNumber, $jobs.Count, $leaf, $folderStatus) -Extra @{ CompletedFolders=$folderNumber }

        if ($cancelled) { break }
    }

    $batchStopwatch.Stop()

    # Keep the aggregate summary with fixed output when possible, rather than
    # writing into the source Takeout tree. Dry runs have no output, so their
    # aggregate summary is written beside the first selected input folder.
    $batchSummaryDir = ''
    if (-not $DryRun) {
        foreach ($r in $results) {
            if (-not [string]::IsNullOrWhiteSpace([string]$r.OutputFolder) -and (Test-Path -LiteralPath ([string]$r.OutputFolder) -PathType Container)) {
                $batchSummaryDir = [string]$r.OutputFolder
                break
            }
        }
        if ([string]::IsNullOrWhiteSpace($batchSummaryDir)) {
            $batchSummaryDir = Split-Path -Parent ([string]$jobs[0].OutputFolder)
        }
    }
    else {
        $batchSummaryDir = Split-Path -Parent ([string]$jobs[0].InputFolder)
        if ([string]::IsNullOrWhiteSpace($batchSummaryDir)) { $batchSummaryDir = [string]$jobs[0].InputFolder }
    }
    if (-not [string]::IsNullOrWhiteSpace($batchSummaryDir) -and (Test-Path -LiteralPath $batchSummaryDir -PathType Container)) {
        $batchSummaryPath = Get-AvailablePath -PreferredPath (Join-Path $batchSummaryDir 'metadata-fix-batch-summary.txt')
    }

    if (-not [string]::IsNullOrWhiteSpace($batchSummaryPath)) {
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add('Google Takeout Metadata Fixer v6.2.3 - Multi-folder Batch Summary')
        $lines.Add('Completed: ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz'))
        $lines.Add('Mode: ' + $(if ($DryRun) { 'DRY RUN' } else { 'CREATE FIXED COPIES' }))
        $lines.Add('Cancelled by user: ' + $(if ($cancelled) { 'YES' } else { 'NO' }))
        $lines.Add('Folder pairs requested: ' + $jobs.Count)
        $lines.Add('Folder pairs completed/attempted: ' + $results.Count)
        $lines.Add('Folder failures: ' + $folderFailures)
        $lines.Add('Workers per folder: ' + $WorkerCount)
        $lines.Add('Processed/matched: ' + $aggregateProcessed)
        $lines.Add('Verified timestamps: ' + $aggregateVerified)
        $lines.Add('Verification failures: ' + $aggregateVerificationFailed)
        $lines.Add('Copied without sidecar: ' + $aggregateUnmatched)
        $lines.Add('Failed/warned: ' + $aggregateFailed)
        $lines.Add('Scan errors/skips: ' + $aggregateScanErrors)
        $lines.Add('Type mismatches handled: ' + $aggregateTypeMismatches)
        $lines.Add('Signature-detected unusual extensions: ' + $aggregateUnknownExtensions)
        $lines.Add('ExifTool commands: ' + $aggregateExifCommands)
        $lines.Add('Elapsed: ' + $batchStopwatch.Elapsed.ToString())
        $lines.Add('')
        $lines.Add('Folder results:')
        foreach ($r in $results) {
            $lines.Add(('- [{0}] {1}' -f $r.Status, $r.InputFolder))
            if (-not $DryRun) { $lines.Add(('    Output: {0}' -f $r.OutputFolder)) }
            if (-not [string]::IsNullOrWhiteSpace($r.LogPath)) { $lines.Add(('    CSV: {0}' -f $r.LogPath)) }
            if (-not [string]::IsNullOrWhiteSpace($r.SummaryPath)) { $lines.Add(('    Summary: {0}' -f $r.SummaryPath)) }
            if (-not [string]::IsNullOrWhiteSpace($r.Message)) { $lines.Add(('    Note: {0}' -f $r.Message)) }
        }
        $lines | Set-Content -LiteralPath $batchSummaryPath -Encoding UTF8
    }

    $stage = if ($cancelled) { 'Cancelled' } else { 'Done' }
    $message = if ($cancelled) {
        'Batch cancelled safely. Completed folder outputs remain intact.'
    }
    elseif ($folderFailures -gt 0) {
        ('Batch finished with ' + $folderFailures + ' folder failure(s). Review the per-folder logs and summaries.')
    }
    else {
        ('Finished ' + $results.Count + ' folder pair(s). Each input was written to its own output folder.')
    }

    Write-BatchProgress -Stage $stage -FolderIndex $results.Count -FolderTotal $jobs.Count -Percent $(if ($cancelled) { [int](($results.Count / [double]$jobs.Count) * 100) } else { 100 }) -Message $message -Extra @{
        Cancelled=$cancelled
        FolderFailures=$folderFailures
        FolderResults=$results.ToArray()
        BatchSummaryPath=$batchSummaryPath
        Processed=$aggregateProcessed
        Verified=$aggregateVerified
        VerificationFailed=$aggregateVerificationFailed
        Unmatched=$aggregateUnmatched
        Failed=$aggregateFailed
        ScanErrors=$aggregateScanErrors
        TypeMismatches=$aggregateTypeMismatches
        UnknownExtensionMedia=$aggregateUnknownExtensions
        ExifToolCommands=$aggregateExifCommands
        WorkerCount=$WorkerCount
        ElapsedSeconds=[Math]::Round($batchStopwatch.Elapsed.TotalSeconds,1)
    } -Force
}
catch {
    $script:FatalMessage = $_.Exception.Message
    Write-BatchProgress -Stage 'Failed' -FolderIndex $results.Count -FolderTotal 0 -Percent 0 -Message $script:FatalMessage -Extra @{ FolderResults=$results.ToArray(); BatchSummaryPath=$batchSummaryPath } -Force
}
finally {
    if ($script:ActiveProcess -and -not $script:ActiveProcess.HasExited) {
        try { Signal-ChildCancel } catch { }
        Start-Sleep -Milliseconds 500
        if (-not $script:ActiveProcess.HasExited) { try { Stop-ProcessTree -Process $script:ActiveProcess } catch { } }
    }
    if (-not [string]::IsNullOrWhiteSpace($jobRoot) -and (Test-Path -LiteralPath $jobRoot -PathType Container)) {
        try { Remove-Item -LiteralPath $jobRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }
    }
}

if (-not [string]::IsNullOrWhiteSpace($script:FatalMessage)) {
    Write-Error -Message $script:FatalMessage -ErrorAction Continue
    exit 1
}
