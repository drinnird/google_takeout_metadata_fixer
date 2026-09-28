<#
.SYNOPSIS
  Windows Forms front end for Google Takeout Metadata Fixer v6.2.3.

.DESCRIPTION
  Supports one or many independent input/output folder pairs. Each input folder
  receives its own output folder (default: "<input> - Fixed"). Folder pairs are
  processed sequentially to avoid multiplying disk contention, while the v6.2.3
  core uses the selected multi-worker pool inside each folder.
#>

[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptDir = $PSScriptRoot
if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
$coreScript = Join-Path $scriptDir 'Fix-GoogleTakeoutMetadata-v6.2.3.ps1'
$workerScript = Join-Path $scriptDir 'Fix-GoogleTakeoutMetadata-v6.2.3-Worker.ps1'
$batchScript = Join-Path $scriptDir 'Fix-GoogleTakeoutMetadata-v6.2.3-Batch.ps1'

function Show-ErrorMessage {
    param([string]$Message)
    [System.Windows.Forms.MessageBox]::Show($Message, 'Google Takeout Metadata Fixer', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
}

function Show-WarningMessage {
    param([string]$Message)
    [System.Windows.Forms.MessageBox]::Show($Message, 'Google Takeout Metadata Fixer', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
}

foreach ($required in @($coreScript, $workerScript, $batchScript)) {
    if (-not (Test-Path -LiteralPath $required -PathType Leaf)) {
        Show-ErrorMessage -Message ('Could not find required script: ' + (Split-Path -Leaf $required))
        exit 1
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
    $parent = Get-NormalizedDirectoryPath $ParentPath
    $child = Get-NormalizedDirectoryPath $ChildPath
    if ($parent.Equals($child, [StringComparison]::OrdinalIgnoreCase)) { return $true }
    $prefix = $parent
    if (-not $prefix.EndsWith([string][IO.Path]::DirectorySeparatorChar)) { $prefix += [IO.Path]::DirectorySeparatorChar }
    return $child.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)
}

function Get-DefaultOutputFolder {
    param([string]$InputFolder)
    if ([string]::IsNullOrWhiteSpace($InputFolder)) { return '' }
    try {
        $normalized = Get-NormalizedDirectoryPath $InputFolder
        $parent = Split-Path -Parent $normalized
        $leaf = Split-Path -Leaf $normalized
        if (-not $parent) { $parent = [IO.Path]::GetPathRoot($normalized) }
        if (-not $leaf) { $leaf = 'Google Photos' }
        return (Join-Path $parent ($leaf + ' - Fixed'))
    }
    catch { return '' }
}

function Find-ExifToolLocal {
    foreach ($candidate in @((Join-Path $scriptDir 'exiftool.exe'), (Join-Path $scriptDir 'exiftool(-k).exe'))) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }
    foreach ($folder in @(Get-ChildItem -LiteralPath $scriptDir -Directory -Filter 'exiftool-*' -ErrorAction SilentlyContinue)) {
        foreach ($name in @('exiftool.exe', 'exiftool(-k).exe')) {
            $candidate = Join-Path $folder.FullName $name
            if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
        }
    }
    $command = Get-Command exiftool.exe -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }
    return ''
}

function Test-CoreScriptSyntax {
    $problems = New-Object System.Collections.Generic.List[string]
    foreach ($scriptPath in @($coreScript, $workerScript, $batchScript)) {
        $tokens = $null
        $parseErrors = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
        if ($parseErrors -and $parseErrors.Count -gt 0) {
            foreach ($parseError in $parseErrors) {
                $problems.Add(('{0} - line {1}, column {2}: {3}' -f (Split-Path -Leaf $scriptPath), $parseError.Extent.StartLineNumber, $parseError.Extent.StartColumnNumber, $parseError.Message))
            }
        }
    }
    return ($problems -join [Environment]::NewLine)
}

function ConvertTo-PowerShellLiteral {
    param([string]$Value)
    return "'" + $Value.Replace("'", "''") + "'"
}

function Show-FolderPicker {
    param([string]$Description, [string]$InitialPath, [bool]$AllowNewFolder)
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Description
    $dialog.ShowNewFolderButton = $AllowNewFolder
    if (-not [string]::IsNullOrWhiteSpace($InitialPath) -and (Test-Path -LiteralPath $InitialPath -PathType Container)) { $dialog.SelectedPath = $InitialPath }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dialog.SelectedPath }
    return $null
}

function Initialize-NativeMultiFolderPicker {
    # FolderBrowserDialog only supports one folder. The Vista+ IFileOpenDialog
    # supports FOS_PICKFOLDERS + FOS_ALLOWMULTISELECT, which gives the familiar
    # Explorer picker with Ctrl-click and Shift-click multi-selection.
    if ('GoogleTakeoutMetadataFixer.NativeMultiFolderPicker' -as [type]) { return $true }

    $source = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace GoogleTakeoutMetadataFixer
{
    [Flags]
    internal enum FOS : uint
    {
        PICKFOLDERS       = 0x00000020,
        FORCEFILESYSTEM   = 0x00000040,
        ALLOWMULTISELECT  = 0x00000200,
        PATHMUSTEXIST     = 0x00000800,
        DONTADDTORECENT   = 0x02000000
    }

    internal enum SIGDN : uint
    {
        FILESYSPATH = 0x80058000
    }

    internal enum FDAP : uint
    {
        BOTTOM = 0x00000000,
        TOP    = 0x00000001
    }

    [ComImport]
    [Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IShellItem
    {
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        void GetDisplayName(SIGDN sigdnName, out IntPtr ppszName);
        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        void Compare(IShellItem psi, uint hint, out int piOrder);
    }

    [ComImport]
    [Guid("B63EA76D-1F85-456F-A19C-48159EFA858B")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IShellItemArray
    {
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppvOut);
        void GetPropertyStore(int flags, ref Guid riid, out IntPtr ppv);
        void GetPropertyDescriptionList(IntPtr keyType, ref Guid riid, out IntPtr ppv);
        void GetAttributes(uint attribFlags, uint sfgaoMask, out uint psfgaoAttribs);
        void GetCount(out uint pdwNumItems);
        void GetItemAt(uint dwIndex, out IShellItem ppsi);
    }

    [ComImport]
    [Guid("D57C7288-D4AD-4768-BE02-9D969532D960")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IFileOpenDialog
    {
        [PreserveSig]
        int Show(IntPtr parent);
        void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
        void SetFileTypeIndex(uint iFileType);
        void GetFileTypeIndex(out uint piFileType);
        void Advise(IntPtr pfde, out uint pdwCookie);
        void Unadvise(uint dwCookie);
        void SetOptions(FOS fos);
        void GetOptions(out FOS pfos);
        void SetDefaultFolder(IShellItem psi);
        void SetFolder(IShellItem psi);
        void GetFolder(out IShellItem ppsi);
        void GetCurrentSelection(out IShellItem ppsi);
        void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetFileName(out IntPtr pszName);
        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
        void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        void GetResult(out IShellItem ppsi);
        void AddPlace(IShellItem psi, FDAP fdap);
        void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
        void Close(int hr);
        void SetClientGuid(ref Guid guid);
        void ClearClientData();
        void SetFilter(IntPtr pFilter);
        void GetResults(out IShellItemArray ppenum);
        void GetSelectedItems(out IShellItemArray ppsai);
    }

    [ComImport]
    [Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")]
    internal class FileOpenDialogRCW
    {
    }

    public static class NativeMultiFolderPicker
    {
        private const int ERROR_CANCELLED_HRESULT = unchecked((int)0x800704C7);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
        private static extern int SHCreateItemFromParsingName(
            [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
            IntPtr pbc,
            ref Guid riid,
            [MarshalAs(UnmanagedType.Interface)] out IShellItem ppv);

        private static string GetPath(IShellItem item)
        {
            if (item == null) return null;
            IntPtr ptr = IntPtr.Zero;
            try
            {
                item.GetDisplayName(SIGDN.FILESYSPATH, out ptr);
                return ptr == IntPtr.Zero ? null : Marshal.PtrToStringUni(ptr);
            }
            finally
            {
                if (ptr != IntPtr.Zero) Marshal.FreeCoTaskMem(ptr);
            }
        }

        public static string[] PickFolders(IntPtr owner, string title, string initialPath)
        {
            IFileOpenDialog dialog = null;
            IShellItem initialItem = null;
            IShellItemArray results = null;
            try
            {
                dialog = (IFileOpenDialog)new FileOpenDialogRCW();
                FOS options;
                dialog.GetOptions(out options);
                options |= FOS.PICKFOLDERS | FOS.FORCEFILESYSTEM | FOS.ALLOWMULTISELECT | FOS.PATHMUSTEXIST | FOS.DONTADDTORECENT;
                dialog.SetOptions(options);
                if (!String.IsNullOrWhiteSpace(title)) dialog.SetTitle(title);
                dialog.SetOkButtonLabel("Add selected folders");

                if (!String.IsNullOrWhiteSpace(initialPath) && Directory.Exists(initialPath))
                {
                    Guid iid = typeof(IShellItem).GUID;
                    if (SHCreateItemFromParsingName(initialPath, IntPtr.Zero, ref iid, out initialItem) == 0 && initialItem != null)
                    {
                        dialog.SetDefaultFolder(initialItem);
                        dialog.SetFolder(initialItem);
                    }
                }

                int hr = dialog.Show(owner);
                if (hr == ERROR_CANCELLED_HRESULT) return new string[0];
                if (hr != 0) Marshal.ThrowExceptionForHR(hr);

                dialog.GetResults(out results);
                if (results == null) return new string[0];

                uint count;
                results.GetCount(out count);
                var folders = new List<string>((int)count);
                for (uint i = 0; i < count; i++)
                {
                    IShellItem item = null;
                    try
                    {
                        results.GetItemAt(i, out item);
                        string path = GetPath(item);
                        if (!String.IsNullOrWhiteSpace(path)) folders.Add(path);
                    }
                    finally
                    {
                        if (item != null && Marshal.IsComObject(item)) Marshal.ReleaseComObject(item);
                    }
                }
                return folders.ToArray();
            }
            finally
            {
                if (results != null && Marshal.IsComObject(results)) Marshal.ReleaseComObject(results);
                if (initialItem != null && Marshal.IsComObject(initialItem)) Marshal.ReleaseComObject(initialItem);
                if (dialog != null && Marshal.IsComObject(dialog)) Marshal.ReleaseComObject(dialog);
            }
        }
    }
}
'@

    try {
        Add-Type -TypeDefinition $source -Language CSharp -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

function Show-MultiFolderPicker {
    param([string]$Description, [string]$InitialPath)

    if (Initialize-NativeMultiFolderPicker) {
        try {
            $owner = [IntPtr]::Zero
            if ($form -and $form.IsHandleCreated) { $owner = $form.Handle }
            $picked = [GoogleTakeoutMetadataFixer.NativeMultiFolderPicker]::PickFolders($owner, $Description, $InitialPath)
            if ($null -eq $picked) { return [string[]]@() }
            return [string[]]$picked
        }
        catch {
            Show-WarningMessage -Message ('Windows multi-folder selection could not be opened. Falling back to the single-folder picker.' + [Environment]::NewLine + [Environment]::NewLine + $_.Exception.Message)
        }
    }

    $single = Show-FolderPicker -Description $Description -InitialPath $InitialPath -AllowNewFolder $false
    if ($single) { return [string[]]@($single) }
    return [string[]]@()
}

function Show-MultiSubfolderPicker {
    $parent = Show-FolderPicker -Description 'Select the parent folder containing the Takeout folders you want to add' -InitialPath '' -AllowNewFolder $false
    if (-not $parent) { return @() }
    $children = @(Get-ChildItem -LiteralPath $parent -Directory -Force -ErrorAction SilentlyContinue | Sort-Object Name)
    if ($children.Count -eq 0) {
        Show-WarningMessage -Message 'The selected parent folder has no immediate subfolders.'
        return @()
    }

    $picker = New-Object System.Windows.Forms.Form
    $picker.Text = 'Select multiple input folders'
    $picker.StartPosition = 'CenterParent'
    $picker.ClientSize = New-Object System.Drawing.Size(680, 520)
    $picker.MinimizeBox = $false
    $picker.MaximizeBox = $false
    $picker.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $picker.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = 'Parent: ' + $parent
    $label.Location = New-Object System.Drawing.Point(14, 12)
    $label.Size = New-Object System.Drawing.Size(650, 36)
    $picker.Controls.Add($label)

    $list = New-Object System.Windows.Forms.CheckedListBox
    $list.CheckOnClick = $true
    $list.Location = New-Object System.Drawing.Point(14, 52)
    $list.Size = New-Object System.Drawing.Size(650, 390)
    foreach ($child in $children) { [void]$list.Items.Add($child.Name, $false) }
    $picker.Controls.Add($list)

    $selectAll = New-Object System.Windows.Forms.Button
    $selectAll.Text = 'Select all'
    $selectAll.Location = New-Object System.Drawing.Point(14, 458)
    $selectAll.Size = New-Object System.Drawing.Size(95, 32)
    $picker.Controls.Add($selectAll)

    $clear = New-Object System.Windows.Forms.Button
    $clear.Text = 'Clear'
    $clear.Location = New-Object System.Drawing.Point(118, 458)
    $clear.Size = New-Object System.Drawing.Size(90, 32)
    $picker.Controls.Add($clear)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Add selected'
    $ok.Location = New-Object System.Drawing.Point(452, 458)
    $ok.Size = New-Object System.Drawing.Size(105, 32)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $picker.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancel'
    $cancel.Location = New-Object System.Drawing.Point(566, 458)
    $cancel.Size = New-Object System.Drawing.Size(98, 32)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $picker.Controls.Add($cancel)

    $selectAll.Add_Click({ for ($i=0; $i -lt $list.Items.Count; $i++) { $list.SetItemChecked($i, $true) } })
    $clear.Add_Click({ for ($i=0; $i -lt $list.Items.Count; $i++) { $list.SetItemChecked($i, $false) } })
    $picker.AcceptButton = $ok
    $picker.CancelButton = $cancel

    $selected = New-Object System.Collections.Generic.List[string]
    if ($picker.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($item in $list.CheckedItems) { $selected.Add((Join-Path $parent ([string]$item))) }
    }
    $picker.Dispose()
    return $selected.ToArray()
}

function Stop-ProcessTree {
    param([System.Diagnostics.Process]$Process)
    if ($null -eq $Process) { return }
    try { if ($Process.HasExited) { return } } catch { return }
    try {
        $taskkill = Join-Path $env:SystemRoot 'System32\taskkill.exe'
        if (Test-Path -LiteralPath $taskkill -PathType Leaf) { & $taskkill /PID $Process.Id /T /F 2>$null | Out-Null; return }
    }
    catch { }
    try { $Process.Kill() } catch { }
}

function Request-GracefulCancel {
    if ([string]::IsNullOrWhiteSpace($script:CancelFile)) { return $false }
    try { 'cancel' | Set-Content -LiteralPath $script:CancelFile -Encoding ASCII -Force; return $true } catch { return $false }
}

# ---------------- UI ----------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Google Takeout Metadata Fixer v6.2.3'
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.MaximizeBox = $true
$form.MinimizeBox = $true
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.AutoScroll = $true

# Keep comfortable visible space above and below the window on normal displays.
# On smaller/high-DPI displays the form becomes shorter and AutoScroll provides access
# to the lower controls instead of forcing the window beyond the usable desktop.
$workingArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$targetClientWidth = [Math]::Min(980, [Math]::Max(860, $workingArea.Width - 120))
$targetClientHeight = [Math]::Min(850, [Math]::Max(620, $workingArea.Height - 140))
$form.ClientSize = New-Object System.Drawing.Size($targetClientWidth, $targetClientHeight)
$form.MinimumSize = New-Object System.Drawing.Size(900, 650)

$title = New-Object System.Windows.Forms.Label
$title.Text = 'Google Takeout Metadata Fixer'
$title.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 16)
$title.AutoSize = $true
$title.Location = New-Object System.Drawing.Point(20, 14)
$form.Controls.Add($title)

$safety = New-Object System.Windows.Forms.Label
$safety.Text = 'Multi-folder mode: every input gets its own fixed-copy output. Folders run sequentially; files inside each folder use the selected worker pool.'
$safety.AutoSize = $true
$safety.Location = New-Object System.Drawing.Point(22, 50)
$form.Controls.Add($safety)

$pairsLabel = New-Object System.Windows.Forms.Label
$pairsLabel.Text = 'Input / output folder pairs'
$pairsLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
$pairsLabel.AutoSize = $true
$pairsLabel.Location = New-Object System.Drawing.Point(22, 82)
$form.Controls.Add($pairsLabel)

$pairGrid = New-Object System.Windows.Forms.DataGridView
$pairGrid.Location = New-Object System.Drawing.Point(22, 106)
$pairGrid.Size = New-Object System.Drawing.Size(936, 165)
$pairGrid.AllowUserToAddRows = $false
$pairGrid.AllowUserToDeleteRows = $false
$pairGrid.AllowUserToResizeRows = $false
$pairGrid.RowHeadersVisible = $false
$pairGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$pairGrid.MultiSelect = $true
$pairGrid.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::None
$pairGrid.BackgroundColor = [System.Drawing.SystemColors]::Window
$pairGrid.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D

$inputCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$inputCol.Name = 'InputFolder'
$inputCol.HeaderText = 'Input folder'
$inputCol.Width = 420
$inputCol.ReadOnly = $true
[void]$pairGrid.Columns.Add($inputCol)

$outputCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$outputCol.Name = 'OutputFolder'
$outputCol.HeaderText = 'Output folder (editable)'
$outputCol.Width = 390
[void]$pairGrid.Columns.Add($outputCol)

$statusCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$statusCol.Name = 'Status'
$statusCol.HeaderText = 'Status'
$statusCol.Width = 120
$statusCol.ReadOnly = $true
[void]$pairGrid.Columns.Add($statusCol)
$pairGrid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($pairGrid)

$addOneButton = New-Object System.Windows.Forms.Button
$addOneButton.Text = 'Add folder(s)...'
$addOneButton.Location = New-Object System.Drawing.Point(22, 280)
$addOneButton.Size = New-Object System.Drawing.Size(128, 31)
$form.Controls.Add($addOneButton)

$addManyButton = New-Object System.Windows.Forms.Button
$addManyButton.Text = 'Add child folders...'
$addManyButton.Location = New-Object System.Drawing.Point(158, 280)
$addManyButton.Size = New-Object System.Drawing.Size(140, 31)
$form.Controls.Add($addManyButton)

$removeButton = New-Object System.Windows.Forms.Button
$removeButton.Text = 'Remove selected'
$removeButton.Location = New-Object System.Drawing.Point(306, 280)
$removeButton.Size = New-Object System.Drawing.Size(130, 31)
$form.Controls.Add($removeButton)

$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = 'Clear list'
$clearButton.Location = New-Object System.Drawing.Point(444, 280)
$clearButton.Size = New-Object System.Drawing.Size(100, 31)
$form.Controls.Add($clearButton)

$browseOutputButton = New-Object System.Windows.Forms.Button
$browseOutputButton.Text = 'Browse selected output...'
$browseOutputButton.Location = New-Object System.Drawing.Point(754, 280)
$browseOutputButton.Size = New-Object System.Drawing.Size(204, 31)
$browseOutputButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($browseOutputButton)

$pairHint = New-Object System.Windows.Forms.Label
$pairHint.Text = 'Tip: Add folder(s)... supports Ctrl-click / Shift-click multi-selection. Add child folders... is also available for checkbox selection. Each input gets its own sibling "<input> - Fixed" output.'
$pairHint.Location = New-Object System.Drawing.Point(22, 317)
$pairHint.Size = New-Object System.Drawing.Size(936, 34)
$pairHint.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($pairHint)

$exifLabel = New-Object System.Windows.Forms.Label
$exifLabel.Text = 'ExifTool executable'
$exifLabel.AutoSize = $true
$exifLabel.Location = New-Object System.Drawing.Point(22, 354)
$form.Controls.Add($exifLabel)

$exifText = New-Object System.Windows.Forms.TextBox
$exifText.Location = New-Object System.Drawing.Point(22, 375)
$exifText.Size = New-Object System.Drawing.Size(814, 24)
$exifText.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($exifText)

$exifBrowse = New-Object System.Windows.Forms.Button
$exifBrowse.Text = 'Browse...'
$exifBrowse.Location = New-Object System.Drawing.Point(848, 373)
$exifBrowse.Size = New-Object System.Drawing.Size(110, 28)
$exifBrowse.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($exifBrowse)

$modeGroup = New-Object System.Windows.Forms.GroupBox
$modeGroup.Text = 'Processing mode'
$modeGroup.Location = New-Object System.Drawing.Point(22, 412)
$modeGroup.Size = New-Object System.Drawing.Size(936, 124)
$modeGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($modeGroup)

$dryRadio = New-Object System.Windows.Forms.RadioButton
$dryRadio.Text = 'Dry run - scan each input and create audit reports only'
$dryRadio.AutoSize = $true
$dryRadio.Location = New-Object System.Drawing.Point(18, 26)
$dryRadio.Checked = $true
$modeGroup.Controls.Add($dryRadio)

$safeRadio = New-Object System.Windows.Forms.RadioButton
$safeRadio.Text = 'Create fixed copies - preserve plausible existing local EXIF times'
$safeRadio.AutoSize = $true
$safeRadio.Location = New-Object System.Drawing.Point(18, 56)
$modeGroup.Controls.Add($safeRadio)

$forceRadio = New-Object System.Windows.Forms.RadioButton
$forceRadio.Text = 'Create fixed copies - force Google JSON photoTakenTime to win'
$forceRadio.AutoSize = $true
$forceRadio.Location = New-Object System.Drawing.Point(18, 86)
$modeGroup.Controls.Add($forceRadio)

$forceHint = New-Object System.Windows.Forms.Label
$forceHint.Text = 'Force mode is mainly for dates you intentionally corrected inside Google Photos.'
$forceHint.AutoSize = $true
$forceHint.Location = New-Object System.Drawing.Point(37, 108)
$modeGroup.Controls.Add($forceHint)

$optionsGroup = New-Object System.Windows.Forms.GroupBox
$optionsGroup.Text = 'Options'
$optionsGroup.Location = New-Object System.Drawing.Point(22, 546)
$optionsGroup.Size = New-Object System.Drawing.Size(936, 90)
$optionsGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($optionsGroup)

$recurseCheck = New-Object System.Windows.Forms.CheckBox
$recurseCheck.Text = 'Include all subfolders (recommended)'
$recurseCheck.Checked = $true
$recurseCheck.AutoSize = $true
$recurseCheck.Location = New-Object System.Drawing.Point(18, 25)
$optionsGroup.Controls.Add($recurseCheck)

$overwriteCheck = New-Object System.Windows.Forms.CheckBox
$overwriteCheck.Text = 'Allow re-run/overwrite when an output folder is not empty'
$overwriteCheck.Checked = $false
$overwriteCheck.AutoSize = $true
$overwriteCheck.Location = New-Object System.Drawing.Point(18, 53)
$optionsGroup.Controls.Add($overwriteCheck)

$verifyCheck = New-Object System.Windows.Forms.CheckBox
$verifyCheck.Text = 'Verify written timestamps (recommended)'
$verifyCheck.Checked = $true
$verifyCheck.AutoSize = $true
$verifyCheck.Location = New-Object System.Drawing.Point(520, 25)
$optionsGroup.Controls.Add($verifyCheck)

$keepAwakeCheck = New-Object System.Windows.Forms.CheckBox
$keepAwakeCheck.Text = 'Keep computer awake during long runs'
$keepAwakeCheck.Checked = $true
$keepAwakeCheck.AutoSize = $true
$keepAwakeCheck.Location = New-Object System.Drawing.Point(520, 53)
$optionsGroup.Controls.Add($keepAwakeCheck)

$performanceGroup = New-Object System.Windows.Forms.GroupBox
$performanceGroup.Text = 'Performance / worker pool (per folder)'
$performanceGroup.Location = New-Object System.Drawing.Point(22, 646)
$performanceGroup.Size = New-Object System.Drawing.Size(936, 76)
$performanceGroup.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($performanceGroup)

$workerModeLabel = New-Object System.Windows.Forms.Label
$workerModeLabel.Text = 'Worker mode'
$workerModeLabel.AutoSize = $true
$workerModeLabel.Location = New-Object System.Drawing.Point(18, 25)
$performanceGroup.Controls.Add($workerModeLabel)

$workerModeCombo = New-Object System.Windows.Forms.ComboBox
$workerModeCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$workerModeCombo.Location = New-Object System.Drawing.Point(105, 21)
$workerModeCombo.Size = New-Object System.Drawing.Size(480, 25)
[void]$workerModeCombo.Items.Add('Conservative - 1 worker (HDD / compatibility)')
[void]$workerModeCombo.Items.Add('Balanced - 4 workers (recommended for SSD/NVMe)')
[void]$workerModeCombo.Items.Add('Custom')
$workerModeCombo.SelectedIndex = 1
$performanceGroup.Controls.Add($workerModeCombo)

$workerCountLabel = New-Object System.Windows.Forms.Label
$workerCountLabel.Text = 'Workers'
$workerCountLabel.AutoSize = $true
$workerCountLabel.Location = New-Object System.Drawing.Point(610, 25)
$performanceGroup.Controls.Add($workerCountLabel)

$workerCountNumeric = New-Object System.Windows.Forms.NumericUpDown
$workerCountNumeric.Minimum = 1
$workerCountNumeric.Maximum = 8
$workerCountNumeric.Value = 4
$workerCountNumeric.Location = New-Object System.Drawing.Point(672, 21)
$workerCountNumeric.Size = New-Object System.Drawing.Size(60, 24)
$workerCountNumeric.Enabled = $false
$performanceGroup.Controls.Add($workerCountNumeric)

$workerHint = New-Object System.Windows.Forms.Label
$workerHint.Text = 'Folder pairs run one after another; each folder can use several persistent ExifTool workers in parallel.'
$workerHint.AutoSize = $true
$workerHint.Location = New-Object System.Drawing.Point(18, 52)
$performanceGroup.Controls.Add($workerHint)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(22, 733)
$progressBar.Size = New-Object System.Drawing.Size(936, 20)
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($progressBar)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = 'Ready. Add one or more input folders. Each input has its own output folder.'
$statusLabel.Location = New-Object System.Drawing.Point(22, 760)
$statusLabel.Size = New-Object System.Drawing.Size(936, 34)
$statusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($statusLabel)

$startButton = New-Object System.Windows.Forms.Button
$startButton.Text = 'Run dry check'
$startButton.Location = New-Object System.Drawing.Point(22, 803)
$startButton.Size = New-Object System.Drawing.Size(130, 34)
$form.Controls.Add($startButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = 'Cancel'
$cancelButton.Location = New-Object System.Drawing.Point(160, 803)
$cancelButton.Size = New-Object System.Drawing.Size(105, 34)
$cancelButton.Enabled = $false
$form.Controls.Add($cancelButton)

$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = 'Close'
$closeButton.Location = New-Object System.Drawing.Point(498, 803)
$closeButton.Size = New-Object System.Drawing.Size(100, 34)
$closeButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($closeButton)

$openLogButton = New-Object System.Windows.Forms.Button
$openLogButton.Text = 'Open selected CSV'
$openLogButton.Location = New-Object System.Drawing.Point(607, 803)
$openLogButton.Size = New-Object System.Drawing.Size(160, 34)
$openLogButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($openLogButton)

$openFolderButton = New-Object System.Windows.Forms.Button
$openFolderButton.Text = 'Open selected output'
$openFolderButton.Location = New-Object System.Drawing.Point(776, 803)
$openFolderButton.Size = New-Object System.Drawing.Size(182, 34)
$openFolderButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($openFolderButton)

$form.AcceptButton = $startButton

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($addOneButton, 'Explorer-style folder picker: hold Ctrl to select individual folders or Shift to select a range, then click Select Folder.')
$toolTip.SetToolTip($addManyButton, 'Alternative batch picker: choose a parent folder, then check multiple immediate child folders in one dialog. Useful for Photos from 2021, Photos from 2022, and similar Takeout folders.')
$toolTip.SetToolTip($browseOutputButton, 'Change the output folder for the currently selected input. The default is a sibling named "<input> - Fixed".')
$toolTip.SetToolTip($workerModeCombo, 'The worker count applies inside each folder. Multiple folder pairs are processed sequentially to avoid multiplying disk contention.')

$script:Process = $null
$script:StdOutTask = $null
$script:StdErrTask = $null
$script:ProgressFile = ''
$script:CancelFile = ''
$script:ManifestFile = ''
$script:CloseAfterCancel = $false
$script:WasCancelled = $false
$script:ResultByInput = @{}
$script:LastBatchSummaryPath = ''

function Add-FolderPair {
    param([string]$InputPath)
    if ([string]::IsNullOrWhiteSpace($InputPath) -or -not (Test-Path -LiteralPath $InputPath -PathType Container)) { return }
    $normalized = Get-NormalizedDirectoryPath $InputPath
    foreach ($row in $pairGrid.Rows) {
        if ([string]$row.Cells['InputFolder'].Value -and (Get-NormalizedDirectoryPath ([string]$row.Cells['InputFolder'].Value)).Equals($normalized, [StringComparison]::OrdinalIgnoreCase)) { return }
    }
    $output = Get-DefaultOutputFolder $normalized
    $index = $pairGrid.Rows.Add($normalized, $output, 'Ready')
    $pairGrid.Rows[$index].Selected = $true
    $pairGrid.CurrentCell = $pairGrid.Rows[$index].Cells['InputFolder']
}

function Get-SelectedPairRow {
    if ($pairGrid.CurrentRow -and -not $pairGrid.CurrentRow.IsNewRow) { return $pairGrid.CurrentRow }
    if ($pairGrid.SelectedRows.Count -gt 0) { return $pairGrid.SelectedRows[0] }
    return $null
}

function Update-ModeControls {
    $copyMode = -not $dryRadio.Checked
    $pairGrid.Columns['OutputFolder'].ReadOnly = -not $copyMode
    $browseOutputButton.Enabled = $copyMode -and $pairGrid.Rows.Count -gt 0
    $overwriteCheck.Enabled = $copyMode
    $verifyCheck.Enabled = $copyMode
    if ($dryRadio.Checked) { $startButton.Text = 'Run dry check' }
    elseif ($verifyCheck.Checked) { $startButton.Text = 'Create & verify copies' }
    else { $startButton.Text = 'Create fixed copies' }
}

function Set-UiBusy {
    param([bool]$Busy)
    $pairGrid.Enabled = -not $Busy
    $addOneButton.Enabled = -not $Busy
    $addManyButton.Enabled = -not $Busy
    $removeButton.Enabled = -not $Busy
    $clearButton.Enabled = -not $Busy
    $exifText.Enabled = -not $Busy
    $exifBrowse.Enabled = -not $Busy
    $dryRadio.Enabled = -not $Busy
    $safeRadio.Enabled = -not $Busy
    $forceRadio.Enabled = -not $Busy
    $recurseCheck.Enabled = -not $Busy
    $keepAwakeCheck.Enabled = -not $Busy
    $workerModeCombo.Enabled = -not $Busy
    $workerCountNumeric.Enabled = (-not $Busy -and $workerModeCombo.SelectedIndex -eq 2)
    $startButton.Enabled = -not $Busy
    $cancelButton.Enabled = $Busy
    $openLogButton.Enabled = -not $Busy
    $openFolderButton.Enabled = -not $Busy
    if ($Busy) {
        $browseOutputButton.Enabled = $false
        $overwriteCheck.Enabled = $false
        $verifyCheck.Enabled = $false
    }
    else { Update-ModeControls }
}

function Read-ProgressState {
    if ([string]::IsNullOrWhiteSpace($script:ProgressFile) -or -not (Test-Path -LiteralPath $script:ProgressFile -PathType Leaf)) { return $null }
    try {
        $raw = Get-Content -LiteralPath $script:ProgressFile -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
        return ($raw | ConvertFrom-Json)
    }
    catch { return $null }
}

function Get-ValidatedPairs {
    param([bool]$DryRun)
    if ($pairGrid.Rows.Count -eq 0) { throw 'Add at least one Google Takeout input folder.' }
    $pairs = New-Object System.Collections.Generic.List[object]
    $inputs = New-Object System.Collections.Generic.List[string]
    $outputs = New-Object System.Collections.Generic.List[string]

    for ($i=0; $i -lt $pairGrid.Rows.Count; $i++) {
        $row = $pairGrid.Rows[$i]
        $input = ([string]$row.Cells['InputFolder'].Value).Trim()
        $output = ([string]$row.Cells['OutputFolder'].Value).Trim()
        if ([string]::IsNullOrWhiteSpace($input) -or -not (Test-Path -LiteralPath $input -PathType Container)) { throw ('Row ' + ($i+1) + ': input folder does not exist.') }
        $input = Get-NormalizedDirectoryPath $input
        if ([string]::IsNullOrWhiteSpace($output)) { $output = Get-DefaultOutputFolder $input; $row.Cells['OutputFolder'].Value = $output }
        $output = Get-NormalizedDirectoryPath $output
        $inputs.Add($input)
        $outputs.Add($output)
        $pairs.Add([PSCustomObject]@{ InputFolder=$input; OutputFolder=$output })
    }

    for ($i=0; $i -lt $pairs.Count; $i++) {
        for ($j=$i+1; $j -lt $pairs.Count; $j++) {
            if ($inputs[$i].Equals($inputs[$j], [StringComparison]::OrdinalIgnoreCase)) { throw ('Duplicate input folder: ' + $inputs[$i]) }
            if ($recurseCheck.Checked -and ((Test-PathContainsPath $inputs[$i] $inputs[$j]) -or (Test-PathContainsPath $inputs[$j] $inputs[$i]))) {
                throw ('Input folders must not contain one another when subfolder processing is enabled:' + [Environment]::NewLine + $inputs[$i] + [Environment]::NewLine + $inputs[$j])
            }
            if ($outputs[$i].Equals($outputs[$j], [StringComparison]::OrdinalIgnoreCase) -or (Test-PathContainsPath $outputs[$i] $outputs[$j]) -or (Test-PathContainsPath $outputs[$j] $outputs[$i])) {
                throw ('Each input needs a separate, non-overlapping output folder:' + [Environment]::NewLine + $outputs[$i] + [Environment]::NewLine + $outputs[$j])
            }
        }
    }

    for ($i=0; $i -lt $pairs.Count; $i++) {
        if ((Test-PathContainsPath $inputs[$i] $outputs[$i]) -or (Test-PathContainsPath $outputs[$i] $inputs[$i])) {
            throw ('Input and output must be separate and must not contain one another:' + [Environment]::NewLine + $inputs[$i] + [Environment]::NewLine + $outputs[$i])
        }
        for ($j=0; $j -lt $pairs.Count; $j++) {
            if ((Test-PathContainsPath $inputs[$j] $outputs[$i]) -or (Test-PathContainsPath $outputs[$i] $inputs[$j])) {
                throw ('An output folder overlaps an input folder from the batch:' + [Environment]::NewLine + 'Output: ' + $outputs[$i] + [Environment]::NewLine + 'Input: ' + $inputs[$j])
            }
        }
        if (-not $DryRun -and (Test-Path -LiteralPath $outputs[$i])) {
            if (-not (Test-Path -LiteralPath $outputs[$i] -PathType Container)) { throw ('Output path exists but is not a folder: ' + $outputs[$i]) }
            $existing = Get-ChildItem -LiteralPath $outputs[$i] -Force -ErrorAction Stop | Select-Object -First 1
            if ($existing -and -not $overwriteCheck.Checked) { throw ('Output folder is not empty:' + [Environment]::NewLine + $outputs[$i] + [Environment]::NewLine + [Environment]::NewLine + 'Choose an empty folder or enable re-run/overwrite.') }
        }
    }
    # Windows PowerShell can throw "Argument types do not match" for @($list)
    # when the list is a Generic.List[object] created by New-Object. Materialize
    # explicitly to object[] instead.
    return $pairs.ToArray()
}

function Update-ResultRows {
    param($FolderResults)
    $script:ResultByInput = @{}
    foreach ($result in @($FolderResults)) {
        if ($null -eq $result) { continue }
        $input = [string]$result.InputFolder
        $script:ResultByInput[$input] = $result
        foreach ($row in $pairGrid.Rows) {
            $rowInput = [string]$row.Cells['InputFolder'].Value
            if ($rowInput.Equals($input, [StringComparison]::OrdinalIgnoreCase)) {
                $row.Cells['Status'].Value = [string]$result.Status
                break
            }
        }
    }
}

$addOneButton.Add_Click({
    $initial = ''
    $row = Get-SelectedPairRow
    if ($row) {
        $existingInput = [string]$row.Cells['InputFolder'].Value
        if (-not [string]::IsNullOrWhiteSpace($existingInput) -and (Test-Path -LiteralPath $existingInput -PathType Container)) {
            $initial = Split-Path -Parent (Get-NormalizedDirectoryPath $existingInput)
        }
    }
    foreach ($selected in [string[]](Show-MultiFolderPicker -Description 'Select one or more Google Photos Takeout folders (Ctrl-click or Shift-click for multiple folders)' -InitialPath $initial)) {
        Add-FolderPair $selected
    }
    Update-ModeControls
})

$addManyButton.Add_Click({
    foreach ($selected in @(Show-MultiSubfolderPicker)) { Add-FolderPair $selected }
    Update-ModeControls
})

$removeButton.Add_Click({
    $rows = @($pairGrid.SelectedRows | Sort-Object Index -Descending)
    foreach ($row in $rows) { if (-not $row.IsNewRow) { $pairGrid.Rows.RemoveAt($row.Index) } }
    Update-ModeControls
})

$clearButton.Add_Click({ $pairGrid.Rows.Clear(); $script:ResultByInput=@{}; Update-ModeControls })

$browseOutputButton.Add_Click({
    $row = Get-SelectedPairRow
    if (-not $row) { Show-WarningMessage -Message 'Select a folder row first.'; return }
    $current = [string]$row.Cells['OutputFolder'].Value
    $selected = Show-FolderPicker -Description 'Select or create the output folder for the selected input' -InitialPath $current -AllowNewFolder $true
    if ($selected) { $row.Cells['OutputFolder'].Value = $selected }
})

$exifBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = 'Select ExifTool executable'
    $dialog.Filter = 'ExifTool executable (exiftool*.exe)|exiftool*.exe|Executable files (*.exe)|*.exe|All files (*.*)|*.*'
    $dialog.CheckFileExists = $true
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $exifText.Text = $dialog.FileName }
})

$workerModeCombo.Add_SelectedIndexChanged({
    if ($workerModeCombo.SelectedIndex -eq 0) { $workerCountNumeric.Value=1; $workerCountNumeric.Enabled=$false }
    elseif ($workerModeCombo.SelectedIndex -eq 1) { $workerCountNumeric.Value=4; $workerCountNumeric.Enabled=$false }
    else { $workerCountNumeric.Enabled=$true }
})

$dryRadio.Add_CheckedChanged({ Update-ModeControls })
$safeRadio.Add_CheckedChanged({ Update-ModeControls })
$forceRadio.Add_CheckedChanged({ Update-ModeControls })
$verifyCheck.Add_CheckedChanged({ Update-ModeControls })
$pairGrid.Add_SelectionChanged({ Update-ModeControls })
$closeButton.Add_Click({ $form.Close() })

$openFolderButton.Add_Click({
    $row = Get-SelectedPairRow
    if (-not $row) { Show-WarningMessage -Message 'Select a folder row first.'; return }
    $path = [string]$row.Cells['OutputFolder'].Value
    if (Test-Path -LiteralPath $path -PathType Container) { Start-Process -FilePath 'explorer.exe' -ArgumentList ('"' + $path + '"') }
    else { Show-WarningMessage -Message 'The selected output folder does not exist yet.' }
})

$openLogButton.Add_Click({
    $row = Get-SelectedPairRow
    if (-not $row) { Show-WarningMessage -Message 'Select a folder row first.'; return }
    $input = [string]$row.Cells['InputFolder'].Value
    $logPath = ''
    if ($script:ResultByInput.ContainsKey($input)) { $logPath = [string]$script:ResultByInput[$input].LogPath }
    if ([string]::IsNullOrWhiteSpace($logPath) -or -not (Test-Path -LiteralPath $logPath -PathType Leaf)) {
        $searchDir = if ($dryRadio.Checked) { Split-Path -Parent $input } else { [string]$row.Cells['OutputFolder'].Value }
        if (Test-Path -LiteralPath $searchDir -PathType Container) {
            $pattern = if ($dryRadio.Checked) { 'metadata-fix-dry-run*.csv' } else { 'metadata-fix-log*.csv' }
            $latest = Get-ChildItem -LiteralPath $searchDir -File -Filter $pattern -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($latest) { $logPath = $latest.FullName }
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($logPath) -and (Test-Path -LiteralPath $logPath -PathType Leaf)) { Start-Process -FilePath $logPath }
    else { Show-WarningMessage -Message 'No CSV log was found for the selected folder.' }
})

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 250
$timer.Add_Tick({
    $state = Read-ProgressState
    if ($state) {
        if ($state.PSObject.Properties['Percent']) {
            $pct = [Math]::Max(0, [Math]::Min(100, [int]$state.Percent))
            $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
            $progressBar.Value = $pct
        }
        if ($state.PSObject.Properties['Message'] -and -not [string]::IsNullOrWhiteSpace([string]$state.Message)) { $statusLabel.Text = [string]$state.Message }
        if ($state.PSObject.Properties['FolderIndex'] -and [int]$state.FolderIndex -gt 0) {
            $idx = [int]$state.FolderIndex - 1
            if ($idx -ge 0 -and $idx -lt $pairGrid.Rows.Count) {
                if ([string]$pairGrid.Rows[$idx].Cells['Status'].Value -eq 'Ready') { $pairGrid.Rows[$idx].Cells['Status'].Value = 'Running' }
            }
        }
    }

    if ($script:Process -and $script:Process.HasExited) {
        $timer.Stop()
        $exitCode = $script:Process.ExitCode
        $stdout=''; $stderr=''
        try { if ($script:StdOutTask) { $stdout=[string]$script:StdOutTask.Result } } catch { }
        try { if ($script:StdErrTask) { $stderr=[string]$script:StdErrTask.Result } } catch { }
        $finalState = Read-ProgressState
        if ($finalState -and $finalState.PSObject.Properties['FolderResults']) { Update-ResultRows $finalState.FolderResults }
        if ($finalState -and $finalState.PSObject.Properties['BatchSummaryPath']) { $script:LastBatchSummaryPath = [string]$finalState.BatchSummaryPath }

        $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
        if ($script:WasCancelled) {
            $statusLabel.Text = 'Cancelled safely. Completed folder outputs remain intact; source Takeout folders were not changed.'
        }
        elseif ($exitCode -eq 0) {
            $progressBar.Value = 100
            $folderFailures = 0
            if ($finalState -and $finalState.PSObject.Properties['FolderFailures']) { $folderFailures = [int]$finalState.FolderFailures }
            if ($folderFailures -gt 0) {
                $statusLabel.Text = ('Finished with {0} folder failure(s). Review the Status column and per-folder logs.' -f $folderFailures)
                Show-WarningMessage -Message ('The batch completed, but {0} folder(s) failed. Successful folder outputs are intact. Review the per-folder CSV/summary files.' -f $folderFailures)
            }
            elseif ($finalState -and $finalState.PSObject.Properties['Verified']) {
                $statusLabel.Text = ('Finished {0} folder pair(s). Processed: {1}; verified: {2}; no sidecar: {3}; failures: {4}.' -f $pairGrid.Rows.Count, $finalState.Processed, $finalState.Verified, $finalState.Unmatched, $finalState.Failed)
                if (($finalState.PSObject.Properties['VerificationFailed'] -and [int]$finalState.VerificationFailed -gt 0) -or ($finalState.PSObject.Properties['Failed'] -and [int]$finalState.Failed -gt 0) -or ($finalState.PSObject.Properties['ScanErrors'] -and [int]$finalState.ScanErrors -gt 0)) {
                    Show-WarningMessage -Message 'The batch finished, but one or more files failed verification, warned/failed, or were skipped. Review the per-folder logs before uploading.'
                }
            }
            else { $statusLabel.Text = 'Finished. Review each output folder and its CSV log before uploading.' }
        }
        else {
            $progressBar.Value = 0
            $message = ''
            if ($finalState -and $finalState.PSObject.Properties['Message']) { $message=[string]$finalState.Message }
            if ([string]::IsNullOrWhiteSpace($message)) { $message=$stderr.Trim() }
            if ([string]::IsNullOrWhiteSpace($message)) { $message=$stdout.Trim() }
            if ([string]::IsNullOrWhiteSpace($message)) { $message='The batch coordinator stopped with an unexpected error.' }
            $statusLabel.Text = 'Stopped because of an error. Source media was not modified.'
            Show-ErrorMessage -Message $message
        }

        Set-UiBusy $false
        foreach ($temp in @($script:ProgressFile,$script:CancelFile,$script:ManifestFile)) { if (-not [string]::IsNullOrWhiteSpace($temp)) { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue } }
        $script:Process.Dispose(); $script:Process=$null; $script:StdOutTask=$null; $script:StdErrTask=$null
        $shouldClose=$script:CloseAfterCancel; $script:WasCancelled=$false; $script:CloseAfterCancel=$false
        if ($shouldClose) { $form.Close() }
    }
})

$cancelButton.Add_Click({
    if (-not $script:Process -or $script:Process.HasExited) { return }
    $answer = [System.Windows.Forms.MessageBox]::Show('Stop the batch safely? Active files in the current folder will finish first. Completed folder outputs remain intact.', 'Cancel processing', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
        $script:WasCancelled=$true
        if (Request-GracefulCancel) { $cancelButton.Enabled=$false; $statusLabel.Text='Cancellation requested. Finishing active files safely...' }
        else { Stop-ProcessTree $script:Process; $statusLabel.Text='Safe cancellation could not be signaled; stopping the process tree...' }
    }
})

$startButton.Add_Click({
    try {
        $syntaxProblem = Test-CoreScriptSyntax
        if (-not [string]::IsNullOrWhiteSpace($syntaxProblem)) { Show-ErrorMessage -Message ('PowerShell syntax self-check failed:' + [Environment]::NewLine + [Environment]::NewLine + $syntaxProblem); return }

        $isDryRun = $dryRadio.Checked
        $pairs = @(Get-ValidatedPairs -DryRun $isDryRun)

        $exifTool = $exifText.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($exifTool)) { $exifTool=Find-ExifToolLocal; if ($exifTool) { $exifText.Text=$exifTool } }
        if ([string]::IsNullOrWhiteSpace($exifTool) -or -not (Test-Path -LiteralPath $exifTool -PathType Leaf)) { Show-ErrorMessage -Message 'ExifTool was not found. Extract the official ExifTool package next to this fixer, or browse to exiftool.exe / exiftool(-k).exe.'; return }

        foreach ($row in $pairGrid.Rows) { $row.Cells['Status'].Value='Queued' }
        $script:ResultByInput=@{}
        $progressBar.Value=0; $progressBar.Style=[System.Windows.Forms.ProgressBarStyle]::Marquee
        $statusLabel.Text=('Starting batch with {0} folder pair(s)...' -f $pairs.Count)
        $script:WasCancelled=$false; $script:CloseAfterCancel=$false; $script:LastBatchSummaryPath=''

        $jobId=[Guid]::NewGuid().ToString('N')
        $tempDir=[IO.Path]::GetTempPath()
        $script:ProgressFile=Join-Path $tempDir ('GoogleTakeoutMetadataFixer-v6.2.3-' + $jobId + '.json')
        $script:CancelFile=Join-Path $tempDir ('GoogleTakeoutMetadataFixer-v6.2.3-' + $jobId + '.cancel')
        $script:ManifestFile=Join-Path $tempDir ('GoogleTakeoutMetadataFixer-v6.2.3-' + $jobId + '-folders.json')
        Remove-Item -LiteralPath $script:CancelFile -Force -ErrorAction SilentlyContinue
        ($pairs | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $script:ManifestFile -Encoding UTF8

        $parts=New-Object System.Collections.Generic.List[string]
        $parts.Add('& ' + (ConvertTo-PowerShellLiteral $batchScript))
        $parts.Add('-ManifestPath ' + (ConvertTo-PowerShellLiteral $script:ManifestFile))
        $parts.Add('-ExifToolPath ' + (ConvertTo-PowerShellLiteral $exifTool))
        $parts.Add('-ProgressFile ' + (ConvertTo-PowerShellLiteral $script:ProgressFile))
        $parts.Add('-CancelFile ' + (ConvertTo-PowerShellLiteral $script:CancelFile))
        $parts.Add('-WorkerCount ' + [int]$workerCountNumeric.Value)
        if ($isDryRun) { $parts.Add('-DryRun') }
        if ($forceRadio.Checked) { $parts.Add('-ForceJsonTime') }
        if (-not $verifyCheck.Checked -and -not $isDryRun) { $parts.Add('-SkipVerification') }
        if (-not $recurseCheck.Checked) { $parts.Add('-TopLevelOnly') }
        if ($overwriteCheck.Checked -and -not $isDryRun) { $parts.Add('-OverwriteOutput') }
        if (-not $keepAwakeCheck.Checked) { $parts.Add('-AllowSystemSleep') }

        $commandText=$parts -join ' '
        $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($commandText))
        $powershellExe=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        if (-not (Test-Path -LiteralPath $powershellExe -PathType Leaf)) { $powershellExe='powershell.exe' }
        $startInfo=New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName=$powershellExe
        $startInfo.Arguments='-NoProfile -ExecutionPolicy Bypass -EncodedCommand ' + $encoded
        $startInfo.UseShellExecute=$false; $startInfo.CreateNoWindow=$true; $startInfo.RedirectStandardOutput=$true; $startInfo.RedirectStandardError=$true
        $process=New-Object System.Diagnostics.Process; $process.StartInfo=$startInfo
        if (-not $process.Start()) { Show-ErrorMessage -Message 'Windows could not start the batch metadata process.'; return }
        $script:Process=$process; $script:StdOutTask=$process.StandardOutput.ReadToEndAsync(); $script:StdErrTask=$process.StandardError.ReadToEndAsync()
        Set-UiBusy $true; $timer.Start()
    }
    catch { Set-UiBusy $false; Show-ErrorMessage -Message $_.Exception.Message }
})

$form.Add_FormClosing({
    param($sender,$eventArgs)
    if ($script:Process -and -not $script:Process.HasExited) {
        $answer=[System.Windows.Forms.MessageBox]::Show('A batch is still running. Stop it safely and close the application?', 'Close application', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { $eventArgs.Cancel=$true; return }
        $script:WasCancelled=$true; $script:CloseAfterCancel=$true; $eventArgs.Cancel=$true
        if (Request-GracefulCancel) { $cancelButton.Enabled=$false; $statusLabel.Text='Cancellation requested. The window will close after active files finish safely...' }
        else { Stop-ProcessTree $script:Process }
    }
})

$detectedExifTool=Find-ExifToolLocal
if ($detectedExifTool) { $exifText.Text=$detectedExifTool }
Update-ModeControls
[void]$form.ShowDialog()
