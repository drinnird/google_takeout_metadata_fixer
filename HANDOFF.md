# Google Takeout Metadata Fixer - Maintainer Handoff

## Current state

**Current known-good release: v6.2.3**

The project is a Windows PowerShell / WinForms application that reconstructs Google Photos Takeout sidecar metadata into **copies** of exported media. ExifTool performs metadata reads/writes. The current release supports multiple input/output folder pairs, parallel processing inside each active folder, persistent ExifTool sessions, read-back verification, signature-based file-type detection, GPS precision preservation, and hard completeness checks.

The most important development principle is:

> Never sacrifice source preservation, timestamp correctness, or auditability for speed.

The original Takeout must remain untouched.

## Architecture

```text
Run-GoogleTakeoutMetadataFixer-v6.2.3.bat
        |
        v
GoogleTakeoutMetadataFixer-GUI.ps1
        |
        | creates/validates folder-pair manifest
        v
Fix-GoogleTakeoutMetadata-v6.2.3-Batch.ps1
        |
        | folder pairs are processed sequentially
        v
Fix-GoogleTakeoutMetadata-v6.2.3.ps1
        |
        | scans one input folder, matches sidecars,
        | distributes media by estimated byte load
        |
        +------> Worker 1 ----> persistent ExifTool
        +------> Worker 2 ----> persistent ExifTool
        +------> Worker 3 ----> persistent ExifTool
        +------> Worker 4 ----> persistent ExifTool
                      |
                      v
             copy -> write -> verify
        |
        v
merged CSV + per-folder summary
        |
        v
batch summary
```

### Files

#### `GoogleTakeoutMetadataFixer-GUI.ps1`

Windows Forms front end.

Responsibilities:

- ExifTool discovery/browse.
- Native multi-folder picker using `IFileOpenDialog` with folder + multi-select flags.
- Fallback `FolderBrowserDialog`.
- Ctrl/Shift multi-folder selection.
- Input/output folder-pair table.
- Batch validation.
- Processing mode and worker-mode controls.
- AST syntax check of core/batch/worker scripts before processing.
- Progress/status display.
- Cancellation signal.
- Keep-awake option exposure.
- Open selected CSV/output helpers.

The form is intentionally shorter than the desktop work area, centered, resizable, and scrollable on smaller/high-DPI screens.

#### `Fix-GoogleTakeoutMetadata-v6.2.3-Batch.ps1`

Multi-folder coordinator.

Responsibilities:

- Reads folder-pair manifest.
- Processes folder pairs sequentially.
- Invokes the per-folder orchestrator.
- Aggregates per-folder results.
- Produces the batch summary.
- Must propagate per-folder failure/completeness status accurately.

Do **not** parallelize whole folders by default. Four workers per ten folders should not become forty simultaneous copy/ExifTool jobs.

#### `Fix-GoogleTakeoutMetadata-v6.2.3.ps1`

Per-folder orchestrator.

Responsibilities:

- Scans the selected folder recursively unless `-TopLevelOnly` is used.
- Fast-path known extensions.
- Signature-sniffs unusual-extension candidates.
- Computes total media bytes.
- Disk-space preflight.
- Creates output tree.
- Assigns files to workers by approximate byte load rather than simple file count.
- Starts/monitors worker processes.
- Retries a failed worker bucket once in direct/compatibility mode.
- Merges worker CSV logs.
- Produces per-folder summary.
- Enforces `Files attempted by workers == Media files scanned` completeness.

#### `Fix-GoogleTakeoutMetadata-v6.2.3-Worker.ps1`

Metadata engine used by each worker.

Responsibilities:

- Persistent or direct ExifTool execution.
- Sidecar matching.
- Google timestamp extraction.
- Embedded metadata reads.
- Time decision logic.
- GPS decision logic.
- Content-type/extension mismatch handling.
- Copying source media to output.
- Metadata writing.
- Filesystem timestamp update.
- Read-back verification.
- Per-worker CSV/result/summary output.

This is the highest-risk file for metadata correctness changes.

## Critical invariants

These rules should be treated as regression requirements.

### 1. Never modify source Takeout media

All metadata changes happen in the output tree. A feature that writes in place is out of scope unless explicitly designed as a separate mode with very prominent safeguards.

### 2. `photoTakenTime` is the primary Google capture instant

Google `creationTime` often represents upload/library creation rather than shutter time.

Current behavior:

- prefer `photoTakenTime`;
- use `creationTime` only as a fallback;
- if only `creationTime` exists and a valid embedded capture date exists, preserve the embedded capture date;
- `-ForceJsonTime` is intended to make `photoTakenTime` win, not to replace a good camera date with an upload timestamp.

### 3. Keep UTC instants separate from local wall-clock time

This project previously had a real +7 hour MP4 bug on a Pacific daylight-time machine.

Do not reintroduce it.

For QuickTime core timestamps, Google's Unix timestamp is already an absolute UTC instant. The worker intentionally writes the UTC clock value without applying `QuickTimeUTC=1` during the write path. UTC-aware XMP/Keys dates receive an explicit `Z`.

For photos, an existing local EXIF time may be preserved if its explicit/inferred offset resolves to the same Google UTC instant.

### 4. Verification is part of correctness

ExifTool returning success is not sufficient.

When enabled, the output timestamp must be read back and compared to the expected instant. A mismatch must appear as a verification failure.

### 5. Completeness is part of correctness

The parent must verify:

```text
Files attempted by workers == Media files scanned
```

If not, the run is incomplete and must fail. Never label a partial worker run `DONE`.

### 6. Unmatched media is preserved

If no reliable sidecar match exists, copy the media unchanged and log it. Do not silently omit it and do not guess ambiguous Live Photo associations.

### 7. Content wins over extension when necessary

Takeout/phone exports can contain files such as:

- `.HEIC` containing JPEG bytes;
- `.PNG` containing JPEG bytes;
- `.MP` containing QuickTime/MP4 bytes.

Do not rely exclusively on the extension.

### 8. GPS edits and precision must remain distinguishable

Current default threshold: 30 m.

- prefer Google `geoData` over `geoDataExif`;
- ignore `0,0`;
- preserve more precise embedded coordinates when they essentially agree with Google;
- if `geoData` materially differs from `geoDataExif`, treat it as evidence of a Google Photos location edit and prefer the edited Google location.

## Sidecar matching strategy

The worker intentionally delays expensive JSON parsing.

Current order:

1. Exact filename candidates, including supplemental metadata and duplicate-counter variants.
2. Sidecar filename basename map for common Live Photo companions.
3. Lazily built truncated supplemental-metadata prefix index.
4. JSON `title` / title-basename fallback as the last resort.

Do not eagerly parse every JSON in the entire Takeout; that was intentionally avoided for memory and speed.

Folder indexes use a bounded cache (`128` folders at the time of this handoff).

## Parallel-worker model

v6 uses multiple PowerShell worker processes. Each worker owns its own ExifTool `-stay_open` process.

A media item should stay with one worker for the full operation:

```text
copy -> embedded read -> metadata write -> verification
```

The final CSV is assembled by the parent rather than allowing several workers to write the same final log concurrently.

### Worker count guidance

- 1: HDD/USB compatibility.
- 4: current balanced default for SSD/NVMe.
- 1-8: supported custom range.

Do not automatically use every logical CPU. The workload is heavily storage/metadata-I/O bound.

## Windows PowerShell compatibility traps

### Generic `List[object]` array conversion

Windows PowerShell 5.1 can throw:

```text
Argument types do not match
```

for patterns such as `@($list)` when `$list` is a `List[object]` created through `New-Object`.

The multi-folder path was patched to use `.ToArray()` where appropriate.

**Do not casually replace this with `@($list)` again.**

### Worker manifests

v6.2.2 passed worker file buckets as JSON arrays. On Windows PowerShell 5.1, workers assigned multiple files could exit before producing their result file, while one-file workers appeared fine.

v6.2.3 changed worker manifests to a deliberately simple format:

```text
C:\absolute\path\file1.jpg
C:\absolute\path\file2.mov
C:\absolute\path\file3.heic
```

One absolute media path per line.

**Do not revert the worker manifest to a PowerShell/JSON collection without a Windows PowerShell 5.1 regression test containing multiple files per worker.**

## Important regression history

### Early versions: naive Google JSON merge

Initial work established sidecar matching, copying, date/GPS writes, and non-destructive behavior.

### v3: safer date policy and GUI

Important changes included protecting existing embedded capture dates from Google `creationTime`, output-overlap checks, dry-run workflow, and clearer UI.

### v4: QuickTime timezone fix + verification

A concrete regression test showed a Pacific machine converting an already-UTC Google timestamp a second time:

```text
expected 16:25:01Z
stored   23:25:01Z
```

The root cause was host-local conversion combined with QuickTime UTC behavior. v4 corrected the write path and added read-back verification.

### v4.1: EXIF/XMP cleanup

- upgrade old EXIF declarations to 2.31 only when adding OffsetTime tags;
- write UTC-aware XMP video dates;
- avoid unnecessary EXIF-version warnings.

### v4.2: content/extension mismatch handling

Real files named `.HEIC` contained JPEG data. ExifTool rejected them as invalid HEIC. The worker learned to detect actual content, process through a temporary correctly typed alias, and preserve the original output filename without recompressing pixels.

### v5.0: full-library performance work

- persistent ExifTool `-stay_open`;
- bounded sidecar indexes;
- streamed CSV;
- GPS precision preservation;
- safer cancellation;
- disk-space preflight;
- keep-awake support.

### v5.1: unusual-extension media discovery

A Pixel `.MP` file containing MP4 data was missed by extension-only scanning. v5.1 added signature sniffing for unusual extensions while skipping obvious sidecars/docs/archives/audio-only files.

### v6.0: worker pool

Introduced parallel per-folder workers, each with its own persistent ExifTool session.

### v6.1: multiple input/output folder pairs

Each input folder gets a corresponding independent output. Folder pairs are sequential; workers are parallel only within the active pair.

### v6.2 / v6.2.1 / v6.2.2

- v6.2: shorter/resizable UI with screen margins.
- v6.2.1: fixed the Windows PowerShell `List[object]` / `Argument types do not match` issue.
- v6.2.2: native Explorer-style Ctrl/Shift multi-folder picker.

### v6.2.3: worker-manifest reliability + completeness

v6.2.2 produced a partial 49-file output because three of four workers exited without result files when given multi-file manifests. v6.2.3 moved to one-path-per-line manifests and added a hard completeness check.

Validated v6.2.3 regression result:

```text
Media files scanned: 49
Processed/matched: 49
Verified timestamps: 49
Verification failures: 0
Failed/warned: 0
Scan errors/skips: 0
Workers used: 4
Worker fallback retries: 0
Files attempted by workers: 49 / 49
Completeness check: PASS
```

## Larger-run evidence

A v5.1 full-folder run over a 2021 export produced:

```text
Media files scanned: 4205
Media bytes scanned: about 9.07 GB
Processed/matched: 4126
Verified timestamps: 4126
Verification failures: 0
Copied without sidecar: 79
Failed/warned: 0
Scan errors/skips: 0
Elapsed: about 12m 27s
Average rate: 5.63 files/sec
ExifTool engine: persistent stay_open
```

The 79 unmatched files were preserved rather than omitted. Many were companion MP4 files with same-basename stills, but ambiguous associations were intentionally not guessed.

The large run predates the v6 worker pool, so it is useful as a correctness/performance baseline but not as a v6 parallel benchmark.

## Media-integrity testing performed

Representative tests compared media content before and after metadata rewriting.

Observed behaviors included:

- JPEG decoded pixel hashes unchanged.
- H.264/HEVC encoded video streams unchanged.
- AAC/audio streams unchanged.
- Live Photo content identifiers preserved in tested pairs.
- Pixel motion-photo embedded payload retained.

When changing write logic, repeat media-integrity checks. Metadata-only behavior is a core requirement.

## Known limitations / non-goals

### Google date ambiguity

Google `photoTakenTime` is the date Google currently associates with an item. For downloaded/vendor images with stripped EXIF, it may not be the true shutter time.

Example observed during development: a vendor filename contained a plausible `2023-06-18-09-59-13` timestamp while Google `photoTakenTime` represented a substantially different time. The tool correctly preserves Google's library state rather than treating filename dates as authoritative.

Do not add filename-date overrides by default. If implemented, they should be opt-in/audit-only unless there is very strong evidence.

### Google-specific library state

EXIF/XMP cannot necessarily reconstruct:

- album membership;
- Google face-recognition assignments;
- all favorite/organization semantics;
- all Google-internal edits.

### Ambiguous companion media

Do not automatically attach a still's sidecar to a same-basename MP4 when multiple stills make the association ambiguous. Preserve the file unchanged instead.

### GPU acceleration

Not useful for the current workload. The program intentionally avoids media transcoding. Metadata I/O, file copies, JSON work, and ExifTool are better optimized through worker parallelism and storage-aware tuning.

## Current areas worth future work

These are enhancements, not known blockers in v6.2.3:

1. Benchmark 1/2/4/6/8 workers on a several-thousand-file SSD/NVMe folder and determine whether 4 remains the best default.
2. Add automated Pester tests for pure PowerShell functions such as:
   - path overlap validation;
   - sidecar name variants;
   - timestamp/offset conversion;
   - GPS distance/decision policy;
   - file signature detection.
3. Add a scripted Windows integration fixture containing tiny representative media + JSON sidecars.
4. Consider a filename-date conflict **warning** report, without changing the default metadata source-of-truth policy.
5. Consider cleaner separation between versioned script filenames and internal module code to reduce version-bump churn.
6. Add GitHub Actions for static PowerShell parsing/linting on Windows. Full ExifTool integration tests should run on Windows, not Linux.
7. Decide and add a project license before public distribution if one has not already been selected.

## Regression checklist before a release

At minimum test the following on Windows PowerShell 5.1:

### GUI

- App opens without filling the entire vertical screen.
- Form is resizable.
- ExifTool auto-detection works.
- `Add folder(s)...` supports Ctrl-click and Shift-click multiple folders.
- `Add child folders...` works.
- Each input has its own default `<input> - Fixed` output.
- Duplicate/overlapping mappings are rejected.
- Non-empty output is rejected unless overwrite/re-run is explicitly enabled.

### Single-folder engine

Use a fixture containing:

- ordinary JPEG with sidecar;
- image with existing local time + offset;
- image missing date;
- MP4/MOV with bad/1970 time;
- GPS-tagged photo;
- dummy `0,0` GPS sidecar;
- JPEG bytes named `.HEIC`;
- JPEG bytes named `.PNG`;
- unusual-extension MP4 such as `.MP`;
- Live Photo pair;
- unmatched media file.

Expected:

```text
Verification failures: 0
Failed/warned: 0
Scan errors/skips: 0
Completeness check: PASS
```

### Parallel engine

The test must contain enough files that **each worker receives more than one file**. A one-file-per-worker test will not catch the manifest regression fixed in v6.2.3.

Check:

```text
Files attempted by workers == Media files scanned
Worker fallback retries == 0   (for the normal happy path)
Completeness check == PASS
```

### Multi-folder batch

Select at least three folders in one run. Confirm:

- each has the intended independent output;
- each per-folder CSV/summary is present;
- the batch summary reports actual failures rather than blindly `DONE`;
- cancelling folder 2 does not start later folders;
- completed folder 1 remains intact.

### Media integrity

For representative files compare:

- decoded image pixel hash;
- encoded video stream hash;
- encoded audio stream hash;
- Live Photo identifiers when present.

No recompression/transcoding should occur.

## Packaging notes

Two release variants have been used:

### Full

Includes the tested ExifTool Windows distribution.

When bundling ExifTool:

- preserve its README/license files;
- do not present ExifTool as part of this project's own code;
- update checksums when the ExifTool version changes.

### Scripts Only

Contains the project scripts and docs but not ExifTool.

For GitHub source control, **Scripts Only is the cleaner default**. Attach the Full ZIP as a GitHub Release asset if desired rather than committing the entire third-party ExifTool distribution to normal source history.

## GitHub repository recommendation

Suggested root:

```text
README.md
HANDOFF.md
GoogleTakeoutMetadataFixer-GUI.ps1
Fix-GoogleTakeoutMetadata-v6.2.3-Batch.ps1
Fix-GoogleTakeoutMetadata-v6.2.3.ps1
Fix-GoogleTakeoutMetadata-v6.2.3-Worker.ps1
Run-GoogleTakeoutMetadataFixer-v6.2.3.bat
docs/
    AUDIT-GoogleTakeoutMetadataFixer-v6.2.3.txt
```

Optional `.gitignore` entries:

```gitignore
*-Fixed/
metadata-fix-log*.csv
metadata-fix-summary*.txt
metadata-fix-batch-summary*.txt
*.zip
exiftool-*/
```

If release binaries are hosted through GitHub Releases, keep release ZIP checksums in the release notes or a generated checksum file.

## Release-version note

The v6.2.3 release text inherited some wording that describes Ctrl/Shift folder selection as the v6.2.3 change. The **actual version history** is:

- v6.2.2: native Ctrl/Shift multi-folder picker;
- v6.2.3: worker manifest reliability fix and hard completeness check.

Use this corrected history in future release notes.

## Development environment caveat

The authoring environment used during development was not a native Windows desktop environment, so WinForms and Windows PowerShell integration could not always be executed end-to-end there. The scripts therefore include an on-machine PowerShell AST parse before jobs begin, and important runtime behavior was validated through repeated user-side Windows regression runs.

Future maintainers should prefer actual Windows PowerShell 5.1 integration tests for release qualification.

