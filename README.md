# Google Takeout Metadata Fixer

A Windows PowerShell tool for restoring Google Photos Takeout sidecar metadata into copies of photos and videos before re-uploading them to Google Photos or archiving them elsewhere.

**Current version: v6.2.3**

> This project is not affiliated with Google. Metadata writing is performed with Phil Harvey's ExifTool.

## Why this exists

Google Photos Takeout commonly exports media and metadata separately. A photo or video may be accompanied by a JSON sidecar such as:

```text
IMG_1234.JPG
IMG_1234.JPG.supplemental-metadata.json
```

The JSON can contain the Google Photos capture time, location, caption, favorite state, and other information that is not necessarily embedded in the exported media file.

This tool matches the media to its sidecar, copies the media into a separate output tree, writes portable metadata with ExifTool, and then reads the result back to verify the timestamp.

## Highlights

- Non-destructive: original Takeout media is never modified.
- Each input folder has its own independent output folder.
- Ctrl-click / Shift-click multi-folder selection in the Windows folder picker.
- Recursive folder processing by default.
- Parallel worker pool: 1-8 ExifTool workers per active folder.
- Persistent ExifTool `-stay_open` sessions for lower process-start overhead.
- Post-write timestamp verification.
- Correct UTC handling for MP4/MOV/QuickTime timestamps.
- Preserves plausible existing local EXIF wall-clock times and offsets.
- Restores GPS while ignoring dummy `0,0` coordinates.
- Preserves more precise embedded GPS when it agrees with Google's rounded location.
- Detects likely Google Photos location edits.
- Handles Google Takeout supplemental-metadata filename quirks and truncated names.
- Handles Live Photo / motion-photo companion matching.
- Detects content/extension mismatches such as JPEG data named `.HEIC` or `.PNG`.
- Signature-sniffs unusual extensions so real media such as Pixel `.MP` components are not silently omitted.
- Unmatched media is copied unchanged rather than discarded.
- CSV audit log plus per-folder summary and multi-folder batch summary.
- Safe cancellation and optional keep-awake behavior for long runs.

## Requirements

- Windows 10 or Windows 11.
- Windows PowerShell 5.1 is the primary supported runtime.
- ExifTool 13.59 or newer is recommended; 13.59 is the version used during development and regression testing.
- Enough free disk space for a second copy of the media being processed.

The **Full** release package includes the tested Windows ExifTool distribution. The **Scripts Only** package expects ExifTool to be supplied separately.

The GUI can locate ExifTool when it is:

- next to the scripts as `exiftool.exe` or `exiftool(-k).exe`;
- inside a sibling `exiftool-*` directory; or
- available on `PATH` as `exiftool.exe`.

## Quick start

1. Keep your original Google Takeout intact.
2. Extract the release to a normal Windows folder.
3. Double-click:

   ```text
   Run-GoogleTakeoutMetadataFixer-v6.2.3.bat
   ```

4. Add one or more Google Photos Takeout folders.
   - **Add folder(s)...** supports Ctrl-click and Shift-click multi-selection.
   - **Add child folders...** lets you choose several immediate child folders under one parent.
5. Confirm the automatically generated output paths. By default:

   ```text
   Photos from 2021  ->  Photos from 2021 - Fixed
   Photos from 2022  ->  Photos from 2022 - Fixed
   ```

6. Recommended settings for normal use:

   ```text
   Create fixed copies - preserve plausible existing local EXIF times
   [x] Include all subfolders
   [x] Verify written timestamps
   [x] Keep computer awake during long runs
   Balanced - 4 workers
   ```

7. For a new or unfamiliar Takeout, run **Dry run** first.
8. After a real run, review `metadata-fix-summary.txt` in every output folder.

A healthy summary should normally end with values like:

```text
Verification failures: 0
Failed/warned: 0
Scan errors/skips: 0
Completeness check: PASS
```

Do not delete the original Takeout until you have spot-checked the fixed files and, if applicable, confirmed that Google Photos imports them as expected.

## Processing modes

### Dry run

Scans media, matches JSON sidecars, and creates audit information without copying or modifying media.

### Create fixed copies - preserve plausible existing local EXIF times

Recommended mode.

Google `photoTakenTime` is treated as the known absolute capture instant. If a photo already contains a valid local camera time and offset that resolve to the same instant, the local wall-clock representation is preserved.

### Create fixed copies - force Google JSON `photoTakenTime` to win

Use this when you intentionally corrected dates in Google Photos and want Google's `photoTakenTime` to override existing embedded capture dates.

This mode does **not** blindly replace a valid embedded capture date with Google `creationTime` when `photoTakenTime` is absent.

## Time policy

The tool deliberately distinguishes between an absolute UTC instant and a local camera wall-clock time.

For photos:

1. `photoTakenTime` is the preferred Google timestamp.
2. A valid existing local EXIF time plus explicit/inferable offset may be preserved if it represents the same instant.
3. If no usable local time exists, the Google instant is written with an explicit UTC offset rather than guessing the computer's local timezone.
4. `creationTime` is only a fallback and is not allowed to casually overwrite a real embedded capture date.

For QuickTime/MP4/MOV files, the core movie/track/media timestamps are written as the known UTC instant. This avoids the historical double-timezone conversion bug where a Pacific machine could shift a UTC value by 7 or 8 hours.

## GPS policy

Google Takeout can contain both `geoData` and `geoDataExif`.

- `geoData` is preferred because it can represent the current Google Photos location, including edits.
- `geoDataExif` is treated as the original embedded location fallback.
- `0,0` coordinates are ignored as missing data.
- When the existing embedded coordinates and Google coordinates agree within the configured threshold (30 m by default), the more precise embedded coordinates are preserved.
- If `geoData` differs materially from `geoDataExif`, the tool can treat that as a Google Photos location edit and use the edited Google location.

## Sidecar matching

Matching is intentionally layered from cheapest/most deterministic to more expensive fallbacks:

1. Exact expected JSON filename.
2. Supplemental-metadata naming variants and duplicate counters.
3. Sidecar filename basename / Live Photo companion matching.
4. Truncated supplemental-metadata prefix matching.
5. JSON `title` matching as a last resort.

Folder indexes are cached with a bounded cache so large Takeouts do not retain every parsed JSON object in memory.

## File-type handling

Filename extensions are not always trustworthy in Takeout exports.

The fixer handles cases such as:

```text
IMG_1019.HEIC   -> actual bytes are JPEG
some-photo.PNG  -> actual bytes are JPEG
PXL_....MP      -> actual bytes are QuickTime/MP4
```

Known media extensions use the fast path. Unknown candidates are signature-sniffed using a small header read. Common sidecars, archives, documents, executables, and audio-only files are excluded from signature probing for performance.

## Multiple folders

Each input folder has a corresponding output folder. Folder pairs are processed **sequentially** so selecting ten folders does not accidentally create ten times the requested worker count.

Inside the active folder, media can be processed in parallel by the worker pool.

Example:

```text
Photos from 2021  -> Photos from 2021 - Fixed
    4 workers in parallel
then
Photos from 2022  -> Photos from 2022 - Fixed
    4 workers in parallel
```

The GUI validates all mappings before starting and rejects duplicate or overlapping input/output paths.

## Worker modes

- **Conservative - 1 worker**: recommended for HDDs, slower USB storage, or troubleshooting.
- **Balanced - 4 workers**: default starting point for SSD/NVMe storage.
- **Custom - 1 to 8 workers**: useful for benchmarking your hardware.

Each worker owns its own persistent ExifTool process and keeps a media item with the same worker through copy, metadata write, and verification.

More workers are not always faster. Storage throughput and random I/O usually become the limiting factors before CPU usage does.

## Verification and completeness

A successful ExifTool exit code alone is not treated as proof that a timestamp was written correctly.

When verification is enabled, the tool reads the output metadata back and compares the stored instant with the expected Google timestamp.

v6.2.3 also performs a hard completeness check:

```text
Files attempted by workers: 49 / 49
Completeness check: PASS
```

A worker-pool failure that leaves files unattempted causes the folder run to fail instead of being incorrectly reported as complete.

## Output files

Each output folder receives:

```text
metadata-fix-log.csv
metadata-fix-summary.txt
```

The CSV contains per-file matching, type detection, timestamp, GPS, verification, and status information.

A multi-folder run also creates a batch summary describing each folder pair and its result.

## Command-line use

The GUI is the recommended interface, but the per-folder orchestrator can be run directly:

```powershell
.\Fix-GoogleTakeoutMetadata-v6.2.3.ps1 `
  -InputFolder 'D:\Takeout\Google Photos\Photos from 2021' `
  -OutputFolder 'D:\Takeout\Google Photos\Photos from 2021 - Fixed' `
  -ExifToolPath 'D:\Tools\exiftool.exe' `
  -WorkerCount 4
```

Dry run:

```powershell
.\Fix-GoogleTakeoutMetadata-v6.2.3.ps1 `
  -InputFolder 'D:\Takeout\Google Photos\Photos from 2021' `
  -ExifToolPath 'D:\Tools\exiftool.exe' `
  -DryRun `
  -WorkerCount 4
```

Useful switches include:

```text
-DryRun
-ForceJsonTime
-TopLevelOnly
-OverwriteOutput
-SkipVerification
-DisablePersistentExifTool
-AllowSystemSleep
-WorkerCount 1..8
-GpsAgreementMeters <number>
```

## Repository layout

```text
GoogleTakeoutMetadataFixer-GUI.ps1
Fix-GoogleTakeoutMetadata-v6.2.3-Batch.ps1
Fix-GoogleTakeoutMetadata-v6.2.3.ps1
Fix-GoogleTakeoutMetadata-v6.2.3-Worker.ps1
Run-GoogleTakeoutMetadataFixer-v6.2.3.bat
README.md
HANDOFF.md
```

Release packages may additionally include:

```text
AUDIT-GoogleTakeoutMetadataFixer-v6.2.3.txt
SHA256SUMS.txt
exiftool-13.59_64\...
```

## Tested behavior

Regression testing during development covered representative JPEG, HEIC, mislabeled JPEG-as-HEIC, PNG-named JPEG, MP4, MOV, GPS-tagged media, iPhone Live Photos, Pixel motion-photo components, missing timestamps, existing timezone-aware EXIF, and QuickTime timestamps.

Observed test runs included:

- a 2021 folder with 4,205 media files / about 9.07 GB: 4,126 matched files verified, 79 unmatched media copied unchanged, and zero verification failures;
- a v6.2.3 four-worker regression with 49/49 files attempted, 49/49 processed and verified, zero fallback retries, and `Completeness check: PASS`.

The project is metadata-oriented and does not intentionally resize/recompress images or transcode video/audio. Representative regression tests confirmed unchanged decoded image pixels and unchanged encoded video/audio streams.

## Important limitations

- Google `photoTakenTime` is the date Google Photos associates with an item; for media with stripped/no original EXIF, it is not guaranteed to be the historically true shutter time.
- A timestamp embedded in a filename is treated as a heuristic, not authoritative metadata.
- Some Google Photos concepts do not have a perfect portable EXIF/XMP equivalent. Face recognition, album membership, and Google-specific organization may not reconstruct on re-upload.
- Unmatched motion/Live Photo companion files are preserved rather than guessed when the association is ambiguous.
- Always keep the original Takeout until the migration has been verified.

## ExifTool

ExifTool is a third-party project by Phil Harvey and has its own license and documentation. If you distribute a release that bundles ExifTool, preserve the ExifTool license/readme files included with the official Windows package.

## Contributing / development

See [`HANDOFF.md`](HANDOFF.md) for architecture, invariants, regression history, Windows PowerShell compatibility traps, testing guidance, and release packaging notes.

