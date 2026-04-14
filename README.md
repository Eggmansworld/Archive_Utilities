# Eggman's Archive Utilities

A multi-tab Python/tkinter desktop utility for archive extraction and compression — built around workflows that standard consumer archive tools don't handle cleanly out of the box.

![Python](https://img.shields.io/badge/Python-3.10%2B-blue) ![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey) ![License](https://img.shields.io/badge/License-MIT-green)

---

## Overview

| Tab | Function |
|---|---|
| 📦 **Extractor** | Recursive ZIP / 7Z / RAR extraction with nested archive detection and flexible post-extract actions |
| 🗜 **ZIP Store Packer** | Wrap files in zero-compression ZIP containers for downstream recompression pipelines |

---

## Requirements

**Python 3.10 or later**

**Required:**
- [7-Zip-Zstandard](https://github.com/mcmilk/7-Zip-zstd/releases) installed at:
  ```
  C:\Program Files\7-Zip-Zstandard\7z.exe
  ```
  > Edit the `SEVEN_ZIP` constant at the top of the script to use a different path. Standard 7-Zip (`C:\Program Files\7-Zip\7z.exe`) works fine if you don't need ZSTD support.

**Optional but recommended:**
```
pip install tkinterdnd2    # drag-and-drop on all folder inputs
pip install send2trash     # Recycle Bin delete in the Extractor tab
```

The application runs without these — affected features are disabled and flagged with a warning in the title bar.

---

## Usage

```
python eggmans_archive_utility.py
```

No installer required.

---

## Tab Reference

---

### 📦 Extractor

Recursively scans a source folder for archives and extracts each one into its own named subfolder. Files from different archives never mix together in the same directory.

**Supported formats:** `.zip`  `.7z`  `.rar`

**Source folder** — Drop a folder onto the drop zone or use Browse.

#### Destination

| Mode | Behaviour |
|---|---|
| Same as source | Extracted folder is created beside the archive |
| Mirror to custom destination | Full source folder structure is mirrored at a new root — e.g. `D:\source\sub\file.zip` extracts to `E:\dest\source\sub\file\` |

#### Extraction rules

- ZIP archives containing exactly one file at the root with no subdirectories are extracted flat — no wrapper folder is created
- All other archives extract into a folder named after the archive stem
- Double-nesting is automatically detected and collapsed — if `game.zip` extracts a single folder also named `game`, the contents are promoted one level up, eliminating the redundant wrapper
- Folder merging handles collisions — existing content is never silently overwritten

#### After extraction

| Option | Behaviour |
|---|---|
| Keep archive | Original archive left in place |
| → Recycle Bin | Archive sent to Windows Recycle Bin *(requires `send2trash`)* |
| → Permanent delete | Archive permanently deleted |
| → Move (mirror structure) | Archive moved to a separate folder, preserving its path relative to the source root |
| → Move (flat dump) | Archive moved into a single flat folder regardless of where it came from; collisions renamed `name(1).ext`, `name(2).ext` etc. |

The two Move options are particularly useful when combined with nested archive detection — processed archives are cleared out of the way before their children are handled.

#### Nested archive detection

After every successful extraction the output folder is scanned for files matching the selected formats. If any are found, a prominent alert fires in the log:

```
▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
  ⚠  NESTED ARCHIVES DETECTED in: GOLD
  ↳  4 archive(s) found after extracting GOLD.RAR
       [FOUND]   part1.rar  (not auto-extracting)
▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
```

Enabling **Auto-extract nested archives** pushes discovered archives onto the back of the processing queue so they are handled automatically in the same run. The stat bar tracks `X/Y (+Z queued)` while the queue is growing. Already-queued files are deduplicated.

#### Other options

- **Recursive** — scan subdirectories (default on)
- **Format checkboxes** — enable/disable `.zip`, `.7z`, `.rar` independently
- **Stop** — halts after the current file finishes; no partial extractions are left behind

---

### 🗜 ZIP Store Packer

Recursively scans a folder for files matching one or more extensions, wraps each file individually in a **ZIP_STORED** (zero compression) archive in-place, verifies the archive integrity, and deletes the original file.

Designed for data preservation and archival workflows where files need to be containerized in a neutral zip wrapper before being passed to a downstream compression tool — such as a deterministic recompressor like ZipTorrent or RomVault's ZStandard (RVZSTD).

**Target folder** — Drop or browse. Processing is recursive by default.

#### Why not just use 7-Zip or WinRAR?

The core problem is **one archive per file**. Both 7-Zip and WinRAR are designed to collect files into an archive, not to wrap each file individually in its own archive and leave it in place. That workflow simply doesn't exist in their interfaces.

| Feature | ZIP Store Packer | 7-Zip GUI | 7-Zip CLI | WinRAR GUI |
|---|:---:|:---:|:---:|:---:|
| ZIP STORE (no compression) | ✅ | ✅ | ✅ | ✅ |
| Recursive folder scan | ✅ | ✅ | ✅ | ✅ |
| Filter by file extension | ✅ | ❌ | ⚠️¹ | ❌ |
| One ZIP per file, in-place | ✅ | ❌ | ⚠️¹ | ❌ |
| Verify archive before delete | ✅ | ❌ | ⚠️¹ | ❌ |
| Delete original after verify | ✅ | ❌ | ⚠️¹ | ✅² |
| No scripting required | ✅ | ✅ | ❌ | ✅ |
| Multiple extensions in one run | ✅ | ❌ | ❌ | ❌ |

> ¹ Possible via a PowerShell `foreach` loop calling `7z.exe` per file — but that's writing a script to replicate what this tool already does.
> ² WinRAR can delete after archiving, but the delete is not conditional on a separate integrity test.

This workflow — recursive, filtered, one-zip-per-file, in-place, verify-before-delete — falls squarely in the gap between what GUI archivers do and what a one-line command can do. Field-tested on folders containing tens of thousands of subfolders with over 200,000 files in a single operation.

#### Extensions

File extensions to target are managed as removable pill tags. Type one or more into the input box and press **Add** or **Enter**:

| Input | Result |
|---|---|
| `exe` | adds `.exe` |
| `.dll` | adds `.dll` |
| `bat com` | adds `.bat` and `.com` |
| `bin, rom` | adds `.bin` and `.rom` |

Click **✕** on any pill to remove it. `.exe` is loaded by default.

#### How it works

```
For each matching file:
  ├── Create  →  filename.zip  (ZIP_STORED, allowZip64)
  ├── Verify  →  testzip() + file size cross-check
  ├── Delete  →  original file removed only if verify passes
  └── Log     →  OK / SKIP / FAIL with file size reported
```

If creation or verification fails for any reason, the incomplete zip is deleted and the original file is left untouched.

#### Options

- **Recursive** — scan subdirectories (default on)
- **Verify before delete** — run integrity check before removing original (default on)
- **Skip if .zip already exists** — avoids reprocessing files already wrapped; the original is left untouched (default on)

#### Intended workflow

ZIP Store Packer is the first stage of a two-pass archival pipeline:

1. **ZIP Store Packer** — wraps files in uncompressed ZIP containers *(this tool)*
2. **Downstream recompressor** — applies deterministic ZSTD or other compression to the zip files

Keeping the two stages separate allows compression parameters to be changed or re-applied at any time without needing to re-wrap the original files.

---

## Log Pane

All tabs share the same log pane:

- **Colour-coded output** — green OK · red fail · amber warn · cyan info · purple nested alerts
- **Horizontal + vertical scrollbars** — long paths are fully readable without resizing the window
- **Save Log** — writes the full log as UTF-8 text on demand
- **Clear Log** — clears between runs

---

## Configuration

The only value likely to need changing is the 7-Zip executable path at the top of the script:

```python
SEVEN_ZIP = r"C:\Program Files\7-Zip-Zstandard\7z.exe"
```

---

## Notes

- Processing is single-threaded and sequential within each run
- The Stop button always completes the current file before halting — no partial results are left behind
- Window opens at 900×900 with a minimum size of 900×600; fully resizable

---

## Related

**[Eggman's Scene Tools](https://github.com/)** — companion utility for working with scene releases (ZIP comment stripping, EOCD repair, and more).

---

## License

MIT
