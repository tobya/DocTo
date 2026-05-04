# DocTo – Agent Guide

## Project Overview

**DocTo** is a Windows command-line utility written in **Delphi** that converts Microsoft Office documents (Word `.doc`/`.docx`, Excel `.xls`/`.xlsx`, PowerPoint `.ppt`/`.pptx`, Visio `.vsd`) to other formats (PDF, TXT, RTF, CSV, etc.) using **COM Automation** against the locally installed Office applications.

- Repository: https://github.com/tobya/DocTo
- Website: https://tobya.github.io/DocTo/

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Language | Delphi (tested with 10.3; compatible with XE4+) |
| Office integration | Windows COM / Office Interop (Word, Excel, PowerPoint, Visio) |
| Build system | Delphi IDE / `.dproj` project file |
| Tests | Batch scripts (`.bat`) in `/test/` |
| Docs / companion site | Markdown + PHP (`/pages/`, `/companion/`) |

## Repository Layout

```
src/           # Delphi source (.pas, .dpr, .dproj, TLB type libraries)
  MainUtils.pas
  ExcelUtils.pas
  PowerPointUtils.pas
  ResourceUtils.pas
  PathUtils.pas
  Visio_TLB.pas
  Word_TLB_Constants.pas
  Excel_TLB_Constants.pas
  PowerPoint_TLB_Constants.pas
  Exceptions/    # Custom exception classes
  shared/        # Shared utilities
test/           # Functional / integration tests (batch scripts)
  InputFiles/        # Sample Word/Excel/PP input docs
  inputfilesxl/      # Excel-specific inputs
  inputfilespp/      # PowerPoint-specific inputs
  inputfilesvs/      # Visio-specific inputs
  GeneratedFiles/    # Output from test runs (gitignored)
pages/          # GitHub Pages / documentation site
companion/      # Companion tooling and generator templates
exe/            # Pre-built release executables
Releases/       # Release archives
```

## Building

- Open `src/docto.dproj` in Delphi IDE and build, **or** use the Delphi command-line compiler (`dcc32`/`dcc64`).
- The project is **Windows-only** — it relies on COM and Office automation which have no Linux equivalents.

## Running Tests

Tests are `.bat` scripts in `/test/`. They call the compiled `docto.exe` and verify output files are produced. Run them directly from a Windows command prompt with Office installed:

```bat
cd test
testDocTo_basic.bat
```

There is no automated test runner — tests must be run manually on a machine with Microsoft Office installed.

## Key Concepts for Agents

- **Application flags**: `-WD` (Word), `-XL` (Excel), `-PP` (PowerPoint), `-VS` (Visio). Word is the default.
- **Three required parameters**: `-F` (input file/dir), `-O` (output file/dir), `-T` (format type, e.g. `wdFormatPDF`).
- **Format types**: Passed as named constants (e.g. `wdFormatPDF`, `xlCSV`) or integers matching the Office Interop enums.
- **COM errors**: Office automation can raise `EOleException`. The `-X` flag controls whether DocTo halts or continues on COM errors.
- **TLB constants**: `Word_TLB_Constants.pas`, `Excel_TLB_Constants.pas`, and `PowerPoint_TLB_Constants.pas` define the Office format enum values.

## Contribution Guidelines

- Open an issue before large PRs to avoid wasted effort.
- The main development branch is `DocTo` (note: not `main`).
- Looking for help with: Delphi/VBA features, PHP/Laravel/Pest tests, and documentation.
- PRs are welcome.
