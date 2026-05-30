# DocTo — Agent Guide

## Project Overview

DocTo is a Windows command-line utility written in **Delphi (Object Pascal)** that converts Microsoft Office documents (Word `.doc`/`.docx`, Excel `.xls`/`.xlsx`, PowerPoint `.ppt`/`.pptx`) to other formats (PDF, CSV, TXT, RTF, etc.) via COM Automation. Microsoft Word, Excel, or PowerPoint must be installed on the host machine.

## Repository Layout

```
docTo/
├── src/                    # All Delphi source files (.pas, .dpr, .dproj)
│   ├── docto.dpr           # Project file (entry point)
│   ├── MainUtils.pas       # Core TDocumentConverter class and shared utilities
│   ├── baseConfig.pas      # Abstract TParamLoader base class for CLI params
│   ├── configInput.pas     # -F / --inputfile parameter handler
│   ├── configOutput.pas    # -OX / --outputextension parameter handler
│   ├── WordUtils.pas       # Word COM interop helpers
│   ├── ExcelUtils.pas      # Excel COM interop helpers
│   ├── PowerPointUtils.pas # PowerPoint COM interop helpers
│   ├── PathUtils.pas       # Path/directory utilities
│   ├── ResourceUtils.pas   # String resource helpers
│   ├── datamodSSL.*        # Data module for SSL/webhook support
│   ├── shared/             # Shared/common units
│   ├── Exceptions/         # Custom exception types
│   └── res/                # Resource files
├── test/                   # Manual test scripts and fixture files
│   ├── testDocTo.bat       # Main test runner batch script
│   ├── InputFiles/         # Sample Word/RTF/CSV/XLS input files
│   ├── inputfilesxl/       # Excel-specific input fixtures
│   ├── inputfilespp/       # PowerPoint-specific input fixtures
│   ├── GeneratedFiles/     # Output directory for test conversions
│   └── GeneratedTestputFiles/
├── .github/
│   ├── workflows/          # GitHub Actions (greetings bot)
│   └── ISSUE_TEMPLATE/
├── pages/                  # GitHub Pages / documentation content
├── companion/              # Companion tooling
├── exe/                    # Pre-built binaries
├── readme.md
└── changes.md
```

## Architecture

### Core Pattern: TParamLoader

CLI parameters follow a **registration/dispatch** pattern:

- `TParamLoader` (`baseConfig.pas`) — abstract base class with three responsibilities:
  - `RegisterParams(List)` — adds the parameter key(s) it handles (e.g. `-F`, `--INPUTFILE`) to a lookup list
  - `Load(Converter, Param, Value)` — applies the parsed value to the `TDocumentConverter` instance
  - `ShouldDec` — whether parsing should decrement the argument index after processing
- Each parameter has a dedicated subclass (e.g. `TParamInput`, `TParamOutputExtension`)
- `TDocumentConverter` (`MainUtils.pas`) is the central domain object passed through all param loaders

### Converters

Three COM-based converter paths:
- **Word** (default): use `-WD` flag or omit; format constants from `wdSaveFormat`
- **Excel**: use `-XL` flag; format constants from `xlFileFormat`
- **PowerPoint**: use `-PP` flag

### Logging

Log levels are integers: `1` ERRORS, `2` STANDARD (default), `5` CHATTY, `9` DEBUG, `10` VERBOSE.
Use `Converter.logdebug(msg, LEVEL)` for diagnostic output.

## Building

- **Compiler**: Embarcadero Delphi (tested with 10.3+; XE4 and XE7 also supported)
- **Platform**: Windows only — relies on COM, Word/Excel/PowerPoint interop
- Open `src/docto.dproj` in the Delphi IDE and build, or use the Delphi command-line compiler (`dcc32`)
- Output is a single `docto.exe` binary

No external package manager or build script is present. The project has no Linux/macOS build path.

## Testing

Tests are manual batch scripts in `test/`:

```bat
# Run the main test suite (requires Word/Excel/PowerPoint installed)
.\test\testDocTo.bat
```

- Input fixtures live in `test/InputFiles/`, `test/inputfilesxl/`, `test/inputfilespp/`
- Outputs are written to `test/GeneratedFiles/` and `test/GeneratedTestputFiles/`
- There is no automated unit-test framework; correctness is verified by inspecting generated files

When adding a new conversion feature, add a corresponding test case to `testDocTo.bat` and provide a sample input file under the appropriate `InputFiles*` directory.

## Error Codes

| Code | Meaning |
|------|---------|
| 200 | Invalid file format specified |
| 201 | Insufficient inputs (need -F, -O, -T at minimum) |
| 202 | Switch requires a value |
| 203 | Unknown switch |
| 204 | Input file does not exist |
| 205 | Invalid parameter value |
| 220 | Word/Excel/PowerPoint COM error |
| 221 | Word/Excel/PowerPoint not installed |
| 400 | Unknown error |

## Adding a New CLI Parameter

1. Create a new unit in `src/` (e.g. `configMyParam.pas`)
2. Declare a class that extends `TParamLoader`
3. Implement `RegisterParams` — add the short and long flag names
4. Implement `Load` — read `Value` and set the appropriate field on `TDocumentConverter`
5. Implement `ShouldDec` — return `false` unless the param consumes an extra token
6. Register the new class in the parameter dispatch table in `MainUtils.pas`
7. Add documentation for the new flag to `readme.md` under "Command Line Help"

## Key Conventions

- Path handling: always call `ExpandFileName` on user-supplied paths to resolve relative references; use `IncludeTrailingBackslash` for directory paths
- Input validation: use `Converter.HaltWithError(code, message)` to exit with a defined error code
- COM errors: wrap COM calls in try/except and honour the `-X` (halt-on-error) flag
- String constants and user-visible messages should be placed in `ResourceUtils.pas` (resource strings), not inlined
- The main branch is named `DocTo` (not `main`)

## External References

- Word SaveAs format constants: https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdsaveformat
- Excel file format constants: https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlfileformat
- Word compatibility mode values: https://msdn.microsoft.com/en-us/library/office/ff192388.aspx
- Releases: https://github.com/tobya/DocTo/releases
- Wiki & examples: https://github.com/tobya/DocTo/wiki
