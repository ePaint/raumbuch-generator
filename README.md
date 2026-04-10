# Raumbuch Generator

Generates PDF specification sheets from room data using Word templates with placeholders. Supports both Excel files and dRofus API as data sources.

## Features

- Generate individual PDFs or merge all rooms into a single document
- Fast OpenXML-based placeholder replacement (~0.6s per room)
- Support for Excel and dRofus API data sources

## Requirements

- Windows with Microsoft Word installed
- PowerShell 5.1+
- [ImportExcel](https://github.com/dfinke/ImportExcel) module

## Quick Start

```powershell
# Install dependency
Install-Module ImportExcel -Scope CurrentUser

# Generate PDFs for all rooms
.\RoomToPDF.ps1

# Generate a single merged PDF
.\RoomToPDF.ps1 -Merge

# Process specific rooms
.\RoomToPDF.ps1 -RoomCode "RT.001,RT.017"

# Specific rooms merged into one PDF
.\RoomToPDF.ps1 -RoomCode "RT.001,RT.002,RT.003" -Merge
```

## Project Structure

```
raumbuch-generator/
├── RoomToPDF.ps1             # Main script
├── config.psd1               # Configuration
├── api-key.txt               # API key (not in repo)
├── api-call.txt              # API endpoint URL (not in repo)
├── DataDictionary.xlsx       # API field reference (dyn_rfp_* → readable names)
├── Input/
│   ├── Raumbuch_Vorlage_API.docx  # Template with API placeholders
│   └── ...
└── Output/
    └── 2026-04-10_22-28-33/
        ├── AllRooms_Merged.pdf    # When using -Merge
        ├── RT.001.pdf             # Individual PDFs
        └── ...
```

## Command Line Options

| Parameter | Description |
|-----------|-------------|
| `-RoomCode` | Process specific room(s), comma-separated |
| `-Source` | Data source: `Excel` or `API` (overrides config) |
| `-Template` | Template file path (overrides config) |
| `-ExcelFile` | Excel data file path (overrides config) |
| `-ConfigPath` | Use a different config file |
| `-Merge` | Combine all rooms into a single PDF |

```powershell
.\RoomToPDF.ps1                                    # All rooms, individual PDFs
.\RoomToPDF.ps1 -Merge                             # All rooms, single PDF
.\RoomToPDF.ps1 -RoomCode "RT.001" -Source API     # Single room from API
.\RoomToPDF.ps1 -RoomCode "RT.001,RT.002" -Merge   # Multiple rooms, merged
```

## Configuration

Edit `config.psd1`:

```powershell
@{
    TemplateFile = 'Input/Raumbuch_Vorlage_API.docx'
    OutputFolder = 'Output'
    DataSource   = 'API'    # 'Excel' or 'API'

    # Excel settings (when DataSource = 'Excel')
    Excel = @{
        DataFile       = 'Input/data.xlsx'
        RoomCodeColumn = 'Code'
    }

    # API settings (when DataSource = 'API')
    API = @{
        EndpointFile  = 'api-call.txt'
        KeyFile       = 'api-key.txt'
        RoomCodeField = 'room_func_no'
    }

    # Value replacements (case-insensitive)
    ValueMap = @{
        'true'  = 'ja'
        'false' = 'nein'
    }
}
```

## Template Format

Use `<<fieldname>>` placeholders anywhere in your Word document. For API integration, use the raw identifiers:

```
<<dyn_rfp_01010201>>    # max. ständige Arbeitsplätze
<<dyn_rfp_01010401>>    # Tageslicht direkt
```

To look up which identifier corresponds to which field, see `DataDictionary.xlsx`.

### Formatting Notes

- Placeholders can be inside tables, paragraphs, headers, anywhere
- Text formatting (bold, colors, fonts) is preserved
- Values are mapped via `ValueMap` in config (e.g., `true` → `ja`)
- Values over 255 characters are truncated

## API Setup (dRofus)

1. Create `api-key.txt` with your API key
2. Create `api-call.txt` with the endpoint URL
3. Set `DataSource = 'API'` in config.psd1

Both files should be in the project root directory and contain only the key/URL (no extra whitespace).

## Output

Each run creates a timestamped folder in `Output/`:

```
Output/
└── 2026-04-10_22-28-33/
    ├── AllRooms_Merged.pdf    # With -Merge flag
    ├── RT.001.pdf             # Without -Merge
    ├── RT.002.pdf
    └── ...
```

## Performance

| Rooms | Time | Per Room |
|-------|------|----------|
| 3 | ~4s | 0.9s |
| 182 | ~107s | 0.6s |

## Troubleshooting

**"ImportExcel module is required"**
```powershell
Install-Module ImportExcel -Scope CurrentUser
```

**"No <<placeholder>> markers found in template"**
- Ensure the template has `<<fieldname>>` placeholders
- For API: use `dyn_rfp_*` identifiers (see DataDictionary.xlsx)

**Placeholder not replaced**
- Ensure placeholder name exactly matches the data field
- For API: use `dyn_rfp_*` identifiers, not descriptive names

**Word process hangs**
- End `WINWORD.EXE` in Task Manager
- The script normally cleans up automatically

## License

MIT
