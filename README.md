# Raumbuch Generator

Generates PDF specification sheets from room data using Word templates with placeholders.

## How It Works

1. You create a Word template with placeholders like `<<room_func_no>>`, `<<name>>`, `<<programmed_area>>`
2. The script loads room data from Excel or an API
3. For each room, it replaces placeholders with actual values and exports to PDF

## Requirements

- Windows with Microsoft Word installed
- PowerShell 5.1+
- [ImportExcel](https://github.com/dfinke/ImportExcel) module

## Quick Start

```powershell
# Install dependency
Install-Module ImportExcel -Scope CurrentUser

# Run with defaults
.\RoomToPDF.ps1

# Process specific rooms
.\RoomToPDF.ps1 -RoomCode "RT.001,RT.017"

# Use API instead of Excel
.\RoomToPDF.ps1 -Source API

# Use a different template
.\RoomToPDF.ps1 -Template "Input/my-template.docx"
```

## Project Structure

```
raumbuch-generator/
├── RoomToPDF.ps1           # Main script
├── config.psd1             # Configuration
├── api-key.txt             # API key (not in repo)
├── Input/
│   ├── sample-template.docx  # Example template
│   ├── sample-data.xlsx      # Example data
│   └── ...                   # Your files
└── Output/
    └── 2026-04-04_16-00-00/
        ├── RT.001.pdf
        └── ...
```

## Configuration

Edit `config.psd1`:

```powershell
@{
    TemplateFile = 'Input/template.docx'
    OutputFolder = 'Output'
    DataSource   = 'Excel'   # or 'API'

    Excel = @{
        DataFile       = 'Input/data.xlsx'
        RoomCodeColumn = 'Code'
    }

    API = @{
        EndpointFile  = 'api-endpoint.txt'
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

Use `<<fieldname>>` placeholders anywhere in your Word document:

```
Room Specification Sheet

Code: <<room_func_no>>
Name: <<name>>
Area: <<programmed_area>> m²

Description:
<<description>>
```

Placeholder names must match the column headers in your Excel file or field names from the API.

### Formatting

- Placeholders can be inside tables, paragraphs, headers, anywhere
- Text formatting (bold, colors, fonts) is preserved
- Values are mapped via `ValueMap` in config (e.g., `true` → `ja`)
- Values over 255 characters are truncated (Word limitation)

## Command Line Options

| Parameter | Description |
|-----------|-------------|
| `-RoomCode` | Process specific room(s), comma-separated |
| `-Source` | Data source: `Excel` or `API` (overrides config) |
| `-Template` | Template file path (overrides config) |
| `-ConfigPath` | Use a different config file |

Examples:

```powershell
# All rooms from Excel
.\RoomToPDF.ps1

# Single room from API
.\RoomToPDF.ps1 -RoomCode "RT.001" -Source API

# Multiple rooms with custom template
.\RoomToPDF.ps1 -RoomCode "RT.001,RT.002" -Template "Input/simple.docx"
```

## Data Sources

### Excel

Set `DataSource = 'Excel'` in config. The script reads all columns from the Excel file. Each column header becomes a placeholder name.

### API (dRofus)

Set `DataSource = 'API'` in config. Requires:
- `api-key.txt` with your API key
- Endpoint URL in a file (configured via `EndpointFile`)

## Output

Each run creates a timestamped folder in `Output/`:

```
Output/
└── 2026-04-04_16-00-00/
    ├── RT.001.pdf
    ├── RT.002.pdf
    └── ...
```

## Troubleshooting

**"ImportExcel module is required"**
```powershell
Install-Module ImportExcel -Scope CurrentUser
```

**"Template file not found"**
- Check paths in config.psd1 (relative to script directory)

**Placeholder not replaced**
- Check that placeholder name exactly matches field/column name
- Placeholder format: `<<fieldname>>` with double angle brackets

**Word process hangs**
- End `WINWORD.EXE` in Task Manager
- The script normally cleans up automatically

## Performance

Typical processing time: 1-3 seconds per room depending on template complexity.

## License

MIT
