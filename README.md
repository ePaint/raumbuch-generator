# Raumbuch Generator

Converts Excel room data to PDF specification sheets using a Word template.

## Requirements

- Windows
- Microsoft Word (installed)
- PowerShell 5.1 or later
- [ImportExcel](https://github.com/dfinke/ImportExcel) module

## Installation

1. Clone this repository:
   ```powershell
   git clone https://github.com/ePaint/raumbuch-generator.git
   cd raumbuch-generator
   ```

2. Install the ImportExcel module:
   ```powershell
   Install-Module ImportExcel -Scope CurrentUser
   ```

3. Place your input files in the `Input/` folder:
   - Your Excel data file (e.g., `ZB3.0.xlsx`)
   - Your Word template (e.g., `Raumbuch_Vorlage.docx`)

4. Update `config.psd1` to match your file names.

## Project Structure

```
raumbuch-generator/
├── RoomToPDF.ps1        # Main script
├── config.psd1          # Configuration
├── MappingTable.xlsx    # Field mappings (Excel column → Word label)
├── Input/               # Your data files
│   ├── *.xlsx           # Room data spreadsheet
│   └── *.docx           # Word template
└── Output/              # Generated PDFs (timestamped folders)
    └── 2026-03-20_19-08-29/
        ├── RT.001.pdf
        ├── RT.002.pdf
        └── ...
```

## Configuration

Edit `config.psd1`:

```powershell
@{
    DataFile       = 'Input/ZB3.0.xlsx'              # Excel file with room data
    TemplateFile   = 'Input/Raumbuch_Vorlage.docx'   # Word template
    MappingFile    = 'MappingTable.xlsx'             # Field mapping table
    OutputFolder   = 'Output'                        # Output directory
    RoomCodeColumn = 'Code'                          # Column name for room codes
}
```

| Setting | Description |
|---------|-------------|
| `DataFile` | Path to your Excel file containing room data (one row per room) |
| `TemplateFile` | Path to the Word template with tables to fill |
| `MappingFile` | Path to the mapping table that links Excel columns to Word labels |
| `OutputFolder` | Where PDFs will be saved (in timestamped subfolders) |
| `RoomCodeColumn` | Name of the Excel column containing room identifiers |

## Mapping Table

The `MappingTable.xlsx` file defines how Excel columns map to Word template labels:

| ExcelColumn | WordLabel | UnitSuffix |
|-------------|-----------|------------|
| `Heizung - min. Raumtemperatur Winter` | `min. Raumtemperatur Winter` | `°C` |
| `Beleuchtung - Grundbeleuchtung` | `Grundbeleuchtung nach SIA 387/4` | `lux` |
| `Architektur - max. Belegung` | `max. Belegung` | |

- **ExcelColumn**: Exact column header from your Excel file
- **WordLabel**: Label text in the Word template (matched case-insensitively)
- **UnitSuffix**: Optional unit to append (e.g., `°C`, `lux`, `%`)

## Usage

Process all rooms:
```powershell
.\RoomToPDF.ps1
```

Process a single room:
```powershell
.\RoomToPDF.ps1 -RoomCode "RT.017"
```

Use a different config file:
```powershell
.\RoomToPDF.ps1 -ConfigPath "path\to\other-config.psd1"
```

## Output

Each run creates a timestamped folder inside `Output/`:

```
Output/
└── 2026-03-20_19-08-29/
    ├── RT.001.pdf
    ├── RT.002.pdf
    ├── RT.003.pdf
    └── ...
```

Previous runs are preserved.

## Word Template Requirements

The Word template must use tables with this structure:

**Two-column tables** (most fields):
```
┌─────────────────────────┬─────────┐
│ max. Belegung           │ [value] │
│ Tageslicht direkt       │ [value] │
└─────────────────────────┴─────────┘
```
- Column 1: Label
- Column 2: Value (filled by script)

**Single-column tables** (remarks/notes):
```
┌─────────────────────────┐
│ Beleuchtung Bemerkungen │
├─────────────────────────┤
│ [value]                 │
└─────────────────────────┘
```
- Row 1: Label
- Row 2: Value (filled by script)

Labels are matched case-insensitively against the `WordLabel` column in the mapping table.

## Troubleshooting

**"ImportExcel module is required"**
```powershell
Install-Module ImportExcel -Scope CurrentUser
```

**"Data file not found" / "Template file not found"**
- Check that the paths in `config.psd1` are correct
- Paths are relative to the script directory

**"Warning: X labels not found in template"**
- Some `WordLabel` values in `MappingTable.xlsx` don't match any label in the Word template
- Check for typos or extra spaces

**Word process hangs**
- The script cleans up Word automatically
- If Word hangs, end `WINWORD.EXE` in Task Manager

## Performance

The script uses position caching for fast processing:
- ~1-2 seconds per room
- 181 rooms in ~4 minutes

## License

MIT
