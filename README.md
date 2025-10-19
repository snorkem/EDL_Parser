# EDL Parser

A Python tool for parsing, analyzing, and converting Edit Decision List (EDL) files to Excel spreadsheets with advanced formatting and categorization features.

## Features

- **EDL to Excel Conversion** - Convert CMX3600 EDL files to formatted Excel spreadsheets
- **Intelligent Categorization** - Automatically categorize clips based on customizable patterns
- **Marker Extraction** - Extract and display markers from EDL files
- **Multiple File Formats** - Support for EDL, SRT, and STL subtitle files
- **GUI Interface** - User-friendly PyQt5 graphical interface (optional)
- **Batch Processing** - Process multiple EDL files at once
- **Custom Formatting** - Configure output styling with YAML configuration files

### Requirements
- Python 3.7+

### Install Dependencies

```bash
pip install -r requirements.txt
```

Core dependencies:
- `pandas` - Data manipulation and Excel export
- `pycmx` - CMX3600 EDL parsing
- `openpyxl` - Excel formatting and styling
- `timecode` - Timecode calculations
- `PyYAML` - YAML configuration file parsing
- `pysrt` - SRT subtitle file parsing
- `PyQt5` - GUI interface (optional)

## Quick Start

### Command Line

```bash
# Basic conversion
python edl_parse.py input.edl output.xlsx

# With categorization config
python edl_parse.py input.edl output.xlsx --config format_config.yaml

# With statistics
python edl_parse.py input.edl output.xlsx --stats

# Advanced mode
python edl_advanced.py input.edl output.xlsx
```

### GUI Interface

```bash
python edl_gui.py
```

The GUI provides an intuitive interface for:
- File selection and batch processing
- Output format configuration
- Real-time processing feedback
- Advanced options and filters

## Project Structure

```
EDL_Parser/
├── edl_parse.py           # Core EDL parsing and conversion logic
├── edl_advanced.py        # Advanced parsing with additional features
├── edl_gui.py            # PyQt5 graphical interface
├── srt_parser.py         # SRT subtitle file parser
├── stl_parser.py         # STL subtitle file parser
├── format_config.yaml    # Default formatting configuration
├── requirements.txt      # Python dependencies
├── USER_GUIDE.md        # Comprehensive documentation
├── Sample Files/        # Example EDL files for testing
├── Test Outputs/        # Test output files
└── Logs/               # Processing logs

```

## Usage Examples

### Basic Conversion
```bash
python edl_parse.py timeline.edl timeline.xlsx
```

### With Custom Categories
Create a `format_config.yaml` file to define categorization rules:

```yaml
categories:
  - name: A-Cam
    description: Footage from A-Camera (main camera)
    priority: 1
    patterns:
      - type: glob
        field: Source File
        pattern: "A*.*"
      - type: glob
        field: Clip Name
        pattern: "A_*"
      - type: regex
        field: Source File
        pattern: "A\d{3}C\d{3}"
    formatting:
      cell_color: "E6F3FF"  # Light blue
      text_color: "000000"  # Black
      bold: false
      italic: false
      entire_row: true
```

Then run:
```bash
python edl_parse.py timeline.edl timeline.xlsx --config format_config.yaml
```

### Generate Statistics
```bash
python edl_parse.py timeline.edl timeline.xlsx --stats
```

This adds analysis sheets including:
- Edit counts and durations
- Category breakdowns
- Timeline statistics

## Documentation

For comprehensive documentation, see [USER_GUIDE.md](USER_GUIDE.md) which covers:
- Detailed command reference
- Categorization and formatting options
- Analysis features
- Data manipulation
- Advanced workflows
- Troubleshooting

## Contributing

This is a personal project. Feel free to fork and modify for your own use.

## License

This project is provided as-is for personal and professional use.
