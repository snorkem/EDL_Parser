# EDL Parser - Complete User Guide

## Table of Contents
1. [Quick Start](#quick-start)
2. [Core Features](#core-features)
3. [Categorization & Formatting](#categorization--formatting)
4. [Analysis Features](#analysis-features)
5. [Data Manipulation](#data-manipulation)
6. [Advanced Operations](#advanced-operations)
7. [Command Reference](#command-reference)
8. [Examples & Workflows](#examples--workflows)

---

## Quick Start

### Installation
```bash
pip install -r requirements.txt
```

**Dependencies:**
- `pandas` - DataFrame manipulation and Excel export
- `pycmx` - CMX3600 EDL parsing
- `openpyxl` - Excel formatting and styling
- `python-timecode` - Timecode calculations
- `PyYAML` - YAML configuration file parsing
- `pysrt` - SRT subtitle file parsing

### Basic Usage
```bash
# Convert EDL to Excel
python edl_parse.py input.edl output.xlsx

# With categorization
python edl_parse.py input.edl output.xlsx --config format_config.yaml

# With statistics
python edl_parse.py input.edl output.xlsx --stats
```

---

## Core Features

### 1. EDL to Excel Conversion
Convert CMX3600 EDL files to Excel spreadsheets with automatic formatting.

**Command:**
```bash
python edl_parse.py input.edl output.xlsx
```

**Features:**
- Parses all event metadata (timecodes, clip names, source files, etc.)
- Extracts motion adapter FPS data
- Auto-sized columns (enabled by default)
- Alternating row colors (enabled by default)

**Disable formatting:**
```bash
python edl_parse.py input.edl output.xlsx --no-table --no-colored
```

---

## Categorization & Formatting

### Pattern-Based Categorization
Automatically categorize events using YAML configuration files.

**Config File:** `format_config.yaml`
```yaml
categories:
  - name: A-Camera
    priority: 1
    patterns:
      - type: glob
        field: Source File
        pattern: "A*.*"
      - type: regex
        field: Clip Name
        pattern: "^A_.*"
    formatting:
      cell_color: "E6F3FF"  # Light blue
      text_color: "000000"  # Black
      bold: false
      entire_row: true
```

**Command:**
```bash
python edl_parse.py input.edl output.xlsx --config format_config.yaml
```

**Features:**
- Glob and regex pattern matching
- Priority-based category assignment
- Custom cell/row colors
- Font styling (bold, italic)
- Match on any field (Source File, Clip Name, Video, etc.)

See `CATEGORIZATION_GUIDE.md` for detailed documentation.

---

## Analysis Features

### 1. Statistics Report
Generate comprehensive statistics about your EDL.

**Command:**
```bash
# Add statistics sheet to Excel
python edl_parse.py input.edl output.xlsx --stats

# Export only statistics
python edl_parse.py input.edl stats.xlsx --stats-only
```

**Statistics Included:**
- Total events count
- Unique source files/clips
- Timeline duration
- Average/min/max shot length
- Video vs. audio-only events
- Category distribution
- FPS distribution
- Transition types
- Reel/tape count

---

### 2. Duplicate Detection
Find and handle duplicate events automatically.

**Commands:**
```bash
# Highlight duplicates in red
python edl_parse.py input.edl output.xlsx --highlight-duplicates

# Remove duplicates entirely
python edl_parse.py input.edl output.xlsx --remove-duplicates
```

**Detection Criteria:**
Duplicates are events with identical:
- Record In timecode
- Record Out timecode
- Clip Name

---

### 3. Validation
Validate EDL timecode order and detect issues.

**Command:**
```bash
python edl_parse.py input.edl output.xlsx --validate
```

**Checks:**
- Timecode sequence validation
- Overlapping events detection
- Gap detection

---

## Data Manipulation

### 1. Sorting
Sort events by various criteria.

**Commands:**
```bash
# Sort by clip name
python edl_parse.py input.edl output.xlsx --sort-by clip_name

# Sort by duration (longest first)
python edl_parse.py input.edl output.xlsx --sort-by duration --fps 30

# Sort by category
python edl_parse.py input.edl output.xlsx --sort-by category --config format_config.yaml

# Sort by source file
python edl_parse.py input.edl output.xlsx --sort-by source_file
```

**Sort Options:**
- `timecode` - Keep original order (default)
- `clip_name` - Alphabetical by clip name
- `source_file` - Alphabetical by source file
- `duration` - Longest to shortest (requires --fps)
- `category` - By category (requires --config)

---

### 2. Filtering
Filter events using pandas query expressions.

**Commands:**
```bash
# Filter by category
python edl_parse.py input.edl output.xlsx --filter "Category == 'A-Camera'" --config format_config.yaml

# Complex filters
python edl_parse.py input.edl output.xlsx --filter "Video == 'V' and Category != 'Audio Only'"

# Numeric comparisons (requires duration calculation)
python edl_parse.py input.edl output.xlsx --filter "`Event #` > 100"
```

**Filter Syntax:**
Uses pandas `.query()` syntax:
- String comparison: `Category == 'A-Camera'`
- Numeric: `\`Event #\` > 10`
- Boolean: `Video == 'V'`
- Combine with `and`, `or`, `not`

---

### 3. Search
Search for events matching patterns.

**Commands:**
```bash
# Search all text fields (glob pattern)
python edl_parse.py input.edl output.xlsx --search "A_*_01"

# Search specific field
python edl_parse.py input.edl output.xlsx --search "INTERVIEW" --search-field "Clip Name"

# Use regex
python edl_parse.py input.edl output.xlsx --search "^A[0-9]{3}" --search-regex
```

**Search Fields:**
- Clip Name
- Source File
- Reel
- (All text fields if no field specified)

---

## Advanced Operations

### 1. Multi-EDL Merge
Combine multiple EDL files into one, with optional subtitle file integration.

**Commands:**
```bash
# Merge multiple EDL files (sorted by timecode)
python edl_parse.py --merge edl1.edl edl2.edl edl3.edl --output output.xlsx

# Merge single EDL with subtitle file
python edl_parse.py --merge edit.edl --output output.xlsx --subtitle-file subtitles.srt --subtitle-fps 30
```

**Features:**
- Events are merged and sorted by Record In timecode (interleave mode)
- Events are automatically renumbered sequentially
- Optional subtitle matching (STL or SRT format)

---

### 2. Subtitle File Support (STL & SRT)
Add subtitle text to EDL events based on timecode overlap.

**Supported Formats:**
- **STL** - EBU N19 STL (EBU Tech 3264) binary subtitle files
- **SRT** - SubRip text-based subtitle files

**Commands:**
```bash
# Merge EDL with STL subtitle file (auto-detect FPS)
python edl_parse.py --merge edit.edl --output output.xlsx \
  --subtitle-file subtitles.stl

# Merge EDL with SRT subtitle file (specify FPS)
python edl_parse.py --merge edit.edl --output output.xlsx \
  --subtitle-file subtitles.srt --subtitle-fps 30

# Override STL auto-detection with manual FPS
python edl_parse.py --merge edit.edl --output output.xlsx \
  --subtitle-file subtitles.stl --subtitle-fps 24

# Align subtitles to specific timeline position
python edl_parse.py --merge edit.edl --output output.xlsx \
  --subtitle-file subtitles.srt --subtitle-fps 30 \
  --subtitle-start-time "00:00:05:00"
```

**How It Works:**
- Subtitles are matched against **Record In/Out** (timeline positions) rather than source timecodes
- Events can have multiple overlapping subtitles (joined with " | ")
- Empty subtitle column for events with no overlap

**FPS Handling:**
- **STL**: Auto-detects FPS from file metadata (can be overridden with `--subtitle-fps`)
- **SRT**: Uses `--subtitle-fps` value (defaults to 30 fps if not specified)

**Subtitle Start Time:**
The `--subtitle-start-time` parameter aligns the first subtitle to a specific timeline position:
- Specify Record In timecode where first subtitle should appear (HH:MM:SS:FF format)
- Useful when subtitle file timecodes don't match EDL timeline
- Example: `--subtitle-start-time "00:02:00:00"` moves all subtitles to start at 2 minutes

**Use Cases:**
- Add burned-in subtitle text to shot lists
- Generate dialogue continuity reports
- Create subtitle timing sheets for localization
- Track subtitle coverage across edit

**Output:**
Creates Excel with "Subtitles" column containing matched subtitle text.

**Backwards Compatibility:**
Old argument names still work:
```bash
--stl-file subtitles.stl        # Same as --subtitle-file
--stl-fps 24                    # Same as --subtitle-fps
--stl-start-time "01:00:00:00"  # Same as --subtitle-start-time
```

---

### 3. Split by Category
Split events into separate files by category.

**Commands:**
```bash
# Split to Excel files by category
python edl_parse.py input.edl output.xlsx --config format_config.yaml --split-by-category

# Custom output directory
python edl_parse.py input.edl output.xlsx --config format_config.yaml --split-by-category --split-output-dir ./categorized_edls
```

**Output:**
Creates one file per category in the output directory.

---

### 4. EDL Comparison (Changelog)
Compare two EDL versions and generate a changelog report.

**Command:**
```bash
python edl_parse.py --compare original.edl revised.edl --changelog-output changelog.xlsx
```

**Changelog Includes:**
- **Summary sheet** - Overview of changes
- **Added Events** - Green highlighting
- **Removed Events** - Red highlighting
- **Modified Events** - Yellow highlighting (same timecode, different source)

**Use Cases:**
- Track editorial changes between versions
- Identify VFX shot updates
- Generate conform notes

---

### 5. Event Grouping
Group events by time intervals.

**Command:**
```bash
# Group events within 15-second intervals
python edl_parse.py input.edl output.xlsx --group 15 --fps 30
```

**Output:**
Creates "Grouped Events" sheet with:
- Event count per group
- Combined timecode ranges
- Unique clip counts
- Duration calculations

---

## Command Reference

### Input/Output
```
input_edl               Input EDL file path
output_xlsx             Output Excel file path
```

### Formatting (Enabled by Default)
```
--no-table              Disable table formatting
--no-colored            Disable alternating row colors
--config FILE           JSON config for categorization
```

### Analysis
```
--stats                 Add statistics sheet
--stats-only            Export only statistics
--highlight-duplicates  Highlight duplicates in red
--remove-duplicates     Remove duplicate events
--validate              Validate timecode order
```

### Data Manipulation
```
--sort-by CRITERION     Sort events (timecode|clip_name|source_file|duration|category)
--filter EXPRESSION     Filter using pandas query
--search TERM           Search for matching events
--search-field FIELD    Specific field to search
--search-regex          Use regex for search
```

### Advanced Operations
```
--merge FILE [FILE ...]            Merge multiple EDLs
--output FILE                      Output file path (use with --merge)
--subtitle-file FILE               Subtitle file (.stl or .srt) to merge with EDL
--subtitle-fps FPS                 Frame rate for subtitle parsing (STL: overrides auto-detect, SRT: default 30)
--subtitle-start-time TIMECODE     Record In timecode (HH:MM:SS:FF) where first subtitle should align
--split-by-category                Split by category into Excel files
--split-output-dir DIR             Output directory for split files
--compare EDL1 EDL2                Compare two EDLs
--changelog-output FILE            Changelog report path
--group SECONDS                    Group events by time
--fps FPS                          Frame rate for calculations
```

### Legacy Subtitle Arguments (Backwards Compatible)
```
--stl-file FILE                    Alias for --subtitle-file
--stl-fps FPS                      Alias for --subtitle-fps
--stl-start-time TIMECODE          Alias for --subtitle-start-time
```

---

## Examples & Workflows

### Workflow 1: Categorize and Split by Camera
```bash
# 1. Convert with categorization
python edl_parse.py shoot_edit.edl analyzed.xlsx --config camera_config.json --stats

# 2. Split into camera-specific Excel files
python edl_parse.py shoot_edit.edl analyzed.xlsx --config camera_config.json --split-by-category
```

**Result:**
- `analyzed.xlsx` - Main file with categories and statistics
- `split_output/A-Camera.xlsx` - A-camera shots only
- `split_output/B-Camera.xlsx` - B-camera shots only

---

### Workflow 2: Clean and Filter VFX Shots
```bash
# Remove duplicates and filter VFX category
python edl_parse.py master.edl vfx_shots.xlsx --config vfx_config.json --remove-duplicates --filter "Category == 'VFX'"
```

---

### Workflow 3: Track Editorial Changes
```bash
# Compare two versions
python edl_parse.py --compare v1_edit.edl v2_edit.edl --changelog-output editorial_changes.xlsx
```

**Result:**
Excel file with color-coded changes (added/removed/modified)

---

### Workflow 4: Merge Multiple Reels
```bash
# Merge all reels sequentially
python edl_parse.py --merge reel1.edl reel2.edl reel3.edl --merge-mode append master_edit.xlsx --stats
```

---

### Workflow 5: Find and Export Long Takes
```bash
# Sort by duration and export longest shots
python edl_parse.py interview.edl long_takes.xlsx --sort-by duration --fps 30 --stats
```

---

### Workflow 6: Search and Extract
```bash
# Find all interview shots
python edl_parse.py master.edl interviews.xlsx --search "*INTERVIEW*" --search-field "Clip Name"

# Find clips from specific camera
python edl_parse.py master.edl a_cam_shots.xlsx --search "A001*" --search-field "Source File"
```

---

### Workflow 7: Add Subtitles to EDL
```bash
# Merge EDL with SRT subtitle file
python edl_parse.py --merge final_edit.edl --output edit_with_subs.xlsx \
  --subtitle-file dialogue_subtitles.srt --subtitle-fps 24

# Generate dialogue continuity report with categories
python edl_parse.py --merge final_edit.edl --output continuity_report.xlsx \
  --subtitle-file dialogue.srt --subtitle-fps 24 \
  --config scene_config.yaml --stats

# Multiple EDLs with aligned STL subtitles
python edl_parse.py --merge reel1.edl reel2.edl --output merged_with_subs.xlsx \
  --subtitle-file master_subtitles.stl \
  --subtitle-start-time "01:00:00:00"
```

**Use Cases:**
- Create dialogue sheets for script supervisors
- Generate subtitle timing reports for localization teams
- Track dialogue coverage across multiple takes
- Export subtitle text for closed captioning QC

---

### Workflow 8: Subtitle Timing Alignment
```bash
# Align subtitle file to specific timeline position
python edl_parse.py --merge edit.edl --output aligned_subs.xlsx \
  --subtitle-file subtitles.srt --subtitle-fps 30 \
  --subtitle-start-time "00:02:00:00"

# This moves all subtitles so the first subtitle appears at 00:02:00:00
# Useful when subtitle timecodes don't match your EDL timeline
```

**Result:**
- Subtitles are offset to align with your timeline
- Events show which subtitles overlap their timecode ranges
- Multiple overlapping subtitles are combined with " | " separator

---

### Workflow 9: Group Events with Subtitles
```bash
# Create shot groupings with subtitle text
python edl_parse.py --merge edit.edl --output grouped_with_subs.xlsx \
  --subtitle-file subtitles.srt --subtitle-fps 24 \
  --group 30 --fps 24

# This creates:
# - Events sheet with individual events and subtitle matches
# - Grouped Events sheet with 30-second intervals including subtitle text
```

**Use Cases:**
- Create scene-based dialogue reports
- Generate subtitle coverage by time segment
- Track subtitle density across edit

---

## Tips & Best Practices

### 1. Use Configuration Files
Create category configs for different project types:
- `camera_config.json` - Camera-based categories
- `content_config.json` - Content type categories
- `vfx_config.json` - VFX-specific categories

### 2. Combine Features
```bash
# Full analysis pipeline
python edl_parse.py input.edl output.xlsx \
  --config format_config.yaml \
  --remove-duplicates \
  --validate \
  --stats \
  --sort-by category
```

### 3. Validate Timecodes
Always use `--validate` to check timecode order:
```bash
python edl_parse.py input.edl output.xlsx --validate
```

### 4. Use Stats for QC
Generate statistics reports for all deliverables:
```bash
python edl_parse.py delivery.edl delivery_stats.xlsx --stats-only
```

### 5. Backup Original Files
When using `--remove-duplicates` or other modifications, always keep original:
```bash
cp original.edl original_backup.edl
python edl_parse.py original.edl cleaned.xlsx --remove-duplicates
```

### 6. Working with Subtitles
**Format Selection:**
- Use **STL** for broadcast/professional subtitle files (EBU N19 standard)
- Use **SRT** for web/consumer subtitle files (SubRip format)
- STL files auto-detect FPS from metadata (more reliable)
- SRT files require manual FPS specification

**FPS Settings:**
- Always verify FPS matches your timeline (23.976, 24, 25, 29.97, or 30)
- For STL: Let auto-detection work unless you know it's incorrect
- For SRT: Explicitly set `--subtitle-fps` to match your timeline

**Alignment Tips:**
- Use `--subtitle-start-time` when subtitle timecodes don't match EDL timeline
- Check first few events to verify subtitle alignment is correct
- Multiple overlapping subtitles are normal for long events (music beds, ambient audio)

**Typical Workflow:**
```bash
# 1. Test with small EDL first
python edl_parse.py --merge test.edl --output test_subs.xlsx \
  --subtitle-file subtitles.srt --subtitle-fps 24

# 2. Verify alignment in Excel

# 3. Process full EDL with same settings
python edl_parse.py --merge full_edit.edl --output final_with_subs.xlsx \
  --subtitle-file subtitles.srt --subtitle-fps 24 \
  --config scene_config.yaml --stats
```

---

## Troubleshooting

### Issue: "No module named 'edl_advanced'"
**Solution:** Ensure `edl_advanced.py` is in the same directory as `edl_parse.py`.

### Issue: Filter expression errors
**Solution:** Use backticks for column names with spaces:
```bash
--filter "\`Event #\` > 10"
```

### Issue: Timecode validation failures
**Solution:** Check for:
- Mixed frame rates
- Incorrect timecode format
- Overlapping events

### Issue: Category not appearing
**Solution:**
- Verify config file path
- Check pattern syntax
- Test patterns individually
- Use `--config format_config.yaml` flag

### Issue: Subtitles not matching / No subtitles found
**Possible causes and solutions:**

1. **FPS Mismatch**
   - STL: Check auto-detected FPS in log output
   - SRT: Verify `--subtitle-fps` matches your timeline
   - Solution: Override with correct FPS: `--subtitle-fps 24`

2. **Timeline Offset**
   - Subtitle timecodes may not align with EDL timeline
   - Solution: Use `--subtitle-start-time` to align first subtitle
   - Example: `--subtitle-start-time "01:00:00:00"`

3. **Wrong Timecode Type**
   - Subtitles match against Record In/Out (timeline), not source timecodes
   - Verify your EDL has Record In/Out values
   - Check log output: "Matching subtitles against timeline positions"

4. **File Format Issues**
   - STL: Ensure file is EBU N19 STL format (binary file)
   - SRT: Ensure file has .srt extension and proper encoding (UTF-8)
   - Check log for parsing errors

### Issue: "No module named 'pysrt'"
**Solution:**
```bash
pip install pysrt
```
Required for SRT subtitle parsing.

### Issue: "Timecode.frames should be a positive integer"
**Solution:**
- This occurs when subtitle timecode starts at 00:00:00:00 with frame 0
- Already handled automatically - update to latest code
- If error persists, verify subtitle file format

---

## Getting Help

```bash
# View all options
python edl_parse.py --help

# Check version
python edl_parse.py --version  # (if implemented)
```

---

## Additional Resources

- `CATEGORIZATION_GUIDE.md` - Detailed categorization documentation
- `FEATURE_ROADMAP.md` - Complete feature roadmap
- `FEATURE_PLAN_SUMMARY.md` - Quick reference for features
- `format_config.yaml` - Example configuration file
- `test_categorization.py` - Test suite

---

## Keyboard Shortcuts & Aliases

Create bash aliases for common operations:

```bash
# Add to ~/.bashrc or ~/.zshrc
alias edl2xlsx='python /path/to/edl_parse.py'
alias edl-stats='python /path/to/edl_parse.py --stats'
alias edl-compare='python /path/to/edl_parse.py --compare'
```

Usage:
```bash
edl2xlsx input.edl output.xlsx --config myconfig.json
edl-stats input.edl
edl-compare v1.edl v2.edl
```
