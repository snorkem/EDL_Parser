#!/usr/bin/env python3
"""
EDL Parser GUI - PyQt5 Frontend
Graphical interface for edl_parse.py
"""

import sys
import subprocess
import os
from pathlib import Path
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QLabel, QPushButton, QLineEdit, QTextEdit, QCheckBox,
    QComboBox, QFileDialog, QGroupBox, QGridLayout, QMessageBox,
    QProgressBar, QSpinBox, QDoubleSpinBox, QScrollArea
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QProcess
from PyQt5.QtGui import QFont, QTextCursor


class EDLProcessThread(QThread):
    """Thread for running EDL processing without blocking GUI"""
    output_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(int)

    def __init__(self, command):
        super().__init__()
        self.command = command

    def run(self):
        """Execute the command and emit output"""
        try:
            process = subprocess.Popen(
                self.command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                bufsize=1,
                universal_newlines=True
            )

            # Read output in real-time
            for line in process.stdout:
                self.output_signal.emit(line.strip())

            # Wait for completion
            process.wait()

            # Get any remaining stderr
            stderr = process.stderr.read()
            if stderr:
                self.error_signal.emit(stderr)

            self.finished_signal.emit(process.returncode)

        except Exception as e:
            self.error_signal.emit(f"Error: {str(e)}")
            self.finished_signal.emit(1)


class ConvertTab(QWidget):
    """Tab for standard EDL conversion"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Input/Output files
        file_group = QGroupBox("Files")
        file_layout = QGridLayout()

        # Input EDL
        file_layout.addWidget(QLabel("Input EDL:"), 0, 0)
        self.input_edl = QLineEdit()
        file_layout.addWidget(self.input_edl, 0, 1)
        input_btn = QPushButton("Browse...")
        input_btn.clicked.connect(self.browse_input)
        file_layout.addWidget(input_btn, 0, 2)

        # Output file
        file_layout.addWidget(QLabel("Output File:"), 1, 0)
        self.output_file = QLineEdit()
        file_layout.addWidget(self.output_file, 1, 1)
        output_btn = QPushButton("Browse...")
        output_btn.clicked.connect(self.browse_output)
        file_layout.addWidget(output_btn, 1, 2)

        # Config file
        file_layout.addWidget(QLabel("Config File:"), 2, 0)
        self.config_file = QLineEdit()
        file_layout.addWidget(self.config_file, 2, 1)
        config_btn = QPushButton("Browse...")
        config_btn.clicked.connect(self.browse_config)
        file_layout.addWidget(config_btn, 2, 2)

        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # Options
        options_group = QGroupBox("Options")
        options_layout = QVBoxLayout()

        # Format options
        format_layout = QHBoxLayout()
        self.no_table = QCheckBox("Disable table formatting")
        self.no_colored = QCheckBox("Disable row colors")
        format_layout.addWidget(self.no_table)
        format_layout.addWidget(self.no_colored)
        options_layout.addLayout(format_layout)

        # Analysis options
        analysis_layout = QHBoxLayout()
        self.stats = QCheckBox("Generate statistics")
        self.stats_only = QCheckBox("Statistics only")
        self.validate = QCheckBox("Validate timecodes")
        analysis_layout.addWidget(self.stats)
        analysis_layout.addWidget(self.stats_only)
        analysis_layout.addWidget(self.validate)
        options_layout.addLayout(analysis_layout)

        # Duplicate handling
        dup_layout = QHBoxLayout()
        self.remove_dups = QCheckBox("Remove duplicates")
        self.highlight_dups = QCheckBox("Highlight duplicates")
        dup_layout.addWidget(self.remove_dups)
        dup_layout.addWidget(self.highlight_dups)
        options_layout.addLayout(dup_layout)

        options_group.setLayout(options_layout)
        layout.addWidget(options_group)

        # Sorting and Filtering
        sort_filter_group = QGroupBox("Sorting & Filtering")
        sort_filter_layout = QGridLayout()

        # Sort by
        sort_filter_layout.addWidget(QLabel("Sort by:"), 0, 0)
        self.sort_by = QComboBox()
        self.sort_by.addItems(['None', 'timecode', 'clip_name', 'source_file', 'duration', 'category'])
        sort_filter_layout.addWidget(self.sort_by, 0, 1)

        # FPS (for duration sort)
        sort_filter_layout.addWidget(QLabel("FPS:"), 0, 2)
        self.fps = QDoubleSpinBox()
        self.fps.setRange(1, 240)
        self.fps.setValue(30)
        self.fps.setDecimals(3)
        sort_filter_layout.addWidget(self.fps, 0, 3)

        # Filter
        sort_filter_layout.addWidget(QLabel("Filter:"), 1, 0)
        self.filter_expr = QLineEdit()
        self.filter_expr.setPlaceholderText('e.g., Category == "A-Camera"')
        sort_filter_layout.addWidget(self.filter_expr, 1, 1, 1, 3)

        # Search
        sort_filter_layout.addWidget(QLabel("Search:"), 2, 0)
        self.search_term = QLineEdit()
        self.search_term.setPlaceholderText('e.g., INTERVIEW*')
        sort_filter_layout.addWidget(self.search_term, 2, 1)

        self.search_field = QLineEdit()
        self.search_field.setPlaceholderText('Field (optional)')
        sort_filter_layout.addWidget(self.search_field, 2, 2)

        self.search_regex = QCheckBox("Use regex")
        sort_filter_layout.addWidget(self.search_regex, 2, 3)

        sort_filter_group.setLayout(sort_filter_layout)
        layout.addWidget(sort_filter_group)

        # Split options
        split_group = QGroupBox("Split by Category")
        split_layout = QGridLayout()

        self.split_by_category = QCheckBox("Split by category into Excel files")
        split_layout.addWidget(self.split_by_category, 0, 0, 1, 2)

        split_layout.addWidget(QLabel("Output Dir:"), 1, 0)
        self.split_output_dir = QLineEdit("./split_output")
        split_layout.addWidget(self.split_output_dir, 1, 1)

        split_group.setLayout(split_layout)
        layout.addWidget(split_group)

        # Grouping
        group_box = QGroupBox("Event Grouping")
        group_layout = QHBoxLayout()

        group_layout.addWidget(QLabel("Group events (seconds):"))
        self.group_interval = QDoubleSpinBox()
        self.group_interval.setRange(0, 1000)
        self.group_interval.setValue(0)
        self.group_interval.setSpecialValueText("Disabled")
        group_layout.addWidget(self.group_interval)

        group_box.setLayout(group_layout)
        layout.addWidget(group_box)

        layout.addStretch()
        self.setLayout(layout)

    def browse_input(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, "Select Input EDL", "", "EDL Files (*.edl);;All Files (*)"
        )
        if filename:
            self.input_edl.setText(filename)
            # Auto-suggest output name
            if not self.output_file.text():
                base = Path(filename).stem
                self.output_file.setText(f"{base}_output.xlsx")

    def browse_output(self):
        filename, _ = QFileDialog.getSaveFileName(
            self, "Save Output File", "", "Excel Files (*.xlsx);;All Files (*)"
        )
        if filename:
            self.output_file.setText(filename)

    def browse_config(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, "Select Config File", "", "YAML Files (*.yaml);;All Files (*)"
        )
        if filename:
            self.config_file.setText(filename)

    def get_command(self):
        """Build command from UI options"""
        if not self.input_edl.text():
            raise ValueError("Input EDL file is required")

        cmd = ['python', 'edl_parse.py', self.input_edl.text()]

        # Output file
        if self.output_file.text():
            cmd.append(self.output_file.text())

        # Config
        if self.config_file.text():
            cmd.extend(['--config', self.config_file.text()])

        # Format options
        if self.no_table.isChecked():
            cmd.append('--no-table')
        if self.no_colored.isChecked():
            cmd.append('--no-colored')

        # Analysis
        if self.stats.isChecked():
            cmd.append('--stats')
        if self.stats_only.isChecked():
            cmd.append('--stats-only')
        if self.validate.isChecked():
            cmd.append('--validate')

        # Duplicates
        if self.remove_dups.isChecked():
            cmd.append('--remove-duplicates')
        if self.highlight_dups.isChecked():
            cmd.append('--highlight-duplicates')

        # Sorting
        if self.sort_by.currentText() != 'None':
            cmd.extend(['--sort-by', self.sort_by.currentText()])
            if self.sort_by.currentText() == 'duration':
                cmd.extend(['--fps', str(self.fps.value())])

        # FPS for grouping
        if self.group_interval.value() > 0:
            cmd.extend(['--group', str(self.group_interval.value())])
            cmd.extend(['--fps', str(self.fps.value())])

        # Filter
        if self.filter_expr.text():
            cmd.extend(['--filter', self.filter_expr.text()])

        # Search
        if self.search_term.text():
            cmd.extend(['--search', self.search_term.text()])
            if self.search_field.text():
                cmd.extend(['--search-field', self.search_field.text()])
            if self.search_regex.isChecked():
                cmd.append('--search-regex')

        # Split
        if self.split_by_category.isChecked():
            cmd.append('--split-by-category')
            if self.split_output_dir.text():
                cmd.extend(['--split-output-dir', self.split_output_dir.text()])

        return cmd


class CompareTab(QWidget):
    """Tab for comparing two EDLs"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Files
        file_group = QGroupBox("EDL Files to Compare")
        file_layout = QGridLayout()

        # Original EDL
        file_layout.addWidget(QLabel("Original EDL:"), 0, 0)
        self.original_edl = QLineEdit()
        file_layout.addWidget(self.original_edl, 0, 1)
        orig_btn = QPushButton("Browse...")
        orig_btn.clicked.connect(self.browse_original)
        file_layout.addWidget(orig_btn, 0, 2)

        # Revised EDL
        file_layout.addWidget(QLabel("Revised EDL:"), 1, 0)
        self.revised_edl = QLineEdit()
        file_layout.addWidget(self.revised_edl, 1, 1)
        rev_btn = QPushButton("Browse...")
        rev_btn.clicked.connect(self.browse_revised)
        file_layout.addWidget(rev_btn, 1, 2)

        # Changelog output
        file_layout.addWidget(QLabel("Changelog Output:"), 2, 0)
        self.changelog_output = QLineEdit("changelog.xlsx")
        file_layout.addWidget(self.changelog_output, 2, 1)
        out_btn = QPushButton("Browse...")
        out_btn.clicked.connect(self.browse_output)
        file_layout.addWidget(out_btn, 2, 2)

        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # Info
        info_label = QLabel(
            "This will generate a changelog report showing:\n"
            "• Added events (green)\n"
            "• Removed events (red)\n"
            "• Modified events (yellow)"
        )
        info_label.setStyleSheet("padding: 10px; background-color: #f0f0f0; border-radius: 5px;")
        layout.addWidget(info_label)

        layout.addStretch()
        self.setLayout(layout)

    def browse_original(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, "Select Original EDL", "", "EDL Files (*.edl);;All Files (*)"
        )
        if filename:
            self.original_edl.setText(filename)

    def browse_revised(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, "Select Revised EDL", "", "EDL Files (*.edl);;All Files (*)"
        )
        if filename:
            self.revised_edl.setText(filename)

    def browse_output(self):
        filename, _ = QFileDialog.getSaveFileName(
            self, "Save Changelog", "", "Excel Files (*.xlsx);;All Files (*)"
        )
        if filename:
            self.changelog_output.setText(filename)

    def get_command(self):
        """Build command from UI options"""
        if not self.original_edl.text() or not self.revised_edl.text():
            raise ValueError("Both original and revised EDL files are required")

        cmd = [
            'python', 'edl_parse.py',
            '--compare', self.original_edl.text(), self.revised_edl.text()
        ]

        if self.changelog_output.text():
            cmd.extend(['--changelog-output', self.changelog_output.text()])

        return cmd


class MergeTab(QWidget):
    """Tab for merging multiple EDLs"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Files to merge
        file_group = QGroupBox("EDL Files to Merge")
        file_layout = QVBoxLayout()

        self.file_list = QTextEdit()
        self.file_list.setPlaceholderText("Add EDL files to merge (one per line)")
        self.file_list.setMaximumHeight(150)
        file_layout.addWidget(self.file_list)

        btn_layout = QHBoxLayout()
        add_btn = QPushButton("Add Files...")
        add_btn.clicked.connect(self.add_files)
        btn_layout.addWidget(add_btn)

        clear_btn = QPushButton("Clear All")
        clear_btn.clicked.connect(lambda: self.file_list.clear())
        btn_layout.addWidget(clear_btn)

        file_layout.addLayout(btn_layout)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # Options
        options_group = QGroupBox("Merge Options")
        options_layout = QGridLayout()

        options_layout.addWidget(QLabel("Output File:"), 0, 0)
        self.output_file = QLineEdit("merged_output.xlsx")
        options_layout.addWidget(self.output_file, 0, 1)
        out_btn = QPushButton("Browse...")
        out_btn.clicked.connect(self.browse_output)
        options_layout.addWidget(out_btn, 0, 2)

        self.add_stats = QCheckBox("Generate statistics")
        options_layout.addWidget(self.add_stats, 1, 0, 1, 3)

        options_group.setLayout(options_layout)
        layout.addWidget(options_group)

        # Subtitle Options (STL or SRT)
        subtitle_group = QGroupBox("Subtitle File (Optional - STL or SRT)")
        subtitle_layout = QGridLayout()

        subtitle_layout.addWidget(QLabel("Subtitle File:"), 0, 0)
        self.subtitle_file = QLineEdit()
        self.subtitle_file.setPlaceholderText("Select subtitle file (.stl or .srt)")
        subtitle_layout.addWidget(self.subtitle_file, 0, 1)
        subtitle_btn = QPushButton("Browse...")
        subtitle_btn.clicked.connect(self.browse_subtitle)
        subtitle_layout.addWidget(subtitle_btn, 0, 2)

        # FPS controls
        subtitle_layout.addWidget(QLabel("Frame Rate:"), 1, 0)
        fps_layout = QHBoxLayout()

        self.stl_auto_fps = QCheckBox("Auto-detect (STL only)")
        self.stl_auto_fps.setChecked(True)
        self.stl_auto_fps.toggled.connect(self.toggle_fps_controls)
        fps_layout.addWidget(self.stl_auto_fps)

        fps_layout.addWidget(QLabel("FPS:"))
        self.stl_fps = QDoubleSpinBox()
        self.stl_fps.setDecimals(3)
        self.stl_fps.setRange(1.0, 120.0)
        self.stl_fps.setValue(23.976)
        self.stl_fps.setEnabled(False)  # Disabled by default (auto-detect on)
        fps_layout.addWidget(self.stl_fps)

        self.stl_fps_preset = QComboBox()
        self.stl_fps_preset.addItems([
            '23.976 (Film/NTSC)',
            '24 (Film)',
            '25 (PAL)',
            '29.97 (NTSC)',
            '30 (NTSC)'
        ])
        self.stl_fps_preset.setEnabled(False)  # Disabled by default (auto-detect on)
        self.stl_fps_preset.currentTextChanged.connect(self.update_fps_from_preset)
        fps_layout.addWidget(self.stl_fps_preset)

        fps_layout.addStretch()
        subtitle_layout.addLayout(fps_layout, 1, 1, 1, 2)

        # Subtitle Start Time (Record In timecode where first subtitle should align)
        subtitle_layout.addWidget(QLabel("Subtitle Start Time:"), 2, 0)
        self.stl_start_time = QLineEdit()
        self.stl_start_time.setPlaceholderText("HH:MM:SS:FF (e.g., 00:02:00:00)")
        self.stl_start_time.setToolTip("Record In timecode where the first subtitle should align in the timeline")
        subtitle_layout.addWidget(self.stl_start_time, 2, 1, 1, 2)

        subtitle_group.setLayout(subtitle_layout)
        layout.addWidget(subtitle_group)

        # Formatting & Categorization
        format_group = QGroupBox("Formatting & Categorization")
        format_layout = QGridLayout()

        # Config file
        format_layout.addWidget(QLabel("Config File:"), 0, 0)
        self.config_file = QLineEdit()
        self.config_file.setPlaceholderText("YAML config for categorization/formatting")
        format_layout.addWidget(self.config_file, 0, 1)
        config_btn = QPushButton("Browse...")
        config_btn.clicked.connect(self.browse_config)
        format_layout.addWidget(config_btn, 0, 2)

        # Format options
        format_options_layout = QHBoxLayout()
        self.no_table = QCheckBox("Disable table formatting")
        self.no_colored = QCheckBox("Disable row colors")
        format_options_layout.addWidget(self.no_table)
        format_options_layout.addWidget(self.no_colored)
        format_layout.addLayout(format_options_layout, 1, 0, 1, 3)

        format_group.setLayout(format_layout)
        layout.addWidget(format_group)

        # Advanced Options
        advanced_group = QGroupBox("Advanced Options")
        advanced_layout = QVBoxLayout()

        # Analysis options
        analysis_layout = QHBoxLayout()
        self.validate = QCheckBox("Validate timecodes")
        self.remove_dups = QCheckBox("Remove duplicates")
        self.highlight_dups = QCheckBox("Highlight duplicates")
        analysis_layout.addWidget(self.validate)
        analysis_layout.addWidget(self.remove_dups)
        analysis_layout.addWidget(self.highlight_dups)
        advanced_layout.addLayout(analysis_layout)

        # Sort by
        sort_layout = QHBoxLayout()
        sort_layout.addWidget(QLabel("Sort by:"))
        self.sort_by = QComboBox()
        self.sort_by.addItems(['None', 'timecode', 'clip_name', 'source_file', 'duration', 'category'])
        sort_layout.addWidget(self.sort_by)
        sort_layout.addWidget(QLabel("FPS:"))
        self.fps = QDoubleSpinBox()
        self.fps.setRange(1, 240)
        self.fps.setValue(30)
        self.fps.setDecimals(3)
        sort_layout.addWidget(self.fps)
        sort_layout.addStretch()
        advanced_layout.addLayout(sort_layout)

        # Filter
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Filter:"))
        self.filter_expr = QLineEdit()
        self.filter_expr.setPlaceholderText('e.g., Category == "A-Camera"')
        filter_layout.addWidget(self.filter_expr)
        advanced_layout.addLayout(filter_layout)

        # Split by category
        split_layout = QHBoxLayout()
        self.split_by_category = QCheckBox("Split by category into separate Excel files")
        split_layout.addWidget(self.split_by_category)
        split_layout.addWidget(QLabel("Output Dir:"))
        self.split_output_dir = QLineEdit("./split_output")
        split_layout.addWidget(self.split_output_dir)
        advanced_layout.addLayout(split_layout)

        advanced_group.setLayout(advanced_layout)
        layout.addWidget(advanced_group)

        # Grouping
        grouping_box = QGroupBox("Event Grouping")
        grouping_layout = QHBoxLayout()

        grouping_layout.addWidget(QLabel("Group events (seconds):"))
        self.group_interval = QDoubleSpinBox()
        self.group_interval.setRange(0, 1000)
        self.group_interval.setValue(0)
        self.group_interval.setSpecialValueText("Disabled")
        self.group_interval.setToolTip("Group events into continuous time intervals (requires FPS setting above)")
        grouping_layout.addWidget(self.group_interval)

        grouping_box.setLayout(grouping_layout)
        layout.addWidget(grouping_box)

        # Info
        info_label = QLabel(
            "EDL files will be merged and sorted by timecode (interleaved)"
        )
        info_label.setStyleSheet("padding: 5px; background-color: #f0f0f0; border-radius: 3px; font-size: 10pt;")
        layout.addWidget(info_label)

        layout.addStretch()
        self.setLayout(layout)

    def add_files(self):
        filenames, _ = QFileDialog.getOpenFileNames(
            self, "Select EDL Files to Merge", "", "EDL Files (*.edl);;All Files (*)"
        )
        if filenames:
            current = self.file_list.toPlainText()
            if current:
                current += "\n"
            current += "\n".join(filenames)
            self.file_list.setPlainText(current)

    def browse_output(self):
        filename, _ = QFileDialog.getSaveFileName(
            self, "Save Merged Output", "", "Excel Files (*.xlsx);;All Files (*)"
        )
        if filename:
            self.output_file.setText(filename)

    def browse_subtitle(self):
        """Browse for subtitle file (STL or SRT)"""
        filename, _ = QFileDialog.getOpenFileName(
            self, "Select Subtitle File", "",
            "Subtitle Files (*.stl *.srt);;STL Files (*.stl);;SRT Files (*.srt);;All Files (*)"
        )
        if filename:
            self.subtitle_file.setText(filename)

    def browse_config(self):
        """Browse for YAML config file"""
        filename, _ = QFileDialog.getOpenFileName(
            self, "Select Config File", "", "YAML Files (*.yaml *.yml);;All Files (*)"
        )
        if filename:
            self.config_file.setText(filename)

    def toggle_fps_controls(self, checked):
        """Enable/disable FPS controls based on auto-detect checkbox"""
        # When auto-detect is ON (checked), disable manual FPS controls
        # When auto-detect is OFF (unchecked), enable manual FPS controls
        self.stl_fps.setEnabled(not checked)
        self.stl_fps_preset.setEnabled(not checked)

    def update_fps_from_preset(self, preset_text):
        """Update FPS value from preset dropdown"""
        # Extract FPS value from preset text (e.g., "23.976 (Film/NTSC)" -> 23.976)
        fps_map = {
            '23.976 (Film/NTSC)': 23.976,
            '24 (Film)': 24.0,
            '25 (PAL)': 25.0,
            '29.97 (NTSC)': 29.97,
            '30 (NTSC)': 30.0
        }
        if preset_text in fps_map:
            self.stl_fps.setValue(fps_map[preset_text])

    def get_command(self):
        """Build command from UI options"""
        files = [f.strip() for f in self.file_list.toPlainText().split('\n') if f.strip()]
        subtitle_file_specified = bool(self.subtitle_file.text().strip())

        # Validation: need at least 2 EDLs for merging, OR 1 EDL with subtitle file
        if len(files) < 1:
            raise ValueError("At least 1 EDL file is required")
        if len(files) < 2 and not subtitle_file_specified:
            raise ValueError("At least 2 EDL files are required for merging, or 1 EDL file with subtitle file (.stl or .srt)")

        # Build command: python edl_parse.py --merge file1.edl file2.edl --output output.xlsx ...
        cmd = ['python', 'edl_parse.py', '--merge'] + files

        # Add output file using --output flag (not as positional argument)
        cmd.extend(['--output', self.output_file.text()])

        # Config file (for categorization and formatting)
        if self.config_file.text().strip():
            cmd.extend(['--config', self.config_file.text().strip()])

        # Format options
        if self.no_table.isChecked():
            cmd.append('--no-table')
        if self.no_colored.isChecked():
            cmd.append('--no-colored')

        # Statistics
        if self.add_stats.isChecked():
            cmd.append('--stats')

        # Validation and duplicates
        if self.validate.isChecked():
            cmd.append('--validate')
        if self.remove_dups.isChecked():
            cmd.append('--remove-duplicates')
        if self.highlight_dups.isChecked():
            cmd.append('--highlight-duplicates')

        # Sorting
        if self.sort_by.currentText() != 'None':
            cmd.extend(['--sort-by', self.sort_by.currentText()])
            if self.sort_by.currentText() == 'duration':
                cmd.extend(['--fps', str(self.fps.value())])

        # Filter
        if self.filter_expr.text().strip():
            cmd.extend(['--filter', self.filter_expr.text().strip()])

        # Split by category
        if self.split_by_category.isChecked():
            cmd.append('--split-by-category')
            if self.split_output_dir.text().strip():
                cmd.extend(['--split-output-dir', self.split_output_dir.text().strip()])

        # Add subtitle file if specified (STL or SRT)
        if subtitle_file_specified:
            cmd.extend(['--subtitle-file', self.subtitle_file.text().strip()])

            # Only add --subtitle-fps if auto-detect is OFF (user wants to override)
            # For STL: overrides auto-detection. For SRT: sets FPS (default 30 if not specified)
            if not self.stl_auto_fps.isChecked():
                cmd.extend(['--subtitle-fps', str(self.stl_fps.value())])

            # Add --subtitle-start-time if specified
            if self.stl_start_time.text().strip():
                cmd.extend(['--subtitle-start-time', self.stl_start_time.text().strip()])

        # Grouping
        if self.group_interval.value() > 0:
            cmd.extend(['--group', str(self.group_interval.value())])
            cmd.extend(['--fps', str(self.fps.value())])

        return cmd


class MainWindow(QMainWindow):
    """Main application window"""

    def __init__(self):
        super().__init__()
        self.process_thread = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("EDL Parser - Professional EDL Analysis Tool")
        self.setGeometry(100, 100, 1000, 750)

        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout()
        layout.setSpacing(5)  # Reduce spacing between widgets

        # Title
        title = QLabel("EDL Parser GUI")
        title_font = QFont()
        title_font.setPointSize(14)  # Reduced from 16
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # Tabs
        self.tabs = QTabWidget()
        self.convert_tab = self._make_scrollable(ConvertTab())
        self.compare_tab = self._make_scrollable(CompareTab())
        self.merge_tab = self._make_scrollable(MergeTab())

        self.tabs.addTab(self.convert_tab, "Convert EDL")
        self.tabs.addTab(self.compare_tab, "Compare EDLs")
        self.tabs.addTab(self.merge_tab, "Merge EDLs")

        layout.addWidget(self.tabs)

        # Output log
        log_group = QGroupBox("Output Log")
        log_layout = QVBoxLayout()

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setMaximumHeight(150)  # Reduced from 200
        self.log_output.setStyleSheet("font-family: monospace; background-color: #2b2b2b; color: #f0f0f0;")
        log_layout.addWidget(self.log_output)

        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # Buttons
        btn_layout = QHBoxLayout()

        self.run_btn = QPushButton("Run")
        self.run_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
        self.run_btn.clicked.connect(self.run_command)
        btn_layout.addWidget(self.run_btn)

        self.clear_log_btn = QPushButton("Clear Log")
        self.clear_log_btn.clicked.connect(self.log_output.clear)
        btn_layout.addWidget(self.clear_log_btn)

        help_btn = QPushButton("Help")
        help_btn.clicked.connect(self.show_help)
        btn_layout.addWidget(help_btn)

        layout.addLayout(btn_layout)

        central_widget.setLayout(layout)

    def _make_scrollable(self, widget):
        """Wrap a widget in a scroll area to make it scrollable"""
        scroll = QScrollArea()
        scroll.setWidget(widget)
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        return scroll

    def run_command(self):
        """Execute the EDL parser with current options"""
        try:
            # Get command from active tab
            current_tab = self.tabs.currentWidget()
            # If tab is wrapped in scroll area, get the actual widget
            if isinstance(current_tab, QScrollArea):
                current_tab = current_tab.widget()
            cmd = current_tab.get_command()

            # Log command
            self.log_output.append(f"\n{'='*60}")
            self.log_output.append(f"Executing: {' '.join(cmd)}")
            self.log_output.append(f"{'='*60}\n")

            # Disable run button
            self.run_btn.setEnabled(False)
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 0)  # Indeterminate progress

            # Start process thread
            self.process_thread = EDLProcessThread(cmd)
            self.process_thread.output_signal.connect(self.append_output)
            self.process_thread.error_signal.connect(self.append_error)
            self.process_thread.finished_signal.connect(self.process_finished)
            self.process_thread.start()

        except ValueError as e:
            QMessageBox.warning(self, "Input Error", str(e))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to execute command: {str(e)}")

    def append_output(self, text):
        """Append output to log"""
        self.log_output.append(text)
        # Auto-scroll to bottom
        cursor = self.log_output.textCursor()
        cursor.movePosition(QTextCursor.End)
        self.log_output.setTextCursor(cursor)

    def append_error(self, text):
        """Append error to log"""
        self.log_output.append(f"<span style='color: red;'>{text}</span>")

    def process_finished(self, return_code):
        """Handle process completion"""
        self.run_btn.setEnabled(True)
        self.progress_bar.setVisible(False)

        if return_code == 0:
            self.log_output.append(f"\n{'='*60}")
            self.log_output.append("✓ Process completed successfully!")
            self.log_output.append(f"{'='*60}\n")
            QMessageBox.information(self, "Success", "EDL processing completed successfully!")
        else:
            self.log_output.append(f"\n{'='*60}")
            self.log_output.append(f"✗ Process failed with exit code: {return_code}")
            self.log_output.append(f"{'='*60}\n")
            QMessageBox.warning(self, "Process Failed", f"EDL processing failed (exit code: {return_code})")

    def show_help(self):
        """Show help dialog"""
        help_text = """
<h2>EDL Parser GUI - Help</h2>

<h3>Convert EDL Tab</h3>
<p>Convert EDL files to Excel with various options:</p>
<ul>
<li><b>Input EDL:</b> Source EDL file to process</li>
<li><b>Output File:</b> Destination Excel file</li>
<li><b>Config File:</b> YAML configuration for categorization</li>
<li><b>Statistics:</b> Generate comprehensive statistics</li>
<li><b>Remove Duplicates:</b> Automatically remove duplicate events</li>
<li><b>Sort by:</b> Sort events by various criteria</li>
<li><b>Filter:</b> Filter events using pandas query syntax</li>
<li><b>Search:</b> Search for specific events (glob or regex)</li>
</ul>

<h3>Compare EDLs Tab</h3>
<p>Compare two EDL versions and generate a changelog report showing added, removed, and modified events.</p>

<h3>Merge EDLs Tab</h3>
<p>Combine multiple EDL files into one:</p>
<ul>
<li><b>Append mode:</b> Sequential merging with timecode adjustment</li>
<li><b>Interleave mode:</b> Merge and sort by timecode</li>
</ul>

<h3>Documentation</h3>
<p>For complete documentation, see:</p>
<ul>
<li>USER_GUIDE.md - Complete user guide</li>
<li>QUICK_REFERENCE.md - Quick reference card</li>
<li>CATEGORIZATION_GUIDE.md - Categorization details</li>
</ul>
        """

        msg = QMessageBox()
        msg.setWindowTitle("Help")
        msg.setTextFormat(Qt.RichText)
        msg.setText(help_text)
        msg.exec_()


def main():
    app = QApplication(sys.argv)

    # Set application style
    app.setStyle('Fusion')

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
