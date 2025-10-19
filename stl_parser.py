"""
STL Parser Module for EBU N19 STL Subtitle Files

This module provides functionality to parse EBU N19 STL (EBU Tech 3264) subtitle files
and match subtitle text with EDL events based on timecode overlap.
ßß
Author: Claude Code
Date: 2025-10-18
"""

import struct
import logging
from pathlib import Path
from timecode import Timecode

logger = logging.getLogger(__name__)


class STLSubtitle:
    """Represents a single subtitle from an STL file"""

    def __init__(self, subtitle_number, time_in, time_out, text):
        """
        Initialize a subtitle entry.

        Args:
            subtitle_number: Sequential subtitle number
            time_in: Timecode object for subtitle start
            time_out: Timecode object for subtitle end
            text: Subtitle text content
        """
        self.subtitle_number = subtitle_number
        self.time_in = time_in  # Timecode object
        self.time_out = time_out  # Timecode object
        self.text = text

    def overlaps(self, record_in, record_out, fps=30, offset=None):
        """
        Check if subtitle overlaps with EDL event timecode range.

        Args:
            record_in: Record In timecode string (HH:MM:SS:FF)
            record_out: Record Out timecode string (HH:MM:SS:FF)
            fps: Frame rate for timecode calculations
            offset: Frame offset to apply to subtitle timecodes (integer, can be negative)
                   Positive offset = subtract from subtitles (move earlier)
                   Negative offset = add to subtitles (move later)

        Returns:
            True if there is any overlap between subtitle and event
        """
        try:
            event_in = Timecode(fps, record_in)
            event_out = Timecode(fps, record_out)

            # Apply offset to subtitle timecodes if provided
            if offset is not None:
                # Subtract offset (positive offset moves subtitles earlier, negative moves later)
                sub_time_in_frames = self.time_in.frames - offset
                sub_time_out_frames = self.time_out.frames - offset

                # Ensure we don't go below frame 1 (Timecode library minimum)
                if sub_time_in_frames < 1 or sub_time_out_frames < 1:
                    return False
            else:
                sub_time_in_frames = self.time_in.frames
                sub_time_out_frames = self.time_out.frames

            # Check for any overlap
            # No overlap if: subtitle ends before event starts OR subtitle starts after event ends
            return not (sub_time_out_frames <= event_in.frames or
                       sub_time_in_frames >= event_out.frames)
        except Exception as e:
            logger.warning(f"Error checking overlap for subtitle {self.subtitle_number}: {e}")
            return False

    def __repr__(self):
        return f"STLSubtitle({self.subtitle_number}, {self.time_in}, {self.time_out}, '{self.text[:30]}...')"


def detect_stl_fps(stl_path):
    """
    Attempt to detect FPS from STL file GSI block.

    This function tries multiple methods to detect the frame rate:
    1. Parse DFC (Disk Format Code) field from GSI block
    2. Infer from maximum frame numbers in TTI blocks (less reliable)

    Args:
        stl_path: Path to STL file

    Returns:
        Detected FPS as float, or None if detection fails
    """
    stl_path = Path(stl_path)

    try:
        with open(stl_path, 'rb') as f:
            # Read GSI block (1024 bytes)
            gsi = f.read(1024)

            if len(gsi) < 1024:
                logger.warning(f"STL file too small (< 1024 bytes): {stl_path}")
                return None

            # Method 1: Check DFC field (bytes 3-10, 8 bytes)
            # DFC should contain strings like "STL25.01", "STL30.01", etc.
            dfc = gsi[3:11].decode('ascii', errors='ignore').strip()
            logger.debug(f"DFC field: '{dfc}'")

            if 'STL25' in dfc:
                logger.info(f"Detected 25 fps from DFC field")
                return 25.0
            elif 'STL30' in dfc:
                logger.info(f"Detected 30 fps from DFC field")
                return 30.0
            elif 'STL24' in dfc:
                logger.info(f"Detected 24 fps from DFC field")
                return 24.0
            elif 'STL23' in dfc:
                logger.info(f"Detected 23.976 fps from DFC field")
                return 23.976
            elif 'STL29' in dfc:
                logger.info(f"Detected 29.97 fps from DFC field")
                return 29.97

            # Method 2: Infer from max frame numbers (less reliable)
            logger.debug("DFC field empty or invalid, attempting frame number inference")

            max_frame = 0
            block_count = 0

            # Read TTI blocks (128 bytes each)
            while True:
                tti = f.read(128)
                if len(tti) < 128:
                    break

                block_count += 1

                # Extract frame numbers from TCI (bytes 5-9) and TCO (bytes 9-13)
                try:
                    _, _, _, frames_in = struct.unpack('4B', tti[5:9])
                    _, _, _, frames_out = struct.unpack('4B', tti[9:13])
                    max_frame = max(max_frame, frames_in, frames_out)
                except:
                    continue

            logger.debug(f"Analyzed {block_count} TTI blocks, max frame: {max_frame}")

            # Infer FPS from max frame number
            if max_frame <= 23:
                logger.warning(f"Inferred 23.976 fps from max frame {max_frame} (ambiguous - consider using --stl-fps)")
                return 23.976
            elif max_frame <= 24:
                logger.warning(f"Inferred 24 fps from max frame {max_frame} (ambiguous - consider using --stl-fps)")
                return 24.0
            elif max_frame <= 25:
                logger.warning(f"Inferred 25 fps from max frame {max_frame} (ambiguous - consider using --stl-fps)")
                return 25.0
            elif max_frame <= 29:
                logger.warning(f"Inferred 29.97 fps from max frame {max_frame} (ambiguous - consider using --stl-fps)")
                return 29.97
            else:
                logger.warning(f"Inferred 30 fps from max frame {max_frame}")
                return 30.0

    except Exception as e:
        logger.error(f"Error detecting STL FPS: {e}")
        return None

    logger.warning("Could not detect FPS from STL file")
    return None


def parse_stl_file(stl_path, fps=None):
    """
    Parse an EBU N19 STL file and extract subtitles.

    Args:
        stl_path: Path to STL file
        fps: Frame rate (optional). If None, attempts auto-detection.
             User-provided FPS ALWAYS overrides auto-detection.

    Returns:
        Tuple of (subtitles_list, effective_fps)
        - subtitles_list: List of STLSubtitle objects
        - effective_fps: The FPS used for parsing

    Raises:
        ValueError: If FPS cannot be determined
        FileNotFoundError: If STL file doesn't exist
    """
    stl_path = Path(stl_path)

    if not stl_path.exists():
        raise FileNotFoundError(f"STL file not found: {stl_path}")

    # Determine effective FPS
    if fps is not None:
        # User override - always takes precedence
        logger.info(f"Using user-specified FPS: {fps}")
        effective_fps = fps
    else:
        # Attempt auto-detection
        logger.info(f"Attempting to auto-detect FPS from STL file: {stl_path}")
        detected_fps = detect_stl_fps(stl_path)

        if detected_fps:
            logger.info(f"Auto-detected FPS: {detected_fps}")
            effective_fps = detected_fps
        else:
            raise ValueError(
                f"Cannot auto-detect FPS for {stl_path}. "
                "Please specify FPS using --stl-fps argument or GUI FPS controls."
            )

    # Parse STL file
    subtitles = []

    try:
        with open(stl_path, 'rb') as f:
            # Skip GSI block (1024 bytes)
            gsi = f.read(1024)

            if len(gsi) < 1024:
                raise ValueError(f"Invalid STL file (too small): {stl_path}")

            # Parse TTI blocks (128 bytes each)
            block_num = 0
            while True:
                tti = f.read(128)
                if len(tti) < 128:
                    break

                block_num += 1

                try:
                    # Parse TTI block structure
                    sgn = tti[0]  # Subtitle Group Number
                    sn = struct.unpack('>H', tti[1:3])[0]  # Subtitle Number (big-endian)
                    ebn = tti[3]  # Extension Block Number
                    cs = tti[4]   # Cumulative Status

                    # Time Code In (TCI): bytes 5-9 (HH:MM:SS:FF)
                    tci_h, tci_m, tci_s, tci_f = struct.unpack('4B', tti[5:9])

                    # Time Code Out (TCO): bytes 9-13 (HH:MM:SS:FF)
                    tco_h, tco_m, tco_s, tco_f = struct.unpack('4B', tti[9:13])

                    # Vertical Position, Justification Code, Comment Flag
                    vp = tti[13]
                    jc = tti[14]
                    cf = tti[15]

                    # Text Field: bytes 16-127 (112 bytes)
                    text_bytes = tti[16:128]

                    # Skip if this is an extension block (EBN != 255)
                    if ebn != 255:
                        continue

                    # Decode text (Latin-1 encoding, remove 0x8F padding)
                    text = text_bytes.replace(b'\x8f', b'').decode('latin-1', errors='ignore').strip()

                    # Skip empty subtitles
                    if not text:
                        continue

                    # Create Timecode objects
                    tci_str = f"{tci_h:02d}:{tci_m:02d}:{tci_s:02d}:{tci_f:02d}"
                    tco_str = f"{tco_h:02d}:{tco_m:02d}:{tco_s:02d}:{tco_f:02d}"

                    time_in = Timecode(effective_fps, tci_str)
                    time_out = Timecode(effective_fps, tco_str)

                    # Create subtitle object
                    subtitle = STLSubtitle(sn, time_in, time_out, text)
                    subtitles.append(subtitle)

                    logger.debug(f"Parsed subtitle {sn}: {tci_str} -> {tco_str}: {text[:50]}")

                except Exception as e:
                    logger.warning(f"Error parsing TTI block {block_num}: {e}")
                    continue

            logger.info(f"Successfully parsed {len(subtitles)} subtitles from {stl_path} at {effective_fps} fps")

    except Exception as e:
        logger.error(f"Error parsing STL file {stl_path}: {e}")
        raise

    return subtitles, effective_fps


def match_subtitles_to_events(df, subtitles, fps, stl_start_time=None):
    """
    Match subtitles to EDL events based on timeline timecode overlap.

    Subtitles are matched against 'Record In' and 'Record Out' (timeline positions)
    rather than 'Timecode In' and 'Timecode Out' (source media timecodes), since STL files
    are aligned to the final timeline edit.

    Args:
        df: DataFrame with EDL events (must have 'Record In' and 'Record Out' columns)
        subtitles: List of STLSubtitle objects
        fps: Frame rate for timecode calculations
        stl_start_time: Record In timecode (HH:MM:SS:FF string) where the first STL subtitle should align.
                        The offset is automatically calculated based on the first subtitle's timecode.

    Returns:
        DataFrame with new 'Subtitles' column containing matched subtitle text
    """
    if subtitles is None or len(subtitles) == 0:
        logger.warning("No subtitles to match")
        return df

    # Calculate offset if STL start time is provided
    offset = None
    if stl_start_time:
        # Get the timecode of the first subtitle
        first_subtitle_tc = subtitles[0].time_in

        # Parse the user-defined start time
        start_time_tc = Timecode(fps, stl_start_time)

        # Calculate offset as frame count (can be negative)
        # Positive offset = subtract from subtitles (move earlier)
        # Negative offset = add to subtitles (move later)
        offset = first_subtitle_tc.frames - start_time_tc.frames

        logger.info(f"STL Start Time: {stl_start_time}")
        logger.info(f"First STL subtitle at: {first_subtitle_tc}")
        if offset > 0:
            logger.info(f"Calculated offset: -{offset} frames (subtitles will be moved earlier)")
        elif offset < 0:
            logger.info(f"Calculated offset: +{-offset} frames (subtitles will be moved later)")
        else:
            logger.info(f"No offset needed (already aligned)")

    # Match against record timecodes (timeline positions)
    if 'Record In' in df.columns and 'Record Out' in df.columns:
        timecode_in_col = 'Record In'
        timecode_out_col = 'Record Out'
        if offset:
            logger.info(f"Matching subtitles against timeline positions (Record In/Out) with offset {offset}")
        else:
            logger.info("Matching subtitles against timeline positions (Record In/Out)")
    # Fall back to source timecodes if record timecodes not available
    elif 'Timecode In' in df.columns and 'Timecode Out' in df.columns:
        timecode_in_col = 'Timecode In'
        timecode_out_col = 'Timecode Out'
        logger.warning("Record timecodes not available, falling back to source timecodes (Timecode In/Out)")
    else:
        logger.error("DataFrame missing required timecode columns")
        return df

    # Create new column for subtitles
    subtitle_texts = []

    for idx, row in df.iterrows():
        tc_in = str(row[timecode_in_col])
        tc_out = str(row[timecode_out_col])

        # Find all subtitles that overlap with this event
        matching_subs = []
        for sub in subtitles:
            if sub.overlaps(tc_in, tc_out, fps, offset=offset):
                matching_subs.append(sub.text)

        # Join multiple subtitles with " | "
        if matching_subs:
            subtitle_text = " | ".join(matching_subs)
            logger.debug(f"Event {idx+1} ({tc_in} - {tc_out}): Matched {len(matching_subs)} subtitle(s)")
        else:
            subtitle_text = ""

        subtitle_texts.append(subtitle_text)

    # Add column to DataFrame
    df['Subtitles'] = subtitle_texts

    matched_count = sum(1 for text in subtitle_texts if text)
    logger.info(f"Matched subtitles to {matched_count} of {len(df)} events")

    return df
