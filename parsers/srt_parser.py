"""
SRT Parser Module for SubRip Subtitle Files

This module provides functionality to parse SRT subtitle files
and match subtitle text with EDL events based on timecode overlap.

Author: Claude Code
Date: 2025-10-19
"""

import pysrt
import logging
from pathlib import Path
from timecode import Timecode

logger = logging.getLogger(__name__)


class SRTSubtitle:
    """Represents a single subtitle from an SRT file"""

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
            return not (sub_time_out_frames <= event_in.frames or
                       sub_time_in_frames >= event_out.frames)
        except Exception as e:
            logger.warning(f"Error checking overlap for subtitle {self.subtitle_number}: {e}")
            return False

    def __repr__(self):
        return f"SRTSubtitle({self.subtitle_number}, {self.time_in}, {self.time_out}, '{self.text[:30]}...')"


def parse_srt_file(srt_path, fps=30):
    """
    Parse an SRT (SubRip) subtitle file and extract subtitles.

    SRT timecodes use milliseconds (HH:MM:SS,MS) which are converted to
    frame-based timecodes based on the specified FPS.

    Args:
        srt_path: Path to SRT file
        fps: Frame rate for timecode conversion (default: 30)

    Returns:
        Tuple of (subtitles_list, fps)
        - subtitles_list: List of SRTSubtitle objects
        - fps: The FPS used for parsing

    Raises:
        FileNotFoundError: If SRT file doesn't exist
        Exception: If parsing fails
    """
    srt_path = Path(srt_path)

    if not srt_path.exists():
        raise FileNotFoundError(f"SRT file not found: {srt_path}")

    logger.info(f"Parsing SRT file: {srt_path} at {fps} fps")

    subtitles = []

    try:
        # Parse SRT file using pysrt
        srt_subs = pysrt.open(srt_path, encoding='utf-8')

        for srt_sub in srt_subs:
            # Convert SRT timecode (hours, minutes, seconds, milliseconds) to frames
            # SRT format: SubRipTime(hours, minutes, seconds, milliseconds)

            # Time In
            start = srt_sub.start
            start_seconds = start.hours * 3600 + start.minutes * 60 + start.seconds + start.milliseconds / 1000.0
            start_frames = int(start_seconds * fps)
            # Timecode library requires frames >= 1, so use max(1, frames)
            time_in = Timecode(fps, frames=max(1, start_frames))

            # Time Out
            end = srt_sub.end
            end_seconds = end.hours * 3600 + end.minutes * 60 + end.seconds + end.milliseconds / 1000.0
            end_frames = int(end_seconds * fps)
            # Timecode library requires frames >= 1, so use max(1, frames)
            time_out = Timecode(fps, frames=max(1, end_frames))

            # Get text (remove any formatting tags if present)
            text = srt_sub.text.strip()

            # Create subtitle object
            subtitle = SRTSubtitle(srt_sub.index, time_in, time_out, text)
            subtitles.append(subtitle)

            logger.debug(f"Parsed subtitle {srt_sub.index}: {time_in} -> {time_out}: {text[:50]}")

        logger.info(f"Successfully parsed {len(subtitles)} subtitles from {srt_path} at {fps} fps")

    except Exception as e:
        logger.error(f"Error parsing SRT file {srt_path}: {e}")
        raise

    return subtitles, fps


def match_subtitles_to_events(df, subtitles, fps, srt_start_time=None):
    """
    Match SRT subtitles to EDL events based on timeline timecode overlap.

    Subtitles are matched against 'Record In' and 'Record Out' (timeline positions)
    rather than 'Timecode In' and 'Timecode Out' (source media timecodes).

    Args:
        df: DataFrame with EDL events (must have 'Record In' and 'Record Out' columns)
        subtitles: List of SRTSubtitle objects
        fps: Frame rate for timecode calculations
        srt_start_time: Record In timecode (HH:MM:SS:FF string) where the first SRT subtitle should align.
                        The offset is automatically calculated based on the first subtitle's timecode.

    Returns:
        DataFrame with new 'Subtitles' column containing matched subtitle text
    """
    if subtitles is None or len(subtitles) == 0:
        logger.warning("No subtitles to match")
        return df

    # Calculate offset if SRT start time is provided
    offset = None
    if srt_start_time:
        # Get the timecode of the first subtitle
        first_subtitle_tc = subtitles[0].time_in

        # Parse the user-defined start time
        start_time_tc = Timecode(fps, srt_start_time)

        # Calculate offset as frame count (can be negative)
        # Positive offset = subtract from subtitles (move earlier)
        # Negative offset = add to subtitles (move later)
        offset = first_subtitle_tc.frames - start_time_tc.frames

        logger.info(f"SRT Start Time: {srt_start_time}")
        logger.info(f"First SRT subtitle at: {first_subtitle_tc}")
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
