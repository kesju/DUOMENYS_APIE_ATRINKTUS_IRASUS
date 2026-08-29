from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path

PROJECT_DIR = Path(
    "/Users/kesju/DI/CODEX_PROJECTS/DUOMENYS_APIE_ATRINKTUS_IRASUS"
)
DEFAULT_INPUT_PATH = PROJECT_DIR / "Anotuotu_sarasas_is_OneNote.md"
DEFAULT_OUTPUT_PATH = PROJECT_DIR / "results" / "selected_records_notes.txt"

RECORD_AT_LINE_START_RE = re.compile(
    r"^\s*(?:[-*+]\s+)?(?:~~|<s(?:\s[^>]*)?>|<del(?:\s[^>]*)?>)?"
    r"(?P<file_name>\d{7}\.\d{3})(?!\d)",
    flags=re.IGNORECASE,
)
ANNOTATION_LINE_RE = re.compile(
    r"^\s*(?:User-Defined|Automatic\s+ML)\s+Annotations\s*\(S/V/U\)\s*:",
    flags=re.IGNORECASE,
)
QUALITY_LINE_RE = re.compile(
    r"^\s*\d+(?:[.,]\d+)?\s+"
    r"(?:Relatively\s+clean(?:asas)?|Moderate\s+distortion|Noisy\s*/\s*suspicious)\b",
    flags=re.IGNORECASE,
)
HTML_STRIKE_OPEN_RE = re.compile(r"<(?:s|del)(?:\s[^>]*)?>", flags=re.IGNORECASE)
HTML_STRIKE_CLOSE_RE = re.compile(r"</(?:s|del)>", flags=re.IGNORECASE)


@dataclass
class RecordNotes:
    file_name: str
    lines: list[str]


def remove_unicode_strike_marks(text: str) -> str:
    """Remove combining strike marks so a struck filename can still be detected."""
    return text.replace("\u0335", "").replace("\u0336", "")


def is_struck_through(line: str, file_name: str) -> bool:
    """Return whether a filename is marked as struck through on its source line."""
    if "\u0335" in line or "\u0336" in line:
        return file_name in remove_unicode_strike_marks(line)

    file_start = line.find(file_name)
    if file_start < 0:
        return False
    file_end = file_start + len(file_name)

    if line[:file_start].count("~~") % 2 == 1 and "~~" in line[file_end:]:
        return True

    html_openings = list(HTML_STRIKE_OPEN_RE.finditer(line, 0, file_start))
    if html_openings:
        last_opening = html_openings[-1]
        closing = HTML_STRIKE_CLOSE_RE.search(line, file_end)
        if closing and not HTML_STRIKE_CLOSE_RE.search(
            line, last_opening.end(), file_start
        ):
            return True

    return False


def find_record_at_line_start(line: str) -> re.Match[str] | None:
    """Find a record filename at the start of a line, including struck variants."""
    return RECORD_AT_LINE_START_RE.match(remove_unicode_strike_marks(line))


def clean_note_lines(lines: list[str]) -> list[str]:
    """Remove generated annotation/status information and trim blank edges."""
    cleaned = [
        line.rstrip()
        for line in lines
        if not ANNOTATION_LINE_RE.match(line) and not QUALITY_LINE_RE.match(line)
    ]

    while cleaned and not cleaned[0].strip():
        cleaned.pop(0)
    while cleaned and not cleaned[-1].strip():
        cleaned.pop()

    compacted: list[str] = []
    for line in cleaned:
        if line.strip() or not compacted or compacted[-1].strip():
            compacted.append(line)
    return compacted


def extract_notes(markdown: str) -> list[RecordNotes]:
    """Extract one note block for each unique, non-struck record filename."""
    records: list[RecordNotes] = []
    seen_file_names: set[str] = set()
    current: RecordNotes | None = None

    for source_line in markdown.splitlines():
        match = find_record_at_line_start(source_line)
        if match:
            file_name = match.group("file_name")
            first_occurrence = file_name not in seen_file_names

            if first_occurrence:
                seen_file_names.add(file_name)
                if is_struck_through(source_line, file_name):
                    current = None
                    continue

                current = RecordNotes(file_name=file_name, lines=[])
                records.append(current)

                remainder = remove_unicode_strike_marks(source_line)[match.end() :]
                remainder = remainder.lstrip(" \t.-–—:")
                if remainder:
                    current.lines.append(remainder)
                continue

        if current is not None:
            current.lines.append(source_line)

    for record in records:
        record.lines = clean_note_lines(record.lines)

    return records


def format_notes(records: list[RecordNotes]) -> str:
    """Format records with the filename first for easy manual control."""
    blocks: list[str] = []
    for record in records:
        block_lines = [record.file_name]
        block_lines.extend(record.lines or ["(no notes)"])
        blocks.append("\n".join(block_lines))
    return "\n\n".join(blocks) + "\n"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract per-record notes from the exported OneNote Markdown file."
    )
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT_PATH)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT_PATH)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    markdown = args.input.read_text(encoding="utf-8")
    records = extract_notes(markdown)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(format_notes(records), encoding="utf-8")
    print(f"Saved notes for {len(records)} records to {args.output}")


if __name__ == "__main__":
    main()
