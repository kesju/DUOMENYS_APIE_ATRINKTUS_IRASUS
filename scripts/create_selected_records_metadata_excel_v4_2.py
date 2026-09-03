"""Create the v4_2 selected-record metadata workbook.

This version preserves the v4_1 extraction and formatting logic while moving
the expanded human/ML annotation-count columns directly after ``h_nz_frac%``.
"""

from __future__ import annotations

from pathlib import Path

if __package__:
    from . import create_selected_records_metadata_excel_v4_1 as base
else:
    import create_selected_records_metadata_excel_v4_1 as base


OUTPUT_PATH = Path(
    "/Users/kesju/DI/CODEX_PROJECTS/DUOMENYS_APIE_ATRINKTUS_IRASUS/"
    "results/selected_records_metadata_v4_2.xlsx"
)
OUTPUT_COLUMNS = [
    "basename",
    "recordingId",
    "userId",
    "comment",
    "tag",
    "rhythm",
    "h_nz_frac%",
    "N",
    "hS",
    "hV",
    "hU",
    "mlS",
    "mlV",
    "mlU",
    "FULLY_ANNOTATED_PROFESSIONALLY",
    "BAD_QUALITY",
    "NOISES_ANNOTATED",
    "Other Flags",
]

# The v4_1 functions read these values from their defining module at runtime.
# Updating them here keeps all extraction/formatting behavior unchanged while
# applying the v4_2 output name and column order.
base.OUTPUT_PATH = OUTPUT_PATH
base.OUTPUT_COLUMNS = OUTPUT_COLUMNS

DATA_DIR = base.DATA_DIR
RECORD_SAMPLES = base.RECORD_SAMPLES
extract_tag_from_comment = base.extract_tag_from_comment
extract_rhythm_from_comment = base.extract_rhythm_from_comment
annotation_count = base.annotation_count
calculate_h_nz_fraction = base.calculate_h_nz_fraction
extract_record_metadata = base.extract_record_metadata
estimate_comment_height = base.estimate_comment_height
save_metadata = base.save_metadata
is_manifest = base.is_manifest
main = base.main


if __name__ == "__main__":
    main()
