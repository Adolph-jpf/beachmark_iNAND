"""Project-wide default paths.

Keep shared operational paths in one place so CLI, GUI, validation, and
PowerShell automation use the same defaults.
"""

from __future__ import annotations

SPOTFIRE_EXPORT_DIR = r"\\cvpfilip03\SDSS_MFG_Data\ENG_Data\Tempfile\ADolph\Spotfire_file"
DEFAULT_DATA_PATH = SPOTFIRE_EXPORT_DIR + r"\PN_TEST_STEPID_YIELD.csv"
DEFAULT_RULE_PATH = "Rule_list.xlsx"
DEFAULT_SRC_GOALS_PATH = "bachmark SRC.xlsx"
DEFAULT_OUTPUT_PATH = "output/INAND_weekly_benchmark.xlsx"
DEFAULT_PPT_PATH = "SDSS INAND YIELD WW45_2026_benchmark.pptx"
DEFAULT_PUBLIC_OUTPUT_DIR = SPOTFIRE_EXPORT_DIR + r"\output"
