"""共享 fixtures."""

from __future__ import annotations

from pathlib import Path
import textwrap

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture(scope="session")
def real_data_path() -> Path:
    return ROOT / "PN_TEST_STEPID_YIELD.txt"


@pytest.fixture(scope="session")
def real_rule_path() -> Path:
    return ROOT / "Rule_list.xlsx"


@pytest.fixture()
def synth_yield_text(tmp_path: Path) -> Path:
    """构造一段合成的 Spotfire 风格 TSV 数据."""
    header = (
        "PYDATE\tWEEK\tSTARTTIME\tENDTIME\tFLAG\tSITE\tPRODUCT_FAMILY\tFAMILY\tTECH"
        "\tDIEQTY\tPACKAGETYPE\tBOD\tMTECH\tDEVICE\tSTEPNAME0\tSTEPNAME1\tSTEPID_YIELD"
        "\tLT_INQTY/Volume\tLT_OUTQTY\tLT_YIELD\tPACKAGECATEGORY\tCARD_TYPE\tCELLBIT"
        "\tPRODUCTLINE\tPACKAGESIZE\tGOAL_AFTER\tGOAL_BEFORE\tGOAL_IN\tGOAL_NONE"
        "\tYIELD_TARGET\tVOLUMN_TARGET\tREF_MTRT"
    )

    rows = []

    def add(week, start, end, flag, family, dieqty, mtech, device, stepid,
            inqty, outqty, goal=0.99, s_goal=0.995, pydate=None):
        pydate = pydate or end.split(" ")[0]
        y = round(outqty / inqty, 9) if inqty else ""
        cells = [
            pydate, week, start, end, flag, "SDSS", "INAND_X", family, "1Znm",
            dieqty, "PKG", "", mtech, device, "ST", "ST", str(stepid),
            inqty, outqty, y, "INAND", "INANDC", "", "1025(iNAND)", "11x12",
            f"{goal*100:.2f}%", f"{goal*100:.2f}%", f"{goal*100:.2f}%", "",
            f"{s_goal*100:.2f}%", inqty, "",
        ]
        rows.append("\t".join(str(c) for c in cells))

    weeks = [
        ("2026FW36", "2026-03-02 0:00", "2026-03-08 23:59"),
        ("2026FW37", "2026-03-09 0:00", "2026-03-15 23:59"),
        ("2026FW38", "2026-03-16 0:00", "2026-03-22 23:59"),
        ("2026FW39", "2026-03-23 0:00", "2026-03-29 23:59"),
        ("2026FW40", "2026-03-30 0:00", "2026-04-05 23:59"),
        ("2026FW41", "2026-04-06 0:00", "2026-04-12 23:59"),
        ("2026FW42", "2026-04-13 0:00", "2026-04-19 23:59"),
        ("2026FW43", "2026-04-20 0:00", "2026-04-26 23:59"),
        ("2026FW44", "2026-04-27 0:00", "2026-05-03 23:59"),
        ("2026FW45", "2026-05-04 0:00", "2026-05-10 23:59"),
        ("2026FW46", "2026-05-11 0:00", "2026-05-17 23:59"),
    ]

    family = "INAND-CONDOR"
    yield_seq = [0.998, 0.997, 0.996, 0.995, 0.994, 0.993, 0.992, 0.998, 0.999, 0.997, 0.998]
    for (wk, s, e), y in zip(weeks, yield_seq):
        outq = int(2000 * y)
        add(wk, s, e, "W", family, "1", "64Gb", "DEV1", 7405, 2000, outq, goal=0.99, s_goal=0.995)
    add("2026FW42", "2026-04-13 0:00", "2026-04-19 23:59", "W",
        family, "1", "64Gb", "DEV1", 7405, 100, 99, goal=0.99, s_goal=0.995)
    add("2026FQ3", "2026-01-04 0:00", "2026-04-05 23:59", "Q",
        family, "1", "64Gb", "DEV1", 7405, 30000, 29850, goal=0.99, s_goal=0.995)
    add("2026FQ4", "2026-04-06 0:00", "2026-07-05 23:59", "Q",
        family, "1", "64Gb", "DEV1", 7405, 30000, 29850, goal=0.99, s_goal=0.995)
    add("2026M04", "2026-03-30 0:00", "2026-04-26 23:59", "M",
        family, "1", "64Gb", "DEV1", 7405, 12000, 11900, goal=0.99, s_goal=0.995)
    add("2026M05", "2026-04-27 0:00", "2026-05-31 23:59", "M",
        family, "1", "64Gb", "DEV1", 7405, 12000, 11900, goal=0.99, s_goal=0.995)
    add("2026M06", "2026-06-01 0:00", "2026-06-28 23:59", "M",
        family, "1", "64Gb", "DEV1", 7405, 12000, 11900, goal=0.99, s_goal=0.995)
    add("2026FW36", "2026-03-02 0:00", "2026-03-08 23:59", "W",
        "INAND-EAGLE-AUTO", "4", "256Gb", "DEV9", 7008, 5000, 4960, goal=0.985, s_goal=0.985)

    p = tmp_path / "synth.txt"
    p.write_text(header + "\n" + "\n".join(rows) + "\n", encoding="utf-8")
    return p
