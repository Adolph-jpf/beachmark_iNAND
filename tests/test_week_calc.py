from __future__ import annotations

from datetime import datetime

import pytest

from benchmark.data_loader import load_yield_data
from benchmark.week_calc import (
    PeriodSelection,
    find_current_week,
    previous_week_label,
    select_periods,
)


def test_find_current_week_synth(synth_yield_text):
    df = load_yield_data(synth_yield_text)
    today = datetime(2026, 5, 7)
    cur = find_current_week(df, today)
    assert cur in {"2026FW46", "2026FW45"}


def test_select_periods_synth(synth_yield_text):
    df = load_yield_data(synth_yield_text)
    today = datetime(2026, 5, 7)
    sel = select_periods(df, today)
    assert isinstance(sel, PeriodSelection)
    assert len(sel.weeks_window) == 9
    assert sel.weeks_window[0] == "2026FW36"
    assert sel.weeks_window[-1] == "2026FW44" or sel.weeks_window[-1] == "2026FW45"
    assert sel.month.startswith("2026M")
    assert sel.quarter.startswith("2026FQ")


def test_previous_week_simple():
    pool = [f"2026FW{i:02d}" for i in range(30, 46)]
    assert previous_week_label("2026FW40", pool) == "2026FW39"
    assert previous_week_label("2026FW45", pool) == "2026FW44"


def test_previous_week_year_wrap():
    pool = ["2025FW51", "2025FW52", "2026FW01", "2026FW02"]
    assert previous_week_label("2026FW01", pool) == "2025FW52"


def test_select_periods_real(real_data_path):
    if not real_data_path.exists():
        pytest.skip("real data not present")
    df = load_yield_data(real_data_path)
    sel = select_periods(df, datetime(2026, 5, 7))
    assert isinstance(sel, PeriodSelection)
    assert sel.prev_week == sel.weeks_window[-1]
    assert 1 <= len(sel.weeks_window) <= 9
    assert sel.month.startswith("2026M")
    assert sel.quarter.startswith("2026FQ")


@pytest.mark.parametrize(
    "today, expected_curr, expected_prev, expected_first, expected_last, expected_month, expected_quarter",
    [
        # 5/7 (W45 中) -> 上一周 W44, 窗口 W36..W44, M04, FQ3
        (
            datetime(2026, 5, 7),
            "2026FW45",
            "2026FW44",
            "2026FW36",
            "2026FW44",
            "2026M04",
            "2026FQ3",
        ),
        # 5/13 (W46 中) -> 上一周 W45, 窗口 W37..W45, M04 (5/31 还没结束), FQ3
        (
            datetime(2026, 5, 13),
            "2026FW46",
            "2026FW45",
            "2026FW37",
            "2026FW45",
            "2026M04",
            "2026FQ3",
        ),
        # 6/15 (5 月已结束) -> month 自动切到 M05; quarter 仍 FQ3 (FQ4 7/5 才结束)
        (
            datetime(2026, 6, 15),
            "2026FW46",  # 数据池只到 W46, fallback 到最新一周
            "2026FW45",
            "2026FW37",
            "2026FW45",
            "2026M05",
            "2026FQ3",
        ),
    ],
    ids=["may-7-W45", "may-13-W46", "jun-15-cross-month"],
)
def test_select_periods_auto_advance(
    synth_yield_text,
    today,
    expected_curr,
    expected_prev,
    expected_first,
    expected_last,
    expected_month,
    expected_quarter,
):
    """覆盖用户描述的"日期变化自动推算"场景:

    * 5/7 (W45 中) -> M04/FQ3, 窗口 W36..W44
    * 5/13 (W46 中) -> 自动滑动到 W37..W45
    * 6/15 (跨月) -> month 自动 M05, 其他自动滑动
    """
    df = load_yield_data(synth_yield_text)
    sel = select_periods(df, today)
    assert sel.current_week == expected_curr
    assert sel.prev_week == expected_prev
    assert sel.weeks_window[0] == expected_first
    assert sel.weeks_window[-1] == expected_last
    assert sel.month == expected_month
    assert sel.quarter == expected_quarter
