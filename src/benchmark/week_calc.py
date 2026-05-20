"""财年周/月/季度选择.

我们不重新实现 SanDisk 的 fiscal calendar，而是直接利用源数据里的
``WEEK`` 列和 ``STARTTIME``/``ENDTIME`` 来构建索引：

    fiscal_week = "2026FW45" (FLAG="W")
    fiscal_month = "2026M05"  (FLAG="M")
    fiscal_quarter = "2026FQ4" (FLAG="Q")

需求：根据当前日期，找出
* 上一周(prev_week)。其 9 周窗口 [prev_week-8, prev_week]
* 上一月 prev_week 所属月
* 上一季 prev_week 所属季
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, date
import re

import pandas as pd

from .data_loader import COL_WEEK, COL_FLAG, COL_STARTTIME, COL_ENDTIME

_WEEK_RE = re.compile(r"^(\d{4})FW(\d{1,2})$", re.IGNORECASE)
_QUARTER_RE = re.compile(r"^(\d{4})FQ([1-4])$", re.IGNORECASE)
_MONTH_RE = re.compile(r"^(\d{4})M(\d{1,2})$", re.IGNORECASE)


@dataclass(frozen=True)
class PeriodSelection:
    """选定的财年周/月/季的标签集合."""

    current_week: str
    prev_week: str
    weeks_window: tuple[str, ...]
    month: str
    quarter: str

    def all_period_labels(self) -> tuple[str, ...]:
        return (self.quarter, self.month, *self.weeks_window)


def _norm_week(label: str) -> tuple[int, int] | None:
    m = _WEEK_RE.match(label.strip())
    return (int(m.group(1)), int(m.group(2))) if m else None


def _norm_month(label: str) -> tuple[int, int] | None:
    m = _MONTH_RE.match(label.strip())
    return (int(m.group(1)), int(m.group(2))) if m else None


def _norm_quarter(label: str) -> tuple[int, int] | None:
    m = _QUARTER_RE.match(label.strip())
    return (int(m.group(1)), int(m.group(2))) if m else None


def _week_key(label: str) -> tuple[int, int]:
    n = _norm_week(label)
    return n if n is not None else (-1, -1)


def find_current_week(
    df: pd.DataFrame, today: date | datetime | None = None
) -> str:
    """根据原始数据中的 STARTTIME/ENDTIME 找包含 today 的 W 行的 WEEK 标签."""
    if today is None:
        today = datetime.now()
    if isinstance(today, date) and not isinstance(today, datetime):
        today = datetime.combine(today, datetime.min.time())
    today_ts = pd.Timestamp(today)

    w = df[df[COL_FLAG] == "W"]
    mask = (w[COL_STARTTIME] <= today_ts) & (w[COL_ENDTIME] >= today_ts)
    candidates = w.loc[mask, COL_WEEK].dropna().unique()
    if len(candidates) > 0:
        return str(candidates[0])

    valid = w.dropna(subset=[COL_ENDTIME])
    earlier = valid[valid[COL_ENDTIME] < today_ts]
    if not earlier.empty:
        latest = earlier.sort_values(COL_ENDTIME).iloc[-1]
        return str(latest[COL_WEEK])

    weeks = sorted({str(x) for x in w[COL_WEEK].dropna().unique() if _norm_week(str(x))}, key=_week_key)
    if not weeks:
        raise ValueError("数据集中没有 W 行的 WEEK 标签")
    return weeks[-1]


def _all_weeks(df: pd.DataFrame) -> list[str]:
    weeks = {str(x) for x in df.loc[df[COL_FLAG] == "W", COL_WEEK].dropna().unique()}
    weeks = {w for w in weeks if _norm_week(w)}
    return sorted(weeks, key=_week_key)


def previous_week_label(week: str, weeks_pool: list[str]) -> str:
    """在 weeks_pool 里找 week 的上一周标签 (按 fiscal year, week# 排序)."""
    if week in weeks_pool:
        idx = weeks_pool.index(week)
        if idx > 0:
            return weeks_pool[idx - 1]
    yr, wk = _norm_week(week) or (None, None)
    if yr is None:
        raise ValueError(f"非法 fiscal week 标签: {week}")
    target_yr, target_wk = (yr, wk - 1) if wk > 1 else (yr - 1, 52)
    target = f"{target_yr}FW{target_wk:02d}"
    if target in weeks_pool:
        return target
    earlier = [w for w in weeks_pool if _week_key(w) < (yr, wk)]
    if earlier:
        return earlier[-1]
    raise ValueError(f"找不到 {week} 之前的周")


def select_periods(
    df: pd.DataFrame,
    today: date | datetime | None = None,
    *,
    weeks_count: int = 9,
) -> PeriodSelection:
    """计算输出报表里需要的所有 period 标签.

    需求(参考 README):
    * current_week  = 包含 today 的 W 标签 (e.g. 2026FW45)
    * prev_week     = current_week 的上一周                  (e.g. 2026FW44)
    * weeks_window  = prev_week 之前 N 周 (含 prev_week)      (e.g. FW36..FW44)
    * month         = 上一个月 (相对 today 的当前月减 1)      (e.g. 2026M04)
    * quarter       = 上一个季度 (相对 today 的当前季减 1)    (e.g. 2026FQ3)

    所有标签都直接来自原始数据 ``WEEK`` 列, 不重新计算 fiscal calendar。
    """
    if today is None:
        today = datetime.now()
    if isinstance(today, date) and not isinstance(today, datetime):
        today = datetime.combine(today, datetime.min.time())
    today_ts = pd.Timestamp(today)

    current = find_current_week(df, today)
    pool = _all_weeks(df)
    prev = previous_week_label(current, pool)

    if prev not in pool:
        raise ValueError(f"上一周 {prev} 不在数据池里, 现有: {pool[-12:]}")
    idx = pool.index(prev)
    start = max(0, idx - weeks_count + 1)
    weeks_window = tuple(pool[start : idx + 1])

    month = _pick_prev_period(df, "M", today_ts)
    quarter = _pick_prev_period(df, "Q", today_ts)

    return PeriodSelection(
        current_week=current,
        prev_week=prev,
        weeks_window=weeks_window,
        month=month,
        quarter=quarter,
    )


def _pick_prev_period(df: pd.DataFrame, flag: str, today_ts: pd.Timestamp) -> str:
    """选 ``flag`` 类型 (M/Q) 的"上一段".

    策略: 取 ``ENDTIME < today_ts`` 的最近一行的 WEEK 标签。这样能自然得到
    "上一个月" / "上一个季度", 且完全依赖原始数据的 STARTTIME/ENDTIME, 我们
    自己不用算 fiscal calendar。
    """
    sub = df[df[COL_FLAG] == flag]
    if sub.empty:
        raise ValueError(f"数据集中没有 FLAG={flag!r} 的行")

    earlier = sub.dropna(subset=[COL_ENDTIME])
    earlier = earlier[earlier[COL_ENDTIME] < today_ts]
    if not earlier.empty:
        last = earlier.sort_values(COL_ENDTIME).iloc[-1]
        return str(last[COL_WEEK])

    last_any = sub.sort_values(COL_ENDTIME).iloc[-1]
    return str(last_any[COL_WEEK])
