"""核心处理流程：将原始数据 + 规则 -> 每个 product table 一组 ``ReportRow``."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterable

import pandas as pd

from .data_loader import (
    COL_DEVICE,
    COL_DIEQTY,
    COL_FAMILY,
    COL_FLAG,
    COL_INQTY,
    COL_MTECH,
    COL_OUTQTY,
    COL_STEPID,
    COL_TECH,
    COL_WEEK,
    COL_YIELD,
    COL_GOAL_AFTER,
    COL_GOAL_BEFORE,
    COL_GOAL_IN,
    COL_GOAL_NONE,
    COL_YIELD_TARGET,
)
from .goal_loader import GoalTable
from .rule_loader import FamilyFilter, ProductRule, RuleSet
from .week_calc import PeriodSelection

VOLUME_THRESHOLD_K = 1.0  # 单位 K, 1K = 1000 颗


@dataclass
class ReportCell:
    """单元格的聚合结果."""

    yield_value: float | None = None  # 0~1 小数
    inqty: float = 0.0
    outqty: float = 0.0

    @property
    def volume_k(self) -> float:
        return self.inqty / 1000.0


@dataclass
class ReportRow:
    """一行: 一个 product table 内部按 (mtech, dieqty, step) 聚合后的结果.

    Rule_list 第一列合并的多个 family 在这里 **跨 family 累加** 后表现为同一行。
    """

    tech: str
    mtech: str
    dieqty: str
    step_code: str
    config: str  # dieqty 的展示, e.g. ``1D``
    assy: str
    goal: float | None
    s_goal: float | None
    group_label: str  # A 列展示, 优先用 SRC 中的 A 列字符串; 否则 fallback
    cells: dict[str, ReportCell] = field(default_factory=dict)
    last_volume_k: float | None = None
    goal_missing: bool = False  # 提供了 SRC 但没在 SRC 中找到 (Goal/S-Goal=None)
    families: tuple[str, ...] = ()  # 该行实际汇总到的 family 列表 (调试用)

    def display_label(self) -> str:
        """第一列显示."""
        if self.group_label:
            return self.group_label
        parts: list[str] = []
        tech = str(self.tech or "").strip()
        if tech and tech.lower() != "nan":
            parts.append(tech)
        mtech = str(self.mtech or "").strip()
        if mtech and mtech.lower() != "nan":
            parts.append(mtech)
        if self.step_code:
            parts.append(self.step_code)
        return " ".join(p for p in parts if p) or "-"


@dataclass
class ProductReport:
    """一个 product table 的完整内容."""

    name: str
    rows: list[ReportRow]
    periods: PeriodSelection


# --------------------------- 内部工具 ---------------------------


def _filter_for_product(df: pd.DataFrame, rule: ProductRule) -> pd.DataFrame:
    """根据 ProductRule 中的多个 FamilyFilter 取并集."""
    parts: list[pd.DataFrame] = []
    for f in rule.filters:
        sub = df[df[COL_FAMILY] == f.family]
        if sub.empty:
            continue
        if f.dieqty:
            sub = sub[sub[COL_DIEQTY].isin(f.dieqty)]
        if f.mtech:
            sub = sub[sub[COL_MTECH].isin(f.mtech)]
        if f.devices:
            sub = sub[sub[COL_DEVICE].isin(f.devices)]
        if f.stepids:
            sub = sub[sub[COL_STEPID].astype(str).isin(f.stepids)]
        parts.append(sub)
    if not parts:
        return df.iloc[0:0]
    return pd.concat(parts, ignore_index=True)


def _aggregate_cell(group: pd.DataFrame) -> ReportCell:
    """对同一 (period, family, dieqty, step) 的多行做加权聚合."""
    inqty = group[COL_INQTY].fillna(0).sum()
    outqty = group[COL_OUTQTY].fillna(0).sum()
    if inqty <= 0:
        return ReportCell(yield_value=None, inqty=0.0, outqty=outqty)
    y = outqty / inqty
    return ReportCell(yield_value=float(y), inqty=float(inqty), outqty=float(outqty))


def _pick_goals(group: pd.DataFrame) -> tuple[float | None, float | None]:
    """从一组行里挑 Goal / S-Goal。

    取最常见的非空值。Goal 用 GOAL_IN / GOAL_BEFORE / GOAL_AFTER / GOAL_NONE 的
    首个非空(因为各 W/M/Q 行可能放在不同列里)。S-Goal 用 YIELD_TARGET。
    """
    if group.empty:
        return None, None

    def first_valid_mode(series: pd.Series) -> float | None:
        s = series.dropna()
        if s.empty:
            return None
        return float(s.mode().iloc[0])

    candidates = [
        COL_GOAL_BEFORE,
        COL_GOAL_IN,
        COL_GOAL_AFTER,
        COL_GOAL_NONE,
    ]
    goal: float | None = None
    for col in candidates:
        if col in group.columns:
            v = first_valid_mode(group[col])
            if v is not None:
                goal = v
                break
    s_goal: float | None = None
    if COL_YIELD_TARGET in group.columns:
        s_goal = first_valid_mode(group[COL_YIELD_TARGET])
    return goal, s_goal


def _row_key(filter_: FamilyFilter, df_row: pd.Series) -> tuple:
    """决定行级聚合的 key。

    默认按 family + dieqty + tech + stepid 分组；若 filter 指定 device 列表，
    则把 device 也作为 key 的一部分(避免不同 device 混在一起)。"""
    base = (
        df_row[COL_FAMILY],
        df_row[COL_TECH],
        str(df_row[COL_DIEQTY]),
        str(df_row[COL_STEPID]),
    )
    return base


def _select_period_data(
    df: pd.DataFrame, label: str, flags: tuple[str, ...]
) -> pd.DataFrame:
    return df[(df[COL_WEEK] == label) & (df[COL_FLAG].isin(flags))]


# --------------------------- 主流程 ---------------------------


def build_product_report(
    df: pd.DataFrame,
    rule: ProductRule,
    rules: RuleSet,
    periods: PeriodSelection,
    *,
    goals: GoalTable | None = None,
) -> ProductReport:
    """根据 product 规则、period 选择，构建一个 ProductReport.

    **跨 family 聚合**: Rule_list 第一列合并(``Macaw`` 关联 4 个 INAND-MACAW-*
    family) 表示这几个 family 在该 product table 中合并展示。这里通过 **不在
    分组 key 中放 FAMILY** 实现 - 同 (TECH, MTECH, DIEQTY, STEPID) 跨 family
    的 LT_INQTY/LT_OUTQTY 自动累加, yield = sum(out)/sum(in).

    Args:
        df: 原始数据.
        rule: product 规则.
        rules: 全局规则集 (用来取 step_code).
        periods: period 选择结果.
        goals: 可选, ``bachmark SRC.xlsx`` 解析出的 Goal/S-Goal 参考表;
            若提供则优先用它, 找不到时记 ``goal_missing=True``.
    """
    sub = _filter_for_product(df, rule)
    rows: list[ReportRow] = []

    if sub.empty:
        return ProductReport(name=rule.name, rows=rows, periods=periods)

    keys = (
        sub[[COL_MTECH, COL_DIEQTY, COL_STEPID]]
        .drop_duplicates()
        .sort_values([COL_MTECH, COL_DIEQTY, COL_STEPID])
    )

    period_label_to_flag = {periods.quarter: ("Q",), periods.month: ("M",)}
    for w in periods.weeks_window:
        period_label_to_flag[w] = ("W",)

    for _, key_row in keys.iterrows():
        dieqty = str(key_row[COL_DIEQTY])
        mtech = str(key_row[COL_MTECH] or "").strip()
        stepid = str(key_row[COL_STEPID])
        step_code = rules.step_code(stepid)

        rgroup = sub[
            (sub[COL_DIEQTY].astype(str) == dieqty)
            & (sub[COL_MTECH].astype(str) == mtech)
            & (sub[COL_STEPID].astype(str) == stepid)
        ]
        families = tuple(sorted(rgroup[COL_FAMILY].astype(str).unique()))
        tech_set = sorted({str(t).strip() for t in rgroup[COL_TECH].dropna() if str(t).strip()})
        tech = tech_set[0] if tech_set else ""

        goal_missing = False
        descriptor = ""
        if goals is not None:
            entry = goals.lookup(rule.name, step_code, dieqty, mtech or None)
            if entry.found:
                goal, s_goal, descriptor = entry.goal, entry.s_goal, entry.descriptor
            else:
                goal, s_goal = None, None
                goal_missing = True
        else:
            goal, s_goal = _pick_goals(rgroup)

        cells: dict[str, ReportCell] = {}
        for label, flags in period_label_to_flag.items():
            chunk = _select_period_data(rgroup, label, flags)
            cells[label] = _aggregate_cell(chunk)

        last_cell = cells.get(periods.prev_week)
        if last_cell is None:
            last_cell = cells.get(periods.weeks_window[-1])

        if descriptor:
            group_label = descriptor
        else:
            parts = [p for p in (tech, mtech, step_code) if p]
            group_label = " ".join(parts) or rule.name

        report_row = ReportRow(
            tech=tech,
            mtech=mtech,
            dieqty=dieqty,
            step_code=step_code,
            config=f"{dieqty}D" if dieqty else "",
            assy="SDSS",
            goal=goal,
            s_goal=s_goal,
            group_label=group_label,
            cells=cells,
            last_volume_k=last_cell.volume_k if last_cell else None,
            goal_missing=goal_missing,
            families=families,
        )
        rows.append(report_row)

    rows.sort(key=lambda r: (r.group_label, _die_sort_key(r.dieqty)))
    return ProductReport(name=rule.name, rows=rows, periods=periods)


def _die_sort_key(d: str) -> int:
    try:
        return int(str(d).strip())
    except (TypeError, ValueError):
        return 999


def build_all_reports(
    df: pd.DataFrame,
    rules: RuleSet,
    periods: PeriodSelection,
    *,
    goals: GoalTable | None = None,
) -> list[ProductReport]:
    return [
        build_product_report(df, p, rules, periods, goals=goals) for p in rules.products
    ]


def yield_or_none(
    cell: ReportCell | None, *, threshold_k: float = VOLUME_THRESHOLD_K
) -> float | None:
    """根据 volume 门槛决定要不要返回 yield."""
    if cell is None or cell.volume_k < threshold_k:
        return None
    return cell.yield_value


def format_volume(cell: ReportCell | None, *, threshold_k: float = VOLUME_THRESHOLD_K) -> object:
    """最后一列 Volume(K) 的展示."""
    if cell is None or cell.volume_k < threshold_k:
        return "No volume"
    return round(cell.volume_k, 1)
