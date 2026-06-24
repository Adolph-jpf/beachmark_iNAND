"""命令行入口: 一条命令完成 Spotfire 数据 -> Excel 报表."""

from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime
from pathlib import Path

from .config import (
    DEFAULT_DATA_PATH,
    DEFAULT_OUTPUT_PATH,
    DEFAULT_RULE_PATH,
    DEFAULT_SRC_GOALS_PATH,
)
from .data_loader import load_yield_data, resolve_data_source, validate_yield
from .excel_writer import write_excel
from .goal_loader import GoalTable, load_goal_table
from .processor import build_all_reports
from .rule_loader import load_rules
from .week_calc import select_periods


def _setup_logging(verbose: bool, log_file: Path | None = None) -> None:
    handlers: list[logging.Handler] = [logging.StreamHandler(sys.stdout)]
    if log_file is not None:
        log_file.parent.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.FileHandler(log_file, encoding="utf-8"))
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s | %(levelname)-7s | %(name)s | %(message)s",
        handlers=handlers,
    )


def run(
    *,
    data_path: Path,
    rule_path: Path,
    output_path: Path,
    src_goals_path: Path | None = None,
    today: datetime | None = None,
    weeks_count: int = 9,
    log_file: Path | None = None,
) -> dict:
    """主流程, 返回元数据字典 (供 GUI 显示).

    Raises:
        ValueError: 数据/规则非法。
        FileNotFoundError: 输入文件缺失。
    """
    log = logging.getLogger("benchmark")
    resolved_data_path = resolve_data_source(data_path)
    if resolved_data_path != data_path:
        log.info("Loading raw data: %s (auto-selected from %s)", resolved_data_path, data_path)
    else:
        log.info("Loading raw data: %s", resolved_data_path)
    df = load_yield_data(resolved_data_path)
    log.info("Loaded %d rows, columns: %d", len(df), len(df.columns))

    df_validated = validate_yield(df)
    bad = (~df_validated["YIELD_OK"]).sum()
    if bad:
        log.warning("LT_YIELD 与 LT_OUTQTY/LT_INQTY 校验失败 %d 行 (容差 1e-3)", int(bad))
    else:
        log.info("LT_YIELD 校验通过 (容差 1e-3)")

    log.info("Loading rules: %s", rule_path)
    rules = load_rules(rule_path)
    log.info("Loaded %d product rule(s)", len(rules.products))

    goals: GoalTable | None = None
    if src_goals_path is not None:
        log.info("Loading SRC goal table: %s", src_goals_path)
        goals = load_goal_table(src_goals_path)
        log.info(
            "Loaded SRC tables: %s",
            ", ".join(sorted(goals.known_tables())) or "<none>",
        )
    else:
        log.info("未提供 --src-goals, 使用原始数据列里的 Goal/S-Goal")

    today_str = today.strftime("%Y-%m-%d") if isinstance(today, datetime) else "now"
    log.info("Selecting fiscal periods (today=%s)", today_str)
    periods = select_periods(df, today, weeks_count=weeks_count)
    weeks_disp = (
        f"{periods.weeks_window[0]}..{periods.weeks_window[-1]} "
        f"({len(periods.weeks_window)} weeks)"
        if periods.weeks_window
        else "<none>"
    )
    log.info(
        "推算结果: 基准日期=%s -> 当前周=%s, 上一周=%s, 9周窗口=%s, 上一月=%s, 上一季=%s",
        today_str,
        periods.current_week,
        periods.prev_week,
        weeks_disp,
        periods.month,
        periods.quarter,
    )

    reports = build_all_reports(df, rules, periods, goals=goals)
    log.info("Built %d product reports", len(reports))
    total_missing = 0
    for r in reports:
        miss = sum(1 for row in r.rows if row.goal_missing)
        total_missing += miss
        log.info("  - %-22s rows=%-3d missing_goal=%d", r.name, len(r.rows), miss)
        for row in r.rows:
            if row.goal_missing:
                log.warning(
                    "    [missing goal] %s | families=%s | step=%s | mtech=%s | %sD",
                    r.name,
                    "/".join(row.families) or "<none>",
                    row.step_code,
                    row.mtech or "<none>",
                    row.dieqty,
                )
    if goals is not None:
        log.info(
            "共 %d 行的 Goal/S-Goal 在 SRC 中未找到 (yield 单元格用绿底黄字标识)",
            total_missing,
        )

    out = write_excel(reports, output_path)
    log.info("Saved Excel: %s", out)

    return {
        "data": str(resolved_data_path),
        "data_requested": str(data_path),
        "rules": str(rule_path),
        "src_goals": str(src_goals_path) if src_goals_path else None,
        "output": str(out),
        "output_actual": str(out),  # 真正写入的路径 (锁占用时可能加了时间戳)
        "rows": len(df),
        "yield_validation_failures": int(bad),
        "missing_goal_rows": total_missing,
        "products": [(r.name, len(r.rows)) for r in reports],
        "periods": {
            "current_week": periods.current_week,
            "prev_week": periods.prev_week,
            "month": periods.month,
            "quarter": periods.quarter,
            "weeks_window": list(periods.weeks_window),
        },
    }


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="benchmark",
        description="Spotfire 原始数据 -> INAND weekly benchmark Excel",
    )
    p.add_argument(
        "--data",
        default=DEFAULT_DATA_PATH,
        help="原始数据 .txt/.csv (同名 .csv 会优先)",
    )
    p.add_argument("--rules", default=DEFAULT_RULE_PATH, help="规则 .xlsx")
    p.add_argument(
        "--src-goals",
        default=DEFAULT_SRC_GOALS_PATH,
        help="Goal/S-Goal 参考 SRC .xlsx; 留空字符串则不使用",
    )
    p.add_argument(
        "--output",
        default=DEFAULT_OUTPUT_PATH,
        help="输出 Excel 路径",
    )
    p.add_argument(
        "--today",
        default=None,
        help="指定 today (YYYY-MM-DD), 默认本机当前时间",
    )
    p.add_argument("--weeks", type=int, default=9, help="window 中包含的 week 数 (默认 9)")
    p.add_argument("--log-file", default="logs/run.log", help="日志文件路径")
    p.add_argument("-v", "--verbose", action="store_true")
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    today = (
        datetime.strptime(args.today, "%Y-%m-%d") if args.today else None
    )
    log_file = Path(args.log_file) if args.log_file else None
    _setup_logging(args.verbose, log_file)
    src_goals_path: Path | None = None
    if args.src_goals:
        candidate = Path(args.src_goals)
        if candidate.exists():
            src_goals_path = candidate
        else:
            logging.getLogger("benchmark").warning(
                "SRC goal 文件 %s 不存在; fallback 用原始数据中的 Goal", candidate
            )
    try:
        run(
            data_path=Path(args.data),
            rule_path=Path(args.rules),
            output_path=Path(args.output),
            src_goals_path=src_goals_path,
            today=today,
            weeks_count=args.weeks,
            log_file=log_file,
        )
    except Exception as exc:  # pragma: no cover - 仅打印
        logging.getLogger("benchmark").exception("运行失败: %s", exc)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
