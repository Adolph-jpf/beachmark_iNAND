"""校验 Spotfire 导出文件是否可被 benchmark 流程消费."""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from benchmark.config import DEFAULT_DATA_PATH  # noqa: E402
from benchmark.data_loader import (  # noqa: E402
    COL_DEVICE,
    COL_DIEQTY,
    COL_ENDTIME,
    COL_FAMILY,
    COL_FLAG,
    COL_MTECH,
    COL_OUTQTY,
    COL_PYDATE,
    COL_STARTTIME,
    COL_STEPID,
    COL_TECH,
    COL_WEEK,
    COL_YIELD,
    load_yield_data,
    resolve_data_source,
)

REQUIRED_COLUMNS = {
    COL_PYDATE,
    COL_WEEK,
    COL_STARTTIME,
    COL_ENDTIME,
    COL_FLAG,
    COL_FAMILY,
    COL_TECH,
    COL_DIEQTY,
    COL_MTECH,
    COL_DEVICE,
    COL_STEPID,
    "LT_INQTY/Volume",  # 经过 loader 归一后应始终存在
    COL_OUTQTY,
    COL_YIELD,
}


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Validate Spotfire export contract")
    p.add_argument("--data", default=DEFAULT_DATA_PATH, help="输入数据路径")
    p.add_argument(
        "--max-age-hours",
        type=float,
        default=36.0,
        help="允许的最大文件时效（小时）",
    )
    p.add_argument(
        "--min-rows",
        type=int,
        default=100,
        help="最小行数门槛（开发环境可低一些）",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    src = resolve_data_source(Path(args.data))
    now = datetime.now(timezone.utc)
    mtime = datetime.fromtimestamp(src.stat().st_mtime, tz=timezone.utc)
    age_h = (now - mtime).total_seconds() / 3600.0
    if age_h > args.max_age_hours:
        print(
            f"[FAIL] file too old: {src} age={age_h:.2f}h > {args.max_age_hours:.2f}h",
            file=sys.stderr,
        )
        return 2

    df = load_yield_data(src)
    if len(df) < args.min_rows:
        print(
            f"[FAIL] too few rows: {len(df)} < {args.min_rows} ({src})",
            file=sys.stderr,
        )
        return 3

    missing = sorted(REQUIRED_COLUMNS - set(df.columns))
    if missing:
        print(f"[FAIL] missing columns: {missing}", file=sys.stderr)
        return 4

    print(f"[OK] source={src}")
    print(f"[OK] rows={len(df)} age_h={age_h:.2f}")
    print(f"[OK] required columns present={len(REQUIRED_COLUMNS)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
