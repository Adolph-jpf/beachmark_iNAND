"""Spotfire 原始数据加载模块.

负责把 Tab 分隔的 PN_TEST_STEPID_YIELD.txt 解析成 ``pandas.DataFrame``。

设计要点
--------
* 文件可能有 BOM、不同换行符、若空字段表示缺失（用 ``""`` / ``"\\N"``）。
* 数值列(LT_INQTY/Volume, LT_OUTQTY, LT_YIELD, GOAL_*) 会被强制成 float，
  非法值变 ``NaN``。
* PYDATE / STARTTIME / ENDTIME 解析成 ``pandas.Timestamp``。
* 列名里有 ``LT_INQTY/Volume`` 这种带斜杠的名字，pandas 也能正常处理，但代码内部
  我们会通过常量引用，避免拼写错误。
"""

from __future__ import annotations

from pathlib import Path
import pandas as pd

COL_PYDATE = "PYDATE"
COL_WEEK = "WEEK"
COL_STARTTIME = "STARTTIME"
COL_ENDTIME = "ENDTIME"
COL_FLAG = "FLAG"
COL_FAMILY = "FAMILY"
COL_TECH = "TECH"
COL_DIEQTY = "DIEQTY"
COL_MTECH = "MTECH"
COL_DEVICE = "DEVICE"
COL_STEPID = "STEPID_YIELD"
COL_INQTY = "LT_INQTY/Volume"
COL_OUTQTY = "LT_OUTQTY"
COL_YIELD = "LT_YIELD"
COL_GOAL_AFTER = "GOAL_AFTER"
COL_GOAL_BEFORE = "GOAL_BEFORE"
COL_GOAL_IN = "GOAL_IN"
COL_GOAL_NONE = "GOAL_NONE"
COL_YIELD_TARGET = "YIELD_TARGET"
COL_VOLUME_TARGET = "VOLUMN_TARGET"

NUMERIC_COLS = (
    COL_INQTY,
    COL_OUTQTY,
    COL_YIELD,
    COL_GOAL_AFTER,
    COL_GOAL_BEFORE,
    COL_GOAL_IN,
    COL_GOAL_NONE,
    COL_YIELD_TARGET,
    COL_VOLUME_TARGET,
)

DATETIME_COLS = (COL_PYDATE, COL_STARTTIME, COL_ENDTIME)

# CSV 版本里字段名和 txt 略有差异，统一映射回内部标准列名。
_COLUMN_ALIASES = {
    "LT_INQTY": COL_INQTY,  # CSV 常见列名
}


def _parse_percent(value: object) -> float | None:
    """允许 99.80% / 99.8 / 0.998 这类输入，统一返回 0~1 的小数."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip()
    if not s or s.upper() in {"NA", "NULL", "\\N"}:
        return None
    if s.endswith("%"):
        try:
            return float(s.rstrip("%")) / 100.0
        except ValueError:
            return None
    try:
        v = float(s)
    except ValueError:
        return None
    return v / 100.0 if v > 1.5 else v


def _sniff_encoding(path: Path) -> str:
    """根据 BOM 推测文件编码; 默认 utf-8-sig."""
    with path.open("rb") as fh:
        head = fh.read(4)
    if head.startswith(b"\xff\xfe"):
        return "utf-16"
    if head.startswith(b"\xfe\xff"):
        return "utf-16"
    if head.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig"
    return "utf-8-sig"


def resolve_data_source(path: str | Path) -> Path:
    """解析原始数据路径, 同名 ``.csv`` 优先于 ``.txt``.

    规则:
    * 当入参是 ``foo.txt`` 且同目录存在 ``foo.csv`` -> 选择 ``foo.csv``.
    * 当入参是 ``foo.csv`` 且存在 -> 选择 ``foo.csv``.
    * 若优先项不存在, 回退到另一种后缀 (``.txt``/``.csv``)。
    * 入参无后缀时, 依次尝试 ``.csv`` -> ``.txt`` -> 原路径。
    """
    p = Path(path)
    suffix = p.suffix.lower()

    if suffix in {".txt", ".csv"}:
        csv_path = p.with_suffix(".csv")
        txt_path = p.with_suffix(".txt")
        if csv_path.exists():
            return csv_path
        if txt_path.exists():
            return txt_path
        if p.exists():
            return p
        raise FileNotFoundError(f"原始数据文件不存在: {p}")

    csv_path = p.with_suffix(".csv")
    txt_path = p.with_suffix(".txt")
    if csv_path.exists():
        return csv_path
    if txt_path.exists():
        return txt_path
    if p.exists():
        return p
    raise FileNotFoundError(f"原始数据文件不存在: {p}")


def _pick_separator(path: Path) -> str:
    """按后缀选分隔符: ``.csv`` 用逗号, 其它默认 tab."""
    return "," if path.suffix.lower() == ".csv" else "\t"


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map: dict[str, str] = {}
    existing = set(df.columns)
    for src, dst in _COLUMN_ALIASES.items():
        if src in existing and dst not in existing:
            rename_map[src] = dst
    if rename_map:
        df = df.rename(columns=rename_map)
    return df


def load_yield_data(path: str | Path) -> pd.DataFrame:
    """加载 Spotfire 导出的原始 yield 文件。

    Args:
        path: txt 文件路径。

    Returns:
        DataFrame，列与文件保持一致；额外保证：
        * ``DIEQTY`` 是字符串(可能是 1/2/4/8 等)。
        * 数值列已经是 float。
        * datetime 列已经是 ``Timestamp``。

    Spotfire 导出的 txt 通常是 UTF-16 LE with BOM (0xFF 0xFE); 我们也兼容
    UTF-8 with/without BOM。
    """
    path = resolve_data_source(path)

    encoding = _sniff_encoding(path)
    sep = _pick_separator(path)

    try:
        df = pd.read_csv(
            path,
            sep=sep,
            dtype=str,
            keep_default_na=False,
            na_values=["", "NA", "\\N"],
            encoding=encoding,
        )
    except UnicodeDecodeError:
        for fallback in ("utf-16", "utf-8-sig", "gbk", "latin-1"):
            if fallback == encoding:
                continue
            try:
                df = pd.read_csv(
                    path,
                    sep=sep,
                    dtype=str,
                    keep_default_na=False,
                    na_values=["", "NA", "\\N"],
                    encoding=fallback,
                )
                break
            except UnicodeDecodeError:
                continue
        else:
            raise

    df.columns = [c.strip() for c in df.columns]
    df = _normalize_columns(df)

    for col in NUMERIC_COLS:
        if col not in df.columns:
            continue
        if col in {
            COL_GOAL_AFTER,
            COL_GOAL_BEFORE,
            COL_GOAL_IN,
            COL_GOAL_NONE,
            COL_YIELD_TARGET,
            COL_YIELD,
        }:
            df[col] = df[col].map(_parse_percent)
        else:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    for col in DATETIME_COLS:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    if COL_DIEQTY in df.columns:
        df[COL_DIEQTY] = df[COL_DIEQTY].fillna("").astype(str).str.strip()

    for col in (COL_FAMILY, COL_MTECH, COL_DEVICE, COL_STEPID, COL_FLAG, COL_WEEK, COL_TECH):
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()

    return df


def validate_yield(
    df: pd.DataFrame, *, tolerance: float = 1e-3
) -> pd.DataFrame:
    """校验 LT_YIELD == LT_OUTQTY / LT_INQTY/Volume.

    Args:
        df: ``load_yield_data`` 的输出。
        tolerance: 允许的相对误差，默认 0.1%。

    Returns:
        新增 ``YIELD_CALC`` 与 ``YIELD_OK`` 两列的 DataFrame 副本。
    """
    out = df.copy()
    inqty = out[COL_INQTY]
    outqty = out[COL_OUTQTY]

    safe_in = inqty.where(inqty.gt(0))
    calc = outqty.where(safe_in.notna()) / safe_in
    out["YIELD_CALC"] = calc
    diff = (out[COL_YIELD] - calc).abs()
    out["YIELD_OK"] = diff.le(tolerance) | out[COL_YIELD].isna() | calc.isna()
    return out
