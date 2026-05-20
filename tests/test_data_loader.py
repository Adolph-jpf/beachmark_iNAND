from __future__ import annotations

import pandas as pd
import pytest

from benchmark.data_loader import (
    COL_FLAG,
    COL_INQTY,
    COL_OUTQTY,
    COL_WEEK,
    COL_YIELD,
    load_yield_data,
    resolve_data_source,
    validate_yield,
)


def test_load_synth(synth_yield_text):
    df = load_yield_data(synth_yield_text)
    assert len(df) > 10
    assert {COL_WEEK, COL_FLAG, COL_INQTY, COL_OUTQTY, COL_YIELD} <= set(df.columns)
    assert pd.api.types.is_float_dtype(df[COL_YIELD])
    assert df[COL_YIELD].between(0, 1).all()


def test_validate_yield_passes(synth_yield_text):
    df = load_yield_data(synth_yield_text)
    out = validate_yield(df)
    assert out["YIELD_OK"].all()


def test_validate_detects_bad_yield(synth_yield_text):
    df = load_yield_data(synth_yield_text)
    df.loc[0, COL_YIELD] = 0.5
    out = validate_yield(df)
    assert not bool(out.loc[0, "YIELD_OK"])


def test_load_real_data(real_data_path):
    if not real_data_path.exists():
        pytest.skip("real data not present")
    df = load_yield_data(real_data_path)
    assert len(df) > 100
    assert df[COL_FLAG].isin(["W", "M", "Q"]).any()


def test_validate_real_data(real_data_path):
    if not real_data_path.exists():
        pytest.skip("real data not present")
    df = load_yield_data(real_data_path)
    out = validate_yield(df)
    bad_ratio = (~out["YIELD_OK"]).mean()
    assert bad_ratio < 0.05, f"超过 5% 行 yield 校验失败: {bad_ratio:.2%}"


def test_load_csv_compatible(tmp_path):
    csv_file = tmp_path / "demo.csv"
    csv_file.write_text(
        "\n".join(
            [
                "PYDATE,WEEK,STARTTIME,ENDTIME,FLAG,FAMILY,TECH,DIEQTY,MTECH,DEVICE,STEPID_YIELD,LT_INQTY/Volume,LT_OUTQTY,LT_YIELD,GOAL_BEFORE,YIELD_TARGET,VOLUMN_TARGET",
                "2026-05-10,2026FW45,2026-05-04 0:00,2026-05-10 23:59,W,INAND-X,1Znm,1,64Gb,DEV1,7405,2000,1990,99.5%,99%,99.5%,2000",
            ]
        ),
        encoding="utf-8",
    )
    df = load_yield_data(csv_file)
    assert len(df) == 1
    assert df.loc[0, COL_WEEK] == "2026FW45"
    assert df.loc[0, COL_YIELD] == pytest.approx(0.995)


def test_resolve_prefers_csv_same_basename(tmp_path):
    txt_file = tmp_path / "PN_TEST_STEPID_YIELD.txt"
    csv_file = tmp_path / "PN_TEST_STEPID_YIELD.csv"
    txt_file.write_text("x", encoding="utf-8")
    csv_file.write_text("x", encoding="utf-8")
    assert resolve_data_source(txt_file) == csv_file
