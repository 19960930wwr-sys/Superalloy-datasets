from __future__ import annotations

import math
from pathlib import Path
from typing import Any

import pandas as pd


ROOT = Path(__file__).resolve().parent
CLEAN_TABLE_DIR = ROOT / "tables"
FINAL_TABLE_DIR = ROOT / "final_valuable_dataset" / "tables"
OUT_DIR = ROOT / "paper_diversity_figures" / "tables"

EMPTY = {
    "",
    "none",
    "nan",
    "null",
    "n/a",
    "na",
    "not specified",
    "not available",
    "[]",
    "['none']",
    "['None']",
}


def clean(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() in EMPTY else text


def read_many(folder: Path, pattern: str) -> pd.DataFrame:
    frames = []
    for path in sorted(folder.glob(pattern)):
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
        df["source_table_file"] = path.name
        frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def value_distribution(df: pd.DataFrame, scope: str) -> pd.DataFrame:
    data = df.copy()
    data["性能"] = data["property_target_canonical"].map(clean)
    data["标准单位"] = data["property_standard_unit"].map(clean)
    data["标准化数值"] = pd.to_numeric(data["property_standard_value"], errors="coerce")
    data = data[(data["性能"] != "") & (data["标准单位"] != "") & data["标准化数值"].notna()]
    data = data[data["标准化数值"].map(lambda x: math.isfinite(float(x)))]
    grouped = (
        data.groupby(["性能", "标准单位"], dropna=False)
        .agg(
            数据量=("标准化数值", "count"),
            最小值=("标准化数值", "min"),
            最大值=("标准化数值", "max"),
            平均值=("标准化数值", "mean"),
            标准差=("标准化数值", "std"),
        )
        .reset_index()
    )
    grouped["标准差"] = grouped["标准差"].fillna(0)
    for col in ["最小值", "最大值", "平均值", "标准差"]:
        grouped[col] = grouped[col].astype(float).round(6)
    grouped["统计口径"] = scope
    return grouped[
        ["统计口径", "性能", "标准单位", "数据量", "最小值", "最大值", "平均值", "标准差"]
    ].sort_values(["性能", "数据量"], ascending=[True, False])


def raw_unit_frequency(df: pd.DataFrame, scope: str) -> pd.DataFrame:
    data = df.copy()
    data["性能"] = data["property_target_canonical"].map(clean)
    data["原始单位"] = data["property_unit_raw"].map(clean)
    data["性能值"] = data["property_value_raw"].map(clean)
    data["标准单位"] = data["property_standard_unit"].map(clean)
    data = data[(data["性能"] != "") & (data["性能值"] != "")]
    data["原始单位"] = data["原始单位"].replace("", "(missing unit)")
    grouped = (
        data.groupby(["性能", "原始单位"], dropna=False)
        .agg(
            频次=("record_id", "count"),
            涉及样品记录数=("record_id", "nunique"),
            映射后的标准单位=(
                "标准单位",
                lambda x: "; ".join(sorted({clean(v) for v in x if clean(v)})) or "(not standardized)",
            ),
        )
        .reset_index()
    )
    totals = grouped.groupby("性能")["频次"].transform("sum")
    grouped["占该性能比例(%)"] = (grouped["频次"] / totals * 100).round(2)
    grouped["统计口径"] = scope
    return grouped[
        ["统计口径", "性能", "原始单位", "频次", "占该性能比例(%)", "涉及样品记录数", "映射后的标准单位"]
    ].sort_values(["性能", "频次"], ascending=[True, False])


def standard_unit_frequency(df: pd.DataFrame, scope: str) -> pd.DataFrame:
    data = df.copy()
    data["性能"] = data["property_target_canonical"].map(clean)
    data["标准单位"] = data["property_standard_unit"].map(clean)
    data["性能值"] = data["property_value_raw"].map(clean)
    data = data[(data["性能"] != "") & (data["性能值"] != "")]
    data["标准单位"] = data["标准单位"].replace("", "(not standardized)")
    grouped = (
        data.groupby(["性能", "标准单位"], dropna=False)
        .agg(频次=("record_id", "count"), 涉及样品记录数=("record_id", "nunique"))
        .reset_index()
    )
    totals = grouped.groupby("性能")["频次"].transform("sum")
    grouped["占该性能比例(%)"] = (grouped["频次"] / totals * 100).round(2)
    grouped["统计口径"] = scope
    return grouped[
        ["统计口径", "性能", "标准单位", "频次", "占该性能比例(%)", "涉及样品记录数"]
    ].sort_values(["性能", "频次"], ascending=[True, False])


def unit_diversity_summary(raw_freq: pd.DataFrame, standard_freq: pd.DataFrame, scope: str) -> pd.DataFrame:
    raw = raw_freq[raw_freq["统计口径"] == scope]
    standard = standard_freq[standard_freq["统计口径"] == scope]
    raw_summary = (
        raw.groupby("性能")
        .agg(原始单位种类数=("原始单位", "nunique"), 原始单位总频次=("频次", "sum"))
        .reset_index()
    )
    standard_summary = (
        standard.groupby("性能")
        .agg(标准单位种类数=("标准单位", "nunique"), 标准单位总频次=("频次", "sum"))
        .reset_index()
    )
    output = raw_summary.merge(standard_summary, on="性能", how="outer").fillna(0)
    output["统计口径"] = scope
    return output[
        ["统计口径", "性能", "原始单位种类数", "标准单位种类数", "原始单位总频次", "标准单位总频次"]
    ]


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    all_long = read_many(CLEAN_TABLE_DIR, "ext-*_properties_long.csv")
    valuable_long = read_many(FINAL_TABLE_DIR, "final_ext-*_valuable_properties_long.csv")

    value_table = pd.concat(
        [
            value_distribution(all_long, "全部清洗后抽取结果"),
            value_distribution(valuable_long, "最终有价值数据"),
        ],
        ignore_index=True,
    )
    raw_unit_table = pd.concat(
        [
            raw_unit_frequency(all_long, "全部清洗后抽取结果"),
            raw_unit_frequency(valuable_long, "最终有价值数据"),
        ],
        ignore_index=True,
    )
    standard_unit_table = pd.concat(
        [
            standard_unit_frequency(all_long, "全部清洗后抽取结果"),
            standard_unit_frequency(valuable_long, "最终有价值数据"),
        ],
        ignore_index=True,
    )
    unit_summary = pd.concat(
        [
            unit_diversity_summary(raw_unit_table, standard_unit_table, "全部清洗后抽取结果"),
            unit_diversity_summary(raw_unit_table, standard_unit_table, "最终有价值数据"),
        ],
        ignore_index=True,
    ).sort_values(["统计口径", "原始单位种类数"], ascending=[True, False])

    value_table.to_csv(OUT_DIR / "property_value_distribution_statistics.csv", index=False, encoding="utf-8-sig")
    raw_unit_table.to_csv(OUT_DIR / "property_raw_unit_frequency_statistics.csv", index=False, encoding="utf-8-sig")
    standard_unit_table.to_csv(OUT_DIR / "property_standard_unit_frequency_statistics.csv", index=False, encoding="utf-8-sig")
    unit_summary.to_csv(OUT_DIR / "property_unit_diversity_summary.csv", index=False, encoding="utf-8-sig")

    with pd.ExcelWriter(OUT_DIR / "property_distribution_and_unit_statistics.xlsx", engine="openpyxl") as writer:
        value_table.to_excel(writer, index=False, sheet_name="value_distribution")
        raw_unit_table.to_excel(writer, index=False, sheet_name="raw_unit_frequency")
        standard_unit_table.to_excel(writer, index=False, sheet_name="standard_unit_frequency")
        unit_summary.to_excel(writer, index=False, sheet_name="unit_diversity_summary")
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

    print("Wrote property statistics tables to", OUT_DIR)


if __name__ == "__main__":
    main()
