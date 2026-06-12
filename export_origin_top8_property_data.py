from __future__ import annotations

import importlib.util
import math
import re
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parent
FINAL_TABLE_DIR = ROOT / "final_valuable_dataset" / "tables"
OUT_DIR = ROOT / "paper_diversity_figures" / "tables"
OUT_XLSX = OUT_DIR / "origin_top8_property_distribution_data.xlsx"


def load_unit_filter_module():
    module_path = ROOT / "filter_valid_property_units.py"
    spec = importlib.util.spec_from_file_location("filter_valid_property_units", module_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load {module_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def read_many(folder: Path, pattern: str) -> pd.DataFrame:
    frames = []
    for path in sorted(folder.glob(pattern)):
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
        df["source_table_file"] = path.name
        frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def clean_sheet_name(name: str, used: set[str]) -> str:
    safe = re.sub(r"[\[\]\:\*\?\/\\]", "_", name).strip() or "sheet"
    safe = safe[:31]
    base = safe
    i = 1
    while safe in used:
        suffix = f"_{i}"
        safe = (base[: 31 - len(suffix)] + suffix)[:31]
        i += 1
    used.add(safe)
    return safe


def make_wide(origin_long: pd.DataFrame, summary: pd.DataFrame) -> pd.DataFrame:
    columns = {}
    max_len = 0
    for _, row in summary.iterrows():
        prop = row["property"]
        unit = row["selected_standard_unit"]
        label = f"{prop} ({unit})"
        values = origin_long.loc[
            (origin_long["property"] == prop) & (origin_long["standard_unit"] == unit),
            "standard_value",
        ].astype(float).reset_index(drop=True)
        columns[label] = values
        max_len = max(max_len, len(values))
    wide = pd.DataFrame(index=range(max_len))
    for label, values in columns.items():
        wide[label] = pd.Series(values)
    return wide


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    unit_filter = load_unit_filter_module()

    final_long = read_many(FINAL_TABLE_DIR, "final_ext-*_valuable_properties_long.csv")
    annotated = unit_filter.annotate_units(final_long, "final_valuable_dataset")
    annotated["standard_value"] = pd.to_numeric(annotated["property_standard_value"], errors="coerce")
    valid = annotated[
        (annotated["单位是否匹配性能"] == "yes")
        & (annotated["性能值"] != "")
        & (annotated["标准单位"] != "")
        & annotated["standard_value"].notna()
    ].copy()
    valid = valid[valid["standard_value"].map(lambda x: math.isfinite(float(x)))]
    valid = valid[~valid["性能"].isin(unit_filter.CATEGORICAL_NO_UNIT_PROPERTIES)]

    unit_counts = (
        valid.groupby(["性能", "标准单位"], dropna=False)
        .agg(data_count=("record_id", "count"), sample_record_count=("record_id", "nunique"))
        .reset_index()
        .sort_values(["性能", "data_count"], ascending=[True, False])
    )
    selected_units = unit_counts.groupby("性能", as_index=False).first()
    selected_units = selected_units.rename(
        columns={
            "性能": "property",
            "标准单位": "selected_standard_unit",
            "data_count": "selected_unit_data_count",
            "sample_record_count": "selected_unit_sample_record_count",
        }
    )
    top8 = selected_units.sort_values("selected_unit_data_count", ascending=False).head(8).copy()
    top8.insert(0, "rank", range(1, len(top8) + 1))

    selected_pairs = set(zip(top8["property"], top8["selected_standard_unit"]))
    origin_rows = valid[
        valid.apply(lambda r: (r["性能"], r["标准单位"]) in selected_pairs, axis=1)
    ].copy()
    origin_rows["property"] = origin_rows["性能"]
    origin_rows["standard_unit"] = origin_rows["标准单位"]
    origin_rows["standard_value"] = origin_rows["standard_value"].astype(float)

    origin_long_cols = [
        "property",
        "standard_unit",
        "standard_value",
        "property_value_raw",
        "property_unit_raw",
        "doi_or_file_id",
        "source_file",
        "source_row",
        "record_id",
        "superalloy_name",
        "sample_name",
        "composition_unit_raw",
        "composition_present_element_count",
        "composition_sum_numeric",
        "distinguishing_factor",
        "synthesis_and_processing_routes",
        "test_route_condition",
        "property_name_raw",
        "property_sourced_figure",
        "property_flags",
        "unit_conversion_note",
    ]
    for col in origin_long_cols:
        if col not in origin_rows.columns:
            origin_rows[col] = ""
    origin_long = origin_rows[origin_long_cols].sort_values(["property", "standard_value"]).reset_index(drop=True)
    origin_wide = make_wide(origin_long, top8)

    stats = (
        origin_long.groupby(["property", "standard_unit"], dropna=False)
        .agg(
            data_count=("standard_value", "count"),
            min_value=("standard_value", "min"),
            max_value=("standard_value", "max"),
            mean_value=("standard_value", "mean"),
            std_value=("standard_value", "std"),
        )
        .reset_index()
    )
    stats["std_value"] = stats["std_value"].fillna(0)
    top8 = top8.merge(
        stats.rename(columns={"standard_unit": "selected_standard_unit"}),
        on=["property", "selected_standard_unit"],
        how="left",
    )

    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
        top8.to_excel(writer, index=False, sheet_name="summary_top8")
        origin_long.to_excel(writer, index=False, sheet_name="origin_long")
        origin_wide.to_excel(writer, index=False, sheet_name="origin_wide")

        used = {"summary_top8", "origin_long", "origin_wide"}
        for _, row in top8.iterrows():
            prop = row["property"]
            unit = row["selected_standard_unit"]
            sheet = clean_sheet_name(f"{int(row['rank'])}_{prop}", used)
            df = origin_long[(origin_long["property"] == prop) & (origin_long["standard_unit"] == unit)].copy()
            df.to_excel(writer, index=False, sheet_name=sheet)

        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

    top8.to_csv(OUT_DIR / "origin_top8_property_distribution_summary.csv", index=False, encoding="utf-8-sig")
    origin_long.to_csv(OUT_DIR / "origin_top8_property_distribution_long.csv", index=False, encoding="utf-8-sig")
    origin_wide.to_csv(OUT_DIR / "origin_top8_property_distribution_wide.csv", index=False, encoding="utf-8-sig")
    print(f"Wrote {OUT_XLSX}")
    print(top8.to_string(index=False))


if __name__ == "__main__":
    main()
