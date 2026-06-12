from __future__ import annotations

import importlib.util
import math
import re
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parent
FINAL_TABLE_DIR = ROOT / "final_valuable_dataset" / "tables"
OUT_DIR = ROOT / "paper_diversity_figures" / "tables"
OUT_XLSX = OUT_DIR / "origin_top8_property_distribution_data_range_filtered.xlsx"


# Conservative ranges for superalloy-related records after unit standardization.
# These are intended to remove impossible extraction/unit-conversion errors while
# keeping broad high-temperature alloy variability.
PROPERTY_RANGES = {
    ("yield strength", "MPa"): (1, 3000),
    ("tensile strength", "MPa"): (1, 3500),
    ("compressive strength", "MPa"): (1, 5000),
    ("total elongation", "%"): (0, 200),
    ("strain rate", "s^-1"): (1e-8, 1e4),
    ("youngs modulus", "GPa"): (20, 350),
    ("creep life", "h"): (1e-4, 1e5),
    ("grain size", "um"): (1e-4, 1e5),
}


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
        .agg(data_count_before_range_filter=("record_id", "count"), sample_record_count_before=("record_id", "nunique"))
        .reset_index()
        .sort_values(["性能", "data_count_before_range_filter"], ascending=[True, False])
    )
    selected_units = unit_counts.groupby("性能", as_index=False).first()
    selected_units = selected_units.rename(
        columns={
            "性能": "property",
            "标准单位": "selected_standard_unit",
        }
    )
    top8 = selected_units.sort_values("data_count_before_range_filter", ascending=False).head(8).copy()
    top8.insert(0, "rank", range(1, len(top8) + 1))
    selected_pairs = set(zip(top8["property"], top8["selected_standard_unit"]))

    selected = valid[valid.apply(lambda r: (r["性能"], r["标准单位"]) in selected_pairs, axis=1)].copy()
    selected["property"] = selected["性能"]
    selected["standard_unit"] = selected["标准单位"]
    selected["range_min"] = selected.apply(
        lambda r: PROPERTY_RANGES.get((r["property"], r["standard_unit"]), (float("-inf"), float("inf")))[0],
        axis=1,
    )
    selected["range_max"] = selected.apply(
        lambda r: PROPERTY_RANGES.get((r["property"], r["standard_unit"]), (float("-inf"), float("inf")))[1],
        axis=1,
    )
    selected["range_filter_pass"] = selected.apply(
        lambda r: r["range_min"] <= float(r["standard_value"]) <= r["range_max"],
        axis=1,
    )
    selected["range_filter_reason"] = selected.apply(
        lambda r: "pass"
        if r["range_filter_pass"]
        else f"outside_common_range_[{r['range_min']}, {r['range_max']}]",
        axis=1,
    )

    filtered = selected[selected["range_filter_pass"]].copy()
    outliers = selected[~selected["range_filter_pass"]].copy()

    origin_long_cols = [
        "property",
        "standard_unit",
        "standard_value",
        "range_min",
        "range_max",
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
        if col not in filtered.columns:
            filtered[col] = ""
        if col not in outliers.columns:
            outliers[col] = ""
    origin_long = filtered[origin_long_cols].sort_values(["property", "standard_value"]).reset_index(drop=True)
    outlier_cols = origin_long_cols + ["range_filter_reason"]
    outliers_out = outliers[outlier_cols].sort_values(["property", "standard_value"]).reset_index(drop=True)

    origin_wide = make_wide(origin_long, top8)

    stats = (
        origin_long.groupby(["property", "standard_unit"], dropna=False)
        .agg(
            data_count_after_range_filter=("standard_value", "count"),
            min_value=("standard_value", "min"),
            max_value=("standard_value", "max"),
            mean_value=("standard_value", "mean"),
            std_value=("standard_value", "std"),
        )
        .reset_index()
    )
    stats["std_value"] = stats["std_value"].fillna(0)
    outlier_counts = (
        outliers_out.groupby(["property", "standard_unit"], dropna=False)
        .size()
        .reset_index(name="removed_by_range_filter")
    )
    top8 = top8.merge(
        stats.rename(columns={"standard_unit": "selected_standard_unit"}),
        on=["property", "selected_standard_unit"],
        how="left",
    ).merge(
        outlier_counts.rename(columns={"standard_unit": "selected_standard_unit"}),
        on=["property", "selected_standard_unit"],
        how="left",
    )
    top8["removed_by_range_filter"] = top8["removed_by_range_filter"].fillna(0).astype(int)
    top8["retention_after_range_filter_percent"] = (
        top8["data_count_after_range_filter"] / top8["data_count_before_range_filter"] * 100
    ).round(2)
    top8["range_rule"] = top8.apply(
        lambda r: str(PROPERTY_RANGES.get((r["property"], r["selected_standard_unit"]), "")),
        axis=1,
    )

    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
        top8.to_excel(writer, index=False, sheet_name="summary_top8_range")
        origin_long.to_excel(writer, index=False, sheet_name="origin_long_filtered")
        origin_wide.to_excel(writer, index=False, sheet_name="origin_wide_filtered")
        outliers_out.to_excel(writer, index=False, sheet_name="removed_outliers")

        used = {"summary_top8_range", "origin_long_filtered", "origin_wide_filtered", "removed_outliers"}
        for _, row in top8.iterrows():
            prop = row["property"]
            unit = row["selected_standard_unit"]
            sheet = clean_sheet_name(f"{int(row['rank'])}_{prop}", used)
            df = origin_long[(origin_long["property"] == prop) & (origin_long["standard_unit"] == unit)].copy()
            df.to_excel(writer, index=False, sheet_name=sheet)

        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

    top8.to_csv(OUT_DIR / "origin_top8_property_distribution_summary_range_filtered.csv", index=False, encoding="utf-8-sig")
    origin_long.to_csv(OUT_DIR / "origin_top8_property_distribution_long_range_filtered.csv", index=False, encoding="utf-8-sig")
    origin_wide.to_csv(OUT_DIR / "origin_top8_property_distribution_wide_range_filtered.csv", index=False, encoding="utf-8-sig")
    outliers_out.to_csv(OUT_DIR / "origin_top8_property_distribution_removed_outliers.csv", index=False, encoding="utf-8-sig")
    print(f"Wrote {OUT_XLSX}")
    print(top8.to_string(index=False))


if __name__ == "__main__":
    main()
