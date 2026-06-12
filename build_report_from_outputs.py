from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd


OUT_DIR = Path(__file__).resolve().parent
TABLE_DIR = OUT_DIR / "tables"


def markdown_table(df: pd.DataFrame) -> str:
    if df.empty:
        return "_No rows._"
    cols = list(df.columns)
    lines = [" | ".join(cols), " | ".join(["---"] * len(cols))]
    for _, row in df.iterrows():
        lines.append(" | ".join(str(row[col]) for col in cols))
    return "\n".join(lines)


def main() -> None:
    category = pd.read_csv(TABLE_DIR / "category_summary.csv")
    composition = pd.read_csv(TABLE_DIR / "composition_summary.csv")
    prop = pd.read_csv(TABLE_DIR / "property_summary.csv")
    flags = pd.read_csv(TABLE_DIR / "quality_flag_summary.csv")
    units = pd.read_csv(TABLE_DIR / "unit_summary.csv")

    total_files = int(category["files_total"].sum())
    total_records = int(category["wide_records"].sum())
    total_property_rows = int(prop["property_rows"].sum())
    total_files_with_records = int(category["files_with_records"].sum())

    top_props = prop.sort_values("property_rows", ascending=False).head(20)
    top_flags = flags.groupby("flag", as_index=False)["count"].sum().sort_values("count", ascending=False).head(20)
    top_unit_diversity = prop.sort_values("unit_variants", ascending=False).head(20)
    common_units = units.groupby(["property_target_canonical", "property_unit_raw"], as_index=False)["count"].sum()
    common_units = common_units.sort_values("count", ascending=False).head(30)

    readme_lines = [
        "# Superalloy Extraction Curation Outputs",
        "",
        f"Generated at: {datetime.now().isoformat(timespec='seconds')}",
        "",
        "## What was created",
        "",
        "- Four independent wide Excel files, one per original `ext-*` category, with consistent column names inside each category.",
        "- Four independent long-format Excel files, one per category, with one row per property mention.",
        "- Summary statistics tables and draft PNG figures for manuscript dataset description.",
        "- Processing logs with header variants and read errors.",
        "",
        "## Conservative automatic checks",
        "",
        "- Missing alloy name, missing composition values, and alloy-name-only records.",
        "- Low or abnormal composition totals after numeric parsing.",
        "- Duplicate element columns merged into one canonical element column.",
        "- Property target/name mismatch based on keyword rules.",
        "- Missing property values or units, non-numeric property values, and unit strings that were not standardized.",
        "- Possible non-superalloy sample/alloy keywords flagged for manual review.",
        "",
        "## Output file guide",
        "",
        "- `tables/*_wide_cleaned.xlsx`: sample-level record table.",
        "- `tables/*_properties_long.xlsx`: property-level table for statistics and plotting.",
        "- `tables/summary_statistics.xlsx`: combined summary workbook.",
        "- `figures/*.png`: draft figure panels.",
        "- `logs/processing_log.json`: machine-readable run log.",
        "",
        "## Category summary",
        "",
        markdown_table(category),
    ]
    (OUT_DIR / "README_整理说明.md").write_text("\n".join(readme_lines), encoding="utf-8")

    report_lines = [
        "# Scientific Data Dataset Assessment Notes",
        "",
        f"Generated at: {datetime.now().isoformat(timespec='seconds')}",
        "",
        "## Overall scale",
        "",
        f"- Source Excel files scanned: {total_files:,}.",
        f"- Files containing at least one extracted row: {total_files_with_records:,}.",
        f"- Sample-level extracted records: {total_records:,}.",
        f"- Property-level rows after long-format normalization: {total_property_rows:,}.",
        "",
        "## Recommended figure panels",
        "",
        "1. Extraction workflow and curation pipeline: raw literature files -> LLM extraction -> automated QC -> wide/long curated tables.",
        "2. Data scale by category: use `figures/records_by_category.png` and `figures/files_with_records_by_category.png`.",
        "3. Property coverage matrix: use `figures/property_record_heatmap_top30.png`.",
        "4. Composition completeness and quality flags: use `figures/missing_composition_by_category.png` and `figures/quality_flags_top20.png`.",
        "5. Unit heterogeneity before standardization: use `figures/unit_variant_count_top20_properties.png` plus `tables/unit_summary.csv`.",
        "6. Element coverage or alloy-family distribution: can be added from the wide tables after deciding how to group alloy designations.",
        "",
        "## What can be automatically detected",
        "",
        "- Structural issues: empty files, header variants, duplicate element columns, missing fields.",
        "- Composition issues: no numeric composition, partial composition, abnormal totals, inconsistent composition units.",
        "- Property issues: target/name mismatch, empty values, non-numeric values, unit anomalies, unstandardized units.",
        "- Dataset description metrics: records per category/property, files with useful records, unit diversity, QC flag frequencies.",
        "",
        "## What should remain semi-automatic",
        "",
        "- Completing compositions from alloy grades such as IN718, CMSX-10, GH4169: this needs a controlled alloy designation/composition dictionary.",
        "- Deciding whether a suspicious property is true signal or model hallucination: the current flags prioritize records for manual audit.",
        "- Merging equivalent alloy names: spelling variants and commercial designations need domain-aware normalization.",
        "- Removing non-superalloy records: keyword flags catch many obvious cases, but borderline substrate/interlayer records should be reviewed.",
        "",
        "## Category summary",
        "",
        markdown_table(category),
        "",
        "## Composition completeness",
        "",
        markdown_table(composition),
        "",
        "## Top property targets",
        "",
        markdown_table(top_props),
        "",
        "## Top quality flags",
        "",
        markdown_table(top_flags),
        "",
        "## Properties with most unit variants",
        "",
        markdown_table(top_unit_diversity),
        "",
        "## Most frequent raw units",
        "",
        markdown_table(common_units),
    ]
    (OUT_DIR / "Scientific_Data_整理分析报告.md").write_text("\n".join(report_lines), encoding="utf-8")


if __name__ == "__main__":
    main()
