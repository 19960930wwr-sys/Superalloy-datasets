from __future__ import annotations

import json
import math
from pathlib import Path
from typing import Any

import pandas as pd
from PIL import Image, ImageDraw, ImageFont


OUT_DIR = Path(__file__).resolve().parent
TABLE_DIR = OUT_DIR / "tables"
FIG_DIR = OUT_DIR / "figures"
FINAL_DIR = OUT_DIR / "final_valuable_dataset"
FINAL_TABLE_DIR = FINAL_DIR / "tables"
FINAL_FIG_DIR = FINAL_DIR / "figures"

ELEMENTS = [
    "H", "B", "C", "N", "O", "F", "Na", "Mg", "Al", "Si", "P", "S", "Cl", "Ar",
    "K", "Ca", "Sc", "Ti", "V", "Cr", "Mn", "Fe", "Co", "Ni", "Cu", "Zn", "Ge",
    "As", "Sr", "Y", "Zr", "Nb", "Mo", "Pd", "Ag", "Cd", "Sn", "Sb", "Te", "Ba",
    "La", "Ce", "Nd", "Er", "Hf", "Ta", "W", "Re", "Os", "Ir", "Pt", "Pb", "Bi",
    "Ru",
]

LONG_PATTERN = "*_properties_long.csv"
WIDE_PATTERN = "*_wide_cleaned.csv"

EMPTY_MARKERS = {
    "", "none", "nan", "null", "n/a", "na", "not specified", "not available",
    "[]", "['none']", "['None']", "[None]",
}


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() in EMPTY_MARKERS else text


def bool_text(value: bool) -> str:
    return "yes" if value else "no"


def write_excel(path: Path, sheets: dict[str, pd.DataFrame]) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe = sheet_name[:31]
            df.to_excel(writer, sheet_name=safe, index=False)
            ws = writer.book[safe]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions


def markdown_table(df: pd.DataFrame) -> str:
    if df.empty:
        return "_No rows._"
    cols = list(df.columns)
    lines = [" | ".join(cols), " | ".join(["---"] * len(cols))]
    for _, row in df.iterrows():
        lines.append(" | ".join(str(row[col]) for col in cols))
    return "\n".join(lines)


def draw_bar_chart(path: Path, title: str, labels: list[str], values: list[float], x_label: str) -> None:
    width, height = 1280, max(420, 90 + 46 * len(labels))
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    font = ImageFont.load_default()
    draw.text((30, 24), title, fill=(20, 20, 20), font=font)
    left, top = 330, 75
    bar_h, gap = 28, 18
    max_value = max(values) if values else 1
    colors = [(41, 111, 145), (202, 126, 45), (79, 145, 98), (145, 93, 142), (178, 76, 76)]
    for i, (label, value) in enumerate(zip(labels, values)):
        y = top + i * (bar_h + gap)
        draw.text((30, y + 8), str(label)[:44], fill=(35, 35, 35), font=font)
        w = int((width - left - 190) * value / max_value) if max_value else 0
        draw.rectangle([left, y, left + w, y + bar_h], fill=colors[i % len(colors)])
        draw.text((left + w + 10, y + 8), f"{value:g}", fill=(35, 35, 35), font=font)
    draw.text((left, height - 30), x_label, fill=(70, 70, 70), font=font)
    image.save(path)


def draw_grouped_bars(path: Path, title: str, df: pd.DataFrame, category_col: str, value_cols: list[str]) -> None:
    labels = df[category_col].tolist()
    width, height = 1500, max(480, 130 + 74 * len(labels))
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    font = ImageFont.load_default()
    draw.text((30, 24), title, fill=(20, 20, 20), font=font)
    left, top = 350, 85
    group_h = 58
    bar_h = 14
    max_value = max([float(df[c].max()) for c in value_cols] + [1.0])
    colors = [(41, 111, 145), (202, 126, 45), (79, 145, 98), (145, 93, 142)]
    for j, col in enumerate(value_cols):
        x = left + j * 260
        draw.rectangle([x, 52, x + 18, 66], fill=colors[j % len(colors)])
        draw.text((x + 24, 51), col, fill=(45, 45, 45), font=font)
    for i, label in enumerate(labels):
        y0 = top + i * group_h
        draw.text((30, y0 + 16), str(label)[:44], fill=(35, 35, 35), font=font)
        for j, col in enumerate(value_cols):
            value = float(df.loc[df.index[i], col])
            y = y0 + j * (bar_h + 2)
            w = int((width - left - 220) * value / max_value)
            draw.rectangle([left, y, left + w, y + bar_h], fill=colors[j % len(colors)])
            draw.text((left + w + 8, y), f"{value:g}", fill=(35, 35, 35), font=font)
    image.save(path)


def draw_heatmap(path: Path, title: str, matrix: pd.DataFrame) -> None:
    cell_w, cell_h = 170, 34
    left, top = 330, 82
    width = left + cell_w * len(matrix.columns) + 80
    height = max(420, top + cell_h * len(matrix.index) + 70)
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    font = ImageFont.load_default()
    draw.text((30, 24), title, fill=(20, 20, 20), font=font)
    max_v = max([float(v) for v in matrix.to_numpy().flatten()] + [1.0])
    for j, col in enumerate(matrix.columns):
        draw.text((left + j * cell_w + 8, top - 28), str(col)[:23], fill=(40, 40, 40), font=font)
    for i, idx in enumerate(matrix.index):
        y = top + i * cell_h
        draw.text((30, y + 9), str(idx)[:48], fill=(40, 40, 40), font=font)
        for j, col in enumerate(matrix.columns):
            x = left + j * cell_w
            v = float(matrix.loc[idx, col])
            shade = int(245 - 170 * (v / max_v))
            draw.rectangle([x, y, x + cell_w - 2, y + cell_h - 2], fill=(shade, 235, 245), outline=(220, 220, 220))
            if v:
                draw.text((x + 8, y + 9), str(int(v)), fill=(20, 20, 20), font=font)
    image.save(path)


def category_from_wide_name(path: Path) -> str:
    stem = path.stem
    return stem.replace("_wide_cleaned", "")


def matching_long_path(wide_path: Path) -> Path:
    return TABLE_DIR / wide_path.name.replace("_wide_cleaned.csv", "_properties_long.csv")


def build_properties_json(prop_df: pd.DataFrame) -> pd.Series:
    cols = [
        "property_target_canonical",
        "property_name_raw",
        "property_value_raw",
        "property_unit_raw",
        "property_standard_value",
        "property_standard_unit",
        "unit_conversion_note",
        "property_flags",
    ]

    def pack(group: pd.DataFrame) -> str:
        rows = []
        for _, row in group[cols].iterrows():
            item = {col: clean_text(row[col]) for col in cols}
            rows.append(item)
        return json.dumps(rows, ensure_ascii=False)

    return prop_df.groupby("record_id", sort=False).apply(pack)


def main() -> None:
    FINAL_TABLE_DIR.mkdir(parents=True, exist_ok=True)
    FINAL_FIG_DIR.mkdir(parents=True, exist_ok=True)

    category_summaries: list[dict[str, Any]] = []
    exclusion_rows: list[dict[str, Any]] = []
    identity_rows: list[dict[str, Any]] = []
    all_final_properties: list[pd.DataFrame] = []

    final_record_columns = [
        "category", "source_folder", "doi_or_file_id", "source_file", "source_row", "record_id",
        "superalloy_name", "sample_name",
        "identity_basis", "has_grade", "has_complete_composition",
        "has_process_info", "has_property_result", "valuable_record",
        "exclusion_reasons",
        "composition_unit_raw", "composition_unit_normalized",
        "composition_present_element_count", "composition_sum_numeric",
        "distinguishing_factor", "synthesis_and_processing_routes", "test_route_condition",
        "property_value_count", "property_targets_with_values", "properties_json",
        "quality_flags", "quality_flag_count",
    ] + [f"element_{e}" for e in ELEMENTS]

    for wide_path in sorted(TABLE_DIR.glob(WIDE_PATTERN)):
        category_file_stem = category_from_wide_name(wide_path)
        long_path = matching_long_path(wide_path)
        print(f"Filtering {category_file_stem}", flush=True)
        wide = pd.read_csv(wide_path, dtype=str, keep_default_na=False)
        long = pd.read_csv(long_path, dtype=str, keep_default_na=False)

        long_value = long[long["property_value_raw"].map(clean_text) != ""].copy()
        prop_counts = long_value.groupby("record_id").size().rename("property_value_count")
        prop_targets = (
            long_value.groupby("record_id")["property_target_canonical"]
            .apply(lambda x: ";".join(sorted({clean_text(v) for v in x if clean_text(v)})))
            .rename("property_targets_with_values")
        )
        properties_json = build_properties_json(long_value).rename("properties_json") if not long_value.empty else pd.Series(dtype=str)

        wide = wide.merge(prop_counts, on="record_id", how="left")
        wide = wide.merge(prop_targets, on="record_id", how="left")
        wide = wide.merge(properties_json, on="record_id", how="left")
        wide["property_value_count"] = wide["property_value_count"].fillna(0).astype(int)
        wide["property_targets_with_values"] = wide["property_targets_with_values"].fillna("")
        wide["properties_json"] = wide["properties_json"].fillna("")

        wide["has_grade_bool"] = wide["superalloy_name"].map(clean_text) != ""
        comp_count = pd.to_numeric(wide["composition_present_element_count"], errors="coerce").fillna(0)
        comp_sum = pd.to_numeric(wide["composition_sum_numeric"], errors="coerce")
        wide["has_complete_composition_bool"] = (comp_count >= 3) & comp_sum.between(90, 110, inclusive="both")
        wide["has_process_info_bool"] = wide["synthesis_and_processing_routes"].map(clean_text) != ""
        wide["has_property_result_bool"] = wide["property_value_count"] > 0
        wide["valuable_record_bool"] = (
            (wide["has_grade_bool"] | wide["has_complete_composition_bool"])
            & wide["has_process_info_bool"]
            & wide["has_property_result_bool"]
        )

        def basis(row: pd.Series) -> str:
            if row["has_grade_bool"] and row["has_complete_composition_bool"]:
                return "grade_and_complete_composition"
            if row["has_grade_bool"]:
                return "grade_only_or_grade_with_partial_composition"
            if row["has_complete_composition_bool"]:
                return "complete_composition_only"
            return "missing_grade_and_complete_composition"

        def reasons(row: pd.Series) -> str:
            out = []
            if not (row["has_grade_bool"] or row["has_complete_composition_bool"]):
                out.append("missing_grade_or_complete_composition")
            if not row["has_process_info_bool"]:
                out.append("missing_process_info")
            if not row["has_property_result_bool"]:
                out.append("missing_property_result")
            return ";".join(out)

        wide["identity_basis"] = wide.apply(basis, axis=1)
        wide["exclusion_reasons"] = wide.apply(reasons, axis=1)
        for src, dst in [
            ("has_grade_bool", "has_grade"),
            ("has_complete_composition_bool", "has_complete_composition"),
            ("has_process_info_bool", "has_process_info"),
            ("has_property_result_bool", "has_property_result"),
            ("valuable_record_bool", "valuable_record"),
        ]:
            wide[dst] = wide[src].map(bool_text)

        final_records = wide[wide["valuable_record_bool"]].copy()
        final_long = long_value[long_value["record_id"].isin(set(final_records["record_id"]))].copy()
        all_final_properties.append(final_long)

        for col in final_record_columns:
            if col not in final_records.columns:
                final_records[col] = ""
        final_records = final_records[final_record_columns]

        for col in final_record_columns:
            if col not in wide.columns:
                wide[col] = ""
        criteria_audit = wide[final_record_columns].copy()

        excel_path = FINAL_TABLE_DIR / f"final_{category_file_stem}_valuable_records.xlsx"
        long_excel_path = FINAL_TABLE_DIR / f"final_{category_file_stem}_valuable_properties_long.xlsx"
        write_excel(excel_path, {"valuable_records": final_records})
        write_excel(long_excel_path, {"valuable_properties": final_long})
        final_records.to_csv(FINAL_TABLE_DIR / f"final_{category_file_stem}_valuable_records.csv", index=False, encoding="utf-8-sig")
        final_long.to_csv(FINAL_TABLE_DIR / f"final_{category_file_stem}_valuable_properties_long.csv", index=False, encoding="utf-8-sig")
        criteria_audit.to_csv(FINAL_TABLE_DIR / f"audit_{category_file_stem}_all_records_with_criteria.csv", index=False, encoding="utf-8-sig")

        category_summaries.append({
            "category_file": category_file_stem,
            "category": wide["category"].iloc[0] if len(wide) else "",
            "input_records": len(wide),
            "valuable_records": len(final_records),
            "retention_rate_percent": round(len(final_records) / len(wide) * 100, 2) if len(wide) else 0,
            "input_property_value_rows": len(long_value),
            "valuable_property_value_rows": len(final_long),
            "has_grade_records": int(wide["has_grade_bool"].sum()),
            "has_complete_composition_records": int(wide["has_complete_composition_bool"].sum()),
            "has_process_records": int(wide["has_process_info_bool"].sum()),
            "has_property_result_records": int(wide["has_property_result_bool"].sum()),
        })

        for reason, count in wide.loc[~wide["valuable_record_bool"], "exclusion_reasons"].str.get_dummies(sep=";").sum().items():
            if count:
                exclusion_rows.append({"category_file": category_file_stem, "reason": reason, "count": int(count)})
        for basis_name, count in final_records["identity_basis"].value_counts().items():
            identity_rows.append({"category_file": category_file_stem, "identity_basis": basis_name, "count": int(count)})

    summary = pd.DataFrame(category_summaries)
    exclusions = pd.DataFrame(exclusion_rows)
    identities = pd.DataFrame(identity_rows)
    final_props = pd.concat(all_final_properties, ignore_index=True) if all_final_properties else pd.DataFrame()

    property_summary = pd.DataFrame()
    if not final_props.empty:
        property_summary = (
            final_props.groupby(["category", "property_target_canonical"], dropna=False)
            .agg(
                valuable_property_rows=("record_id", "count"),
                records=("record_id", "nunique"),
                unit_variants=("property_unit_raw", lambda x: len({clean_text(v) for v in x if clean_text(v)})),
                standardized_rows=("property_standard_unit", lambda x: sum(bool(clean_text(v)) for v in x)),
            )
            .reset_index()
            .sort_values(["category", "valuable_property_rows"], ascending=[True, False])
        )

    write_excel(FINAL_TABLE_DIR / "final_valuable_summary_statistics.xlsx", {
        "category_summary": summary,
        "exclusion_reasons": exclusions,
        "identity_basis": identities,
        "property_summary": property_summary,
    })
    summary.to_csv(FINAL_TABLE_DIR / "final_category_summary.csv", index=False, encoding="utf-8-sig")
    exclusions.to_csv(FINAL_TABLE_DIR / "final_exclusion_reasons.csv", index=False, encoding="utf-8-sig")
    identities.to_csv(FINAL_TABLE_DIR / "final_identity_basis.csv", index=False, encoding="utf-8-sig")
    property_summary.to_csv(FINAL_TABLE_DIR / "final_property_summary.csv", index=False, encoding="utf-8-sig")

    if not summary.empty:
        draw_bar_chart(
            FINAL_FIG_DIR / "valuable_records_by_category.png",
            "Valuable records by category",
            summary["category"].tolist(),
            summary["valuable_records"].astype(float).tolist(),
            "records passing final criteria",
        )
        draw_bar_chart(
            FINAL_FIG_DIR / "retention_rate_by_category.png",
            "Retention rate after final value criteria",
            summary["category"].tolist(),
            summary["retention_rate_percent"].astype(float).tolist(),
            "percent of cleaned records retained",
        )
        draw_grouped_bars(
            FINAL_FIG_DIR / "criteria_funnel_by_category.png",
            "Records satisfying each final-value criterion",
            summary,
            "category",
            ["has_grade_records", "has_complete_composition_records", "has_process_records", "has_property_result_records"],
        )
    if not exclusions.empty:
        ex = exclusions.groupby("reason", as_index=False)["count"].sum().sort_values("count", ascending=False)
        draw_bar_chart(
            FINAL_FIG_DIR / "exclusion_reasons_total.png",
            "Why records failed final value criteria",
            ex["reason"].tolist(),
            ex["count"].astype(float).tolist(),
            "records",
        )
    if not identities.empty:
        ident = identities.groupby("identity_basis", as_index=False)["count"].sum().sort_values("count", ascending=False)
        draw_bar_chart(
            FINAL_FIG_DIR / "identity_basis_total.png",
            "Identity basis among valuable records",
            ident["identity_basis"].tolist(),
            ident["count"].astype(float).tolist(),
            "valuable records",
        )
    if not property_summary.empty:
        pivot = property_summary.pivot_table(
            index="property_target_canonical",
            columns="category",
            values="valuable_property_rows",
            aggfunc="sum",
            fill_value=0,
        )
        pivot = pivot.loc[pivot.sum(axis=1).sort_values(ascending=False).index[:30]]
        draw_heatmap(
            FINAL_FIG_DIR / "valuable_property_heatmap_top30.png",
            "Valuable property-value rows by category",
            pivot,
        )

    report = [
        "# Final Valuable Dataset Filtering",
        "",
        "Final-value rule used here:",
        "",
        "- `has_grade` OR `has_complete_composition` must be true.",
        "- `has_complete_composition` means at least 3 numeric elements and composition sum between 90 and 110.",
        "- `has_process_info` requires non-empty `synthesis_and_processing_routes`.",
        "- `has_property_result` requires at least one non-empty extracted property value.",
        "",
        "## Category Summary",
        "",
        markdown_table(summary),
    ]
    (FINAL_DIR / "final_filtering_report.md").write_text("\n".join(report), encoding="utf-8")
    print(f"Done. Final valuable dataset written to: {FINAL_DIR}", flush=True)


if __name__ == "__main__":
    main()
