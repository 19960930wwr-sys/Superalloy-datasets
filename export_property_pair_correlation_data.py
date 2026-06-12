from __future__ import annotations

import itertools
import math
import re
from pathlib import Path
from typing import Any

import pandas as pd


ROOT = Path(__file__).resolve().parent
FINAL_TABLE_DIR = ROOT / "final_valuable_dataset" / "tables"
OUT_DIR = ROOT / "paper_diversity_figures" / "tables"
OUT_XLSX = OUT_DIR / "property_pair_correlation_candidates.xlsx"


SELECTED_PROPERTIES = {
    "yield strength": {"unit": "MPa", "min": 1.0, "max": 3000.0},
    "tensile strength": {"unit": "MPa", "min": 1.0, "max": 3500.0},
    "total elongation": {"unit": "%", "min": 0.0, "max": 200.0},
    "strain rate": {"unit": "s^-1", "min": 1e-8, "max": 1e4},
    "youngs modulus": {"unit": "GPa", "min": 20.0, "max": 350.0},
    "compressive strength": {"unit": "MPa", "min": 1.0, "max": 5000.0},
    "creep life": {"unit": "h", "min": 1e-4, "max": 1e5},
    "grain size": {"unit": "um", "min": 1e-4, "max": 1e5},
}


PAIR_NOTES = {
    ("yield strength", "tensile strength"): "Strength-strength consistency; usually expected to show positive correlation.",
    ("yield strength", "total elongation"): "Strength-ductility trade-off; useful for mechanical performance map.",
    ("tensile strength", "total elongation"): "Strength-ductility trade-off; useful for mechanical performance map.",
    ("yield strength", "grain size"): "Hall-Petch-type analysis; plot yield strength versus grain_size^-1/2.",
    ("tensile strength", "grain size"): "Microstructure-strength relation; plot tensile strength versus grain_size^-1/2.",
    ("hardness", "yield strength"): "Hardness-strength relation, if hardness is included in a later selection.",
    ("creep life", "grain size"): "Creep performance may correlate with grain size, but test stress/temperature should be controlled.",
    ("youngs modulus", "yield strength"): "Elastic-plastic property relation; may be weak but can reveal extraction consistency.",
    ("compressive strength", "yield strength"): "Compression-tension strength relation; useful if matched testing conditions are comparable.",
}


def clean(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    text = str(value).strip()
    if text.lower() in {"", "none", "nan", "null", "n/a", "na", "not specified", "[]", "['none']", "['None']"}:
        return ""
    return text


def key_text(value: Any) -> str:
    text = clean(value).lower()
    text = re.sub(r"\s+", " ", text)
    return text


def read_many(folder: Path, pattern: str) -> pd.DataFrame:
    frames = []
    for path in sorted(folder.glob(pattern)):
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
        df["source_table_file"] = path.name
        frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def safe_sheet_name(name: str, used: set[str]) -> str:
    sheet = re.sub(r"[\[\]\:\*\?\/\\]", "_", name).strip()[:31] or "sheet"
    base = sheet
    idx = 1
    while sheet in used:
        suffix = f"_{idx}"
        sheet = (base[: 31 - len(suffix)] + suffix)[:31]
        idx += 1
    used.add(sheet)
    return sheet


def pearson(x: pd.Series, y: pd.Series) -> float:
    if len(x) < 3 or x.nunique(dropna=True) < 2 or y.nunique(dropna=True) < 2:
        return float("nan")
    return float(x.corr(y, method="pearson"))


def spearman(x: pd.Series, y: pd.Series) -> float:
    if len(x) < 3 or x.nunique(dropna=True) < 2 or y.nunique(dropna=True) < 2:
        return float("nan")
    return float(x.rank().corr(y.rank(), method="pearson"))


def pair_note(a: str, b: str) -> str:
    key = tuple(sorted([a, b]))
    normalized = {tuple(sorted(k)): v for k, v in PAIR_NOTES.items()}
    return normalized.get(key, "Exploratory pair; interpret with matched processing and testing context.")


def prepare_selected_property_rows() -> pd.DataFrame:
    df = read_many(FINAL_TABLE_DIR, "final_ext-*_valuable_properties_long.csv")
    df["property"] = df["property_target_canonical"].map(clean)
    df["standard_unit"] = df["property_standard_unit"].map(clean)
    df["standard_value"] = pd.to_numeric(df["property_standard_value"], errors="coerce")
    rows = []
    for prop, rule in SELECTED_PROPERTIES.items():
        sub = df[
            (df["property"] == prop)
            & (df["standard_unit"] == rule["unit"])
            & df["standard_value"].notna()
        ].copy()
        sub = sub[sub["standard_value"].map(lambda v: math.isfinite(float(v)))]
        sub = sub[(sub["standard_value"] >= rule["min"]) & (sub["standard_value"] <= rule["max"])]
        rows.append(sub)
    selected = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame()
    selected["strict_key"] = selected.apply(
        lambda r: "|".join(
            [
                key_text(r.get("doi_or_file_id", "")),
                key_text(r.get("superalloy_name", "")),
                key_text(r.get("sample_name", "")),
                key_text(r.get("distinguishing_factor", "")),
                key_text(r.get("synthesis_and_processing_routes", "")),
            ]
        ),
        axis=1,
    )
    selected["relaxed_key"] = selected.apply(
        lambda r: "|".join(
            [
                key_text(r.get("doi_or_file_id", "")),
                key_text(r.get("superalloy_name", "")),
                key_text(r.get("synthesis_and_processing_routes", "")),
            ]
        ),
        axis=1,
    )
    selected = selected[(selected["doi_or_file_id"].map(clean) != "") & (selected["superalloy_name"].map(clean) != "")]
    return selected


def collapse_property_per_key(df: pd.DataFrame, key_col: str) -> pd.DataFrame:
    meta_cols = [
        "doi_or_file_id",
        "superalloy_name",
        "sample_name",
        "distinguishing_factor",
        "synthesis_and_processing_routes",
        "test_route_condition",
        "source_file",
        "source_row",
        "record_id",
        "property_value_raw",
        "property_unit_raw",
        "property_sourced_figure",
    ]
    grouped_rows = []
    for (key, prop), group in df.groupby([key_col, "property"], dropna=False):
        values = group["standard_value"].astype(float)
        first = group.iloc[0]
        row = {
            key_col: key,
            "property": prop,
            "value": float(values.median()),
            "value_mean": float(values.mean()),
            "value_count_for_key": int(len(values)),
            "standard_unit": first["standard_unit"],
        }
        for col in meta_cols:
            row[col] = first.get(col, "")
        grouped_rows.append(row)
    return pd.DataFrame(grouped_rows)


def build_pair_table(collapsed: pd.DataFrame, key_col: str, prop_a: str, prop_b: str) -> pd.DataFrame:
    a = collapsed[collapsed["property"] == prop_a].copy()
    b = collapsed[collapsed["property"] == prop_b].copy()
    merged = a.merge(b, on=key_col, suffixes=("_x", "_y"))
    if merged.empty:
        return merged
    out = pd.DataFrame({
        "match_level": "strict" if key_col == "strict_key" else "relaxed",
        "pair": f"{prop_a} vs {prop_b}",
        "match_key": merged[key_col],
        "doi_or_file_id": merged["doi_or_file_id_x"].combine_first(merged["doi_or_file_id_y"]),
        "superalloy_name": merged["superalloy_name_x"].combine_first(merged["superalloy_name_y"]),
        "sample_name_x": merged["sample_name_x"],
        "sample_name_y": merged["sample_name_y"],
        "distinguishing_factor_x": merged["distinguishing_factor_x"],
        "distinguishing_factor_y": merged["distinguishing_factor_y"],
        "synthesis_and_processing_routes_x": merged["synthesis_and_processing_routes_x"],
        "synthesis_and_processing_routes_y": merged["synthesis_and_processing_routes_y"],
        "test_route_condition_x": merged["test_route_condition_x"],
        "test_route_condition_y": merged["test_route_condition_y"],
        "property_x": prop_a,
        "value_x": merged["value_x"],
        "unit_x": merged["standard_unit_x"],
        "property_y": prop_b,
        "value_y": merged["value_y"],
        "unit_y": merged["standard_unit_y"],
        "source_file_x": merged["source_file_x"],
        "source_row_x": merged["source_row_x"],
        "record_id_x": merged["record_id_x"],
        "raw_value_x": merged["property_value_raw_x"],
        "raw_unit_x": merged["property_unit_raw_x"],
        "source_file_y": merged["source_file_y"],
        "source_row_y": merged["source_row_y"],
        "record_id_y": merged["record_id_y"],
        "raw_value_y": merged["property_value_raw_y"],
        "raw_unit_y": merged["property_unit_raw_y"],
    })
    if prop_a == "grain size":
        out["grain_size_inverse_sqrt_x"] = 1 / (out["value_x"].astype(float) ** 0.5)
    if prop_b == "grain size":
        out["grain_size_inverse_sqrt_y"] = 1 / (out["value_y"].astype(float) ** 0.5)
    return out


def summarize_pair(pair_df: pd.DataFrame, prop_a: str, prop_b: str, match_level: str) -> dict[str, Any]:
    if pair_df.empty:
        return {
            "match_level": match_level,
            "property_x": prop_a,
            "unit_x": SELECTED_PROPERTIES[prop_a]["unit"],
            "property_y": prop_b,
            "unit_y": SELECTED_PROPERTIES[prop_b]["unit"],
            "matched_count": 0,
            "pearson_r": float("nan"),
            "spearman_r": float("nan"),
            "recommended_plot_x": prop_a,
            "recommended_plot_y": prop_b,
            "reason": pair_note(prop_a, prop_b),
        }
    x_col = "value_x"
    y_col = "value_y"
    recommended_x = prop_a
    recommended_y = prop_b
    plot_pearson = pearson(pair_df[x_col].astype(float), pair_df[y_col].astype(float))
    plot_spearman = spearman(pair_df[x_col].astype(float), pair_df[y_col].astype(float))
    if prop_a == "grain size":
        x_col = "grain_size_inverse_sqrt_x"
        recommended_x = "grain size^-1/2"
        plot_pearson = pearson(pair_df[x_col].astype(float), pair_df["value_y"].astype(float))
        plot_spearman = spearman(pair_df[x_col].astype(float), pair_df["value_y"].astype(float))
    elif prop_b == "grain size":
        x_col = "grain_size_inverse_sqrt_y"
        recommended_x = "grain size^-1/2"
        recommended_y = prop_a
        plot_pearson = pearson(pair_df[x_col].astype(float), pair_df["value_x"].astype(float))
        plot_spearman = spearman(pair_df[x_col].astype(float), pair_df["value_x"].astype(float))
    return {
        "match_level": match_level,
        "property_x": prop_a,
        "unit_x": SELECTED_PROPERTIES[prop_a]["unit"],
        "property_y": prop_b,
        "unit_y": SELECTED_PROPERTIES[prop_b]["unit"],
        "matched_count": int(len(pair_df)),
        "pearson_r": round(plot_pearson, 4) if not math.isnan(plot_pearson) else "",
        "spearman_r": round(plot_spearman, 4) if not math.isnan(plot_spearman) else "",
        "recommended_plot_x": recommended_x,
        "recommended_plot_y": recommended_y,
        "reason": pair_note(prop_a, prop_b),
    }


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    selected = prepare_selected_property_rows()
    all_summaries = []
    pair_tables = {}

    for key_col in ["strict_key", "relaxed_key"]:
        collapsed = collapse_property_per_key(selected, key_col)
        match_level = "strict" if key_col == "strict_key" else "relaxed"
        for prop_a, prop_b in itertools.combinations(SELECTED_PROPERTIES.keys(), 2):
            pair_df = build_pair_table(collapsed, key_col, prop_a, prop_b)
            all_summaries.append(summarize_pair(pair_df, prop_a, prop_b, match_level))
            if len(pair_df) >= 20:
                pair_tables[(match_level, prop_a, prop_b)] = pair_df

    summary = pd.DataFrame(all_summaries)
    summary = summary.sort_values(["match_level", "matched_count"], ascending=[True, False]).reset_index(drop=True)
    recommended = summary[
        (
            (summary["matched_count"] >= 50)
            & summary["reason"].str.contains("Strength|Hall|ductility|Creep|Microstructure|Compression", case=False, regex=True)
        )
        | (summary["matched_count"] >= 200)
    ].copy()
    recommended = recommended.sort_values(["match_level", "matched_count"], ascending=[True, False])

    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
        summary.to_excel(writer, index=False, sheet_name="pair_summary_all")
        recommended.to_excel(writer, index=False, sheet_name="recommended_pairs")
        selected.to_excel(writer, index=False, sheet_name="selected_property_rows")
        used = {"pair_summary_all", "recommended_pairs", "selected_property_rows"}
        for (match_level, prop_a, prop_b), df in sorted(pair_tables.items(), key=lambda item: (item[0][0], -len(item[1]))):
            sheet = safe_sheet_name(f"{match_level}_{prop_a[:8]}_{prop_b[:8]}", used)
            df.to_excel(writer, index=False, sheet_name=sheet)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

    summary.to_csv(OUT_DIR / "property_pair_correlation_summary.csv", index=False, encoding="utf-8-sig")
    recommended.to_csv(OUT_DIR / "property_pair_correlation_recommended_pairs.csv", index=False, encoding="utf-8-sig")
    # Export the largest useful pair tables as separate CSV files for easy Origin import.
    for (match_level, prop_a, prop_b), df in pair_tables.items():
        if len(df) >= 50:
            name = re.sub(r"[^A-Za-z0-9]+", "_", f"property_pair_{match_level}_{prop_a}_vs_{prop_b}").strip("_")
            df.to_csv(OUT_DIR / f"{name}.csv", index=False, encoding="utf-8-sig")

    print(f"Wrote {OUT_XLSX}")
    print("Top pair candidates:")
    print(recommended.head(20).to_string(index=False))


if __name__ == "__main__":
    main()
