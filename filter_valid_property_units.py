from __future__ import annotations

import math
import re
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
    "(missing unit)",
    "(not standardized)",
}

CATEGORICAL_NO_UNIT_PROPERTIES = {
    "phases present gamma gamma prime tcp phases carbides",
    "gamma prime morphology",
}

PROPERTY_GROUP_FOR_RULE = {
    "gauge length": "length_mm",
    "strain rate": "strain_rate",
    "tensile strength": "stress",
    "yield strength": "stress",
    "compressive strength": "stress",
    "stress rupture strength": "stress",
    "total elongation": "percent",
    "uniform elongation": "percent",
    "hardness": "hardness",
    "creep life": "time_life",
    "fatigue life": "fatigue_life",
    "oxidation gain": "oxidation_gain",
    "density": "density",
    "gamma prime solvus temperature": "temperature",
    "liquidus temperature": "temperature",
    "solidus temperature": "temperature",
    "gamma prime volume fraction": "percent",
    "gamma prime size": "micro_length",
    "grain size": "micro_length",
    "creep strain rate": "strain_rate",
    "fracture toughness": "fracture_toughness",
    "crack growth rate": "crack_growth_rate",
    "youngs modulus": "modulus",
    "shear modulus": "modulus",
    "thermal conductivity": "thermal_conductivity",
    "thermal expansion coefficient": "thermal_expansion",
    "specific heat capacity": "specific_heat",
}

RULE_DESCRIPTIONS = {
    "stress": "应力/强度单位，例如 MPa、GPa、Pa、ksi、N/mm2；统计时优先使用已转换的 MPa。",
    "modulus": "弹性模量单位，例如 GPa、MPa；统计时优先使用已转换的 GPa。",
    "percent": "百分数单位，例如 %、percent。",
    "hardness": "硬度单位，例如 HV/HV0.1/HRC/HRB/HB/GPa。",
    "time_life": "时间寿命单位，例如 h、s、min；统计时优先使用已转换的 h。",
    "fatigue_life": "疲劳寿命单位，例如 cycles、cycle、Nf。",
    "oxidation_gain": "氧化增重单位，例如 mg/cm2、g/m2、kg/m2、%。",
    "density": "密度或相对密度单位，例如 g/cm3、kg/m3、%。",
    "temperature": "温度单位，例如 degC、°C、K。",
    "micro_length": "显微尺度长度单位，例如 um、μm、µm、nm、mm。",
    "length_mm": "试样长度单位，例如 mm、cm、m、um。",
    "strain_rate": "应变率单位，例如 s^-1、1/s。",
    "fracture_toughness": "断裂韧性单位，例如 MPa·m^0.5、MPa√m、ksi√in。",
    "crack_growth_rate": "裂纹扩展速率单位，例如 m/cycle、mm/cycle、um/cycle。",
    "thermal_conductivity": "热导率单位，例如 W/m/K、W/(m·K)。",
    "thermal_expansion": "热膨胀系数单位，例如 1/K、K^-1、10^-6/K、ppm/K。",
    "specific_heat": "比热容单位，例如 J/kg/K、J/g/K。",
}


def clean(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() in EMPTY else text


def unit_norm(unit: Any) -> str:
    text = clean(unit).lower()
    replacements = {
        "μ": "u",
        "µ": "u",
        "−": "-",
        "–": "-",
        "—": "-",
        "·": " ",
        "⋅": " ",
        "×": "x",
        "⁻": "-",
        "¹": "1",
        "²": "2",
        "³": "3",
        "½": "1/2",
        "℃": "degc",
        "°c": "degc",
        "° c": "degc",
        "√": "sqrt",
        "⁄": "/",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = text.replace("per", "/")
    text = re.sub(r"\s+", "", text)
    text = text.replace("−", "-")
    return text


def any_match(units: list[str], patterns: list[str]) -> bool:
    return any(re.search(pattern, unit) for unit in units if unit for pattern in patterns)


def is_valid_unit(property_name: str, raw_unit: Any, standard_unit: Any) -> tuple[bool, str, str]:
    prop = clean(property_name)
    raw = unit_norm(raw_unit)
    std = unit_norm(standard_unit)
    units = [std, raw]
    rule = PROPERTY_GROUP_FOR_RULE.get(prop, "")

    if prop in CATEGORICAL_NO_UNIT_PROPERTIES:
        if not raw and not std:
            return True, "categorical_no_unit", "分类/文本型性能，无单位是合理的"
        return False, "categorical_no_unit", "分类/文本型性能不应带物理单位"

    if not rule:
        return False, "no_rule", "该性能尚未配置单位白名单"

    if not any(units):
        return False, rule, "缺失单位，无法确认物理量维度"

    valid = False
    if rule == "stress":
        valid = any_match(units, [r"^mpa$", r"^gpa$", r"^kpa$", r"^pa$", r"ksi", r"n/mm2", r"nmm-2", r"kn/mm2"])
    elif rule == "modulus":
        valid = any_match(units, [r"^gpa$", r"^mpa$", r"^pa$", r"ksi"])
    elif rule == "percent":
        valid = any_match(units, [r"^%$", r"percent", r"vol%", r"wt%"])
    elif rule == "hardness":
        valid = any_match(units, [r"^hv", r"vickers", r"^hr[abc]?$", r"rockwell", r"^hb", r"brinell", r"^gpa$"])
    elif rule == "time_life":
        valid = any_match(units, [r"^h$", r"^hr$", r"hour", r"^s$", r"sec", r"^min$"])
    elif rule == "fatigue_life":
        valid = any_match(units, [r"cycle", r"cycles", r"^nf$", r"^n$"])
    elif rule == "oxidation_gain":
        valid = any_match(
            units,
            [
                r"mg/cm2",
                r"mgcm-2",
                r"mg/cm\^2",
                r"mgcm2",
                r"g/m2",
                r"gm-2",
                r"kg/m2",
                r"^%$",
                r"percent",
            ],
        )
    elif rule == "density":
        valid = any_match(units, [r"g/cm3", r"gcm-3", r"kg/m3", r"kgm-3", r"^%$", r"percent"])
    elif rule == "temperature":
        valid = any_match(units, [r"degc", r"celsius", r"^c$", r"^k$"])
    elif rule == "micro_length":
        valid = any_match(units, [r"^um$", r"^u?m$", r"micron", r"^nm$", r"^mm$"])
    elif rule == "length_mm":
        valid = any_match(units, [r"^mm$", r"^cm$", r"^m$", r"^um$", r"^nm$"])
    elif rule == "strain_rate":
        valid = any_match(units, [r"s\^-?1", r"s-1", r"1/s", r"/s"])
    elif rule == "fracture_toughness":
        valid = any_match(
            units,
            [
                r"mpa.*m(\^?0\.5|1/2)",
                r"mpa.*sqrtm",
                r"mpasqrtm",
                r"mpa.?m0\.5",
                r"ksi.*sqrt",
            ],
        )
    elif rule == "crack_growth_rate":
        valid = any_match(units, [r"m/cycle", r"mm/cycle", r"um/cycle", r"nm/cycle", r"m/cycles", r"cycle-1"])
    elif rule == "thermal_conductivity":
        valid = any_match(units, [r"w/\(?m\)?/?k", r"wm-1k-1", r"w/mk", r"w\(m\.?k\)-1"])
    elif rule == "thermal_expansion":
        valid = any_match(units, [r"k-1", r"/k", r"1/k", r"degc-1", r"/degc", r"ppm/k", r"10\^-?6"])
    elif rule == "specific_heat":
        valid = any_match(units, [r"j/kg/k", r"jkg-1k-1", r"j/g/k", r"jg-1k-1", r"j/\(?kg", r"j/\(?g"])

    return valid, rule, RULE_DESCRIPTIONS.get(rule, "")


def read_many(folder: Path, pattern: str) -> pd.DataFrame:
    frames = []
    for path in sorted(folder.glob(pattern)):
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
        df["source_table_file"] = path.name
        frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def annotate_units(df: pd.DataFrame, scope: str) -> pd.DataFrame:
    data = df.copy()
    data["统计口径"] = scope
    data["性能"] = data["property_target_canonical"].map(clean)
    data["原始单位"] = data["property_unit_raw"].map(clean)
    data["标准单位"] = data["property_standard_unit"].map(clean)
    data["性能值"] = data["property_value_raw"].map(clean)
    decisions = data.apply(
        lambda row: is_valid_unit(row["性能"], row["原始单位"], row["标准单位"]),
        axis=1,
        result_type="expand",
    )
    data["单位是否匹配性能"] = decisions[0].map(lambda x: "yes" if bool(x) else "no")
    data["单位规则类型"] = decisions[1]
    data["单位规则说明"] = decisions[2]
    return data


def value_distribution(valid_df: pd.DataFrame) -> pd.DataFrame:
    data = valid_df[valid_df["单位是否匹配性能"] == "yes"].copy()
    data["标准化数值"] = pd.to_numeric(data["property_standard_value"], errors="coerce")
    data = data[(data["标准单位"] != "") & data["标准化数值"].notna()]
    data = data[data["标准化数值"].map(lambda x: math.isfinite(float(x)))]
    data = data[~data["性能"].isin(CATEGORICAL_NO_UNIT_PROPERTIES)]
    grouped = (
        data.groupby(["统计口径", "性能", "标准单位"], dropna=False)
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
    return grouped.sort_values(["统计口径", "性能", "数据量"], ascending=[True, True, False])


def valid_raw_unit_frequency(valid_df: pd.DataFrame) -> pd.DataFrame:
    data = valid_df[(valid_df["性能值"] != "") & (valid_df["单位是否匹配性能"] == "yes")].copy()
    data["原始单位"] = data["原始单位"].replace("", "(missing unit)")
    grouped = (
        data.groupby(["统计口径", "性能", "原始单位"], dropna=False)
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
    totals = grouped.groupby(["统计口径", "性能"])["频次"].transform("sum")
    grouped["占该性能比例(%)"] = (grouped["频次"] / totals * 100).round(2)
    return grouped[
        ["统计口径", "性能", "原始单位", "频次", "占该性能比例(%)", "涉及样品记录数", "映射后的标准单位"]
    ].sort_values(["统计口径", "性能", "频次"], ascending=[True, True, False])


def invalid_unit_frequency(valid_df: pd.DataFrame) -> pd.DataFrame:
    data = valid_df[(valid_df["性能值"] != "") & (valid_df["单位是否匹配性能"] == "no")].copy()
    data["原始单位"] = data["原始单位"].replace("", "(missing unit)")
    data["标准单位"] = data["标准单位"].replace("", "(not standardized)")
    grouped = (
        data.groupby(["统计口径", "性能", "原始单位", "标准单位", "单位规则类型", "单位规则说明"], dropna=False)
        .agg(频次=("record_id", "count"), 涉及样品记录数=("record_id", "nunique"))
        .reset_index()
    )
    return grouped.sort_values(["统计口径", "性能", "频次"], ascending=[True, True, False])


def validation_summary(valid_df: pd.DataFrame) -> pd.DataFrame:
    data = valid_df[valid_df["性能值"] != ""].copy()
    grouped = (
        data.groupby(["统计口径", "性能"], dropna=False)
        .agg(
            有性能值总数=("record_id", "count"),
            单位匹配数量=("单位是否匹配性能", lambda x: int((x == "yes").sum())),
            单位不匹配数量=("单位是否匹配性能", lambda x: int((x == "no").sum())),
            原始单位种类数=("原始单位", lambda x: len({clean(v) or "(missing unit)" for v in x})),
        )
        .reset_index()
    )
    grouped["单位匹配比例(%)"] = (grouped["单位匹配数量"] / grouped["有性能值总数"] * 100).round(2)
    return grouped.sort_values(["统计口径", "单位匹配比例(%)", "有性能值总数"], ascending=[True, False, False])


def rule_table() -> pd.DataFrame:
    rows = []
    for prop, rule in sorted(PROPERTY_GROUP_FOR_RULE.items()):
        rows.append({"性能": prop, "单位规则类型": rule, "单位规则说明": RULE_DESCRIPTIONS.get(rule, "")})
    for prop in sorted(CATEGORICAL_NO_UNIT_PROPERTIES):
        rows.append({"性能": prop, "单位规则类型": "categorical_no_unit", "单位规则说明": "分类/文本型性能，无单位是合理的"})
    return pd.DataFrame(rows)


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    all_long = read_many(CLEAN_TABLE_DIR, "ext-*_properties_long.csv")
    valuable_long = read_many(FINAL_TABLE_DIR, "final_ext-*_valuable_properties_long.csv")

    all_annotated = annotate_units(all_long, "全部清洗后抽取结果")
    valuable_annotated = annotate_units(valuable_long, "最终有价值数据")
    combined = pd.concat([all_annotated, valuable_annotated], ignore_index=True)

    distribution = value_distribution(combined)
    valid_units = valid_raw_unit_frequency(combined)
    invalid_units = invalid_unit_frequency(combined)
    summary = validation_summary(combined)
    rules = rule_table()

    distribution.to_csv(OUT_DIR / "property_value_distribution_statistics_valid_units.csv", index=False, encoding="utf-8-sig")
    valid_units.to_csv(OUT_DIR / "property_valid_raw_unit_frequency_statistics.csv", index=False, encoding="utf-8-sig")
    invalid_units.to_csv(OUT_DIR / "property_invalid_unit_candidates.csv", index=False, encoding="utf-8-sig")
    summary.to_csv(OUT_DIR / "property_unit_validation_summary.csv", index=False, encoding="utf-8-sig")
    rules.to_csv(OUT_DIR / "property_unit_validation_rules.csv", index=False, encoding="utf-8-sig")

    with pd.ExcelWriter(OUT_DIR / "property_distribution_and_unit_statistics_valid_units.xlsx", engine="openpyxl") as writer:
        distribution.to_excel(writer, index=False, sheet_name="value_distribution_valid")
        valid_units.to_excel(writer, index=False, sheet_name="valid_raw_unit_frequency")
        invalid_units.to_excel(writer, index=False, sheet_name="invalid_unit_candidates")
        summary.to_excel(writer, index=False, sheet_name="validation_summary")
        rules.to_excel(writer, index=False, sheet_name="unit_rules")
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

    print("Wrote valid-unit filtered statistics to", OUT_DIR)


if __name__ == "__main__":
    main()
