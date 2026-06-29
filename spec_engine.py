import os
import re
import numpy as np
import pandas as pd


# ============================================================
# 1. 基础配置
# ============================================================

LOT_DIR = "/content/lots"
TEMPLATE_FILE = "/content/template.xlsx"
ASSIGN_FILE = "/content/assign_standard.xlsx"
TARGET_FILE = "/content/target_spec.xlsx"
SIM_FILE = "/content/sim_value.xlsx"
OUTPUT_FILE = "/content/output/spec_assignment_output.xlsx"

INDEX_SHEET_NAME = "Index"

DEFAULT_MARGIN_FACTOR = 1.2

# 仿真值对比阈值
SIM_TOLERANCE_MEDIUM = 0.10
SIM_TOLERANCE_HIGH = 0.20

TEMP_GROUP_90 = [-55, -45, 25, 90]
TEMP_GROUP_110 = [-55, -45, 25, 90, 110]
TEMP_GROUP_130 = [-55, -45, 25, 90, 110, 130]

# 如果你的测试文件里 55 实际代表 -55，45 实际代表 -45，保持 True
ENABLE_NEGATIVE_TEMP_ALIAS = True

TEMP_ALIAS = {
    55: -55,
    45: -45,
}


# ============================================================
# 2. 通用工具函数
# ============================================================

def normalize_text(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def to_number(x):
    if pd.isna(x):
        return np.nan

    if isinstance(x, (int, float, np.number)):
        return float(x)

    s = str(x).strip()

    if s.upper() in ["", "-", "NA", "N/A", "NONE", "NULL"]:
        return np.nan

    s = s.replace(",", "")

    try:
        return float(s)
    except Exception:
        return np.nan


def round_value(x, digits=4):
    if pd.isna(x):
        return ""
    try:
        return round(float(x), digits)
    except Exception:
        return ""


def percent_value(x, digits=2):
    if pd.isna(x):
        return ""
    try:
        return round(float(x) * 100, digits)
    except Exception:
        return ""


def normalize_param_for_match(x):
    s = normalize_text(x)
    s = s.replace(" ", "")
    s = s.replace("-", "_")
    return s.upper()


def clean_column_name(x):
    s = normalize_text(x)
    s = s.replace(" ", "")
    s = s.replace("_", "")
    s = s.replace("-", "")
    s = s.replace("\n", "")
    s = s.replace("\r", "")
    s = s.replace("\t", "")
    s = s.replace("℃", "C")
    s = s.replace("°C", "C")
    s = s.replace("°", "")
    return s.upper()


def parse_temp(x):
    s = normalize_text(x)

    if not s:
        return None

    m = re.search(r"(-?\d+)", s)

    if not m:
        return None

    temp = int(m.group(1))

    if ENABLE_NEGATIVE_TEMP_ALIAS and temp in TEMP_ALIAS:
        temp = TEMP_ALIAS[temp]

    return temp


def clean_stat(x):
    s = normalize_text(x).lower()

    if "min" in s:
        return "Min"
    if "typ" in s or "avg" in s or "mean" in s:
        return "Typ"
    if "max" in s:
        return "Max"

    return ""


def is_prefix_match(actual_key, base_key):
    """
    避免 ICC1 误匹配 ICC10。
    例如：
    tSLCH_DRV10 -> tSLCH
    ICC3_1IO_DRV00_133MHz -> ICC3
    """

    if actual_key == base_key:
        return True

    if not actual_key.startswith(base_key):
        return False

    if len(actual_key) <= len(base_key):
        return False

    next_char = actual_key[len(base_key)]

    if next_char in ["_", "(", "[", "/", "."]:
        return True

    if actual_key.startswith(base_key + "DRV"):
        return True

    return False

def get_first_underscore_base_key(parameter):
    """
    提取第一个下划线前面的基础参数。

    例子：
    tSLCH_DRV10 -> TSLCH
    tCHSH_DRV00 -> TCHSH
    ICC3_1IO_DRV00_133MHz -> ICC3
    tCLQV_133MHz_DRV10 -> TCLQV
    """
    actual_key = normalize_param_for_match(parameter)

    if "_" in actual_key:
        return actual_key.split("_")[0]

    return actual_key
# ============================================================
# 3. 读取赋值标准文件
# ============================================================

def read_assign_standard(assign_file):
    df = pd.read_excel(assign_file, sheet_name=0)

    col_map = {}

    for col in df.columns:
        key = clean_column_name(col)

        if key in ["BASEPARAMETER", "PARAMETER", "SYMBOL", "基础参数", "标准参数", "参数"]:
            col_map[col] = "Base_Parameter"

        elif key in ["SPECTYPE", "TYPE", "规格类型", "参数类型", "赋值类型"]:
            col_map[col] = "Spec_Type"

        elif key in ["UNIT", "单位"]:
            col_map[col] = "Unit"

        elif key in ["ASSIGNRULE", "RULE", "赋值标准", "赋值规则", "取值规则"]:
            col_map[col] = "Assign_Rule"

        elif key in ["USEEDGE", "EDGE", "边沿", "使用边沿", "FERE", "FE_RE"]:
            col_map[col] = "Use_Edge"

    df = df.rename(columns=col_map)

    required_cols = ["Base_Parameter", "Assign_Rule"]

    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"赋值标准文件缺少必要列：{col}")

    # 没有这些列就自动补空
    for col in ["Spec_Type", "Unit", "Use_Edge"]:
        if col not in df.columns:
            df[col] = ""

    df["Base_Parameter"] = df["Base_Parameter"].astype(str).str.strip()
    df["Base_Key"] = df["Base_Parameter"].apply(normalize_param_for_match)

    # Use_Edge 规则：
    # 空白 = ALL
    # FE = 只看 FE
    # RE = 只看 RE
    # IGNORE = 不分析
    df["Use_Edge"] = df["Use_Edge"].fillna("").astype(str).str.upper().str.strip()
    df.loc[df["Use_Edge"] == "", "Use_Edge"] = "ALL"

    df = df[df["Base_Parameter"] != ""].copy()

    print("赋值标准文件读取完成，参数数量：", len(df))

    return df

# ============================================================
# 4. 读取目标规格文件
# ============================================================

def read_target_spec(target_file):
    df = pd.read_excel(target_file, sheet_name=0)

    col_map = {}

    for col in df.columns:
        key = clean_column_name(col)

        if key in ["BASEPARAMETER", "PARAMETER", "SYMBOL", "基础参数", "标准参数", "参数"]:
            col_map[col] = "Base_Parameter"

        elif key in ["SPECTYPE", "TYPE", "规格类型", "参数类型"]:
            col_map[col] = "Spec_Type"

        elif key in ["TARGET90C", "SPEC90C", "90C", "90", "目标规格90C", "90度目标规格"]:
            col_map[col] = "Target_90C"

        elif key in ["TARGET110C", "SPEC110C", "110C", "110", "目标规格110C", "110度目标规格"]:
            col_map[col] = "Target_110C"

        elif key in ["TARGET130C", "SPEC130C", "130C", "130", "目标规格130C", "130度目标规格"]:
            col_map[col] = "Target_130C"

        elif key in ["UNIT", "单位"]:
            col_map[col] = "Unit"

        elif key in ["REMARK", "备注", "说明"]:
            col_map[col] = "Remark"

    df = df.rename(columns=col_map)

    if "Base_Parameter" not in df.columns:
        raise ValueError("目标规格文件缺少必要列：Base_Parameter")

    for col in ["Spec_Type", "Target_90C", "Target_110C", "Target_130C", "Unit", "Remark"]:
        if col not in df.columns:
            df[col] = ""

    df["Base_Parameter"] = df["Base_Parameter"].astype(str).str.strip()
    df["Base_Key"] = df["Base_Parameter"].apply(normalize_param_for_match)

    df = df[df["Base_Parameter"] != ""].copy()

    print("目标规格文件读取完成，参数数量：", len(df))

    return df


# ============================================================
# 5. 读取仿真值文件
# ============================================================

def read_sim_value(sim_file):
    """
    推荐仿真值文件表头：

    Base_Parameter | Spec_Type | Sim_Typ_25C | Sim_Worst_90C | Sim_Worst_110C | Sim_Worst_130C | Unit | Remark
    """

    df = pd.read_excel(sim_file, sheet_name=0)

    col_map = {}

    for col in df.columns:
        key = clean_column_name(col)

        if key in ["BASEPARAMETER", "PARAMETER", "SYMBOL", "基础参数", "标准参数", "参数"]:
            col_map[col] = "Base_Parameter"

        elif key in ["SPECTYPE", "TYPE", "规格类型", "参数类型"]:
            col_map[col] = "Spec_Type"

        elif key in [
            "SIMTYP25C", "SIM25C", "25CSIM", "25CTYP", "SIMTYP",
            "仿真TYP25C", "仿真25C", "25度仿真TYP"
        ]:
            col_map[col] = "Sim_Typ_25C"

        elif key in [
            "SIMWORST90C", "SIM90C", "SIMUPTO90C", "90CSIM",
            "90C仿真", "仿真90C", "90度仿真"
        ]:
            col_map[col] = "Sim_Worst_90C"

        elif key in [
            "SIMWORST110C", "SIM110C", "SIMUPTO110C", "110CSIM",
            "110C仿真", "仿真110C", "110度仿真"
        ]:
            col_map[col] = "Sim_Worst_110C"

        elif key in [
            "SIMWORST130C", "SIM130C", "SIMUPTO130C", "130CSIM",
            "130C仿真", "仿真130C", "130度仿真"
        ]:
            col_map[col] = "Sim_Worst_130C"

        elif key in ["UNIT", "单位"]:
            col_map[col] = "Unit"

        elif key in ["REMARK", "备注", "说明"]:
            col_map[col] = "Remark"

    df = df.rename(columns=col_map)

    if "Base_Parameter" not in df.columns:
        raise ValueError("仿真值文件缺少必要列：Base_Parameter")

    for col in [
        "Spec_Type",
        "Sim_Typ_25C",
        "Sim_Worst_90C",
        "Sim_Worst_110C",
        "Sim_Worst_130C",
        "Unit",
        "Remark",
    ]:
        if col not in df.columns:
            df[col] = ""

    df["Base_Parameter"] = df["Base_Parameter"].astype(str).str.strip()
    df["Base_Key"] = df["Base_Parameter"].apply(normalize_param_for_match)

    df = df[df["Base_Parameter"] != ""].copy()

    print("仿真值文件读取完成，参数数量：", len(df))

    return df


# ============================================================
# 6. 参数匹配函数
# ============================================================

def get_best_match_rows(actual_parameter, df):
    """
    参数匹配优先级：

    1. 精确匹配实际测试参数
       tSLCH_DRV10 -> tSLCH_DRV10

    2. 匹配第一个 "_" 前面的基础参数
       tSLCH_DRV10 -> tSLCH
       ICC3_1IO_DRV00_133MHz -> ICC3

    3. 最长前缀匹配
       用于兼容少数特殊命名
    """

    if df is None or df.empty:
        return pd.DataFrame()

    actual_key = normalize_param_for_match(actual_parameter)
    base_key_from_underscore = get_first_underscore_base_key(actual_parameter)

    # 1. 精确匹配完整测试参数
    exact = df[df["Base_Key"] == actual_key]

    if not exact.empty:
        return exact.copy()

    # 2. 匹配第一个 "_" 前面的基础参数
    # 例如 tSLCH_DRV10 -> tSLCH
    base_exact = df[df["Base_Key"] == base_key_from_underscore]

    if not base_exact.empty:
        return base_exact.copy()

    # 3. 最长前缀匹配
    candidates = []

    for _, row in df.iterrows():
        standard_key = normalize_text(row.get("Base_Key", ""))

        if not standard_key:
            continue

        if is_prefix_match(actual_key, standard_key):
            candidates.append((len(standard_key), row))

    if not candidates:
        return pd.DataFrame()

    max_len = max(x[0] for x in candidates)
    best_rows = [x[1] for x in candidates if x[0] == max_len]

    return pd.DataFrame(best_rows)




def match_one_row(actual_parameter, spec_type, df):
    """
    匹配单行标准数据。
    如果 spec_type 有值，优先匹配相同 Spec_Type。
    """
    matched = get_best_match_rows(actual_parameter, df)

    if matched.empty:
        return None

    spec_type_key = normalize_text(spec_type).lower()

    if spec_type_key and "Spec_Type" in matched.columns:
        same_type = matched[
            matched["Spec_Type"].astype(str).str.lower().str.strip() == spec_type_key
        ]

        if not same_type.empty:
            return same_type.iloc[0]

    return matched.iloc[0]

# ============================================================
# 7. 读取汇总模板
# ============================================================

def read_template(template_file):
    """
    读取汇总模板 template.xlsx。

    支持模板格式：
    A: Parameter
    B: Spec_Type
    C: Simulated_90℃
    D: Simulated_110℃
    E: Simulated_130℃
    F: Target_90℃
    G: Target_110℃
    H: Target_130℃
    I: Unit
    J: Assign_Rule
    """

    df = pd.read_excel(template_file, sheet_name=0, header=None)

    header_row_idx = None

    for i in range(min(20, len(df))):
        row_values = [normalize_text(v) for v in df.iloc[i].tolist()]
        row_text = " ".join(row_values)

        if "Parameter" in row_text or "参数" in row_text:
            header_row_idx = i
            break

    if header_row_idx is None:
        raise ValueError("汇总模板没有找到表头行，请确认 A1 是否为 Parameter。")

    header = df.iloc[header_row_idx].tolist()

    col_map = {}

    for idx, col in enumerate(header):
        key = clean_column_name(col)

        if key in ["PARAMETER", "SYMBOL", "参数"]:
            col_map["Parameter"] = idx

        elif key in ["SPECTYPE", "TYPE", "规格类型", "参数类型"]:
            col_map["Spec_Type"] = idx

        elif key in [
            "SIMULATED90C", "SIMULATED90", "SIM90C", "SIM90",
            "90CSIMULATED", "90CSIM"
        ]:
            col_map["Simulated_90C"] = idx

        elif key in [
            "SIMULATED110C", "SIMULATED110", "SIM110C", "SIM110",
            "110CSIMULATED", "110CSIM"
        ]:
            col_map["Simulated_110C"] = idx

        elif key in [
            "SIMULATED130C", "SIMULATED130", "SIM130C", "SIM130",
            "130CSIMULATED", "130CSIM"
        ]:
            col_map["Simulated_130C"] = idx

        elif key in ["TARGET90C", "TARGET90", "90CTARGET", "90C目标", "目标规格90C"]:
            col_map["Target_90C"] = idx

        elif key in ["TARGET110C", "TARGET110", "110CTARGET", "110C目标", "目标规格110C"]:
            col_map["Target_110C"] = idx

        elif key in ["TARGET130C", "TARGET130", "130CTARGET", "130C目标", "目标规格130C"]:
            col_map["Target_130C"] = idx

        elif key in ["UNIT", "单位"]:
            col_map["Unit"] = idx

        elif key in ["ASSIGNRULE", "RULE", "赋值标准", "赋值规则"]:
            col_map["Assign_Rule"] = idx

    if "Parameter" not in col_map:
        raise ValueError("汇总模板缺少 Parameter 列，请确认 A列表头为 Parameter。")

    required_output_cols = [
        "Parameter",
        "Spec_Type",
        "Simulated_90C",
        "Simulated_110C",
        "Simulated_130C",
        "Target_90C",
        "Target_110C",
        "Target_130C",
        "Unit",
        "Assign_Rule",
    ]

    rows = []

    for i in range(header_row_idx + 1, len(df)):
        row = df.iloc[i]

        parameter = normalize_text(row[col_map["Parameter"]]) if "Parameter" in col_map else ""
        spec_type = normalize_text(row[col_map["Spec_Type"]]) if "Spec_Type" in col_map else ""

        simulated_90 = to_number(row[col_map["Simulated_90C"]]) if "Simulated_90C" in col_map else np.nan
        simulated_110 = to_number(row[col_map["Simulated_110C"]]) if "Simulated_110C" in col_map else np.nan
        simulated_130 = to_number(row[col_map["Simulated_130C"]]) if "Simulated_130C" in col_map else np.nan

        target_90 = to_number(row[col_map["Target_90C"]]) if "Target_90C" in col_map else np.nan
        target_110 = to_number(row[col_map["Target_110C"]]) if "Target_110C" in col_map else np.nan
        target_130 = to_number(row[col_map["Target_130C"]]) if "Target_130C" in col_map else np.nan

        unit = normalize_text(row[col_map["Unit"]]) if "Unit" in col_map else ""
        rule = normalize_text(row[col_map["Assign_Rule"]]) if "Assign_Rule" in col_map else ""

        if not parameter and not spec_type and not rule:
            continue

        rows.append({
            "Parameter": parameter,
            "Spec_Type": spec_type,
            "Simulated_90C": simulated_90,
            "Simulated_110C": simulated_110,
            "Simulated_130C": simulated_130,
            "Target_90C": target_90,
            "Target_110C": target_110,
            "Target_130C": target_130,
            "Unit": unit,
            "Assign_Rule": rule,
        })

    template_df = pd.DataFrame(rows, columns=required_output_cols)

    if not template_df.empty:
        template_df["Parameter"] = template_df["Parameter"].replace("", np.nan).ffill()
        template_df = template_df[template_df["Parameter"].notna()].copy()

    print("汇总模板读取完成，模板参数行数：", len(template_df))

    if len(template_df) == 0:
        print("提示：汇总模板只有表头，没有参数行。程序将从 Lot 数据中自动生成所有参数。")

    return template_df


# ============================================================
# 8. 读取 Lot Index 页
# ============================================================

def read_lot_index(file_path):
    lot_name = os.path.splitext(os.path.basename(file_path))[0]

    try:
        df = pd.read_excel(file_path, sheet_name=INDEX_SHEET_NAME, header=None)
    except Exception as e:
        print(f"[跳过] {lot_name}：没有找到 Index 页。错误信息：{e}")
        return pd.DataFrame()

    stat_row_idx = None

    for i in range(min(20, len(df))):
        row_values = [clean_stat(v) for v in df.iloc[i].tolist()]
        count_stat = sum([1 for v in row_values if v in ["Min", "Typ", "Max"]])

        if count_stat >= 3:
            stat_row_idx = i
            break

    if stat_row_idx is None:
        print(f"[跳过] {lot_name}：Index 页没有识别到 Min/Typ/Max 行。")
        return pd.DataFrame()

    temp_row_idx = max(stat_row_idx - 1, 0)

    temp_row = df.iloc[temp_row_idx].copy().ffill()
    stat_row = df.iloc[stat_row_idx].copy()

    value_cols = []

    for c in range(df.shape[1]):
        stat = clean_stat(stat_row.iloc[c])
        temp = parse_temp(temp_row.iloc[c])

        if stat in ["Min", "Typ", "Max"] and temp is not None:
            value_cols.append(c)

    if not value_cols:
        print(f"[跳过] {lot_name}：没有识别到有效温度数据列。")
        return pd.DataFrame()

    first_value_col = min(value_cols)

    records = []

    for r in range(stat_row_idx + 1, len(df)):
        parameter = normalize_text(df.iloc[r, 0])

        if not parameter:
            continue

        if parameter.lower() in ["symbol", "parameter", "参数", "test item", "item"]:
            continue

        test_condition = normalize_text(df.iloc[r, 2]) if df.shape[1] > 2 else ""

        unit = ""

        for c in range(0, first_value_col):
            cell = normalize_text(df.iloc[r, c])
            if cell in ["uA", "μA", "mA", "A", "V", "mV", "ns", "us", "μs", "ms", "s", "MHz"]:
                unit = cell
                break

        for c in value_cols:
            temp = parse_temp(temp_row.iloc[c])
            stat = clean_stat(stat_row.iloc[c])
            value = to_number(df.iloc[r, c])

            if temp is None or not stat or pd.isna(value):
                continue

            records.append({
                "Lot": lot_name,
                "Parameter": parameter,
                "Test_Condition": test_condition,
                "Unit_From_Lot": unit,
                "Temp": temp,
                "Stat": stat,
                "Value": value,
                "Source_File": os.path.basename(file_path),
            })

    lot_df = pd.DataFrame(records)

    print(f"{lot_name} 读取完成，数据点数量：{len(lot_df)}")

    return lot_df


def read_all_lots(lot_dir):
    all_data = []

    for file in os.listdir(lot_dir):
        if file.startswith("~$"):
            continue

        if not file.lower().endswith((".xlsx", ".xlsm")):
            continue

        file_path = os.path.join(lot_dir, file)

        print("正在读取 Lot 文件：", file)

        lot_df = read_lot_index(file_path)

        if not lot_df.empty:
            all_data.append(lot_df)

    if not all_data:
        raise ValueError("没有读取到任何 Lot 数据，请检查 lots 文件夹和每个 Lot 文件的 Index 页。")

    raw_df = pd.concat(all_data, ignore_index=True)

    print("所有 Lot 读取完成，总数据点数量：", len(raw_df))

    return raw_df


# ============================================================
# 9. 计算规则
# ============================================================

def infer_side(parameter, spec_type, rule):
    p = normalize_param_for_match(parameter)
    b = normalize_text(spec_type).lower()
    r = normalize_text(rule).lower()

    if "avg" in r or "typ" in r or "平均" in r:
        return "avg"

    if "min" in r or "/1.2" in r:
        return "min"

    if "max" in r or "*1.2" in r:
        return "max"

    if b == "min":
        return "min"

    if b == "max":
        return "max"

    if b == "typ":
        return "avg"

    if p.startswith(("VOH", "VIH")):
        return "min"

    return "max"


def parse_margin_factor(rule):
    r = normalize_text(rule)

    m = re.search(r"[\*/]\s*(\d+(\.\d+)?)", r)

    if m:
        return float(m.group(1))

    return DEFAULT_MARGIN_FACTOR


def get_avg_25(param_df):
    d = param_df[(param_df["Temp"] == 25) & (param_df["Stat"] == "Typ")]

    if d.empty:
        d = param_df[param_df["Temp"] == 25]

    if d.empty:
        return np.nan

    return d["Value"].mean()


def get_worst(param_df, temps, side):
    d = param_df[param_df["Temp"].isin(temps)]

    if d.empty:
        return np.nan

    if side == "max":
        d2 = d[d["Stat"] == "Max"]
        if d2.empty:
            d2 = d
        return d2["Value"].max()

    if side == "min":
        d2 = d[d["Stat"] == "Min"]
        if d2.empty:
            d2 = d
        return d2["Value"].min()

    return np.nan


def suggest_spec(value, side, factor):
    if pd.isna(value):
        return np.nan

    if side == "max":
        return value * factor

    if side == "min":
        return value / factor

    return value


def calc_delta_percent(test_value, sim_value):
    if pd.isna(test_value) or pd.isna(sim_value):
        return np.nan

    if sim_value == 0:
        return np.nan

    return (test_value - sim_value) / abs(sim_value)


def judge_target_risk_one(side, suggest_value, target_value, temp_label):
    if pd.isna(suggest_value) or pd.isna(target_value):
        return "Review", f"{temp_label} 缺少建议值或目标规格"

    if side == "max":
        if suggest_value > target_value:
            return "High", f"{temp_label} 建议值 {round_value(suggest_value)} > 目标规格 {round_value(target_value)}"
        if suggest_value > target_value * 0.9:
            return "Medium", f"{temp_label} 建议值 {round_value(suggest_value)} 接近目标规格 {round_value(target_value)}"
        return "Low", f"{temp_label} 建议值 {round_value(suggest_value)} 满足目标规格 {round_value(target_value)}"

    if side == "min":
        if suggest_value < target_value:
            return "High", f"{temp_label} 建议值 {round_value(suggest_value)} < 目标规格 {round_value(target_value)}"
        if suggest_value < target_value * 1.1:
            return "Medium", f"{temp_label} 建议值 {round_value(suggest_value)} 接近目标规格 {round_value(target_value)}"
        return "Low", f"{temp_label} 建议值 {round_value(suggest_value)} 满足目标规格 {round_value(target_value)}"

    return "Review", f"{temp_label} Typ/Avg 类型暂不做目标规格风险判定"


def judge_sim_risk_one(side, test_value, sim_value, temp_label):
    if pd.isna(test_value) or pd.isna(sim_value):
        return "Review", f"{temp_label} 缺少实测值或仿真值"

    if sim_value == 0:
        return "Review", f"{temp_label} 仿真值为 0，无法计算偏差比例"

    delta = calc_delta_percent(test_value, sim_value)

    if side == "max":
        if test_value > sim_value * (1 + SIM_TOLERANCE_HIGH):
            return "High", f"{temp_label} 实测 {round_value(test_value)} > 仿真 {round_value(sim_value)} +20%"
        if test_value > sim_value * (1 + SIM_TOLERANCE_MEDIUM):
            return "Medium", f"{temp_label} 实测 {round_value(test_value)} > 仿真 {round_value(sim_value)} +10%"
        return "Low", f"{temp_label} 实测与仿真差异可接受，偏差 {percent_value(delta)}%"

    if side == "min":
        if test_value < sim_value / (1 + SIM_TOLERANCE_HIGH):
            return "High", f"{temp_label} 实测 {round_value(test_value)} < 仿真 {round_value(sim_value)} /1.2"
        if test_value < sim_value / (1 + SIM_TOLERANCE_MEDIUM):
            return "Medium", f"{temp_label} 实测 {round_value(test_value)} < 仿真 {round_value(sim_value)} /1.1"
        return "Low", f"{temp_label} 实测与仿真差异可接受，偏差 {percent_value(delta)}%"

    if abs(delta) > SIM_TOLERANCE_HIGH:
        return "High", f"{temp_label} Typ 实测与仿真偏差超过 20%，偏差 {percent_value(delta)}%"

    if abs(delta) > SIM_TOLERANCE_MEDIUM:
        return "Medium", f"{temp_label} Typ 实测与仿真偏差超过 10%，偏差 {percent_value(delta)}%"

    return "Low", f"{temp_label} Typ 实测与仿真差异可接受，偏差 {percent_value(delta)}%"


def combine_risk(risk_items):
    risks = [x[0] for x in risk_items]
    reasons = [x[1] for x in risk_items if x[1]]

    if "High" in risks:
        return "High", "；".join([r for risk, r in risk_items if risk == "High"])

    if "Medium" in risks:
        return "Medium", "；".join([r for risk, r in risk_items if risk == "Medium"])

    if all(x == "Low" for x in risks):
        return "Low", "；".join(reasons)

    return "Review", "；".join(reasons)


# ============================================================
# 10. 行信息补充
# ============================================================

def apply_assign_info(row_dict, parameter, assign_df):
    assign_row = match_one_row(parameter, row_dict.get("Spec_Type", ""), assign_df)

    assign_parameter = ""

    if assign_row is None:
        return row_dict, assign_parameter

    assign_parameter = assign_row.get("Base_Parameter", "")

    if not normalize_text(row_dict.get("Spec_Type", "")):
        row_dict["Spec_Type"] = assign_row.get("Spec_Type", "")

    if not normalize_text(row_dict.get("Unit", "")):
        row_dict["Unit"] = assign_row.get("Unit", "")

    if not normalize_text(row_dict.get("Assign_Rule", "")):
        row_dict["Assign_Rule"] = assign_row.get("Assign_Rule", "")

    return row_dict, assign_parameter


def apply_target_info(row_dict, parameter, target_df):
    target_row = match_one_row(parameter, row_dict.get("Spec_Type", ""), target_df)

    target_parameter = ""

    if target_row is None:
        return row_dict, target_parameter

    target_parameter = target_row.get("Base_Parameter", "")

    if not normalize_text(row_dict.get("Spec_Type", "")):
        row_dict["Spec_Type"] = target_row.get("Spec_Type", "")

    if not pd.isna(to_number(target_row.get("Target_90C", np.nan))):
        row_dict["Target_90C"] = to_number(target_row.get("Target_90C"))

    if not pd.isna(to_number(target_row.get("Target_110C", np.nan))):
        row_dict["Target_110C"] = to_number(target_row.get("Target_110C"))

    if not pd.isna(to_number(target_row.get("Target_130C", np.nan))):
        row_dict["Target_130C"] = to_number(target_row.get("Target_130C"))

    if not normalize_text(row_dict.get("Unit", "")):
        row_dict["Unit"] = target_row.get("Unit", "")

    return row_dict, target_parameter


def get_sim_info(parameter, spec_type, sim_df):
    sim_row = match_one_row(parameter, spec_type, sim_df)

    if sim_row is None:
        return {
            "Sim_Parameter": "",
            "Sim_Typ_25C": np.nan,
            "Sim_Worst_90C": np.nan,
            "Sim_Worst_110C": np.nan,
            "Sim_Worst_130C": np.nan,
        }

    return {
        "Sim_Parameter": sim_row.get("Base_Parameter", ""),
        "Sim_Typ_25C": to_number(sim_row.get("Sim_Typ_25C", np.nan)),
        "Sim_Worst_90C": to_number(sim_row.get("Sim_Worst_90C", np.nan)),
        "Sim_Worst_110C": to_number(sim_row.get("Sim_Worst_110C", np.nan)),
        "Sim_Worst_130C": to_number(sim_row.get("Sim_Worst_130C", np.nan)),
    }


# ============================================================
# 11. 生成 Summary
# ============================================================

def build_summary(template_df, raw_df, assign_df, target_df, sim_df):
    all_params_from_lots = sorted(raw_df["Parameter"].dropna().unique().tolist())
    all_params_from_template = template_df["Parameter"].dropna().unique().tolist()

    template_keys = set([normalize_param_for_match(p) for p in all_params_from_template])

    add_rows_list = []

    for p in all_params_from_lots:
        if normalize_param_for_match(p) not in template_keys:

            assign_rows = get_best_match_rows(p, assign_df)

            if not assign_rows.empty:
                for _, assign_row in assign_rows.iterrows():
                    add_rows_list.append({
                        "Parameter": p,
                        "Spec_Type": assign_row.get("Spec_Type", ""),
                        "Simulated_90C": np.nan,
                        "Simulated_110C": np.nan,
                        "Simulated_130C": np.nan,
                        "Target_90C": np.nan,
                        "Target_110C": np.nan,
                        "Target_130C": np.nan,
                        "Unit": assign_row.get("Unit", ""),
                        "Assign_Rule": assign_row.get("Assign_Rule", ""),
                    })
            else:
                add_rows_list.append({
                    "Parameter": p,
                    "Spec_Type": "",
                    "Simulated_90C": np.nan,
                    "Simulated_110C": np.nan,
                    "Simulated_130C": np.nan,
                    "Target_90C": np.nan,
                    "Target_110C": np.nan,
                    "Target_130C": np.nan,
                    "Unit": "",
                    "Assign_Rule": "AUTO_MAX*1.2",
                })

    if add_rows_list:
        add_rows = pd.DataFrame(add_rows_list)
        template_df = pd.concat([template_df, add_rows], ignore_index=True)

    summary_rows = []

    for _, row in template_df.iterrows():
        parameter = row["Parameter"]

        row_dict = {
            "Parameter": parameter,
            "Spec_Type": row.get("Spec_Type", ""),
            "Simulated_90C": row.get("Simulated_90C", np.nan),
            "Simulated_110C": row.get("Simulated_110C", np.nan),
            "Simulated_130C": row.get("Simulated_130C", np.nan),
            "Target_90C": row.get("Target_90C", np.nan),
            "Target_110C": row.get("Target_110C", np.nan),
            "Target_130C": row.get("Target_130C", np.nan),
            "Unit": row.get("Unit", ""),
            "Assign_Rule": row.get("Assign_Rule", ""),
        }

        row_dict, assign_parameter = apply_assign_info(row_dict, parameter, assign_df)
        row_dict, target_parameter = apply_target_info(row_dict, parameter, target_df)

        spec_type = row_dict["Spec_Type"]
        rule = row_dict["Assign_Rule"]

        if not normalize_text(rule):
            rule = "AUTO_MAX*1.2"
            row_dict["Assign_Rule"] = rule

        param_df = raw_df[raw_df["Parameter"] == parameter]

        side = infer_side(parameter, spec_type, rule)
        factor = parse_margin_factor(rule)

        avg_25 = get_avg_25(param_df)

        worst_90 = get_worst(param_df, TEMP_GROUP_90, side)
        worst_110 = get_worst(param_df, TEMP_GROUP_110, side)
        worst_130 = get_worst(param_df, TEMP_GROUP_130, side)

        suggest_typ = avg_25
        suggest_90 = suggest_spec(worst_90, side, factor)
        suggest_110 = suggest_spec(worst_110, side, factor)
        suggest_130 = suggest_spec(worst_130, side, factor)

        target_risk_90 = judge_target_risk_one(side, suggest_90, row_dict["Target_90C"], "UpTo90C")
        target_risk_110 = judge_target_risk_one(side, suggest_110, row_dict["Target_110C"], "UpTo110C")
        target_risk_130 = judge_target_risk_one(side, suggest_130, row_dict["Target_130C"], "UpTo130C")

        target_risk, target_risk_reason = combine_risk([
            target_risk_90,
            target_risk_110,
            target_risk_130
        ])

        sim_info = get_sim_info(parameter, spec_type, sim_df)

        sim_typ = sim_info["Sim_Typ_25C"]
        sim_90 = sim_info["Sim_Worst_90C"]
        sim_110 = sim_info["Sim_Worst_110C"]
        sim_130 = sim_info["Sim_Worst_130C"]

        delta_typ = calc_delta_percent(avg_25, sim_typ)
        delta_90 = calc_delta_percent(worst_90, sim_90)
        delta_110 = calc_delta_percent(worst_110, sim_110)
        delta_130 = calc_delta_percent(worst_130, sim_130)

        sim_risk_typ = judge_sim_risk_one("avg", avg_25, sim_typ, "25C Typ")
        sim_risk_90 = judge_sim_risk_one(side, worst_90, sim_90, "UpTo90C")
        sim_risk_110 = judge_sim_risk_one(side, worst_110, sim_110, "UpTo110C")
        sim_risk_130 = judge_sim_risk_one(side, worst_130, sim_130, "UpTo130C")

        sim_risk, sim_risk_reason = combine_risk([
            sim_risk_typ,
            sim_risk_90,
            sim_risk_110,
            sim_risk_130
        ])

        summary_rows.append({
            "Parameter": parameter,
            "Spec_Type": spec_type,
            "Simulated_90C": round_value(row_dict["Simulated_90C"]),
            "Simulated_110C": round_value(row_dict["Simulated_110C"]),
            "Simulated_130C": round_value(row_dict["Simulated_130C"]),

            "Target_90C": round_value(row_dict["Target_90C"]),
            "Target_110C": round_value(row_dict["Target_110C"]),
            "Target_130C": round_value(row_dict["Target_130C"]),
            "Unit": row_dict["Unit"],
            "Assign_Rule": rule,
            "Analysis_Direction": side,

            "Test_Typ_25C": round_value(avg_25),
            "Test_Worst_UpTo90C": round_value(worst_90),
            "Test_Worst_UpTo110C": round_value(worst_110),
            "Test_Worst_UpTo130C": round_value(worst_130),

            "Suggest_Typ_25C": round_value(suggest_typ),
            "Suggest_Spec_UpTo90C": round_value(suggest_90),
            "Suggest_Spec_UpTo110C": round_value(suggest_110),
            "Suggest_Spec_UpTo130C": round_value(suggest_130),

            "Target_Risk": target_risk,
            "Target_Risk_Reason": target_risk_reason,

            "Sim_Parameter": sim_info["Sim_Parameter"],
            "Sim_Typ_25C": round_value(sim_typ),
            "Sim_Worst_90C": round_value(sim_90),
            "Sim_Worst_110C": round_value(sim_110),
            "Sim_Worst_130C": round_value(sim_130),

            "Delta_vs_Sim_Typ_25C_%": percent_value(delta_typ),
            "Delta_vs_Sim_90C_%": percent_value(delta_90),
            "Delta_vs_Sim_110C_%": percent_value(delta_110),
            "Delta_vs_Sim_130C_%": percent_value(delta_130),

            "Sim_Risk": sim_risk,
            "Sim_Risk_Reason": sim_risk_reason,

            "Assign_Parameter": assign_parameter,
            "Target_Parameter": target_parameter,
        })

    summary_df = pd.DataFrame(summary_rows)

    print("Summary 生成完成，参数行数：", len(summary_df))

    return summary_df


# ============================================================
# 12. Raw Pivot 和检查表
# ============================================================

def build_raw_pivot(raw_df):
    df = raw_df.copy()

    df["Raw_Column"] = (
        df["Lot"].astype(str)
        + "_"
        + df["Temp"].astype(str)
        + "C_"
        + df["Stat"].astype(str)
    )

    pivot = df.pivot_table(
        index="Parameter",
        columns="Raw_Column",
        values="Value",
        aggfunc="first"
    ).reset_index()

    pivot.columns.name = None

    print("Raw Pivot 生成完成，参数数量：", len(pivot))

    return pivot


def build_need_check_assign(summary_df):
    return summary_df[
        (summary_df["Assign_Parameter"].astype(str).str.strip() == "") |
        (summary_df["Assign_Rule"].astype(str).str.contains("AUTO", case=False, na=False))
    ].copy()


def build_need_check_target(summary_df):
    return summary_df[
        (summary_df["Target_Parameter"].astype(str).str.strip() == "") |
        (
            (summary_df["Target_90C"].astype(str).str.strip() == "") &
            (summary_df["Target_110C"].astype(str).str.strip() == "") &
            (summary_df["Target_130C"].astype(str).str.strip() == "")
        )
    ].copy()


def build_need_check_sim(summary_df):
    return summary_df[
        (summary_df["Sim_Parameter"].astype(str).str.strip() == "") |
        (
            (summary_df["Sim_Typ_25C"].astype(str).str.strip() == "") &
            (summary_df["Sim_Worst_90C"].astype(str).str.strip() == "") &
            (summary_df["Sim_Worst_110C"].astype(str).str.strip() == "") &
            (summary_df["Sim_Worst_130C"].astype(str).str.strip() == "")
        )
    ].copy()


# ============================================================
# 13. 导出 Excel
# ============================================================

def export_excel(
    summary_df,
    raw_df,
    raw_pivot_df,
    need_assign_df,
    need_target_df,
    need_sim_df,
    output_file
):
    os.makedirs(os.path.dirname(output_file), exist_ok=True)

    final_df = summary_df.merge(raw_pivot_df, on="Parameter", how="left")

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        final_df.to_excel(writer, sheet_name="Summary", index=False)
        raw_df.to_excel(writer, sheet_name="Raw_Long", index=False)
        raw_pivot_df.to_excel(writer, sheet_name="Raw_Pivot", index=False)
        need_assign_df.to_excel(writer, sheet_name="Need_Check_Assign", index=False)
        need_target_df.to_excel(writer, sheet_name="Need_Check_Target", index=False)
        need_sim_df.to_excel(writer, sheet_name="Need_Check_Sim", index=False)

        workbook = writer.book

        header_fmt = workbook.add_format({
            "bold": True,
            "bg_color": "#D9EAD3",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
        })

        high_fmt = workbook.add_format({
            "bg_color": "#F4CCCC",
            "font_color": "#990000",
            "bold": True,
            "border": 1,
            "align": "center",
        })

        medium_fmt = workbook.add_format({
            "bg_color": "#FFF2CC",
            "font_color": "#7F6000",
            "bold": True,
            "border": 1,
            "align": "center",
        })

        low_fmt = workbook.add_format({
            "bg_color": "#D9EAD3",
            "font_color": "#274E13",
            "bold": True,
            "border": 1,
            "align": "center",
        })

        review_fmt = workbook.add_format({
            "bg_color": "#D9EAF7",
            "font_color": "#0B5394",
            "bold": True,
            "border": 1,
            "align": "center",
        })

        sheets = {
            "Summary": final_df,
            "Raw_Long": raw_df,
            "Raw_Pivot": raw_pivot_df,
            "Need_Check_Assign": need_assign_df,
            "Need_Check_Target": need_target_df,
            "Need_Check_Sim": need_sim_df,
        }

        for sheet_name, df_sheet in sheets.items():
            ws = writer.sheets[sheet_name]

            max_row = max(len(df_sheet), 1)
            max_col = max(len(df_sheet.columns), 1)

            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, max_row, max_col - 1)

            for col_num, col_name in enumerate(df_sheet.columns):
                ws.write(0, col_num, col_name, header_fmt)

            ws.set_column(0, 0, 28)
            ws.set_column(1, 8, 16)
            ws.set_column(9, 20, 20)
            ws.set_column(21, 40, 22)
            ws.set_column(41, 300, 18)

        ws = writer.sheets["Summary"]

        for risk_col_name in ["Target_Risk", "Sim_Risk"]:
            if risk_col_name in final_df.columns:
                risk_col = final_df.columns.get_loc(risk_col_name)
                last_row = len(final_df)

                ws.conditional_format(1, risk_col, last_row, risk_col, {
                    "type": "text",
                    "criteria": "containing",
                    "value": "High",
                    "format": high_fmt,
                })

                ws.conditional_format(1, risk_col, last_row, risk_col, {
                    "type": "text",
                    "criteria": "containing",
                    "value": "Medium",
                    "format": medium_fmt,
                })

                ws.conditional_format(1, risk_col, last_row, risk_col, {
                    "type": "text",
                    "criteria": "containing",
                    "value": "Low",
                    "format": low_fmt,
                })

                ws.conditional_format(1, risk_col, last_row, risk_col, {
                    "type": "text",
                    "criteria": "containing",
                    "value": "Review",
                    "format": review_fmt,
                })

    print("Excel 输出完成：", output_file)

# ============================================================
# 14. Streamlit 调用入口
# ============================================================

def run_analysis(
    lot_dir,
    template_file,
    assign_file,
    target_file,
    sim_file,
    output_file
):
    """
    Streamlit app.py 调用的主入口函数。
    所有文件路径由 app.py 在临时目录中生成并传入。
    """
    template_df = read_template(template_file)
    assign_df = read_assign_standard(assign_file)
    target_df = read_target_spec(target_file)
    sim_df = read_sim_value(sim_file)
    raw_df = read_all_lots(lot_dir)

    summary_df = build_summary(
        template_df,
        raw_df,
        assign_df,
        target_df,
        sim_df
    )

    raw_pivot_df = build_raw_pivot(raw_df)

    need_assign_df = build_need_check_assign(summary_df)
    need_target_df = build_need_check_target(summary_df)
    need_sim_df = build_need_check_sim(summary_df)

    export_excel(
        summary_df,
        raw_df,
        raw_pivot_df,
        need_assign_df,
        need_target_df,
        need_sim_df,
        output_file
    )

    return {
        "summary_df": summary_df,
        "raw_df": raw_df,
        "raw_pivot_df": raw_pivot_df,
        "need_assign_df": need_assign_df,
        "need_target_df": need_target_df,
        "need_sim_df": need_sim_df,
        "output_file": output_file,
    }

