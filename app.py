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
    
def detect_edge_from_text(x):
    """
    从参数名、测试条件、表头中识别 FE / RE。
    """
    s = normalize_text(x).upper()

    if not s:
        return ""

    if "FALLING" in s:
        return "FE"

    if "RISING" in s:
        return "RE"

    # 常见写法
    if re.search(r"(^|[^A-Z0-9])FE([^A-Z0-9]|$)", s):
        return "FE"

    if re.search(r"(^|[^A-Z0-9])RE([^A-Z0-9]|$)", s):
        return "RE"

    if "_FE" in s or s.endswith("FE"):
        return "FE"

    if "_RE" in s or s.endswith("RE"):
        return "RE"

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
def get_icc3_freq_base_key(parameter):
    """
    专门处理 ICC3_xIO_DRV00_xxMHz 这类参数。

    例子：
    ICC3_1IO_DRV00_80Mhz  -> ICC3_80MHZ
    ICC3_2IO_DRV00_114Mhz -> ICC3_114MHZ
    ICC3_4IO_DRV00_133Mhz -> ICC3_133MHZ

    这样可以匹配目标规格表里的：
    ICC3_80MHz
    ICC3_114MHz
    ICC3_133MHz
    """
    actual_key = normalize_param_for_match(parameter)

    if not actual_key.startswith("ICC3_"):
        return ""

    m = re.search(r"(80|100|104|108|114|120|133)\s*MHZ", actual_key, re.IGNORECASE)

    if not m:
        return ""

    freq = m.group(1)

    return normalize_param_for_match(f"ICC3_{freq}MHz")


def get_candidate_match_keys(parameter):
    """
    为一个实际测试参数生成多个可能的匹配 Key。
    匹配顺序很重要：越精确的放越前面。
    """
    actual_key = normalize_param_for_match(parameter)

    keys = []

    # 1. 完整参数
    keys.append(actual_key)

    # 2. ICC3 特殊规则：ICC3_4IO_DRV00_80Mhz -> ICC3_80MHz
    icc3_key = get_icc3_freq_base_key(parameter)
    if icc3_key:
        keys.append(icc3_key)

    # 3. 第一个 "_" 前面的基础参数
    first_base_key = get_first_underscore_base_key(parameter)
    if first_base_key:
        keys.append(first_base_key)

    # 去重但保持顺序
    result = []
    for k in keys:
        if k and k not in result:
            result.append(k)

    return result
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
            "SIMULATED90C", "SIMULATED90", "90CSIMULATED",
            "90C仿真", "仿真90C", "90度仿真", "SIMULATED_90C"
        ]:
            col_map[col] = "Sim_Worst_90C"

        elif key in [
            "SIMWORST110C", "SIM110C", "SIMUPTO110C", "110CSIM",
            "SIMULATED110C", "SIMULATED110", "110CSIMULATED",
            "110C仿真", "仿真110C", "110度仿真", "SIMULATED_110C"
        ]:
            col_map[col] = "Sim_Worst_110C"

        elif key in [
            "SIMWORST130C", "SIM130C", "SIMUPTO130C", "130CSIM",
            "SIMULATED130C", "SIMULATED130", "130CSIMULATED",
            "130C仿真", "仿真130C", "130度仿真", "SIMULATED_130C"
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

    1. 精确匹配完整参数
       ICC3_4IO_DRV00_80MHz -> ICC3_4IO_DRV00_80MHz

    2. 特殊规则匹配
       ICC3_4IO_DRV00_80MHz -> ICC3_80MHz

    3. 第一个 "_" 前基础参数
       tSLCH_DRV10 -> tSLCH

    4. 最长前缀匹配
    """

    if df is None or df.empty:
        return pd.DataFrame()

    actual_key = normalize_param_for_match(actual_parameter)

    # 1/2/3. 按候选 Key 顺序精确匹配
    candidate_keys = get_candidate_match_keys(actual_parameter)

    for key in candidate_keys:
        exact = df[df["Base_Key"] == key]

        if not exact.empty:
            return exact.copy()

    # 4. 最长前缀匹配
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


def is_typ_spec_value(x):
    spec = normalize_text(x).lower()
    return spec in ["typ", "type", "avg", "mean", "典型", "典型值"]


def has_typ_definition(parameter, assign_df, target_df, sim_df):
    """
    判断该参数是否在赋值标准 / 目标规格 / 仿真值中定义了 typ。
    如果三份文件都没有 typ 定义，则 Suggest_Typ_25C 留空。
    """
    for df in [assign_df, target_df, sim_df]:
        matched = get_best_match_rows(parameter, df)

        if matched.empty:
            continue

        if "Spec_Type" in matched.columns:
            if matched["Spec_Type"].apply(is_typ_spec_value).any():
                return True

        if "Sim_Typ_25C" in matched.columns:
            sim_typ_values = matched["Sim_Typ_25C"].apply(to_number)
            if sim_typ_values.notna().any():
                return True

    return False


def is_mhz_unit(unit):
    return normalize_text(unit).replace(" ", "").upper() == "MHZ"


def format_suggest_value(value, unit):
    """MHz 单位的 Suggest 输出取整，其余按默认小数位。"""
    if pd.isna(value):
        return ""

    if is_mhz_unit(unit):
        try:
            return int(round(float(value)))
        except Exception:
            return ""

    return round_value(value)

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

    value_infos = []

    for c in range(df.shape[1]):
        stat = clean_stat(stat_row.iloc[c])
        temp = parse_temp(temp_row.iloc[c])

        if stat in ["Min", "Typ", "Max"] and temp is not None:
            # 从该列上方所有表头信息里识别 FE / RE
            header_text = " ".join([
                normalize_text(df.iloc[rr, c])
                for rr in range(0, stat_row_idx + 1)
            ])

            edge = detect_edge_from_text(header_text)

            value_infos.append({
                 "col": c,
                 "temp": temp,
                 "stat": stat,
                 "edge": edge,
            })
  
    if not value_infos:
         print(f"[跳过] {lot_name}：没有识别到有效温度数据列。")
         return pd.DataFrame()

    first_value_col = min([x["col"] for x in value_infos])
    # 识别 Edge 列，一般表头为 Edge / Edge.
    edge_col = None

    for c in range(0, first_value_col):
        header_text = " ".join([
            normalize_text(df.iloc[rr, c])
            for rr in range(0, stat_row_idx + 1)
        ])

        header_key = clean_column_name(header_text)

        if "EDGE" in header_key or "边沿" in header_text:
            edge_col = c
            break

    if edge_col is not None:
        print(f"{lot_name} 识别到 Edge 列：第 {edge_col + 1} 列")
    else:
        print(f"{lot_name} 未识别到 Edge 列，将尝试从行描述中识别 FE/RE")
        
    records = []
    last_parameter = ""

    for r in range(stat_row_idx + 1, len(df)):
        raw_parameter = normalize_text(df.iloc[r, 0])

        # 当前行 Edge 列的值，例如 RE / FE
        edge_cell = ""
        if edge_col is not None:
            edge_cell = normalize_text(df.iloc[r, edge_col])

        edge_from_edge_col = detect_edge_from_text(edge_cell)

        # 如果当前行 A列有参数名，则更新 last_parameter
        if raw_parameter and raw_parameter.lower() not in ["symbol", "parameter", "参数", "test item", "item"]:
            last_parameter = raw_parameter

        # 如果当前行 A列为空，但是 Edge 列有 FE/RE，则沿用上一行参数名
        parameter = raw_parameter if raw_parameter else last_parameter

        if not parameter:
            continue

        if parameter.lower() in ["symbol", "parameter", "参数", "test item", "item"]:
            continue

        # 从当前行左侧信息中读取测试条件
        row_meta_cells = []

        for cc in range(0, first_value_col):
            row_meta_cells.append(normalize_text(df.iloc[r, cc]))

        row_meta_text = " ".join(row_meta_cells)

        test_condition = normalize_text(df.iloc[r, 2]) if df.shape[1] > 2 else row_meta_text

        # FE/RE 优先从 Edge 列读取
        row_edge = edge_from_edge_col if edge_from_edge_col else detect_edge_from_text(row_meta_text)

        unit = ""

        for c in range(0, first_value_col):
            cell = normalize_text(df.iloc[r, c])
            if cell in ["uA", "μA", "mA", "A", "V", "mV", "ns", "us", "μs", "ms", "s", "MHz"]:
                unit = cell
                break

        for info in value_infos:
            c = info["col"]
            temp = info["temp"]
            stat = info["stat"]
            edge = info["edge"]

            value = to_number(df.iloc[r, c])

            if temp is None or not stat or pd.isna(value):
                continue
 
            # 如果行测试条件里也有 FE / RE，优先用行里的
            condition_edge = detect_edge_from_text(test_condition)
            parameter_edge = detect_edge_from_text(parameter)

            if edge_from_edge_col:
                final_edge = edge_from_edge_col
            elif parameter_edge:
                final_edge = parameter_edge
            elif row_edge:
                final_edge = row_edge
            elif condition_edge:
                final_edge = condition_edge
            else:
                final_edge = edge

            records.append({
                 "Lot": lot_name,
                 "Parameter": parameter,
                 "Test_Condition": test_condition,
                 "Edge": final_edge,
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


def judge_target_risk_one(side, test_worst, suggest_value, target_value, sim_value, temp_label):
    """
    Target_Risk 评估逻辑：
    - 使用 Suggest_Spec、Test_Worst 与 limit 对比。
    - limit = Worst(目标规格, 仿真值)：
      * max 类型：目标规格/仿真值中取较大的值作为 limit。
      * min 类型：目标规格/仿真值中取较小的值作为 limit。
    - 目标规格为空则忽略目标规格，仿真值为空则忽略仿真值。
    - 两者都为空，该温度段返回 Review。

    max 类型：
      limit > Suggest_Spec                 -> Low
      Suggest_Spec > limit >= Test_Worst   -> Medium
      limit < Test_Worst                   -> High

    min 类型：
      limit < Suggest_Spec                 -> Low
      Suggest_Spec > limit >= Test_Worst   -> Medium
      limit > Test_Worst                   -> High
    """
    if pd.isna(test_worst) and pd.isna(suggest_value):
        return "Review", f"{temp_label} 缺少测试最差值和建议规格值"

    if pd.isna(suggest_value):
        return "Review", f"{temp_label} 缺少建议规格值"

    if pd.isna(test_worst):
        return "Review", f"{temp_label} 缺少测试最差值"

    limits = []

    if not pd.isna(target_value):
        limits.append(("目标规格", target_value))

    if not pd.isna(sim_value):
        limits.append(("仿真值", sim_value))

    if not limits:
        return "Review", f"{temp_label} 目标规格和仿真值均为空，跳过风险判断"

    if side == "max":
        limit_label, limit = max(limits, key=lambda x: x[1])

        if limit < test_worst:
            return (
                "High",
                f"{temp_label} limit({limit_label}) {round_value(limit)} < Test_Worst {round_value(test_worst)}"
            )

        if suggest_value > limit >= test_worst:
            return (
                "Medium",
                f"{temp_label} Suggest_Spec {round_value(suggest_value)} > limit({limit_label}) {round_value(limit)} >= Test_Worst {round_value(test_worst)}"
            )

        return (
            "Low",
            f"{temp_label} limit({limit_label}) {round_value(limit)} > Suggest_Spec {round_value(suggest_value)}"
        )

    if side == "min":
        limit_label, limit = min(limits, key=lambda x: x[1])

        if limit > test_worst:
            return (
                "High",
                f"{temp_label} limit({limit_label}) {round_value(limit)} > Test_Worst {round_value(test_worst)}"
            )

        if suggest_value > limit >= test_worst:
            return (
                "Medium",
                f"{temp_label} Suggest_Spec {round_value(suggest_value)} > limit({limit_label}) {round_value(limit)} >= Test_Worst {round_value(test_worst)}"
            )

        return (
            "Low",
            f"{temp_label} limit({limit_label}) {round_value(limit)} < Suggest_Spec {round_value(suggest_value)}"
        )

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

def get_use_edge_from_assign(parameter, spec_type, assign_df):
    """
    从 assign_standard.xlsx 获取 Use_Edge。
    如果没有填写，则返回 ALL。
    """
    assign_row = match_one_row(parameter, spec_type, assign_df)

    if assign_row is None:
        return "ALL"

    use_edge = normalize_text(assign_row.get("Use_Edge", "")).upper()

    if use_edge in ["FE", "RE", "IGNORE"]:
        return use_edge

    return "ALL"


def has_fe_re_target(parameter, target_df):
    """
    判断目标规格表中这个参数是否存在 FE / RE 两种边沿。
    如果有 FE/RE，默认优先取 FE。
    """
    matched = get_best_match_rows(parameter, target_df)

    if matched.empty or "Spec_Type" not in matched.columns:
        return False

    spec_values = matched["Spec_Type"].astype(str).str.upper().str.strip().tolist()

    return ("FE" in spec_values) or ("RE" in spec_values)


def get_preferred_edge(parameter, spec_type, assign_df, target_df):
    """
    决定当前参数取哪个边沿：
    1. assign_standard.xlsx 的 Use_Edge 优先
    2. 如果目标规格表有 FE / RE，则默认取 FE
    3. 其他参数取 ALL
    """
    use_edge = get_use_edge_from_assign(parameter, spec_type, assign_df)

    if use_edge in ["FE", "RE", "IGNORE"]:
        return use_edge

    if has_fe_re_target(parameter, target_df):
        return "FE"

    return "ALL"


def filter_raw_by_edge(param_df, preferred_edge):
    """
    根据 FE / RE 过滤 Raw 数据。
    """
    if preferred_edge == "ALL":
        return param_df

    if preferred_edge == "IGNORE":
        return param_df.iloc[0:0].copy()

    if param_df.empty:
        return param_df

    if "Edge" not in param_df.columns:
        return param_df

    df = param_df.copy()

    # 如果识别到了 FE/RE，只保留指定边沿
    edge_mask = df["Edge"].astype(str).str.upper().str.strip() == preferred_edge

    if edge_mask.any():
        return df[edge_mask].copy()

    # 如果完全没有识别到边沿，则不强行清空，避免误删
    return df

def select_target_row(parameter, spec_type, target_df):
    """
    目标规格行选择逻辑：
    1. 如果目标规格里有与当前 Spec_Type 完全一致的行，优先取该行。
    2. 如果目标规格存在 FE/RE 两种边沿，默认优先取 FE 行。
    3. 其他情况取第一条匹配行。
    """
    matched = get_best_match_rows(parameter, target_df)

    if matched.empty:
        return None

    spec_type_key = normalize_text(spec_type).lower()

    if spec_type_key and "Spec_Type" in matched.columns:
        same_type = matched[
            matched["Spec_Type"].astype(str).str.lower().str.strip() == spec_type_key
        ]

        if not same_type.empty:
            return same_type.iloc[0]

    if "Spec_Type" in matched.columns:
        fe_rows = matched[matched["Spec_Type"].astype(str).str.upper().str.strip() == "FE"]

        if not fe_rows.empty:
            return fe_rows.iloc[0]

    return matched.iloc[0]


def apply_target_info(row_dict, parameter, target_df):
    target_row = select_target_row(parameter, row_dict.get("Spec_Type", ""), target_df)

    target_parameter = ""
    target_spec_type = ""

    if target_row is None:
        row_dict["Target_Spec_Type"] = ""
        return row_dict, target_parameter

    target_parameter = target_row.get("Base_Parameter", "")
    target_spec_type = normalize_text(target_row.get("Spec_Type", ""))

    # 目标规格中的 Spec_Type 优先级最高。
    # 如果目标规格有定义 Spec_Type，则覆盖模板/赋值标准中的 Spec_Type。
    if target_spec_type:
        row_dict["Spec_Type"] = target_spec_type

    row_dict["Target_Spec_Type"] = target_spec_type

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
            "Sim_Spec_Type": "",
            "Sim_Typ_25C": np.nan,
            "Sim_Worst_90C": np.nan,
            "Sim_Worst_110C": np.nan,
            "Sim_Worst_130C": np.nan,
        }

    return {
        "Sim_Parameter": sim_row.get("Base_Parameter", ""),
        "Sim_Spec_Type": normalize_text(sim_row.get("Spec_Type", "")),
        "Sim_Typ_25C": to_number(sim_row.get("Sim_Typ_25C", np.nan)),
        "Sim_Worst_90C": to_number(sim_row.get("Sim_Worst_90C", np.nan)),
        "Sim_Worst_110C": to_number(sim_row.get("Sim_Worst_110C", np.nan)),
        "Sim_Worst_130C": to_number(sim_row.get("Sim_Worst_130C", np.nan)),
    }


# ============================================================
# 11. 生成 Summary
# ============================================================

def build_summary(template_df, raw_df, assign_df, target_df, sim_df):
    # 保持 Lot 原始参数顺序，不做 sorted 排序
    all_params_from_lots = list(dict.fromkeys(raw_df["Parameter"].dropna().tolist()))
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

        preferred_edge = get_preferred_edge(parameter, spec_type, assign_df, target_df)

        param_df = raw_df[raw_df["Parameter"] == parameter]
        param_df = filter_raw_by_edge(param_df, preferred_edge)

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

        sim_info = get_sim_info(parameter, spec_type, sim_df)

        sim_typ = sim_info["Sim_Typ_25C"]
        sim_90 = sim_info["Sim_Worst_90C"]
        sim_110 = sim_info["Sim_Worst_110C"]
        sim_130 = sim_info["Sim_Worst_130C"]
        # Simulated 三列优先使用上传的 sim_value.xlsx
        # 如果 sim_value.xlsx 没有值，再保留 template.xlsx 里的 Simulated 值
        simulated_90 = sim_90
        simulated_110 = sim_110
        simulated_130 = sim_130

        if pd.isna(simulated_90):
            simulated_90 = row_dict.get("Simulated_90C", np.nan)

        if pd.isna(simulated_110):
            simulated_110 = row_dict.get("Simulated_110C", np.nan)

        if pd.isna(simulated_130):
            simulated_130 = row_dict.get("Simulated_130C", np.nan)

        target_risk_90 = judge_target_risk_one(
            side,
            worst_90,
            suggest_90,
            row_dict["Target_90C"],
            simulated_90,
            "UpTo90C"
        )
        target_risk_110 = judge_target_risk_one(
            side,
            worst_110,
            suggest_110,
            row_dict["Target_110C"],
            simulated_110,
            "UpTo110C"
        )
        target_risk_130 = judge_target_risk_one(
            side,
            worst_130,
            suggest_130,
            row_dict["Target_130C"],
            simulated_130,
            "UpTo130C"
        )

        target_risk, target_risk_reason = combine_risk([
            target_risk_90,
            target_risk_110,
            target_risk_130
        ])

        typ_defined = has_typ_definition(parameter, assign_df, target_df, sim_df)
        suggest_typ_output = suggest_typ if typ_defined else np.nan

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
            "Simulated_90C": round_value(simulated_90),
            "Simulated_110C": round_value(simulated_110),
            "Simulated_130C": round_value(simulated_130),

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

            "Suggest_Typ_25C": format_suggest_value(suggest_typ_output, row_dict["Unit"]),
            "Suggest_Spec_UpTo90C": format_suggest_value(suggest_90, row_dict["Unit"]),
            "Suggest_Spec_UpTo110C": format_suggest_value(suggest_110, row_dict["Unit"]),
            "Suggest_Spec_UpTo130C": format_suggest_value(suggest_130, row_dict["Unit"]),

            # 分温度段 Target 风险；同时保留 Target_Risk 作为整行汇总风险。
            "Target_Risk_UpTo90C": target_risk_90[0],
            "Target_Risk_UpTo110C": target_risk_110[0],
            "Target_Risk_UpTo130C": target_risk_130[0],
            "Target_Risk": target_risk,
            "Target_Risk_Reason": target_risk_reason,

            "Sim_Parameter": sim_info["Sim_Parameter"],
            "Sim_Typ_25C": round_value(sim_typ),
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
            (summary_df["Simulated_90C"].astype(str).str.strip() == "") &
            (summary_df["Simulated_110C"].astype(str).str.strip() == "") &
            (summary_df["Simulated_130C"].astype(str).str.strip() == "")
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

        def _excel_col_name(col_idx):
            name = ""
            col_idx += 1
            while col_idx:
                col_idx, rem = divmod(col_idx - 1, 26)
                name = chr(65 + rem) + name
            return name

        def apply_text_risk_format(ws, col_idx, last_row):
            ws.conditional_format(1, col_idx, last_row, col_idx, {
                "type": "text",
                "criteria": "containing",
                "value": "High",
                "format": high_fmt,
            })
            ws.conditional_format(1, col_idx, last_row, col_idx, {
                "type": "text",
                "criteria": "containing",
                "value": "Medium",
                "format": medium_fmt,
            })
            ws.conditional_format(1, col_idx, last_row, col_idx, {
                "type": "text",
                "criteria": "containing",
                "value": "Low",
                "format": low_fmt,
            })
            ws.conditional_format(1, col_idx, last_row, col_idx, {
                "type": "text",
                "criteria": "containing",
                "value": "Review",
                "format": review_fmt,
            })

        last_row = len(final_df)

        # 对风险列本身着色。Target_Risk 是整行汇总，Target_Risk_UpTo* 是分温度段风险。
        risk_col_names = [
            "Target_Risk_UpTo90C",
            "Target_Risk_UpTo110C",
            "Target_Risk_UpTo130C",
            "Target_Risk",
            "Sim_Risk",
        ]

        for risk_col_name in risk_col_names:
            if risk_col_name in final_df.columns:
                apply_text_risk_format(ws, final_df.columns.get_loc(risk_col_name), last_row)

        # 对 Suggest_Spec 三列按对应温度段 Target_Risk 着色，便于直接看到每个建议规格的风险。
        suggest_risk_pairs = [
            ("Suggest_Spec_UpTo90C", "Target_Risk_UpTo90C"),
            ("Suggest_Spec_UpTo110C", "Target_Risk_UpTo110C"),
            ("Suggest_Spec_UpTo130C", "Target_Risk_UpTo130C"),
        ]

        for suggest_col_name, risk_col_name in suggest_risk_pairs:
            if suggest_col_name not in final_df.columns or risk_col_name not in final_df.columns:
                continue

            suggest_col = final_df.columns.get_loc(suggest_col_name)
            risk_col = final_df.columns.get_loc(risk_col_name)
            risk_col_letter = _excel_col_name(risk_col)

            ws.conditional_format(1, suggest_col, last_row, suggest_col, {
                "type": "formula",
                "criteria": f'=${risk_col_letter}2="High"',
                "format": high_fmt,
            })
            ws.conditional_format(1, suggest_col, last_row, suggest_col, {
                "type": "formula",
                "criteria": f'=${risk_col_letter}2="Medium"',
                "format": medium_fmt,
            })
            ws.conditional_format(1, suggest_col, last_row, suggest_col, {
                "type": "formula",
                "criteria": f'=${risk_col_letter}2="Low"',
                "format": low_fmt,
            })
            ws.conditional_format(1, suggest_col, last_row, suggest_col, {
                "type": "formula",
                "criteria": f'=${risk_col_letter}2="Review"',
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



# ============================================================
# 15. V8 Safe Overrides: voltage-group target/sim support
#     基于 V6 稳定版追加，避免启动阶段白屏。
# ============================================================

# 保存 V6 标准读取函数，宽表识别失败时回退使用。
_read_target_spec_standard_v6 = read_target_spec
_read_sim_value_standard_v6 = read_sim_value


def detect_voltage_range_from_text(x):
    """
    从文本中识别电压范围，例如：
    1.65V-2.3V、2.3V-3.6V、1.65V~2.3V。
    """
    s = normalize_text(x)

    if not s:
        return ""

    s = (
        s.replace("－", "-")
        .replace("–", "-")
        .replace("—", "-")
        .replace("~", "-")
        .replace("～", "-")
    )

    m = re.search(
        r"(\d+(?:\.\d+)?)\s*V\s*-\s*(\d+(?:\.\d+)?)\s*V",
        s,
        re.IGNORECASE,
    )

    if m:
        return f"{m.group(1)}V-{m.group(2)}V"

    return ""


def append_voltage_suffix(parameter, voltage_range):
    base = normalize_text(parameter)
    voltage = normalize_text(voltage_range)

    if not base or not voltage:
        return base

    if base.upper().endswith(("_" + voltage).upper()):
        return base

    return f"{base}_{voltage}"


def strip_voltage_suffix(parameter):
    """
    去掉 Parameter 末尾电压后缀，用于兜底匹配。
    例：ILI_1.65V-2.3V -> ILI
    """
    s = normalize_text(parameter)
    if not s:
        return ""

    return re.sub(
        r"_[0-9]+(?:\.[0-9]+)?V-[0-9]+(?:\.[0-9]+)?V$",
        "",
        s,
        flags=re.IGNORECASE,
    )


def add_match_keys(df):
    if df is None or df.empty:
        return df

    if "Base_Parameter" not in df.columns:
        return df

    df = df.copy()
    df["Base_Parameter"] = df["Base_Parameter"].astype(str).str.strip()
    df["Base_Key"] = df["Base_Parameter"].apply(normalize_param_for_match)
    return df


def get_candidate_match_keys(parameter):
    """
    V8 安全版匹配 Key：
    1. 完整参数，保留电压后缀
    2. 去掉电压后缀后的参数
    3. ICC3 频率规则
    4. 第一个下划线前基础参数
    """
    actual_key = normalize_param_for_match(parameter)
    stripped_parameter = strip_voltage_suffix(parameter)
    stripped_key = normalize_param_for_match(stripped_parameter)

    keys = [actual_key]

    if stripped_key and stripped_key != actual_key:
        keys.append(stripped_key)

    for p in [parameter, stripped_parameter]:
        icc3_key = get_icc3_freq_base_key(p)
        if icc3_key:
            keys.append(icc3_key)

    for p in [stripped_parameter, parameter]:
        first_base_key = get_first_underscore_base_key(p)
        if first_base_key:
            keys.append(first_base_key)

    result = []
    for k in keys:
        if k and k not in result:
            result.append(k)

    return result


def get_best_match_rows(actual_parameter, df):
    """
    V8 安全版匹配：
    - 先按带电压后缀的完整 Parameter 精确匹配
    - 再按去电压后缀/ICC3频率/基础参数匹配
    - 最后做最长前缀匹配
    """
    if df is None or df.empty:
        return pd.DataFrame()

    if "Base_Key" not in df.columns and "Base_Parameter" in df.columns:
        df = add_match_keys(df)

    if "Base_Key" not in df.columns:
        return pd.DataFrame()

    actual_key = normalize_param_for_match(actual_parameter)

    exact = df[df["Base_Key"] == actual_key]
    if not exact.empty:
        return exact.copy()

    for key in get_candidate_match_keys(actual_parameter):
        exact = df[df["Base_Key"] == key]
        if not exact.empty:
            return exact.copy()

    candidates = []

    for _, row in df.iterrows():
        standard_key = normalize_text(row.get("Base_Key", ""))

        if not standard_key:
            continue

        # 同时用完整 key 和去电压后的 key 做前缀匹配
        for actual in [actual_key, normalize_param_for_match(strip_voltage_suffix(actual_parameter))]:
            if actual and is_prefix_match(actual, standard_key):
                candidates.append((len(standard_key), row))
                break

    if not candidates:
        return pd.DataFrame()

    max_len = max(x[0] for x in candidates)
    best_rows = [x[1] for x in candidates if x[0] == max_len]

    return pd.DataFrame(best_rows)


def _find_header_row_for_voltage_file(raw_df):
    for r in range(min(20, len(raw_df))):
        row_keys = [clean_column_name(v) for v in raw_df.iloc[r].tolist()]
        if any(k in ["BASEPARAMETER", "PARAMETER", "SYMBOL", "基础参数", "标准参数", "参数"] for k in row_keys):
            return r
    return None


def _find_voltage_row(raw_df, header_row_idx, max_scan_rows=8):
    start = header_row_idx + 1
    end = min(len(raw_df), header_row_idx + 1 + max_scan_rows)

    for r in range(start, end):
        row_values = raw_df.iloc[r].tolist()
        row_text = " ".join([normalize_text(v) for v in row_values])
        voltage_count = sum(1 for v in row_values if detect_voltage_range_from_text(v))

        if ("VOLTAGE" in row_text.upper() or "电压" in row_text) and voltage_count >= 1:
            return r

    return None


def _get_basic_columns_from_header(header_values):
    result = {"base_col": None, "spec_col": None, "unit_col": None}

    for c, value in enumerate(header_values):
        key = clean_column_name(value)

        if key in ["BASEPARAMETER", "PARAMETER", "SYMBOL", "基础参数", "标准参数", "参数"]:
            result["base_col"] = c
        elif key in ["SPECTYPE", "TYPE", "规格类型", "参数类型"]:
            result["spec_col"] = c
        elif key in ["UNIT", "单位"]:
            result["unit_col"] = c

    return result


def _classify_temp_column_for_voltage_file(header_value, mode):
    key = clean_column_name(header_value)

    if mode == "target":
        if key in ["TARGET90C", "TARGET90", "SPEC90C", "SPEC90", "90C", "90"]:
            return "Target_90C"
        if key in ["TARGET110C", "TARGET110", "SPEC110C", "SPEC110", "110C", "110"]:
            return "Target_110C"
        if key in ["TARGET130C", "TARGET130", "SPEC130C", "SPEC130", "130C", "130"]:
            return "Target_130C"
        return ""

    if mode == "sim":
        if key in ["SIMTYP25C", "SIM25C", "25CSIM", "25CTYP", "SIMTYP", "TARGET25C", "25C", "25"]:
            return "Sim_Typ_25C"
        if key in ["SIMWORST90C", "SIM90C", "SIMUPTO90C", "SIMULATED90C", "SIMULATED90", "90CSIM", "TARGET90C", "SPEC90C", "90C", "90"]:
            return "Sim_Worst_90C"
        if key in ["SIMWORST110C", "SIM110C", "SIMUPTO110C", "SIMULATED110C", "SIMULATED110", "110CSIM", "TARGET110C", "SPEC110C", "110C", "110"]:
            return "Sim_Worst_110C"
        if key in ["SIMWORST130C", "SIM130C", "SIMUPTO130C", "SIMULATED130C", "SIMULATED130", "130CSIM", "TARGET130C", "SPEC130C", "130C", "130"]:
            return "Sim_Worst_130C"
        return ""

    return ""


def parse_voltage_wide_spec_file(file_path, mode):
    """
    读取电压分组宽表格式的目标规格/仿真值。

    支持示例：
    Base_Parameter | Spec_Type | Target_90C | Target_110C | Target_130C | Target_90C | Target_110C | Target_130C | Unit
    Voltage        |           | 1.65V-2.3V |              |              | 2.3V-3.6V  |              |              |
    """
    raw = pd.read_excel(file_path, sheet_name=0, header=None)
    header_row_idx = _find_header_row_for_voltage_file(raw)

    if header_row_idx is None:
        return pd.DataFrame()

    voltage_row_idx = _find_voltage_row(raw, header_row_idx)

    if voltage_row_idx is None:
        return pd.DataFrame()

    header_values = raw.iloc[header_row_idx].tolist()
    voltage_values = raw.iloc[voltage_row_idx].tolist()
    basic_cols = _get_basic_columns_from_header(header_values)

    base_col = basic_cols["base_col"]
    spec_col = basic_cols["spec_col"]
    unit_col = basic_cols["unit_col"]

    if base_col is None:
        return pd.DataFrame()

    voltage_by_col = []
    current_voltage = ""

    for value in voltage_values:
        voltage = detect_voltage_range_from_text(value)
        if voltage:
            current_voltage = voltage
        voltage_by_col.append(current_voltage)

    value_columns = []

    for c, header_value in enumerate(header_values):
        output_col = _classify_temp_column_for_voltage_file(header_value, mode)
        voltage = voltage_by_col[c] if c < len(voltage_by_col) else ""

        if output_col and voltage:
            value_columns.append({
                "col": c,
                "voltage": voltage,
                "output_col": output_col,
            })

    if not value_columns:
        return pd.DataFrame()

    voltage_order = []
    for info in value_columns:
        if info["voltage"] not in voltage_order:
            voltage_order.append(info["voltage"])

    rows = []

    for r in range(voltage_row_idx + 1, len(raw)):
        base_parameter = normalize_text(raw.iloc[r, base_col])

        if not base_parameter:
            continue

        if clean_column_name(base_parameter) in ["VOLTAGE", "BASEPARAMETER", "PARAMETER", "SYMBOL"]:
            continue

        spec_type = normalize_text(raw.iloc[r, spec_col]) if spec_col is not None else ""
        unit = normalize_text(raw.iloc[r, unit_col]) if unit_col is not None else ""

        for voltage in voltage_order:
            display_parameter = append_voltage_suffix(base_parameter, voltage)

            if mode == "target":
                row_dict = {
                    "Base_Parameter": display_parameter,
                    "Spec_Type": spec_type,
                    "Target_90C": np.nan,
                    "Target_110C": np.nan,
                    "Target_130C": np.nan,
                    "Unit": unit,
                    "Remark": "",
                }
            else:
                row_dict = {
                    "Base_Parameter": display_parameter,
                    "Spec_Type": spec_type,
                    "Sim_Typ_25C": np.nan,
                    "Sim_Worst_90C": np.nan,
                    "Sim_Worst_110C": np.nan,
                    "Sim_Worst_130C": np.nan,
                    "Unit": unit,
                    "Remark": "",
                }

            for info in value_columns:
                if info["voltage"] != voltage:
                    continue

                value = to_number(raw.iloc[r, info["col"]])
                row_dict[info["output_col"]] = value

            if mode == "sim" and pd.isna(row_dict.get("Sim_Typ_25C", np.nan)) and is_typ_spec_value(spec_type):
                for col_name in ["Sim_Worst_90C", "Sim_Worst_110C", "Sim_Worst_130C"]:
                    value = to_number(row_dict.get(col_name, np.nan))
                    if not pd.isna(value):
                        row_dict["Sim_Typ_25C"] = value
                        break

            numeric_cols = [c for c in row_dict.keys() if c.startswith("Target_") or c.startswith("Sim_")]
            if any(not pd.isna(to_number(row_dict.get(c, np.nan))) for c in numeric_cols):
                rows.append(row_dict)

    result = pd.DataFrame(rows)

    if result.empty:
        return result

    result = add_match_keys(result)
    return result


def read_target_spec(target_file):
    wide_df = parse_voltage_wide_spec_file(target_file, mode="target")

    if not wide_df.empty:
        print("目标规格文件读取完成，识别到电压分组宽表格式，参数数量：", len(wide_df))
        return wide_df

    df = _read_target_spec_standard_v6(target_file)
    df = add_match_keys(df)
    return df


def read_sim_value(sim_file):
    wide_df = parse_voltage_wide_spec_file(sim_file, mode="sim")

    if not wide_df.empty:
        print("仿真值文件读取完成，识别到电压分组宽表格式，参数数量：", len(wide_df))
        return wide_df

    df = _read_sim_value_standard_v6(sim_file)

    # 标准格式兜底：如果 typ 行没有 Sim_Typ_25C，但 90/110/130 有值，则取第一个非空值作为 Sim_Typ_25C。
    if "Spec_Type" in df.columns:
        for idx, row in df.iterrows():
            if pd.isna(to_number(row.get("Sim_Typ_25C", np.nan))) and is_typ_spec_value(row.get("Spec_Type", "")):
                for col_name in ["Sim_Worst_90C", "Sim_Worst_110C", "Sim_Worst_130C"]:
                    value = to_number(row.get(col_name, np.nan))
                    if not pd.isna(value):
                        df.at[idx, "Sim_Typ_25C"] = value
                        break

    df = add_match_keys(df)
    return df


def _build_voltage_by_col(df, stat_row_idx):
    """读取 Lot Index 页表头里的电压分组，并向右填充。"""
    voltage_by_col = []
    current_voltage = ""

    for c in range(df.shape[1]):
        header_text = " ".join([
            normalize_text(df.iloc[rr, c])
            for rr in range(0, stat_row_idx + 1)
        ])
        voltage = detect_voltage_range_from_text(header_text)
        if voltage:
            current_voltage = voltage
        voltage_by_col.append(current_voltage)

    return voltage_by_col


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
    voltage_by_col = _build_voltage_by_col(df, stat_row_idx)

    value_infos = []

    for c in range(df.shape[1]):
        stat = clean_stat(stat_row.iloc[c])
        temp = parse_temp(temp_row.iloc[c])

        if stat in ["Min", "Typ", "Max"] and temp is not None:
            header_text = " ".join([
                normalize_text(df.iloc[rr, c])
                for rr in range(0, stat_row_idx + 1)
            ])

            edge = detect_edge_from_text(header_text)
            voltage_range = voltage_by_col[c] if c < len(voltage_by_col) else ""

            value_infos.append({
                "col": c,
                "temp": temp,
                "stat": stat,
                "edge": edge,
                "voltage_range": voltage_range,
            })

    if not value_infos:
        print(f"[跳过] {lot_name}：没有识别到有效温度数据列。")
        return pd.DataFrame()

    first_value_col = min([x["col"] for x in value_infos])

    edge_col = None

    for c in range(0, first_value_col):
        header_text = " ".join([
            normalize_text(df.iloc[rr, c])
            for rr in range(0, stat_row_idx + 1)
        ])
        header_key = clean_column_name(header_text)

        if "EDGE" in header_key or "边沿" in header_text:
            edge_col = c
            break

    if edge_col is not None:
        print(f"{lot_name} 识别到 Edge 列：第 {edge_col + 1} 列")
    else:
        print(f"{lot_name} 未识别到 Edge 列，将尝试从行描述中识别 FE/RE")

    records = []
    last_parameter = ""

    for r in range(stat_row_idx + 1, len(df)):
        raw_parameter = normalize_text(df.iloc[r, 0])

        edge_cell = ""
        if edge_col is not None:
            edge_cell = normalize_text(df.iloc[r, edge_col])

        edge_from_edge_col = detect_edge_from_text(edge_cell)

        if raw_parameter and raw_parameter.lower() not in ["symbol", "parameter", "参数", "test item", "item"]:
            last_parameter = raw_parameter

        base_parameter = raw_parameter if raw_parameter else last_parameter

        if not base_parameter:
            continue

        if base_parameter.lower() in ["symbol", "parameter", "参数", "test item", "item"]:
            continue

        row_meta_cells = []

        for cc in range(0, first_value_col):
            row_meta_cells.append(normalize_text(df.iloc[r, cc]))

        row_meta_text = " ".join(row_meta_cells)
        test_condition = normalize_text(df.iloc[r, 2]) if df.shape[1] > 2 else row_meta_text
        row_edge = edge_from_edge_col if edge_from_edge_col else detect_edge_from_text(row_meta_text)

        unit = ""

        for c in range(0, first_value_col):
            cell = normalize_text(df.iloc[r, c])
            if cell in ["uA", "μA", "mA", "A", "V", "mV", "ns", "us", "μs", "ms", "s", "MHz"]:
                unit = cell
                break

        for info in value_infos:
            c = info["col"]
            temp = info["temp"]
            stat = info["stat"]
            edge = info["edge"]
            voltage_range = info.get("voltage_range", "")

            value = to_number(df.iloc[r, c])

            if temp is None or not stat or pd.isna(value):
                continue

            condition_edge = detect_edge_from_text(test_condition)
            parameter_edge = detect_edge_from_text(base_parameter)

            if edge_from_edge_col:
                final_edge = edge_from_edge_col
            elif parameter_edge:
                final_edge = parameter_edge
            elif row_edge:
                final_edge = row_edge
            elif condition_edge:
                final_edge = condition_edge
            else:
                final_edge = edge

            display_parameter = append_voltage_suffix(base_parameter, voltage_range)

            records.append({
                "Lot": lot_name,
                "Parameter": display_parameter,
                "Base_Parameter": base_parameter,
                "Voltage_Range": voltage_range,
                "Test_Condition": test_condition,
                "Edge": final_edge,
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

    # 调整 Raw 顺序，让 Summary 参数按电压分组输出：
    # ILI_1.65V-2.3V ... tVSL_1.65V-2.3V，然后 ILI_2.3V-3.6V ...
    if "Voltage_Range" in raw_df.columns and "Base_Parameter" in raw_df.columns:
        raw_df = raw_df.copy()
        raw_df["__orig_order"] = range(len(raw_df))

        voltage_order = [
            v for v in dict.fromkeys(raw_df["Voltage_Range"].fillna("").astype(str).tolist())
            if normalize_text(v)
        ]
        base_order = list(dict.fromkeys(raw_df["Base_Parameter"].fillna("").astype(str).tolist()))

        voltage_rank = {v: i for i, v in enumerate(voltage_order)}
        base_rank = {v: i for i, v in enumerate(base_order)}

        def _sort_key(row):
            voltage = normalize_text(row.get("Voltage_Range", ""))
            base = normalize_text(row.get("Base_Parameter", ""))
            if voltage:
                return (
                    voltage_rank.get(voltage, 10_000),
                    base_rank.get(base, 10_000),
                    row["__orig_order"],
                )
            return (20_000, base_rank.get(base, 10_000), row["__orig_order"])

        sort_keys = raw_df.apply(_sort_key, axis=1)
        raw_df["__v_rank"] = [x[0] for x in sort_keys]
        raw_df["__b_rank"] = [x[1] for x in sort_keys]
        raw_df["__o_rank"] = [x[2] for x in sort_keys]
        raw_df = raw_df.sort_values(["__v_rank", "__b_rank", "__o_rank"]).drop(
            columns=["__orig_order", "__v_rank", "__b_rank", "__o_rank"]
        )
        raw_df = raw_df.reset_index(drop=True)

    print("所有 Lot 读取完成，总数据点数量：", len(raw_df))
    return raw_df
