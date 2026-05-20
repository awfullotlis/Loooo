"""data_loader — 数据源加载抽象层 + Excel 报表读取 + 列名映射 + 多报表合并。

设计：
- DataSourceLoader（抽象）：定义加载报表数据的接口
- LocalDataSourceLoader（当前实现）：从 数据源/ 本地文件夹读取
- DownloadDataSourceLoader（未来占位）：将调用 skill_download

- ExcelReportReader：读 Excel → 原始列名行 dict，支持多行标题配置
- reconcile_columns：将 rule 中的概念列名对齐到实际 Excel 列名
- map_columns_to_fields：原始列名 → rule 字段名
- merge_multi_report_data：按 employee_id 合并多报表
- normalize_value：百分比字符串/数值类型归一化
"""

import json
import logging
from abc import ABC, abstractmethod
from pathlib import Path

import openpyxl

from report_registry import ReportRegistry

logger = logging.getLogger("incentive_bot")

SKILLS_DIR = Path(__file__).parent


# ---------------------------------------------------------------------------
# Value Normalization
# ---------------------------------------------------------------------------

EXCEL_ERROR_TOKENS = frozenset({
    "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!", "#GETTING_DATA",
})


def normalize_value(value) -> float | int | str | None:
    """类型归一化，确保 skill_calculate 能正确消费。

    - "92%" → 0.92（百分比字符串转小数）
    - openpyxl 数字类型 → float / int
    - None / 空单元格 / Excel 错误码 (#DIV/0! 等) → None
    - 纯数字字符串 "3.14" → 3.14
    """
    if value is None:
        return None

    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return None
        if stripped.upper() in EXCEL_ERROR_TOKENS:
            return None
        if stripped.endswith("%"):
            try:
                return float(stripped.rstrip("%")) / 100
            except ValueError:
                pass
        try:
            if "." in stripped:
                return float(stripped)
            return int(stripped)
        except ValueError:
            return stripped

    if isinstance(value, float):
        return value
    if isinstance(value, int):
        return value

    return str(value)


# ---------------------------------------------------------------------------
# Column Reconciliation Config
# ---------------------------------------------------------------------------

def load_column_mapping(filepath: Path | None = None) -> dict:
    """加载列名映射配置。

    配置格式见 column_mapping.json：
    每个报表有 read_config（读取参数）和 column_reconcile（列名对齐）。
    """
    if filepath is None:
        filepath = SKILLS_DIR / "column_mapping.json"
    if not filepath.exists():
        logger.info("未找到 column_mapping.json，使用默认配置（单行 header）")
        return {}
    data = json.loads(filepath.read_text(encoding="utf-8"))
    logger.info(f"加载列名映射配置: {len(data)} 个报表")
    return data


# ---------------------------------------------------------------------------
# Abstract Interface
# ---------------------------------------------------------------------------

class DataSourceLoader(ABC):
    """报表数据加载抽象接口。

    Contract: 给定一组报表名，返回 {报表名: [行dict]}，
    行 dict 使用原始 Excel 列名作为 key。

    未来 skill_download 实现时，只需新增 DownloadDataSourceLoader 子类，
    编排代码零改动。
    """

    @abstractmethod
    def load_reports(
        self,
        report_names: set[str],
        month: str,
    ) -> dict[str, list[dict]]:
        """加载指定报表数据。

        Args:
            report_names: data_source_mapping 中的报表名集合
            month: YYYY-MM，预留月份特定数据场景

        Returns:
            {报表名: [行dict]}，行 dict key 为原始 Excel 列名
        """
        ...


# ---------------------------------------------------------------------------
# Excel Report Reader
# ---------------------------------------------------------------------------

class ExcelReportReader:
    """读 Excel (.xlsx) 文件，每行转为 dict {列名: 值}。

    支持配置化多行标题：
    - header_rows: 标题行数（默认 1）
    - data_start_row: 数据起始行（1-based，默认 header_rows+1）
    - id_column_override: 覆盖标识列名（如 "LP姓名" 替代 "员工编号"）

    多行标题策略：取最后一个包含具体列名的行作为列名 header，
    前面的标题行作为分组标题（用于消除重复列名歧义）。
    """

    def read_file(
        self,
        filepath: Path,
        sheet_name: str | None = None,
        header_rows: int = 1,
        data_start_row: int | None = None,
    ) -> list[dict]:
        """读 Excel 文件为 list[dict]。

        Args:
            filepath: .xlsx 文件路径
            sheet_name: 指定 sheet 名（None = 第一个 sheet）
            header_rows: 标题行数（多行标题时指定）
            data_start_row: 数据起始行（1-based）。None = header_rows+1

        Returns:
            行 dict 列表，key 为原始列名，value 经 normalize_value 处理
        """
        wb = openpyxl.load_workbook(filepath, data_only=True)

        if sheet_name:
            ws = wb[sheet_name]
        else:
            ws = wb.worksheets[0]

        rows = list(ws.iter_rows(values_only=True))

        if not rows:
            logger.warning(f"Excel 文件 {filepath.name} 为空")
            return []

        # 构建列名：从多行标题中提取最具体的列名
        # 策略：扫描 header_rows 行，取每个位置上最具体的名称
        # "最具体" = 非分组标题（不含纯数字、不含分组标记）
        max_col = max(len(r) for r in rows) if rows else 0
        headers = self._build_headers_from_multi_row(rows, header_rows, max_col)

        # 数据起始行
        if data_start_row is not None:
            data_start_idx = data_start_row - 1  # 转为 0-based
        else:
            data_start_idx = header_rows  # header 后紧接着

        data_rows = rows[data_start_idx:]

        result = []
        for row_vals in data_rows:
            row_dict = {}
            for i, val in enumerate(row_vals):
                if i < len(headers) and headers[i]:
                    row_dict[headers[i]] = normalize_value(val)
            # 只跳过全空行，保留所有数据行（包括总计行）
            # 总计/合计行在 map_columns_to_fields 中按 employee_id 过滤
            if any(v is not None and v != "" for v in row_dict.values()):
                result.append(row_dict)

        wb.close()
        logger.info(f"读取 {filepath.name} [{ws.title}]: {len(result)} 行数据, {len(headers)} 列")
        return result

    def _build_headers_from_multi_row(
        self, rows: list, header_rows: int, max_col: int,
    ) -> list[str]:
        """从多行标题中构建唯一列名列表。

        策略：
        1. 对每个列位置，从最后一行往前找第一个"具体列名"
        2. 如果有重复列名，加父级分组名前缀消除歧义
        3. 空位置保持空字符串
        """
        # 提取标题区域的所有行
        header_data = rows[:header_rows]

        # 对每个列位置，收集所有标题行的值
        col_names_by_pos = []
        for col_idx in range(max_col):
            names = []
            for row_idx in range(header_rows):
                if col_idx < len(header_data[row_idx]):
                    val = header_data[row_idx][col_idx]
                    if val is not None and str(val).strip():
                        names.append(str(val).strip())
                else:
                    names.append("")
            col_names_by_pos.append(names)

        # 构建最终列名：取最具体的一个（通常是最后一个非空值）
        headers = []
        for names in col_names_by_pos:
            # 从后往前找最具体的名称
            specific_name = ""
            for name in reversed(names):
                if name and not self._is_group_header(name):
                    specific_name = name
                    break

            if specific_name:
                headers.append(specific_name)
            else:
                headers.append("")

        # 检查重复列名并加前缀
        name_count = {}
        for h in headers:
            if h:
                name_count[h] = name_count.get(h, 0) + 1

        # 对重复列名，添加分组前缀
        if any(c > 1 for c in name_count.values()):
            headers = self._disambiguate_duplicate_headers(
                headers, col_names_by_pos, name_count
            )

        return headers

    def _is_group_header(self, name: str) -> bool:
        """判断是否是分组标题（而非具体列名）。

        分组标题特征：
        - 纯数字（如 "111.0" 是系统导出标记）
        - 常见分组词（月度、今日、首通、首课等 —— 但这些也可能是具体列名的一部分）
        """
        # 纯数字标记（Excel导出的行号标记）
        try:
            float(name)
            return True
        except ValueError:
            pass
        return False

    def _disambiguate_duplicate_headers(
        self, headers: list[str], col_names_by_pos: list[list[str]], name_count: dict,
    ) -> list[str]:
        """对重复列名添加分组前缀消除歧义。

        例如：两个"及时跟进率" → "月度_及时跟进率" 和 "今日_及时跟进率"
        """
        # 找到所有重复名
        duplicates = {h for h, c in name_count.items() if c > 1}

        # 为每个重复名，找其分组标题（前一行中最近的非空分组名）
        result = list(headers)
        seen = {}  # {原始名: 出现次数}

        for i, h in enumerate(headers):
            if h in duplicates:
                count = seen.get(h, 0)
                # 查找该列位置的分组标题（从前面的标题行找）
                parent = ""
                names_at_pos = col_names_by_pos[i]
                for name in names_at_pos:
                    if name and name != h and not self._is_group_header(name):
                        parent = name
                        break

                if parent:
                    result[i] = f"{parent}_{h}"
                else:
                    # 无分组标题，用序号区分
                    result[i] = f"{h}_{count + 1}"

                seen[h] = count + 1

        return result


# ---------------------------------------------------------------------------
# Concrete: Local File-based Loader
# ---------------------------------------------------------------------------

class LocalDataSourceLoader(DataSourceLoader):
    """从本地 数据源/ 目录读取报表 Excel 文件。

    通过 ReportRegistry 映射抽象报表名 → 实际文件名，
    通过 ExcelReportReader 解析 Excel 内容，
    通过 column_mapping 配置处理多行标题。
    """

    def __init__(
        self,
        data_dir: Path,
        registry: ReportRegistry,
        column_mapping: dict | None = None,
    ):
        self._data_dir = data_dir
        self._registry = registry
        self._reader = ExcelReportReader()
        self._column_mapping = column_mapping or {}

    def load_reports(
        self,
        report_names: set[str],
        month: str,
    ) -> dict[str, list[dict]]:
        """从本地文件加载报表数据。

        使用 column_mapping 中的 read_config 配置读取参数。
        """
        results = {}
        for report_name in report_names:
            filepath = self._registry.resolve_path(report_name, month)
            logger.info(f"加载报表 '{report_name}': {filepath.name}")

            # 从 column_mapping 获取读取配置
            read_config = self._column_mapping.get(report_name, {}).get("read_config", {})
            header_rows = read_config.get("header_rows", 1)
            data_start_row = read_config.get("data_start_row")

            results[report_name] = self._reader.read_file(
                filepath,
                header_rows=header_rows,
                data_start_row=data_start_row,
            )
        return results


# ---------------------------------------------------------------------------
# Future: Download-based Loader (stub)
# ---------------------------------------------------------------------------

class DownloadDataSourceLoader(DataSourceLoader):
    """未来实现：调用 skill_download 获取报表数据。

    由其他同事实现。编排代码只调用 loader.load_reports()，
    替换 Loader 子类即可。
    """

    def load_reports(
        self,
        report_names: set[str],
        month: str,
    ) -> dict[str, list[dict]]:
        raise NotImplementedError(
            "DownloadDataSourceLoader 尚未实现，等待 skill_download 开发完成"
        )


# ---------------------------------------------------------------------------
# Column Mapping + Multi-Report Merge
# ---------------------------------------------------------------------------

def collect_field_mapping_by_report(
    rule: dict | list,
) -> dict[str, dict[str, str]]:
    """从 rule 的 data_source_mapping 中收集按报表分组的 字段→列名 映射。"""
    plans = rule if isinstance(rule, list) else [rule]
    report_mapping = {}

    for plan in plans:
        mapping = plan.get("data_source_mapping", {})
        for field_name, source_info in mapping.items():
            report_name = source_info.get("report", "")
            column_name = source_info.get("column", "")

            if not report_name or not column_name:
                continue

            if report_name not in report_mapping:
                report_mapping[report_name] = {}
            report_mapping[report_name][field_name] = column_name

    return report_mapping


def collect_report_names(rule: dict | list) -> set[str]:
    """提取 rule 中所有引用的报表名。"""
    mapping = collect_field_mapping_by_report(rule)
    return set(mapping.keys())


def reconcile_field_mapping(
    field_mapping: dict[str, str],
    column_reconcile: dict[str, str],
) -> dict[str, str]:
    """将 rule 中的概念列名对齐到实际 Excel 列名。

    Args:
        field_mapping: {字段名: 概念列名} —— 来自 data_source_mapping
        column_reconcile: {概念列名: 实际列名} —— 来自 column_mapping.json

    Returns:
        {字段名: 实际列名} —— 可用于 map_columns_to_fields

    对齐策略：
    1. 对每个字段的概念列名，检查 reconcile 中是否有映射
    2. 有映射 → 替换为实际列名
    3. 无映射 → 保留原概念列名（可能是精确匹配的）
    4. 同时处理 id_column_override：如果 reconcile 中有 "员工编号" → "LP姓名" 的映射
       则 employee_id 字段也使用实际列名
    """
    reconciled = {}
    for field_name, conceptual_col in field_mapping.items():
        actual_col = column_reconcile.get(conceptual_col, conceptual_col)
        reconciled[field_name] = actual_col

    # 日志记录对齐情况
    for field_name, conceptual_col in field_mapping.items():
        actual_col = reconciled[field_name]
        if actual_col != conceptual_col:
            logger.debug(f"列名对齐: {field_name}: '{conceptual_col}' → '{actual_col}'")

    return reconciled


def map_columns_to_fields(
    raw_rows: list[dict],
    field_mapping: dict[str, str],
) -> list[dict]:
    """将原始 Excel 列名映射为 rule 定义的字段名。

    Args:
        raw_rows: [{原始列名: 值}]
        field_mapping: {字段名: 实际列名} —— 已经过 reconcile 对齐

    Returns:
        [{字段名: 值}] —— 可直接供 skill_calculate 使用

    只保留 field_mapping 引用的列，其余列丢弃。
    映射中的列名在 Excel 中找不到 → WARNING + 字段设 None。
    """
    reverse_map = {v: k for k, v in field_mapping.items()}

    # 检查映射的列名是否存在于 Excel header 中
    if raw_rows:
        available_cols = set(raw_rows[0].keys())
        for col_name in reverse_map:
            if col_name not in available_cols:
                field_name = reverse_map[col_name]
                logger.warning(
                    f"字段 '{field_name}' 映射到列 '{col_name}'，"
                    f"但报表中未找到该列，该字段将设为 None"
                )

    mapped = []
    summary_keywords = {"总计", "合计", "平均值", "平均", "汇总", "小计"}
    for row in raw_rows:
        new_row = {}
        for col_name, value in row.items():
            if col_name in reverse_map:
                new_row[reverse_map[col_name]] = value
        # 补充缺失字段为 None
        for field_name, col_name in field_mapping.items():
            if field_name not in new_row:
                new_row[field_name] = None
        # 只保留有有效 employee_id 的行（跳过总计/合计等汇总行）
        eid = new_row.get("employee_id")
        if eid is not None and str(eid).strip() and str(eid).strip() not in summary_keywords:
            mapped.append(new_row)

    return mapped


def merge_multi_report_data(
    mapped_reports: dict[str, list[dict]],
) -> list[dict]:
    """合并多个报表的数据，按 employee_id 做 union。

    员工只出现在部分报表 → 缺失字段为 None，
    skill_calculate 用 data.get(field, 0) 处理。
    """
    master: dict[str, dict] = {}

    for report_name, employees in mapped_reports.items():
        for emp in employees:
            eid = emp.get("employee_id")
            if eid is None:
                continue

            if eid not in master:
                master[eid] = {}
            master[eid].update(emp)

    # 检查员工覆盖率
    for eid, fields in master.items():
        none_fields = [k for k, v in fields.items() if v is None and k != "employee_id"]
        if none_fields:
            logger.warning(
                f"员工 '{eid}' 部分字段为 None: {none_fields}，"
                f"可能在该报表中不存在"
            )

    logger.info(f"合并 {len(mapped_reports)} 个报表，共 {len(master)} 名员工")
    return list(master.values())


# ---------------------------------------------------------------------------
# Aggregate Row Extraction
# ---------------------------------------------------------------------------

def extract_aggregate_row(
    raw_rows: list[dict],
    match_column: str,
    match_value: str,
    match_type: str = "contains",
    id_column: str | None = None,
    id_value: str | None = None,
) -> dict | None:
    """在原始 Excel 行中按 match_column 的值匹配汇总行。

    Args:
        raw_rows: [{原始列名: 值}]
        match_column: 分组标识列名（如 LP小组/团队/小组/lp组别）
        match_value: 匹配目标值（如"海外团队"/"海外教学服务部"）
        match_type: "contains"（默认）或 "exact"
        id_column: 员工标识列名（可选，用于辅助验证是汇总行而非个人行）
        id_value: 员工标识列期望值（如"总计"）

    Returns:
        匹配到的行 dict 或 None
    """
    for row in raw_rows:
        # 先验证 id_column（如 id_column="LP个人" 且 id_value="总计"，确保是汇总行）
        if id_column and id_value:
            cell_id = row.get(id_column)
            if cell_id is None or str(cell_id).strip() != id_value:
                continue

        cell_value = row.get(match_column)
        if cell_value is None:
            continue
        cell_str = str(cell_value).strip()
        if match_type == "contains":
            if match_value in cell_str:
                logger.info(f"匹配汇总行: '{match_column}' contains '{match_value}' → '{cell_str}'")
                return row
        elif match_type == "exact":
            if cell_str == match_value:
                logger.info(f"匹配汇总行: '{match_column}' exact '{match_value}'")
                return row

    logger.warning(f"未找到汇总行: '{match_column}' {match_type} '{match_value}'")
    return None


def _resolve_aggregate_config(
    report_name: str,
    aggregate_key: str,
    column_mapping: dict,
) -> dict | None:
    """从 column_mapping.aggregate_row_mapping 解析汇总行定位配置。

    Returns:
        {"match_column": str, "match_value": str, "match_type": str} 或 None
    """
    agg_mapping = column_mapping.get("aggregate_row_mapping", {})
    report_config = agg_mapping.get(report_name)
    if not report_config:
        logger.warning(f"报表 '{report_name}' 未配置 aggregate_row_mapping")
        return None

    keys = report_config.get("keys", {})
    key_config = keys.get(aggregate_key)
    if not key_config:
        logger.warning(
            f"报表 '{report_name}' aggregate_row_mapping 中无 key '{aggregate_key}'，"
            f"可用 keys: {list(keys.keys())}"
        )
        return None

    return {
        "match_column": report_config["match_column"],
        "match_value": key_config["match_value"],
        "match_type": key_config.get("match_type", "contains"),
        "id_column": report_config.get("id_column"),
        "id_value": key_config.get("id_value", report_config.get("id_value")),
    }


def _resolve_composite_aggregate(
    raw_rows: list[dict],
    aggregate_key: str,
    report_name: str,
    column_mapping: dict,
) -> dict | None:
    """合成聚合：通过 parent − exclude 计算派生聚合行。

    用于"广州 = 港澳(地区行) − 港澳组(小组行)"这类场景。
    sum_fields 直接相减，derived_fields 按公式重新计算。
    """
    composite_keys = column_mapping.get("composite_keys", {})
    config = composite_keys.get(aggregate_key)
    if not config:
        return None

    parent_key = config["parent"]
    exclude_keys = config.get("exclude", [])
    sum_fields = config.get("sum_fields", [])
    derived_fields = config.get("derived_fields", {})

    # 定位 parent 行
    parent_config = _resolve_aggregate_config(report_name, parent_key, column_mapping)
    if not parent_config:
        return None
    parent_row = extract_aggregate_row(
        raw_rows,
        parent_config["match_column"],
        parent_config["match_value"],
        parent_config["match_type"],
        id_column=parent_config.get("id_column"),
        id_value=parent_config.get("id_value"),
    )
    if not parent_row:
        return None

    # 定位 exclude 行并累加
    exclude_sums: dict[str, float] = {f: 0.0 for f in sum_fields}
    for exc_key in exclude_keys:
        exc_config = _resolve_aggregate_config(report_name, exc_key, column_mapping)
        if not exc_config:
            continue
        exc_row = extract_aggregate_row(
            raw_rows,
            exc_config["match_column"],
            exc_config["match_value"],
            exc_config["match_type"],
            id_column=exc_config.get("id_column"),
            id_value=exc_config.get("id_value"),
        )
        if not exc_row:
            continue
        for f in sum_fields:
            col = _fuzzy_col(exc_row, f)
            if col:
                val = normalize_value(exc_row[col])
                if isinstance(val, (int, float)):
                    exclude_sums[f] += val

    # 构建合成行：复制 parent，对 sum_fields 做减法
    result = dict(parent_row)
    for f in sum_fields:
        col = _fuzzy_col(parent_row, f)
        if col:
            parent_val = normalize_value(parent_row[col])
            if isinstance(parent_val, (int, float)):
                result[col] = parent_val - exclude_sums[f]
            else:
                result[col] = None

    # 计算 derived_fields
    for field_name, formula_cfg in derived_fields.items():
        formula = formula_cfg.get("formula")
        if formula == "divide":
            num_col = _fuzzy_col(result, formula_cfg["numerator"])
            den_col = _fuzzy_col(result, formula_cfg["denominator"])
            nv_num = normalize_value(result.get(num_col)) if num_col else None
            nv_den = normalize_value(result.get(den_col)) if den_col else None
            if isinstance(nv_num, (int, float)) and isinstance(nv_den, (int, float)) and nv_den != 0:
                derived_val = nv_num / nv_den
            else:
                derived_val = None
            # 写入所有匹配的列名
            target_col = _fuzzy_col(result, field_name)
            if target_col:
                result[target_col] = derived_val
            else:
                result[field_name] = derived_val
        elif formula == "subtract":
            a_col = _fuzzy_col(result, formula_cfg["a"])
            b_col = _fuzzy_col(result, formula_cfg["b"])
            nv_a = normalize_value(result.get(a_col)) if a_col else None
            nv_b = normalize_value(result.get(b_col)) if b_col else None
            if isinstance(nv_a, (int, float)) and isinstance(nv_b, (int, float)):
                derived_val = nv_a - nv_b
            else:
                derived_val = None
            target_col = _fuzzy_col(result, field_name)
            if target_col:
                result[target_col] = derived_val
            else:
                result[field_name] = derived_val

    logger.info(f"composite aggregate '{aggregate_key}' = '{parent_key}' - {exclude_keys}")
    return result


def _fuzzy_col(row: dict, field_name: str) -> str | None:
    """模糊列名匹配：处理多行标题消歧后的列名。

    匹配优先级：
    1. 精确匹配
    2. 以 field_name 结尾（如 "汇总_池子数" 匹配 "池子数"）
    3. 以 field_name 开头 + "_" + 数字（如 "目前续费率_1" 匹配 "目前续费率"）
    """
    if field_name in row:
        return field_name
    suffix = "_" + field_name
    for k in row:
        if k.endswith(suffix):
            return k
    prefix = field_name + "_"
    for k in row:
        if k.startswith(prefix):
            return k
    return None


def build_aggregate_employee_data(
    raw_report_data: dict[str, list[dict]],
    rule: dict,
    column_mapping: dict,
) -> list[dict]:
    """aggregate 模式数据构建（支持 item 级 override）。

    流程：
    1. 读取 rule.data_scope.aggregate_key
    2. 对 rule 引用的每张报表：用 aggregate_row_mapping 定位汇总行
    3. reconcile 字段名 → 抽取字段值 → 合并为单条 agg_record
    4. 检查 items 有无 data_scope_override=individual 的字段
       - 若有：为每个 specific_employee 拉个人行，抽取 override 字段
       - 合并策略：override 字段取个人值，其余取整体值
    5. 按 specific_employees 展开 N 条（employee_id 改写为人名）
    """
    data_scope = rule.get("data_scope", {})
    aggregate_key = data_scope.get("aggregate_key", "")
    specific_employees = rule.get("employee_scope", {}).get("specific_employees", [])

    # 收集按报表分组的字段映射
    field_mapping_by_report = collect_field_mapping_by_report(rule)

    # 应用字段重定向
    field_redirect = column_mapping.get("field_redirect", {})
    if field_redirect:
        for field_name, redirect in field_redirect.items():
            from_report = redirect.get("from_report", "")
            to_report = redirect.get("to_report", "")
            actual_column = redirect.get("actual_column", "")
            # 仅当 rule 实际引用了 from_report 中的该字段时才转移，避免引入幽灵报表
            if from_report not in field_mapping_by_report:
                continue
            if field_name not in field_mapping_by_report[from_report]:
                continue
            del field_mapping_by_report[from_report][field_name]
            if to_report not in field_mapping_by_report:
                field_mapping_by_report[to_report] = {}
            field_mapping_by_report[to_report][field_name] = actual_column

    # 分离 aggregate 字段和 individual override 字段
    override_fields = _collect_override_fields(rule)
    aggregate_fields_by_report = {}
    individual_fields_by_report = {}

    for report_name, field_mapping in field_mapping_by_report.items():
        agg_fields = {}
        ind_fields = {}
        for field_name, col_name in field_mapping.items():
            if field_name in override_fields:
                ind_fields[field_name] = col_name
            else:
                agg_fields[field_name] = col_name
        if agg_fields:
            aggregate_fields_by_report[report_name] = agg_fields
        if ind_fields:
            individual_fields_by_report[report_name] = ind_fields

    # Step A: 从汇总行抽取 aggregate 字段
    agg_record = {}
    for report_name, field_mapping in aggregate_fields_by_report.items():
        if report_name not in raw_report_data:
            logger.warning(f"aggregate: 报表 '{report_name}' 未加载")
            continue

        agg_config = _resolve_aggregate_config(report_name, aggregate_key, column_mapping)
        if not agg_config:
            # Fallback: 尝试 composite_keys 合成聚合
            composite_row = _resolve_composite_aggregate(
                raw_report_data[report_name], aggregate_key, report_name, column_mapping
            )
            if composite_row:
                reconcile = column_mapping.get(report_name, {}).get("column_reconcile", {})
                reconciled = reconcile_field_mapping(field_mapping, reconcile)
                reverse_map = {v: k for k, v in reconciled.items()}
                for col_name, value in composite_row.items():
                    if col_name in reverse_map:
                        agg_record[reverse_map[col_name]] = normalize_value(value)
                for field_name, col_name in reconciled.items():
                    if field_name not in agg_record:
                        agg_record[field_name] = None
            else:
                for field_name in field_mapping:
                    agg_record[field_name] = None
            continue

        agg_row = extract_aggregate_row(
            raw_report_data[report_name],
            agg_config["match_column"],
            agg_config["match_value"],
            agg_config["match_type"],
            id_column=agg_config.get("id_column"),
            id_value=agg_config.get("id_value"),
        )
        if not agg_row:
            for field_name in field_mapping:
                agg_record[field_name] = None
            continue

        # reconcile + 抽字段
        reconcile = column_mapping.get(report_name, {}).get("column_reconcile", {})
        reconciled = reconcile_field_mapping(field_mapping, reconcile)

        reverse_map = {v: k for k, v in reconciled.items()}
        for col_name, value in agg_row.items():
            if col_name in reverse_map:
                agg_record[reverse_map[col_name]] = normalize_value(value)
        # 补缺失字段
        for field_name, col_name in reconciled.items():
            if field_name not in agg_record:
                agg_record[field_name] = None

    # Step B: 为每个 specific_employee 抽取 individual override 字段
    employee_individual_data = {}
    if override_fields and specific_employees:
        for emp_name in specific_employees:
            emp_record = {}
            for report_name, field_mapping in individual_fields_by_report.items():
                if report_name not in raw_report_data:
                    continue

                # 用 id_column_override 或 column_reconcile 中的 员工编号 映射来定位个人行
                reconcile = column_mapping.get(report_name, {}).get("column_reconcile", {})
                id_col = reconcile.get("员工编号", column_mapping.get(report_name, {}).get("read_config", {}).get("id_column_override", ""))

                # 找到该员工的个人行
                emp_row = None
                for row in raw_report_data[report_name]:
                    cell_val = row.get(id_col)
                    if cell_val is not None and str(cell_val).strip() == emp_name:
                        emp_row = row
                        break

                if not emp_row:
                    logger.warning(f"aggregate override: '{emp_name}' 在 '{report_name}' 中未找到个人行")
                    for field_name in field_mapping:
                        emp_record[field_name] = None
                    continue

                reconciled = reconcile_field_mapping(field_mapping, reconcile)
                reverse_map = {v: k for k, v in reconciled.items()}
                for col_name, value in emp_row.items():
                    if col_name in reverse_map:
                        emp_record[reverse_map[col_name]] = normalize_value(value)
                for field_name, col_name in reconciled.items():
                    if field_name not in emp_record:
                        emp_record[field_name] = None

            employee_individual_data[emp_name] = emp_record

    # Step C: 合并 aggregate + individual override，展开为 N 条
    results = []
    if not specific_employees:
        logger.warning("aggregate 模式但 specific_employees 为空，无法展开")
        return results

    for emp_name in specific_employees:
        combined = {"employee_id": emp_name}
        combined.update(agg_record)  # aggregate 字段
        if emp_name in employee_individual_data:
            combined.update(employee_individual_data[emp_name])  # override 字段覆盖
        # 确保 employee_id 被人名覆盖（aggregate 行可能有不同的 id 值）
        combined["employee_id"] = emp_name
        results.append(combined)

    logger.info(f"aggregate 模式: {len(specific_employees)} 名员工, aggregate_key='{aggregate_key}'")
    return results


def _collect_override_fields(rule: dict) -> set[str]:
    """收集 rule 中所有有 data_scope_override=individual 的字段名。"""
    override_fields = set()
    calc_rule = rule.get("calculation_rule", {})
    method = calc_rule.get("method", "")

    if method == "ratio_based":
        for item in calc_rule.get("items", []):
            override = item.get("data_scope_override")
            if override and override.get("mode") == "individual":
                override_fields.add(item.get("target_field", ""))
    elif method == "tiered_commission":
        override = calc_rule.get("data_scope_override")
        if override and override.get("mode") == "individual":
            override_fields.add(calc_rule.get("performance_field", ""))
    elif method == "fixed_bonus":
        override = calc_rule.get("data_scope_override")
        if override and override.get("mode") == "individual":
            cond_field = calc_rule.get("condition", {}).get("field", "")
            override_fields.add(cond_field)
    elif method == "per_unit":
        override = calc_rule.get("data_scope_override")
        if override and override.get("mode") == "individual":
            override_fields.add(calc_rule.get("quantity_field", ""))

    return override_fields


def build_employee_data(
    raw_report_data: dict[str, list[dict]],
    rule: dict | list,
    column_mapping: dict | None = None,
) -> list[dict]:
    """完整管道：原始报表数据 → 字段重定向 → 列名对齐 → 列名映射 → 多报表合并。

    支持 individual 和 aggregate 两种 data_scope 模式：
    - individual（默认）: 原有路径
    - aggregate: 新路径，从汇总行抽取数据并展开给 specific_employees

    对于 list 型 rule（多 plan），不同 plan 可有不同 mode，分别处理后合并。
    """
    column_mapping = column_mapping or {}

    # 处理 list 型 rule：逐 plan 分别处理
    plans = rule if isinstance(rule, list) else [rule]

    all_results = []
    for plan in plans:
        data_scope = plan.get("data_scope", {"mode": "individual"})
        mode = data_scope.get("mode", "individual")

        if mode == "aggregate":
            result = build_aggregate_employee_data(
                raw_report_data, plan, column_mapping
            )
        else:
            # individual 模式：原有路径（对单 plan 处理）
            result = _build_individual_employee_data(
                raw_report_data, plan, column_mapping
            )

        all_results.extend(result)

    # 如果只有 1 个 plan 且是 individual 模式，直接返回原有逻辑的结果
    # 如果有多个 plan，合并所有结果
    if len(plans) == 1 and plans[0].get("data_scope", {"mode": "individual"}).get("mode", "individual") == "individual":
        return all_results

    # 多 plan 合并或 aggregate + individual 混合：按 employee_id union
    if len(all_results) > 0 and len(plans) > 1:
        master = {}
        for emp in all_results:
            eid = emp.get("employee_id")
            if eid is None:
                continue
            if eid not in master:
                master[eid] = {}
            master[eid].update(emp)
        return list(master.values())

    return all_results


def _build_individual_employee_data(
    raw_report_data: dict[str, list[dict]],
    rule: dict,
    column_mapping: dict,
) -> list[dict]:
    """individual 模式的数据构建：原有路径。"""
    # 1. 收集按报表分组的字段映射
    field_mapping_by_report = collect_field_mapping_by_report(rule)

    # 2. 应用字段重定向：将 rule 错误分配的字段移到正确报表
    field_redirect = column_mapping.get("field_redirect", {})
    if field_redirect:
        for field_name, redirect in field_redirect.items():
            from_report = redirect.get("from_report", "")
            to_report = redirect.get("to_report", "")
            actual_column = redirect.get("actual_column", "")

            # 仅当 rule 实际引用了 from_report 中的该字段时才转移
            if from_report not in field_mapping_by_report:
                continue
            if field_name not in field_mapping_by_report[from_report]:
                continue

            del field_mapping_by_report[from_report][field_name]
            logger.info(f"字段重定向: '{field_name}' 从 '{from_report}' → '{to_report}'")

            # 添加到目标报表
            if to_report not in field_mapping_by_report:
                field_mapping_by_report[to_report] = {}
            field_mapping_by_report[to_report][field_name] = actual_column

    # 3. 逐报表做列名对齐 + 列名映射
    mapped_reports = {}
    for report_name, field_mapping in field_mapping_by_report.items():
        if report_name not in raw_report_data:
            logger.warning(f"规则引用了报表 '{report_name}'，但未加载该报表数据")
            continue

        # 列名对齐：概念列名 → 实际列名
        reconcile = column_mapping.get(report_name, {}).get("column_reconcile", {})
        reconciled_mapping = reconcile_field_mapping(field_mapping, reconcile)

        mapped_reports[report_name] = map_columns_to_fields(
            raw_report_data[report_name], reconciled_mapping
        )

    # 4. 合并多报表
    return merge_multi_report_data(mapped_reports)