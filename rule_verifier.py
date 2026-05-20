"""rule_verifier — 解析结果验证器，在 LLM 解析后、计算前进行多层校验。

验证层次：
1. 解析结果可视化对比：将规则JSON转为人类可读摘要，与原文关键信息对照
2. 关键数值双重验证：从原文正则提取金额/比例，与解析结果对比
3. 字段映射预检查：检查 data_source_mapping 中的报表名和列名是否真实存在
4. 试算验证：用真实数据进行1-2条试算，展示计算过程

使用方式：
    verifier = RuleVerifier(data_dir, column_mapping)
    report = verifier.verify(parsed_rule, plan_text)
    report.print_summary()
"""

import logging
import re
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

from report_registry import ReportRegistry, DEFAULT_REGISTRY

logger = logging.getLogger("incentive_bot")


# ---------------------------------------------------------------------------
# Data Classes
# ---------------------------------------------------------------------------

@dataclass
class VerifyItem:
    """单条验证结果"""
    level: str  # "ok", "warn", "error"
    category: str  # "number", "mapping", "trial"
    message: str
    detail: str = ""


@dataclass
class VerifyReport:
    """完整验证报告"""
    items: list[VerifyItem] = field(default_factory=list)

    @property
    def has_errors(self) -> bool:
        return any(i.level == "error" for i in self.items)

    @property
    def has_warnings(self) -> bool:
        return any(i.level == "warn" for i in self.items)

    def print_summary(self):
        symbols = {"ok": "✓", "warn": "⚠", "error": "✗"}
        print("\n" + "=" * 60)
        print("  规则验证报告")
        print("=" * 60)

        for cat in ["compare", "number", "mapping", "trial"]:
            cat_items = [i for i in self.items if i.category == cat]
            if not cat_items:
                continue
            cat_names = {
                "compare": "解析结果可视化对比",
                "number": "关键数值验证",
                "mapping": "字段映射检查",
                "trial": "试算验证",
            }
            print(f"\n--- {cat_names[cat]} ---")
            for item in cat_items:
                sym = symbols[item.level]
                print(f"  {sym} {item.message}")
                if item.detail:
                    for line in item.detail.split("\n"):
                        print(f"    {line}")

        errors = sum(1 for i in self.items if i.level == "error")
        warns = sum(1 for i in self.items if i.level == "warn")
        oks = sum(1 for i in self.items if i.level == "ok")
        print(f"\n结果: {oks} 通过, {warns} 警告, {errors} 错误")
        if errors:
            print("⚠ 存在错误，建议检查后再进行计算")
        print("=" * 60 + "\n")



# ---------------------------------------------------------------------------
# Visual Comparison (解析结果可视化对比)
# ---------------------------------------------------------------------------

_METHOD_NAMES = {
    "ratio_based": "比例达标激励",
    "tiered_commission": "阶梯提成",
    "fixed_bonus": "固定奖金",
    "per_unit": "按件计酬",
}

_ROLE_NAMES = {"LP": "LP（班主任）", "CC": "CC（销售）"}


def generate_comparison(rule: dict | list, plan_text: str | None = None) -> list[VerifyItem]:
    """生成解析结果的可视化对比摘要。

    将每个 plan 的关键信息以人类可读格式展示，方便用户快速确认
    LLM 是否正确理解了方案文档。
    """
    items = []
    plans = rule if isinstance(rule, list) else [rule]

    items.append(VerifyItem("ok", "compare",
        f"共解析出 {len(plans)} 个激励方案（plan）"))

    for idx, plan in enumerate(plans):
        detail = _format_plan_summary(plan)
        plan_name = plan.get("plan_name", f"方案{idx+1}")
        plan_id = plan.get("plan_id", "?")
        items.append(VerifyItem("ok", "compare",
            f"[{idx+1}] {plan_name} ({plan_id})",
            detail))

    return items


def _format_plan_summary(plan: dict) -> str:
    """将单个 plan 格式化为可读摘要"""
    lines = []

    # 基本信息
    month = plan.get("effective_month", "?")
    plan_type = plan.get("plan_type", "?")
    lines.append(f"生效月份: {month} | 类型: {plan_type}")

    # 激励对象
    scope = plan.get("employee_scope", {})
    roles = scope.get("role", [])
    role_str = ", ".join(_ROLE_NAMES.get(r, r) for r in roles)
    lines.append(f"激励对象: {role_str}")
    if scope.get("department"):
        lines.append(f"  部门: {', '.join(scope['department'])}")
    if scope.get("specific_employees"):
        lines.append(f"  指定人员: {', '.join(scope['specific_employees'])}")
    if scope.get("exclude"):
        lines.append(f"  排除: {', '.join(scope['exclude'])}")

    # 数据口径
    data_scope = plan.get("data_scope", {})
    mode = data_scope.get("mode", "individual")
    if mode == "aggregate":
        agg_key = data_scope.get("aggregate_key", "?")
        lines.append(f"数据口径: 整体聚合 (aggregate_key={agg_key})")
    else:
        lines.append(f"数据口径: 个人 (individual)")

    # 前置条件
    prereqs = plan.get("prerequisites", [])
    if prereqs:
        lines.append(f"前置条件 ({len(prereqs)}项):")
        for p in prereqs:
            lines.append(f"  - {p['field']} {p['operator']} {p['value']}")
    else:
        lines.append("前置条件: 无")

    # 计算规则
    calc = plan.get("calculation_rule", {})
    method = calc.get("method", "?")
    method_name = _METHOD_NAMES.get(method, method)
    lines.append(f"计算方法: {method_name}")

    if method == "ratio_based":
        items_list = calc.get("items", [])
        lines.append(f"  子项数量: {len(items_list)}")
        for item in items_list:
            name = item.get("name", "?")
            base = item.get("base_amount", 0)
            target_field = item.get("target_field", "?")
            target_val = item.get("target_value", 0)
            override = item.get("data_scope_override")
            override_str = f" [个人口径]" if override and override.get("mode") == "individual" else ""
            if target_val <= 1:
                lines.append(f"  - {name}: 基数{base}元, 目标{target_field}≥{target_val*100:.0f}%{override_str}")
            else:
                lines.append(f"  - {name}: 基数{base}元, 目标{target_field}≥{target_val}{override_str}")

    elif method == "tiered_commission":
        perf_field = calc.get("performance_field", "?")
        tiers = calc.get("tiers", [])
        lines.append(f"  业绩字段: {perf_field}")
        lines.append(f"  阶梯档位 ({len(tiers)}档):")
        for t in tiers:
            max_str = f"{t['max']}" if t.get("max") is not None else "∞"
            lines.append(f"    {t['min']} ~ {max_str}: {t['rate']*100:.1f}%")
        deductions = calc.get("deductions", [])
        if deductions:
            for d in deductions:
                lines.append(f"  扣减: {d['name']} ({d['field']} × {d['rate']*100:.0f}%)")

    elif method == "fixed_bonus":
        amount = calc.get("amount", 0)
        cond = calc.get("condition", {})
        lines.append(f"  奖金金额: {amount}元")
        if cond:
            lines.append(f"  达成条件: {cond.get('field','?')} {cond.get('operator','?')} {cond.get('value','?')}")

    elif method == "per_unit":
        unit_price = calc.get("unit_price", 0)
        qty_field = calc.get("quantity_field", "?")
        lines.append(f"  单价: {unit_price}元/件")
        lines.append(f"  数量字段: {qty_field}")

    # 封顶
    capping = calc.get("capping", {})
    if capping:
        cap_parts = []
        if capping.get("max_amount"):
            cap_parts.append(f"金额≤{capping['max_amount']}元")
        if capping.get("max_quantity"):
            cap_parts.append(f"数量≤{capping['max_quantity']}")
        if cap_parts:
            lines.append(f"封顶: {', '.join(cap_parts)}")

    # 取整
    rounding = calc.get("rounding_rule", {})
    if rounding:
        direction_map = {"up": "向上", "down": "向下", "half_up": "四舍五入", "none": "不取整"}
        d = direction_map.get(rounding.get("direction", "up"), "?")
        p = rounding.get("precision", 2)
        lines.append(f"取整: {d}取整, 保留{p}位小数")

    # 数据源映射
    mapping = plan.get("data_source_mapping", {})
    lines.append(f"数据源映射 ({len(mapping)}个字段):")
    for field_name, source in mapping.items():
        lines.append(f"  {field_name} ← {source['report']}.{source['column']}")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Number Extraction & Verification
# ---------------------------------------------------------------------------

def extract_amounts_from_text(text: str) -> list[float]:
    """从方案原文中提取所有可能是金额的数值。

    覆盖多种写法：
    - 标准 "X元"
    - 公式中的基数 "=200*..." / "200x..." / "200×..."
    - 封顶 "封顶X" / "上限X"
    - 大于阈值的独立整数（≥50）
    """
    amounts = set()
    # 1. 带"元"
    for m in re.finditer(r'(\d+(?:\.\d+)?)\s*元', text):
        amounts.add(float(m.group(1)))
    # 2. 公式基数 = / × / x / *
    for m in re.finditer(r'[=＝]\s*(\d+(?:\.\d+)?)\s*[\*×x]', text):
        amounts.add(float(m.group(1)))
    # 3. 所有不小于 50 的独立整数（金额起步通常 ≥50）
    for m in re.finditer(r'(?<![\d.])(\d{2,})(?![\d.%])', text):
        v = float(m.group(1))
        if v >= 50:
            amounts.add(v)
    return sorted(amounts)


def extract_percentages_from_text(text: str) -> list[float]:
    """从方案原文中提取所有百分比数值"""
    percentages = []
    for m in re.finditer(r'(\d+(?:\.\d+)?)\s*%', text):
        percentages.append(float(m.group(1)) / 100)
    return percentages


def extract_amounts_from_rule(rule: dict) -> list[float]:
    """从解析后的规则JSON中提取所有金额"""
    amounts = []
    plans = rule if isinstance(rule, list) else [rule]
    for plan in plans:
        calc = plan.get("calculation_rule", {})
        method = calc.get("method", "")
        capping = calc.get("capping", {})
        if capping.get("max_amount"):
            amounts.append(capping["max_amount"])
        if method == "ratio_based":
            for item in calc.get("items", []):
                if item.get("base_amount"):
                    amounts.append(item["base_amount"])
        elif method == "fixed_bonus":
            if calc.get("amount"):
                amounts.append(calc["amount"])
        elif method == "per_unit":
            if calc.get("unit_price"):
                amounts.append(calc["unit_price"])
    return amounts


def extract_percentages_from_rule(rule: dict) -> list[float]:
    """从解析后的规则JSON中提取所有百分比/比例"""
    percentages = []
    plans = rule if isinstance(rule, list) else [rule]
    for plan in plans:
        calc = plan.get("calculation_rule", {})
        method = calc.get("method", "")
        if method == "ratio_based":
            for item in calc.get("items", []):
                tv = item.get("target_value")
                if tv and tv <= 1.0:
                    percentages.append(tv)
        elif method == "tiered_commission":
            for tier in calc.get("tiers", []):
                if tier.get("rate"):
                    percentages.append(tier["rate"])
        for pre in plan.get("prerequisites", []):
            v = pre.get("value")
            if isinstance(v, (int, float)) and 0 < v <= 1.0:
                percentages.append(v)
    return percentages


def verify_numbers(plan_text: str, rule: dict) -> list[VerifyItem]:
    """关键数值双重验证：检查规则中的金额/比例是否都能在原文中找到"""
    items = []

    text_amounts = set(extract_amounts_from_text(plan_text))
    rule_amounts = extract_amounts_from_rule(rule)

    text_pcts = extract_percentages_from_text(plan_text)
    text_pcts_pct = {round(p * 100, 1) for p in text_pcts}
    rule_pcts = extract_percentages_from_rule(rule)

    # 仅做"规则中数值能否在原文找到"的单向检查（反向噪声太多）
    for amt in rule_amounts:
        if amt in text_amounts:
            items.append(VerifyItem("ok", "number", f"金额 {amt}元 在原文中确认"))
        else:
            items.append(VerifyItem("warn", "number",
                f"金额 {amt}元 在原文中未直接找到",
                "可能是LLM推断的值或原文格式特殊，请人工确认"))

    for pct in rule_pcts:
        pct_display = f"{pct*100:.1f}%"
        if round(pct * 100, 1) in text_pcts_pct:
            items.append(VerifyItem("ok", "number", f"比例 {pct_display} 在原文中确认"))
        else:
            items.append(VerifyItem("warn", "number",
                f"比例 {pct_display} 在原文中未直接找到",
                "可能是LLM推断或转换后的值"))

    if not items:
        items.append(VerifyItem("ok", "number", "未提取到需要验证的数值"))

    return items



# ---------------------------------------------------------------------------
# Field Mapping Pre-check
# ---------------------------------------------------------------------------

def verify_mapping(rule: dict, data_dir: Path, column_mapping: dict) -> list[VerifyItem]:
    """检查 data_source_mapping 中的报表名和列名是否真实存在"""
    items = []
    plans = rule if isinstance(rule, list) else [rule]
    registry = ReportRegistry(data_dir=data_dir)

    for plan in plans:
        plan_id = plan.get("plan_id", "unknown")
        mapping = plan.get("data_source_mapping", {})
        month = plan.get("effective_month", "2026-01")

        for field_name, source in mapping.items():
            report_name = source.get("report", "")
            column_name = source.get("column", "")

            # 检查报表是否已注册
            if report_name not in DEFAULT_REGISTRY:
                items.append(VerifyItem("error", "mapping",
                    f"[{plan_id}] 报表 '{report_name}' 未在 report_registry 中注册",
                    f"字段: {field_name}, 已注册报表: {list(DEFAULT_REGISTRY.keys())}"))
                continue

            # 检查报表文件是否存在
            try:
                filepath = registry.resolve_path(report_name, month)
            except (FileNotFoundError, KeyError) as e:
                items.append(VerifyItem("error", "mapping",
                    f"[{plan_id}] 报表 '{report_name}' 文件未找到",
                    f"字段: {field_name}, 错误: {e}"))
                continue

            # 检查列名是否存在（考虑 column_reconcile）
            reconcile = column_mapping.get(report_name, {}).get("column_reconcile", {})
            actual_col = reconcile.get(column_name, column_name)

            # 也检查 field_redirect
            field_redirect = column_mapping.get("field_redirect", {})
            if field_name in field_redirect:
                redirect = field_redirect[field_name]
                actual_col = redirect.get("actual_column", actual_col)
                report_name_actual = redirect.get("to_report", report_name)
                try:
                    filepath = registry.resolve_path(report_name_actual, month)
                except (FileNotFoundError, KeyError):
                    pass

            # 读取Excel表头验证列名
            col_exists = _check_column_in_excel(filepath, actual_col, column_mapping.get(report_name, {}))
            if col_exists:
                items.append(VerifyItem("ok", "mapping",
                    f"[{plan_id}] {field_name} → {report_name}.{actual_col} ✓"))
            else:
                items.append(VerifyItem("warn", "mapping",
                    f"[{plan_id}] 列 '{actual_col}' 在报表 '{report_name}' 中未确认",
                    f"字段: {field_name}, 可能是列名映射问题"))

    return items


def _check_column_in_excel(filepath: Path, column_name: str, report_config: dict) -> bool:
    """检查Excel文件中是否存在指定列名。

    使用 ExcelReportReader 实际读取后的列名集合（含 data_loader 的多行标题处理
    与 _1/_2 等去重后缀），与 data_source_mapping 实际消费的列名一致。
    """
    try:
        from data_loader import ExcelReportReader
        read_config = report_config.get("read_config", {})
        header_rows = read_config.get("header_rows", 1)
        data_start_row = read_config.get("data_start_row")

        reader = ExcelReportReader()
        rows = reader.read_file(
            filepath,
            header_rows=header_rows,
            data_start_row=data_start_row,
        )
        if not rows:
            return False
        headers = set(rows[0].keys())
        return column_name in headers
    except Exception as e:
        logger.debug(f"检查列名时出错: {e}")
        return False



# ---------------------------------------------------------------------------
# Trial Calculation
# ---------------------------------------------------------------------------

def verify_trial_calculation(rule: dict, employee_data: list[dict], sample_size: int = 1) -> list[VerifyItem]:
    """用真实数据进行试算，展示计算过程"""
    items = []

    if not employee_data:
        items.append(VerifyItem("warn", "trial", "无可用员工数据，跳过试算"))
        return items

    from incentive_calculate import skill_calculate, skill_calculate_batch

    plans = rule if isinstance(rule, list) else [rule]
    samples = employee_data[:sample_size]

    for plan in plans:
        plan_id = plan.get("plan_id", "unknown")
        plan_name = plan.get("plan_name", "")
        try:
            results = skill_calculate(plan, samples)
            for result in results:
                emp_id = result.get("employee_id", "?")
                eligible = result.get("eligible", False)
                amount = result.get("amount", 0)
                breakdown = result.get("breakdown", [])

                detail_lines = [f"员工: {emp_id}, 资格: {eligible}, 金额: {amount}元"]
                for bd in breakdown[:5]:
                    name = bd.get("name", "")
                    sub_amount = bd.get("amount", 0)
                    detail_lines.append(f"  - {name}: {sub_amount}元")

                items.append(VerifyItem("ok", "trial",
                    f"[{plan_id}] {plan_name} 试算成功",
                    "\n".join(detail_lines)))
        except Exception as e:
            items.append(VerifyItem("error", "trial",
                f"[{plan_id}] {plan_name} 试算失败",
                f"错误: {e}"))

    return items


# ---------------------------------------------------------------------------
# Main Verifier
# ---------------------------------------------------------------------------

class RuleVerifier:
    """规则验证器，整合数值验证、映射检查和试算"""

    def __init__(self, data_dir: Path, column_mapping: dict):
        self._data_dir = data_dir
        self._column_mapping = column_mapping

    def verify(
        self,
        rule: dict | list,
        plan_text: str | None = None,
        employee_data: list[dict] | None = None,
        skip_compare: bool = False,
        skip_numbers: bool = False,
        skip_mapping: bool = False,
        skip_trial: bool = False,
    ) -> VerifyReport:
        """执行多层验证"""
        report = VerifyReport()

        if not skip_compare:
            report.items.extend(generate_comparison(rule, plan_text))

        if not skip_numbers and plan_text:
            report.items.extend(verify_numbers(plan_text, rule))

        if not skip_mapping:
            report.items.extend(verify_mapping(rule, self._data_dir, self._column_mapping))

        if not skip_trial and employee_data:
            report.items.extend(verify_trial_calculation(rule, employee_data))

        return report
