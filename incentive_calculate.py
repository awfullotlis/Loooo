"""skill_calculate — 每月激励方案的计算执行 Skill。

职责：
1. 接收 rule_phrase 输出的结构化规则 JSON + 报表数据
2. 执行 4 种计算方法：ratio_based / tiered_commission / fixed_bonus / per_unit
3. 处理前置条件、封顶、取整、扣减等通用逻辑
4. 校验计算结果（负值、封顶越界等）
5. 输出每个员工的激励金额明细

使用方式：
    result = skill_calculate(rule=plan_dict, data=report_data_dict)
"""

import math
from typing import Any


# ---------------------------------------------------------------------------
# Rounding
# ---------------------------------------------------------------------------

def _apply_rounding(value: float, rounding_rule: dict) -> float:
    """Apply rounding rule to a calculated value."""
    direction = rounding_rule.get("direction", "up")
    precision = rounding_rule.get("precision", 2)

    if direction == "none":
        return value

    factor = 10 ** precision
    shifted = value * factor

    if direction == "up":
        result = math.ceil(shifted) / factor
    elif direction == "down":
        result = math.floor(shifted) / factor
    elif direction == "half_up":
        result = round(shifted) / factor
    else:
        result = round(shifted) / factor

    return result


# ---------------------------------------------------------------------------
# Condition Evaluation
# ---------------------------------------------------------------------------

def _evaluate_condition(field: str, operator: str, value: Any, data: dict) -> bool:
    """Evaluate a single condition against employee data."""
    actual = data.get(field)
    if actual is None:
        return False

    # 数值比较前的强制转换：处理报表里残留的字符串数字（含百分号、空白等）
    if operator in (">", "<", ">=", "<=") and not isinstance(actual, (int, float)):
        try:
            s = str(actual).strip().rstrip("%")
            if not s:
                return False
            actual = float(s)
        except (TypeError, ValueError):
            return False

    if operator == "==":
        return actual == value
    elif operator == "!=":
        return actual != value
    elif operator == ">":
        return actual > value
    elif operator == "<":
        return actual < value
    elif operator == ">=":
        return actual >= value
    elif operator == "<=":
        return actual <= value
    elif operator == "in":
        return actual in value
    else:
        raise ValueError(f"Unknown operator: {operator}")


def _check_prerequisites(prerequisites: list, data: dict) -> bool:
    """Check all prerequisite conditions. Employee must pass ALL to be eligible."""
    for cond in prerequisites:
        if not _evaluate_condition(cond["field"], cond["operator"], cond["value"], data):
            return False
    return True


# ---------------------------------------------------------------------------
# Calculation Methods
# ---------------------------------------------------------------------------

def _calc_ratio_based(rule: dict, data: dict) -> dict:
    """ratio_based: each item = base_amount * (actual / target_value)"""
    items = rule.get("items", [])
    breakdown = []
    total = 0.0

    for item in items:
        actual = data.get(item["target_field"], 0) or 0
        target = item["target_value"]
        base = item["base_amount"]
        if target <= 0:
            raise ValueError(f"item '{item['name']}': target_value must be > 0")
        item_amount = base * (actual / target)
        breakdown.append({
            "name": item["name"],
            "base_amount": base,
            "actual": actual,
            "target_value": target,
            "raw_amount": item_amount,
        })
        total += item_amount

    return {"breakdown": breakdown, "total_raw": total}


def _calc_tiered_commission(rule: dict, data: dict) -> dict:
    """tiered_commission: apply tier rates to performance amount, with optional deductions."""
    performance = data.get(rule.get("performance_field", ""), 0) or 0
    tiers = rule.get("tiers", [])
    deductions = rule.get("deductions", [])

    # Calculate commission per tier
    tier_breakdown = []
    total_commission = 0.0

    for i, tier in enumerate(tiers):
        min_val = tier["min"]
        max_val = tier.get("max")  # None means unlimited
        rate = tier["rate"]

        # Determine the portion of performance that falls in this tier
        if performance <= min_val:
            tier_amount = 0
        else:
            upper = max_val if max_val is not None else performance
            tier_base = min(performance, upper) - min_val
            if tier_base < 0:
                tier_base = 0
            tier_amount = tier_base * rate

        tier_breakdown.append({
            "tier_index": i,
            "min": min_val,
            "max": max_val,
            "rate": rate,
            "tier_base": min(performance, upper) - min_val if performance > min_val else 0,
            "tier_commission": tier_amount,
        })
        total_commission += tier_amount

    # Apply deductions
    deduction_breakdown = []
    total_deduction = 0.0
    for ded in deductions:
        ded_amount = data.get(ded["field"], 0) * ded["rate"]
        deduction_breakdown.append({
            "name": ded["name"],
            "deduction_raw": data.get(ded["field"], 0),
            "deduction_rate": ded["rate"],
            "deduction_amount": ded_amount,
        })
        total_deduction += ded_amount

    net = total_commission - total_deduction
    return {
        "performance": performance,
        "tier_breakdown": tier_breakdown,
        "total_commission": total_commission,
        "deduction_breakdown": deduction_breakdown,
        "total_deduction": total_deduction,
        "total_raw": net,
    }


def _calc_fixed_bonus(rule: dict, data: dict) -> dict:
    """fixed_bonus: fixed amount if condition is met, otherwise 0."""
    amount = rule.get("amount", 0)
    condition = rule.get("condition", {})

    met = _evaluate_condition(condition["field"], condition["operator"], condition["value"], data)

    return {
        "amount": amount,
        "condition_field": condition["field"],
        "condition_operator": condition["operator"],
        "condition_value": condition["value"],
        "condition_actual": data.get(condition["field"]),
        "condition_met": met,
        "total_raw": amount if met else 0,
    }


def _calc_per_unit(rule: dict, data: dict) -> dict:
    """per_unit: unit_price * quantity, with optional max_quantity cap."""
    unit_price = rule.get("unit_price", 0)
    quantity = data.get(rule.get("quantity_field", ""), 0) or 0
    capping = rule.get("capping", {})

    # Apply quantity cap first
    max_qty = capping.get("max_quantity")
    capped_qty = min(quantity, max_qty) if max_qty is not None else quantity

    total = unit_price * capped_qty

    return {
        "unit_price": unit_price,
        "quantity_raw": quantity,
        "max_quantity": max_qty,
        "quantity_capped": capped_qty,
        "total_raw": total,
    }


def _calc_rank_tiered(rule: dict, all_data: list[dict]) -> dict[str, dict]:
    """rank_tiered: 跨员工排序，按排名给固定档位金额。

    流程:
    1. 用 qualifier 筛出达标员工
    2. 按 rank_field 降序（或升序）排名，tiebreaker_field 处理同分
    3. 按 rank_rewards 给前N名分配金额，未达标或排名外的为 0

    Returns: {employee_id: {rank, amount, qualifier_met, qualifier_actual}}
    """
    qualifier = rule.get("qualifier", {})
    rank_field = rule.get("rank_field")
    rank_order = rule.get("rank_order", "desc")
    rank_rewards = rule.get("rank_rewards", [])
    tiebreaker_field = rule.get("tiebreaker_field")

    if not rank_field:
        raise ValueError("rank_tiered requires rank_field")
    if not qualifier:
        raise ValueError("rank_tiered requires qualifier")

    # 1. Filter by qualifier
    qualified = []
    detail_by_emp = {}
    for emp in all_data:
        emp_id = emp.get("employee_id", "UNKNOWN")
        actual = emp.get(qualifier["field"])
        met = _evaluate_condition(qualifier["field"], qualifier["operator"], qualifier["value"], emp)
        detail_by_emp[emp_id] = {
            "qualifier_field": qualifier["field"],
            "qualifier_actual": actual,
            "qualifier_met": met,
            "rank": None,
            "rank_value": emp.get(rank_field),
            "amount": 0,
            "total_raw": 0,
        }
        if met:
            qualified.append(emp)

    # 2. Sort qualified by rank_field
    reverse = (rank_order == "desc")
    def sort_key(e):
        primary = e.get(rank_field, 0) or 0
        secondary = e.get(tiebreaker_field, 0) or 0 if tiebreaker_field else 0
        # For desc, larger comes first; for tiebreaker also use the same direction
        return (primary, secondary)
    qualified.sort(key=sort_key, reverse=reverse)

    # 3. Assign rewards by rank
    rewards_by_rank = {r["rank"]: r["amount"] for r in rank_rewards}
    for idx, emp in enumerate(qualified):
        emp_id = emp.get("employee_id", "UNKNOWN")
        rank = idx + 1
        amount = rewards_by_rank.get(rank, 0)
        detail_by_emp[emp_id]["rank"] = rank
        detail_by_emp[emp_id]["amount"] = amount
        detail_by_emp[emp_id]["total_raw"] = amount

    return detail_by_emp


# ---------------------------------------------------------------------------
# Result Validation
# ---------------------------------------------------------------------------

def _validate_result(amount: float, rule: dict, employee_id: str) -> list[str]:
    """Check calculated result for anomalies."""
    errors = []

    if amount < 0:
        errors.append(f"[{employee_id}] calculated amount is negative: {amount}")

    capping = rule.get("capping", {})
    max_amount = capping.get("max_amount")
    if max_amount and amount > max_amount:
        errors.append(f"[{employee_id}] amount {amount} exceeds capping {max_amount} — capping should have been applied")

    return errors


# ---------------------------------------------------------------------------
# Main Skill Function
# ---------------------------------------------------------------------------

def skill_calculate(rule: dict, data: list[dict]) -> list[dict]:
    """Calculate incentive amounts for all employees based on a single plan rule.

    Args:
        rule:  Structured rule JSON (output from rule_phrase), must conform to rule_schema.json
        data:  List of employee data dicts. Each dict must contain fields referenced in the rule
               (field names as defined in data_source_mapping keys)

    Returns:
        List of result dicts, one per employee:
        {
            "employee_id": str,
            "plan_id": str,
            "plan_name": str,
            "eligible": bool,
            "amount": float,
            "detail": dict,        # method-specific breakdown
            "validation_errors": list[str]
        }
    """
    calc_rule = rule["calculation_rule"]
    method = calc_rule["method"]
    prerequisites = rule.get("prerequisites", [])
    rounding_rule = calc_rule.get("rounding_rule", {"direction": "up", "precision": 2})
    capping = calc_rule.get("capping", {})
    max_amount = capping.get("max_amount")

    # rank_tiered 需要跨员工计算，先一次性算好排名结果
    rank_results = None
    if method == "rank_tiered":
        # 仅对通过 prerequisites 的员工参与排名
        ranked_pool = [e for e in data if _check_prerequisites(prerequisites, e)]
        rank_results = _calc_rank_tiered(calc_rule, ranked_pool)

    results = []

    for emp_data in data:
        employee_id = emp_data.get("employee_id", "UNKNOWN")

        # 1. Check prerequisites
        eligible = _check_prerequisites(prerequisites, emp_data)

        if not eligible:
            results.append({
                "employee_id": employee_id,
                "plan_id": rule.get("plan_id", ""),
                "plan_name": rule.get("plan_name", ""),
                "eligible": False,
                "amount": 0,
                "detail": {"reason": "prerequisites not met"},
                "validation_errors": [],
            })
            continue

        # 2. Calculate by method
        if method == "ratio_based":
            detail = _calc_ratio_based(calc_rule, emp_data)
        elif method == "tiered_commission":
            detail = _calc_tiered_commission(calc_rule, emp_data)
        elif method == "fixed_bonus":
            detail = _calc_fixed_bonus(calc_rule, emp_data)
        elif method == "per_unit":
            detail = _calc_per_unit(calc_rule, emp_data)
        elif method == "rank_tiered":
            detail = rank_results.get(employee_id, {
                "qualifier_met": False, "rank": None, "amount": 0, "total_raw": 0,
            })
        else:
            raise ValueError(f"Unknown method: {method}")

        raw_amount = detail.get("total_raw", 0)

        # 3. Apply capping
        capped_amount = min(raw_amount, max_amount) if max_amount is not None else raw_amount

        # 4. Apply rounding
        final_amount = _apply_rounding(capped_amount, rounding_rule)

        # 5. Validate result
        validation_errors = _validate_result(final_amount, calc_rule, employee_id)

        results.append({
            "employee_id": employee_id,
            "plan_id": rule.get("plan_id", ""),
            "plan_name": rule.get("plan_name", ""),
            "eligible": True,
            "amount": final_amount,
            "raw_amount": raw_amount,
            "capped_amount": capped_amount,
            "detail": detail,
            "validation_errors": validation_errors,
        })

    return results


# ---------------------------------------------------------------------------
# Batch Calculation (multiple plans)
# ---------------------------------------------------------------------------

def skill_calculate_batch(rules: list[dict], data: list[dict]) -> dict:
    """Calculate incentive amounts for multiple plan rules against the same employee data.

    Args:
        rules: List of structured rule JSONs (from rule_phrase)
        data:  List of employee data dicts

    Returns:
        Dict keyed by plan_id, each value is the result list from skill_calculate

    Note: 当 plan.employee_scope.specific_employees 非空时，仅对该列表中的 employee_id
    进行计算，避免 owner-aggregate 类方案被错误展开到所有 owner 上。
    """
    all_results = {}
    for rule in rules:
        plan_id = rule.get("plan_id", "UNKNOWN")
        specific = rule.get("employee_scope", {}).get("specific_employees", []) or []
        if specific:
            spec_set = {str(s).strip() for s in specific}
            scoped_data = [
                e for e in data
                if str(e.get("employee_id", "")).strip() in spec_set
            ]
        else:
            scoped_data = data
        all_results[plan_id] = skill_calculate(rule, scoped_data)
    return all_results


# ---------------------------------------------------------------------------
# CLI Entry Point (for standalone testing)
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse
    import json

    parser = argparse.ArgumentParser(description="skill_calculate: compute incentive amounts")
    parser.add_argument("rule_file", help="Path to structured rule JSON file")
    parser.add_argument("data_file", help="Path to employee data JSON file")
    args = parser.parse_args()

    rule = json.loads(open(args.rule_file, encoding="utf-8").read())
    data = json.loads(open(args.data_file, encoding="utf-8").read())

    if isinstance(rule, list):
        result = skill_calculate_batch(rule, data)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        result = skill_calculate(rule, data)
        print(json.dumps(result, ensure_ascii=False, indent=2))