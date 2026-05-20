"""incentive_bot — CLI 编排器，串联 rule_phrase → 数据加载 → skill_calculate → 导出。

使用方式：
    python incentive_bot.py run --month 2026-04 --plan 服务 [--force]
    python incentive_bot.py parse --month 2026-04 --plan 服务 [--force]
    python incentive_bot.py calc --month 2026-04 --plan 服务
    python incentive_bot.py list-plans
    python incentive_bot.py list-rules
"""

import argparse
import io
import json
import logging
import sys
from pathlib import Path

# Windows 终端 UTF-8 支持
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

from data_loader import (
    LocalDataSourceLoader,
    collect_report_names,
    build_employee_data,
    load_column_mapping,
)
from employee_directory import load_employee_directory
from incentive_calculate import skill_calculate, skill_calculate_batch
from plan_reader import PlanDocumentReader
from result_exporter import ResultExporter
from report_registry import ReportRegistry
from rule_phrase import skill_phrase, create_client
from rule_verifier import RuleVerifier

SKILLS_DIR = Path(__file__).parent
PLAN_DIR = SKILLS_DIR / "方案"
DATA_DIR = SKILLS_DIR / "数据源"
OUTPUT_DIR = SKILLS_DIR / "output"

logger = logging.getLogger("incentive_bot")
COLUMN_MAPPING = load_column_mapping()


# ---------------------------------------------------------------------------
# File Resolution
# ---------------------------------------------------------------------------

def resolve_plan_file(month: str, plan_type: str) -> Path:
    """根据月份+方案类型定位方案文件。

    "2026-04" + "服务" → 方案/26年4月服务激励.xlsx
    """
    year, mm = month.split("-")
    mm_int = int(mm)  # 兼容 "4月" 不带前导零的写法
    month_keywords = [
        f"{mm}月", f"{mm_int}月",
        f"{year[-2:]}年{mm}月", f"{year[-2:]}年{mm_int}月",
        f"{year}年{mm}月", f"{year}年{mm_int}月",
    ]

    candidates = []
    for f in PLAN_DIR.glob("*.xlsx"):
        if f.name.startswith("~$"):
            continue  # 跳过 Excel 打开时的临时锁文件
        name = f.stem
        if any(kw in name for kw in month_keywords) and plan_type in name:
            candidates.append(f)

    if len(candidates) == 1:
        logger.info(f"方案文件: {candidates[0].name}")
        return candidates[0]
    if len(candidates) > 1:
        logger.warning(f"匹配到多个方案文件: {[c.name for c in candidates]}，取第一个")
        return candidates[0]
    available = [f.name for f in PLAN_DIR.glob("*.xlsx")]
    raise FileNotFoundError(
        f"找不到 {month} {plan_type} 方案文件。可用文件: {available}"
    )


def resolve_rule_cache(month: str, plan_type: str) -> Path:
    """返回缓存规则文件路径: rule_{YYYYMM}_{type}.json"""
    ym = month.replace("-", "")
    return SKILLS_DIR / f"rule_{ym}_{plan_type}.json"


# ---------------------------------------------------------------------------
# Rule Loading / Parsing
# ---------------------------------------------------------------------------

def load_or_parse_rule(
    month: str,
    plan_type: str,
    force: bool,
    provider: str | None = None,
    model: str | None = None,
) -> dict | list:
    """加载缓存规则或解析方案文档。

    1. 检查 rule_{YYYYMM}_{type}.json
    2. 存在 & 非 force → 加载并返回
    3. 不存在或 force → 读方案 Excel → skill_phrase → 保存缓存
    """
    cache_path = resolve_rule_cache(month, plan_type)

    if cache_path.exists() and not force:
        logger.info(f"加载缓存规则: {cache_path.name}")
        return json.loads(cache_path.read_text(encoding="utf-8"))

    # 需要调用 LLM 解析
    logger.info(f"缓存不存在或 --force，开始解析方案文档...")
    plan_path = resolve_plan_file(month, plan_type)
    reader = PlanDocumentReader()
    plan_text = reader.read_plan_as_text(plan_path)

    client = create_client(provider=provider, model=model)
    rule = skill_phrase(llm_client=client, plan_text=plan_text)

    # 保存缓存
    cache_path.write_text(
        json.dumps(rule, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    logger.info(f"规则已缓存: {cache_path.name}")
    return rule


# ---------------------------------------------------------------------------
# Pipeline Steps
# ---------------------------------------------------------------------------

def run_parse(
    month: str, plan_type: str, force: bool,
    provider: str | None, model: str | None,
    verify: bool = False,
) -> dict | list:
    """仅解析步骤：方案 → rule JSON

    verify=True 时，解析后立即执行验证（数值校验 + 映射检查 + 试算）
    """
    rule = load_or_parse_rule(month, plan_type, force, provider, model)
    logger.info("解析完成")

    if verify:
        _run_verification(rule, month, plan_type)

    return rule


def _run_verification(rule: dict | list, month: str, plan_type: str) -> None:
    """对解析后的规则执行验证：数值、映射、试算"""
    # 尝试加载原方案文档（用于数值验证），失败则跳过该层
    plan_text = None
    try:
        plan_path = resolve_plan_file(month, plan_type)
        reader = PlanDocumentReader()
        plan_text = reader.read_plan_as_text(plan_path)
    except FileNotFoundError as e:
        logger.warning(f"无法定位方案文档，跳过数值验证: {e}")

    # 加载员工数据用于试算
    employee_data = []
    try:
        report_names = collect_report_names(rule)
        registry = ReportRegistry(data_dir=DATA_DIR)
        loader = LocalDataSourceLoader(DATA_DIR, registry, COLUMN_MAPPING)
        raw_data = loader.load_reports(report_names, month)
        employee_data = build_employee_data(raw_data, rule, COLUMN_MAPPING)
    except Exception as e:
        logger.warning(f"无法加载员工数据用于试算: {e}")

    verifier = RuleVerifier(DATA_DIR, COLUMN_MAPPING)
    report = verifier.verify(
        rule=rule,
        plan_text=plan_text,
        employee_data=employee_data[:1] if employee_data else None,
    )
    report.print_summary()


def run_calc(
    month: str, plan_type: str, output_dir: Path,
) -> dict[str, list[dict]] | list[dict]:
    """仅计算步骤：加载缓存 rule + 报表数据 → 计算 → 导出"""
    cache_path = resolve_rule_cache(month, plan_type)
    if not cache_path.exists():
        raise ValueError(
            f"缓存规则文件不存在: {cache_path.name}。请先运行 parse 命令生成规则缓存。"
        )
    rule = json.loads(cache_path.read_text(encoding="utf-8"))

    # 收集报表名并加载数据
    report_names = collect_report_names(rule)
    registry = ReportRegistry(data_dir=DATA_DIR)
    loader = LocalDataSourceLoader(DATA_DIR, registry, COLUMN_MAPPING)
    raw_data = loader.load_reports(report_names, month)

    # 列名映射 + 合并
    employee_data = build_employee_data(raw_data, rule, COLUMN_MAPPING)
    logger.info(f"员工数据: {len(employee_data)} 名")

    # 计算
    if isinstance(rule, list):
        results = skill_calculate_batch(rule, employee_data)
    else:
        results = skill_calculate(rule, employee_data)

    # 导出
    directory = load_employee_directory(DATA_DIR)
    exporter = ResultExporter(output_dir, employee_directory=directory)
    paths = exporter.export(results, month, plan_type, rule=rule)
    logger.info(f"结果已导出: {', '.join(p.name for p in paths.values())}")
    return results


def run_pipeline(
    month: str, plan_type: str, force: bool,
    provider: str | None, model: str | None,
    output_dir: Path,
    verify: bool = False,
) -> dict[str, list[dict]] | list[dict]:
    """全流程: parse → calc → export

    verify=True 时，解析后先做验证再继续计算
    """
    # Step 1: 解析（如已有缓存则跳过）
    rule = load_or_parse_rule(month, plan_type, force, provider, model)

    # Step 2: 加载报表数据
    report_names = collect_report_names(rule)
    registry = ReportRegistry(data_dir=DATA_DIR)
    loader = LocalDataSourceLoader(DATA_DIR, registry, COLUMN_MAPPING)
    raw_data = loader.load_reports(report_names, month)

    # Step 3: 列名映射 + 合并
    employee_data = build_employee_data(raw_data, rule, COLUMN_MAPPING)
    logger.info(f"员工数据: {len(employee_data)} 名")

    # Step 3.5: 验证（可选）
    if verify:
        plan_path = resolve_plan_file(month, plan_type)
        plan_reader = PlanDocumentReader()
        plan_text = plan_reader.read_plan_as_text(plan_path)
        verifier = RuleVerifier(DATA_DIR, COLUMN_MAPPING)
        report = verifier.verify(
            rule=rule,
            plan_text=plan_text,
            employee_data=employee_data[:1] if employee_data else None,
        )
        report.print_summary()
        if report.has_errors:
            logger.warning("验证存在错误，但继续执行计算（如需中断请手动取消）")

    # Step 4: 计算
    if isinstance(rule, list):
        plans_count = len(rule)
        results = skill_calculate_batch(rule, employee_data)
    else:
        plans_count = 1
        results = skill_calculate(rule, employee_data)

    logger.info(f"计算完成: {plans_count} 个方案 × {len(employee_data)} 名员工")

    # Step 5: 导出
    directory = load_employee_directory(DATA_DIR)
    exporter = ResultExporter(output_dir, employee_directory=directory)
    paths = exporter.export(results, month, plan_type, rule=rule)
    logger.info(f"结果已导出: {', '.join(p.name for p in paths.values())}")
    return results


# ---------------------------------------------------------------------------
# List Commands
# ---------------------------------------------------------------------------

def list_plans():
    """列出 方案/ 目录中的所有方案文件"""
    files = list(PLAN_DIR.glob("*.xlsx"))
    if not files:
        print("方案目录为空")
        return
    print(f"可用方案文件（{len(files)} 个）:")
    for f in sorted(files):
        print(f"  - {f.name}")


def list_rules():
    """列出所有缓存规则文件"""
    # 匹配缓存规则文件：rule_YYYYMM_类型.json，排除 rule_examples/rule_schema
    import re
    pattern = re.compile(r"^rule_\d{6}_.*\.json$")
    files = [f for f in SKILLS_DIR.glob("rule_*.json") if pattern.match(f.name)]
    if not files:
        print("无缓存规则文件")
        return
    print(f"缓存规则文件（{len(files)} 个）:")
    for f in sorted(files):
        rule = json.loads(f.read_text(encoding="utf-8"))
        if isinstance(rule, list):
            plan_ids = [p.get("plan_id", "?") for p in rule]
            print(f"  - {f.name}: {len(rule)} 个方案 ({', '.join(plan_ids[:3])}{'...' if len(plan_ids) > 3 else ''})")
        else:
            print(f"  - {f.name}: {rule.get('plan_id', '?')}")


# ---------------------------------------------------------------------------
# CLI Entry Point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="incentive_bot: 每月激励方案自动计算编排器"
    )
    parser.add_argument("--verbose", action="store_true", help="启用 DEBUG 日志")
    subparsers = parser.add_subparsers(dest="command", required=True)

    # run
    run_p = subparsers.add_parser("run", help="全流程: 解析 → 加载 → 计算 → 导出")
    run_p.add_argument("--month", required=True, help="YYYY-MM")
    run_p.add_argument("--plan", required=True, help="方案类型关键词")
    run_p.add_argument("--force", action="store_true", help="重新解析（忽略缓存）")
    run_p.add_argument("--output-dir", default=None, help="输出目录（默认 output/）")
    run_p.add_argument("--provider", default=None, help="LLM provider 覆盖")
    run_p.add_argument("--model", default=None, help="LLM model 覆盖")
    run_p.add_argument("--verify", action="store_true", help="计算前执行验证（数值/映射/试算）")

    # parse
    parse_p = subparsers.add_parser("parse", help="仅解析方案 → rule JSON")
    parse_p.add_argument("--month", required=True, help="YYYY-MM")
    parse_p.add_argument("--plan", required=True, help="方案类型关键词")
    parse_p.add_argument("--force", action="store_true", help="重新解析")
    parse_p.add_argument("--provider", default=None, help="LLM provider 覆盖")
    parse_p.add_argument("--model", default=None, help="LLM model 覆盖")
    parse_p.add_argument("--verify", action="store_true", help="解析后执行验证（数值/映射/试算）")

    # calc
    calc_p = subparsers.add_parser("calc", help="仅计算（需已有缓存规则）")
    calc_p.add_argument("--month", required=True, help="YYYY-MM")
    calc_p.add_argument("--plan", required=True, help="方案类型关键词")
    calc_p.add_argument("--output-dir", default=None, help="输出目录")

    # verify
    verify_p = subparsers.add_parser("verify", help="对已缓存规则执行验证（不重新解析、不计算导出）")
    verify_p.add_argument("--month", required=True, help="YYYY-MM")
    verify_p.add_argument("--plan", required=True, help="方案类型关键词")

    # list
    subparsers.add_parser("list-plans", help="列出可用方案文件")
    subparsers.add_parser("list-rules", help="列出缓存规则文件")

    args = parser.parse_args()

    # 配置日志
    level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=level, format="[%(levelname)s] %(message)s")

    # 解析命令
    try:
        if args.command == "run":
            output_dir = Path(args.output_dir) if args.output_dir else OUTPUT_DIR
            run_pipeline(args.month, args.plan, args.force, args.provider, args.model, output_dir, verify=args.verify)

        elif args.command == "parse":
            run_parse(args.month, args.plan, args.force, args.provider, args.model, verify=args.verify)

        elif args.command == "calc":
            output_dir = Path(args.output_dir) if args.output_dir else OUTPUT_DIR
            run_calc(args.month, args.plan, output_dir)

        elif args.command == "verify":
            cache_path = resolve_rule_cache(args.month, args.plan)
            if not cache_path.exists():
                logger.error(f"缓存规则文件不存在: {cache_path.name}。请先运行 parse 命令。")
                sys.exit(1)
            rule = json.loads(cache_path.read_text(encoding="utf-8"))
            _run_verification(rule, args.month, args.plan)

        elif args.command == "list-plans":
            list_plans()

        elif args.command == "list-rules":
            list_rules()

    except FileNotFoundError as e:
        logger.error(str(e))
        sys.exit(1)
    except ValueError as e:
        logger.error(str(e))
        sys.exit(1)
    except KeyError as e:
        logger.error(str(e))
        sys.exit(1)
    except NotImplementedError as e:
        logger.error(str(e))
        sys.exit(1)
    except Exception as e:
        logger.error(f"未预期错误: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()