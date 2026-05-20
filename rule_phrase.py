"""rule_phrase — 理解每月激励方案文档，输出符合 rule_schema.json 的结构化 JSON。

职责：
1. 接收方案文档（纯文本或文件路径）
2. 调用 LLM 将自然语言方案解析为结构化规则
3. 校验输出：schema 合规 + 业务逻辑检查
4. 返回可供 skill_calculate 直接消费的 JSON

使用方式：
    result = skill_phrase(plan_text="...", llm_client=client)
    result = skill_phrase(plan_file="方案.txt", llm_client=client)
"""

import json
import os
import re
from abc import ABC, abstractmethod
from pathlib import Path

try:
    import jsonschema
except ImportError:
    jsonschema = None

# ---------------------------------------------------------------------------
# LLM Client Abstraction
# ---------------------------------------------------------------------------

class LLMClient(ABC):
    @abstractmethod
    def call(self, prompt: str, system: str = "") -> str:
        """Send prompt to LLM, return text response."""
        ...

    @abstractmethod
    def call_json(self, prompt: str, system: str = "") -> str:
        """Send prompt, request JSON-only response."""
        ...


class ClaudeClient(LLMClient):
    def __init__(self, model: str, api_key: str | None = None, base_url: str | None = None, max_tokens: int = 16384):
        try:
            import anthropic
        except ImportError:
            raise ImportError("pip install anthropic")
        key = api_key or os.getenv("API_KEY")
        if not key:
            raise ValueError("API_KEY not configured — set it in .env")
        url = base_url or os.getenv("BASE_URL")
        self._model = model
        self._max_tokens = max_tokens
        client_kwargs = {"api_key": key}
        if url:
            client_kwargs["base_url"] = url
        self._client = anthropic.Anthropic(**client_kwargs)

    def call(self, prompt: str, system: str = "") -> str:
        msg = self._client.messages.create(
            model=self._model, max_tokens=self._max_tokens, system=system,
            messages=[{"role": "user", "content": prompt}],
        )
        return msg.content[0].text

    def call_json(self, prompt: str, system: str = "") -> str:
        prompt += "\n\n请严格以JSON格式输出，不要包含任何其他文字或markdown标记。"
        return self.call(prompt, system)


class GPTClient(LLMClient):
    """OpenAI-compatible client. Works with OpenAI, DeepSeek, MiniMax, Kimi, and any
    service that exposes an OpenAI-compatible API endpoint."""

    def __init__(self, model: str, api_key: str | None = None, base_url: str | None = None, max_tokens: int = 16384):
        try:
            import openai
        except ImportError:
            raise ImportError("pip install openai")
        key = api_key or os.getenv("API_KEY")
        if not key:
            raise ValueError("API_KEY not configured — set it in .env")
        url = base_url or os.getenv("BASE_URL")
        self._model = model
        self._max_tokens = max_tokens
        client_kwargs = {"api_key": key}
        if url:
            client_kwargs["base_url"] = url
        self._client = openai.OpenAI(**client_kwargs)

    def call(self, prompt: str, system: str = "") -> str:
        messages = []
        if system:
            messages.append({"role": "system", "content": system})
        messages.append({"role": "user", "content": prompt})
        resp = self._client.chat.completions.create(
            model=self._model, messages=messages, max_tokens=self._max_tokens,
        )
        return resp.choices[0].message.content

    def call_json(self, prompt: str, system: str = "") -> str:
        prompt += "\n\nPlease output strictly in JSON format with no other text or markdown."
        return self.call(prompt, system)


def create_client(provider: str | None = None, model: str | None = None) -> LLMClient:
    """Factory: create LLMClient based on provider config.

    LLM_MODEL is required — must be set via env var or parameter.
    LLM_PROVIDER defaults to env LLM_PROVIDER, then "claude".
    Supported providers:
      - claude        → Anthropic Claude, model must be specified (e.g. claude-sonnet-4-6)
      - openai        → OpenAI GPT, model must be specified (e.g. gpt-4o)
      - deepseek      → DeepSeek (OpenAI-compatible), model must be specified (e.g. deepseek-chat)
      - minimax       → MiniMax (OpenAI-compatible), model must be specified (e.g. abab6.5s-chat)
      - kimi          → Kimi/Moonshot (OpenAI-compatible), model must be specified (e.g. moonshot-v1-8k)
    Any OpenAI-compatible API can be used via the "openai" provider with custom base_url and model.
    """
    provider = provider or os.getenv("LLM_PROVIDER")
    if not provider:
        raise ValueError("LLM_PROVIDER is required — set it in .env (e.g. LLM_PROVIDER=claude) or pass provider parameter")
    model = model or os.getenv("LLM_MODEL")
    if not model:
        raise ValueError("LLM_MODEL is required — set it in .env (LLM_MODEL=your-model-name) or pass model parameter")

    # Provider shortcuts: pre-configured base_url for common Chinese models
    # All providers use unified API_KEY and BASE_URL from .env
    # For deepseek/minimax/kimi, base_url is auto-set unless BASE_URL is explicitly configured
    provider_configs = {
        "claude":    {"client_cls": ClaudeClient},
        "openai":    {"client_cls": GPTClient},
        "deepseek":  {"client_cls": GPTClient,
                       "preset_base_url": "https://api.deepseek.com/v1"},
        "minimax":   {"client_cls": GPTClient,
                       "preset_base_url": "https://api.minimax.chat/v1"},
        "kimi":      {"client_cls": GPTClient,
                       "preset_base_url": "https://api.moonshot.cn/v1"},
    }

    if provider not in provider_configs:
        raise ValueError(f"Unknown provider: {provider}. Supported: {', '.join(provider_configs.keys())}")

    config = provider_configs[provider]

    if config["client_cls"] == ClaudeClient:
        return ClaudeClient(model=model)
    else:
        # OpenAI-compatible clients: use preset_base_url if BASE_URL env var is not set
        extra_kwargs = {}
        preset_url = config.get("preset_base_url")
        env_url = os.getenv("BASE_URL")
        if env_url:
            extra_kwargs["base_url"] = env_url
        elif preset_url:
            extra_kwargs["base_url"] = preset_url
        return GPTClient(model=model, **extra_kwargs)


# ---------------------------------------------------------------------------
# File & Resource Loading
# ---------------------------------------------------------------------------

SKILLS_DIR = Path(__file__).parent


def _load_schema() -> dict:
    path = SKILLS_DIR / "rule_schema.json"
    return json.loads(path.read_text(encoding="utf-8"))


def _load_examples() -> dict:
    path = SKILLS_DIR / "rule_examples.json"
    return json.loads(path.read_text(encoding="utf-8"))


def _read_plan_file(path: str) -> str:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Plan file not found: {path}")
    return p.read_text(encoding="utf-8")


# ---------------------------------------------------------------------------
# JSON Extraction from LLM Response
# ---------------------------------------------------------------------------

def _extract_json(raw: str) -> dict | list:
    """Extract JSON from LLM response. Handles markdown fences and truncated output.
    Returns a dict (single plan) or list of dicts (multi-direction plan)."""
    # Strip markdown code fences
    cleaned = re.sub(r"^```(?:json)?\s*", "", raw)
    cleaned = re.sub(r"\s*```$", "", cleaned.strip())

    # First try: parse directly
    try:
        parsed = json.loads(cleaned)
        if isinstance(parsed, (dict, list)):
            return parsed
        raise ValueError(f"Expected dict or list, got {type(parsed).__name__}")
    except json.JSONDecodeError:
        pass

    # Second try: fix control characters inside string literals (LLM sometimes
    # emits raw newlines/tabs inside JSON strings instead of escaping them)
    fixed = _escape_control_chars_in_strings(cleaned)
    try:
        parsed = json.loads(fixed)
        if isinstance(parsed, (dict, list)):
            return parsed
    except json.JSONDecodeError:
        pass

    # Third try: repair truncated JSON by closing open brackets
    repaired = _repair_truncated_json(fixed)
    try:
        parsed = json.loads(repaired)
        if isinstance(parsed, (dict, list)):
            return parsed
        raise ValueError(f"Expected dict or list, got {type(parsed).__name__}")
    except json.JSONDecodeError as e:
        raise ValueError(f"LLM response is not valid JSON (even after repair): {e}\nRaw:\n{raw[:500]}")


def _escape_control_chars_in_strings(text: str) -> str:
    """Walk through JSON text and escape raw \\n / \\r / \\t inside string literals."""
    result = []
    in_string = False
    escape = False
    for ch in text:
        if escape:
            result.append(ch)
            escape = False
            continue
        if ch == "\\" and in_string:
            result.append(ch)
            escape = True
            continue
        if ch == '"':
            in_string = not in_string
            result.append(ch)
            continue
        if in_string and ch in ("\n", "\r", "\t"):
            result.append({"\n": "\\n", "\r": "\\r", "\t": "\\t"}[ch])
            continue
        result.append(ch)
    return "".join(result)


def _repair_truncated_json(text: str) -> str:
    """Close truncated JSON: find the last complete object, then close remaining brackets."""
    # Find the last complete } and truncate incomplete content after it
    last_close_pos = -1
    for i in range(len(text) - 1, -1, -1):
        if text[i] == "}":
            last_close_pos = i
            break

    if last_close_pos > 0:
        truncated = text[:last_close_pos + 1]
    else:
        truncated = text

    # Count remaining open brackets and close them in reverse order
    remaining = []
    for ch in truncated:
        if ch in ("{", "["):
            remaining.append(ch)
        elif ch == "}" and remaining and remaining[-1] == "{":
            remaining.pop()
        elif ch == "]" and remaining and remaining[-1] == "[":
            remaining.pop()

    closers = "".join("}" if b == "{" else "]" for b in reversed(remaining))
    return truncated + closers


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

def _validate_schema(data: dict | list, schema: dict) -> list[str]:
    """Validate data against JSON Schema. Return list of error messages.
    If data is a list, validate each item separately."""
    if jsonschema is None:
        return ["jsonschema package not installed — skipping schema validation (pip install jsonschema)"]

    items = data if isinstance(data, list) else [data]
    errors = []
    for idx, item in enumerate(items):
        prefix = f"[{idx}]" if isinstance(data, list) else ""
        validator = jsonschema.Draft7Validator(schema)
        for err in sorted(validator.iter_errors(item), key=lambda e: list(e.path)):
            path = ".".join(str(p) for p in err.path) if err.path else "(root)"
            errors.append(f"{prefix}{path}: {err.message}")
    return errors


def _validate_business_logic(data: dict | list) -> list[str]:
    """Check business rules that schema can't enforce.
    If data is a list, validate each item separately."""
    items = data if isinstance(data, list) else [data]
    all_errors = []
    for idx, item in enumerate(items):
        prefix = f"[{idx}]" if isinstance(data, list) else ""
        errors = _validate_single_plan(item, prefix)
        all_errors.extend(errors)
    return all_errors


def _validate_single_plan(data: dict, prefix: str = "") -> list[str]:
    """Validate business rules for a single plan object. prefix marks list index."""
    errors = []
    rule = data.get("calculation_rule", {})
    method = rule.get("method")

    # --- ratio_based ---
    if method == "ratio_based":
        for item in rule.get("items", []):
            if item.get("base_amount", 0) <= 0:
                errors.append(f"{prefix}item '{item.get('name')}': base_amount must be > 0")
            if item.get("target_value", 0) <= 0:
                errors.append(f"{prefix}item '{item.get('name')}': target_value must be > 0")

    # --- tiered_commission ---
    elif method == "tiered_commission":
        tiers = rule.get("tiers", [])
        if not tiers:
            errors.append(f"{prefix}tiers must have at least one entry")
        for i, tier in enumerate(tiers):
            if tier.get("rate", 0) < 0:
                errors.append(f"{prefix}tier[{i}]: rate must be >= 0")
        for i in range(1, len(tiers)):
            prev_max = tiers[i - 1].get("max")
            cur_min = tiers[i].get("min", 0)
            if prev_max is not None and prev_max != cur_min:
                errors.append(f"{prefix}tier gap: tier[{i-1}].max={prev_max} != tier[{i}].min={cur_min}")
        if tiers and tiers[0].get("min", 0) != 0:
            errors.append(f"{prefix}first tier min must be 0")

    # --- fixed_bonus ---
    elif method == "fixed_bonus":
        if rule.get("amount", 0) <= 0:
            errors.append(f"{prefix}amount must be > 0")

    # --- per_unit ---
    elif method == "per_unit":
        if rule.get("unit_price", 0) <= 0:
            errors.append(f"{prefix}unit_price must be > 0")
        capping = rule.get("capping", {})
        if capping.get("max_quantity") and capping["max_quantity"] <= 0:
            errors.append(f"{prefix}capping.max_quantity must be > 0")

    # --- rank_tiered ---
    elif method == "rank_tiered":
        rewards = rule.get("rank_rewards", [])
        if not rewards:
            errors.append(f"{prefix}rank_rewards must have at least one entry")
        # ranks 必须从 1 起、连续递增
        ranks = [r.get("rank", 0) for r in rewards]
        if ranks and ranks != list(range(1, len(ranks) + 1)):
            errors.append(f"{prefix}rank_rewards: ranks must start at 1 and be consecutive (got {ranks})")
        for r in rewards:
            if r.get("amount", 0) <= 0:
                errors.append(f"{prefix}rank_rewards rank={r.get('rank')}: amount must be > 0")
        if not rule.get("rank_field"):
            errors.append(f"{prefix}rank_field is required")
        if not rule.get("qualifier"):
            errors.append(f"{prefix}qualifier is required")

    # --- Common: capping ---
    capping = rule.get("capping", {})
    if capping.get("max_amount") and capping["max_amount"] <= 0:
        errors.append(f"{prefix}capping.max_amount must be > 0")

    # --- Common: data_source_mapping must include employee_id ---
    mapping = data.get("data_source_mapping", {})
    if "employee_id" not in mapping:
        errors.append(f"{prefix}data_source_mapping must include 'employee_id'")

    # --- Common: all fields referenced in rule must exist in mapping ---
    referenced_fields = _collect_referenced_fields(data)
    for field in referenced_fields:
        if field not in mapping:
            errors.append(f"{prefix}field '{field}' referenced in rule but missing from data_source_mapping")

    return errors


def _collect_referenced_fields(data: dict) -> set[str]:
    """Collect all field names referenced in calculation_rule, prerequisites, etc."""
    fields = set()
    rule = data.get("calculation_rule", {})
    method = rule.get("method")

    if method == "ratio_based":
        for item in rule.get("items", []):
            fields.add(item.get("target_field"))
    elif method == "tiered_commission":
        fields.add(rule.get("performance_field"))
        for ded in rule.get("deductions", []):
            fields.add(ded.get("field"))
    elif method == "fixed_bonus":
        cond = rule.get("condition", {})
        fields.add(cond.get("field"))
    elif method == "per_unit":
        fields.add(rule.get("quantity_field"))
    elif method == "rank_tiered":
        fields.add(rule.get("rank_field"))
        if rule.get("tiebreaker_field"):
            fields.add(rule.get("tiebreaker_field"))
        qualifier = rule.get("qualifier", {})
        if qualifier.get("field"):
            fields.add(qualifier["field"])

    for pre in data.get("prerequisites", []):
        fields.add(pre.get("field"))

    fields.discard(None)
    return fields


# ---------------------------------------------------------------------------
# Prompt Construction
# ---------------------------------------------------------------------------

def _build_report_list() -> str:
    """从 report_registry 动态生成可用报表列表，新增报表只需改 registry 一处。"""
    from report_registry import DEFAULT_REGISTRY
    lines = []
    for name in DEFAULT_REGISTRY:
        lines.append(f"- {name}\n")
    return "".join(lines)


def _build_system_prompt(schema: dict, examples: dict) -> str:
    examples_text = json.dumps(examples, ensure_ascii=False, indent=2)
    return (
        "你是一个激励方案解析专家。你的任务是读取每月激励方案文档，"
        "将其转换为严格符合 schema 的结构化 JSON。\n\n"
        "关键规则：\n"
        "1. 岗位角色只能使用标准编码：LP（班主任/服务老师）、CC（销售/课程顾问）\n"
        "2. 取整规则默认为 {\"direction\": \"up\", \"precision\": \"2\"}（保留两位小数向上取整），"
        "除非方案文档明确指定例外\n"
        "3. data_source_mapping 中每个字段必须指定来源报表名和列名\n"
        "4. data_source_mapping 必须包含 employee_id\n"
        "5. 所有 calculation_rule 中引用的字段名必须在 data_source_mapping 中有对应条目\n"
        "6. condition 和 prerequisites 必须使用结构化格式 {field, operator, value}\n"
        "7. 不要输出任何 _ 开头的注释字段\n"
        "8. 如果方案文档包含多个并列的激励方向，输出一个 JSON 数组（list），"
        "每个方向是一个独立的 plan 对象；如果只有单个激励方向，输出一个 plan 对象（dict）\n"
        "9. effective_month 必须为 YYYY-MM 格式。如果方案文档正文中明确标注了月份，"
        "使用该月份；如果没有标注，从文件名或上下文推断（如'4月'→'2026-04'，"
        "'26年4月'→'2026-04'）。当前年份为2026年，月份以方案所属年份为准\n"
        "10. employee_scope.specific_employees：当方案文档明确标出具名个人（如'激励对象=李文韬'、"
        "'XXX（服务 owner）'、'负责人 XXX'、'TL【刘洋】'、'TL【向梦清】'、"
        "'李文韬特殊申请参与个人激励'）时，填入 specific_employees 数组。"
        "role 仍应保留（李文韬是 LP 就填 ['LP']，刘洋是 TL 就填 ['LP']——TL 也是 LP 岗位）。\n"
        "11. 方案表格中'激励对象'/'激励人群'列若留空，继承上一行的值，直到遇到新值为止\n"
        "12. data_scope（plan 级默认）判定：\n"
        "    - 指标描述涉及'整体'、'海外大区'、'XX 团队整体'、'XX组整体'、'大盘'、'部门整体' → mode=aggregate，"
        "aggregate_key 填相应关键字\n"
        "    - aggregate_key 必须与方案文字里的小组/大区名一致，常见值："
        "'海外团队'、'非港澳台团队'、'港澳台'、'港澳'、'台湾'、'欧美澳'、"
        "'美澳1组'、'美澳2组'、'美澳3组'、'美澳4组'、'港澳1组'、'港澳2组'、'台湾组'、'广州'\n"
        "    - owner 激励（具名个人 + 整体/小组指标）一律 mode=aggregate，"
        "aggregate_key 取该 owner 负责的小组或大区名\n"
        "    - 默认 mode=individual\n"
        "13. data_scope_override（item 级覆盖）：\n"
        "    - 当某个子项明确写'个人'/'该员工本人'/'个人业绩' → data_scope_override={mode:'individual'}\n"
        "    - 不写 override 的子项跟 plan 默认\n"
        "14. 多个并列子方案独立决定 data_scope 和 specific_employees，输出为 plan list\n"
        "15. '瓜分即止'/'瓜分池'/'发完即止'类奖金池：method=per_unit，unit_price=单价（如每单30元），"
        "capping.max_amount=池子总金额（如3000）；scope.mode=aggregate（数据来自团队订单数）；"
        "若文档指定具名 owner 参与（如港澳台李文韬），specific_employees 填 owner 名单，"
        "aggregate_key 取相应大区/小组（如'港澳台'）。在 plan_name 中保留'瓜分'字样。\n"
        "16. TL/小组之间的排名激励（如'TL 续费率前2名'、'M0-LP小组排名前三'）：\n"
        "    - method=rank_tiered，qualifier 填门槛条件，rank_field 填排名指标\n"
        "    - data_scope.mode 保持 individual（按现有员工/小组数据排序）\n"
        "    - specific_employees 留空（[]）——表示『全体参赛』，由数据源中所有同等级行参与排名\n"
        "    - 即使竞赛主体是'小组'而非个人，也用此结构；若数据缺失导致无法排序，下游会报告\n"
        "17. 大盘性'第一个达标小组获得X元'：method=rank_tiered，qualifier=GMV达成率100%，"
        "rank_field=订单完成时间或GMV达成率（取首个达标，rank_rewards 仅 rank=1），"
        "并在 plan_name 注明'达成时间排序'\n"
        "\n"
        f"【可用报表列表（data_source_mapping.report 只能使用以下名称）】\n"
        f"{_build_report_list()}"
        "注意：不要虚构报表名称，只能使用上述列表中的名称。\n"
        "\n"
        "【calculation_rule.method 选择决策树（非常重要，决定计算正确性）】\n"
        "判断顺序：先看是否有'排名'，再看是否'按比例'，再看是否'按件计酬'，最后才用 fixed_bonus。\n"
        "\n"
        "A. rank_tiered（排名档位激励）—— 当原文出现'排名前N'、'第一名X元，第二名Y元'、'前2名'等表述时\n"
        "   特征：达成门槛 + 按排名给不同金额（一档一档下降）\n"
        "   示例原文：'达成激励门槛且续费率前2名：第一名500元，第二名400元'\n"
        "   必填：qualifier（门槛条件）、rank_field（排名字段）、rank_rewards（每档金额）\n"
        "   可选：tiebreaker_field（同分时的次级排序字段，如'排名重复则按X降序'）\n"
        "   不要用 fixed_bonus 模拟此场景，会丢失第二/第三名等档位信息！\n"
        "\n"
        "B. ratio_based（比例达标激励）—— 当原文出现'按达标率/比例'、'激励金额=基数×达成率/目标值'等表述时\n"
        "   特征：金额随实际值线性变化，typically 写成'X元×actual/target'\n"
        "   不要在'达标即给固定金额'场景下用此方法，会导致超额或不足！\n"
        "\n"
        "C. tiered_commission（阶梯提成）—— 当原文出现'X万以下提成A%，X-Y万提成B%'等阶梯式提成时\n"
        "\n"
        "D. per_unit（按件计酬）—— 当原文出现'每单/每件X元'，金额=单价×数量\n"
        "   注意：'瓜分池/发完即止/按时间先后'类激励，可以近似为 per_unit + max_amount，"
        "   但需要在 plan_name 中保留'瓜分'字样，提示用户此方法是近似实现\n"
        "\n"
        "E. fixed_bonus（固定奖金）—— 仅当原文是'达到门槛即获得X元'这种单一固定金额、无排名、无档位时使用\n"
        "   示例原文：'小组续费GMV达成率100%，且第一个达标小组获得1000元'（注意：'第一个达标'是排名"
        "前1名，应该用 rank_tiered，rank_field=达标时间或GMV达成率）\n"
        "   常见误用：把'第一名500/第二名400'拆成两个 fixed_bonus —— 错！应使用 rank_tiered\n"
        "\n"
        f"Schema 定义：\n{json.dumps(schema, ensure_ascii=False, indent=2)}\n\n"
        f"参考示例（照此结构填入参数值）：\n{examples_text}"
    )


def _build_user_prompt(plan_text: str) -> str:
    return (
        "请将以下激励方案文档解析为符合 schema 的结构化 JSON。"
        "只输出 JSON，不要输出任何解释文字。\n\n"
        f"--- 激励方案文档 ---\n{plan_text}\n--- 文档结束 ---"
    )


# ---------------------------------------------------------------------------
# Main Skill Function
# ---------------------------------------------------------------------------

def skill_phrase(
    llm_client: LLMClient,
    plan_text: str | None = None,
    plan_file: str | None = None,
    max_retries: int = 2,
) -> dict | list:
    """Parse a monthly incentive plan document into structured rule JSON.

    Args:
        llm_client:  LLM client instance (ClaudeClient / GPTClient)
        plan_text:   Plan document content as text
        plan_file:   Path to plan document file (txt/pdf etc.)
        max_retries: Max retry attempts when validation fails

    Returns:
        dict — single plan object (one incentive direction)
        list — list of plan objects (multiple incentive directions)

    Raises:
        ValueError:  If plan input is missing or validation fails after retries
    """
    if not plan_text and not plan_file:
        raise ValueError("Must provide either plan_text or plan_file")

    if plan_file and not plan_text:
        plan_text = _read_plan_file(plan_file)

    schema = _load_schema()
    examples = _load_examples()
    system_prompt = _build_system_prompt(schema, examples)
    user_prompt = _build_user_prompt(plan_text)

    last_errors = []
    for _ in range(max_retries):
        raw = llm_client.call_json(user_prompt, system_prompt)
        data = _extract_json(raw)

        # Validate
        schema_errors = _validate_schema(data, schema)
        biz_errors = _validate_business_logic(data)
        all_errors = schema_errors + biz_errors

        if not all_errors:
            return data

        last_errors = all_errors
        # Retry with error feedback
        user_prompt = _build_user_prompt(plan_text) + (
            f"\n\n上一轮输出存在以下问题，请修正后重新输出：\n"
            + "\n".join(f"- {e}" for e in all_errors)
        )

    raise ValueError(
        f"Validation failed after {max_retries} retries. Errors:\n"
        + "\n".join(f"- {e}" for e in last_errors)
    )


# ---------------------------------------------------------------------------
# CLI Entry Point (for standalone testing)
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="rule_phrase: parse incentive plan into structured JSON")
    parser.add_argument("plan_file", help="Path to plan document file")
    parser.add_argument("--provider", default=None, help="LLM provider: claude or openai")
    parser.add_argument("--model", default=None, help="LLM model name override")
    args = parser.parse_args()

    client = create_client(provider=args.provider, model=args.model)
    result = skill_phrase(llm_client=client, plan_file=args.plan_file)
    print(json.dumps(result, ensure_ascii=False, indent=2))