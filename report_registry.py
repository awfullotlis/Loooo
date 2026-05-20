"""report_registry — 报表抽象名 → 实际文件名映射。

data_source_mapping 中的报表名（如"新生跟进报表"）与 数据源/ 目录中的
实际文件名（如"益智海外新生首通监控 (29).xlsx"）不同。
Registry 通过前缀匹配桥接这个差异。

当 skill_download 到位后，Registry 不再需要——download 直接用抽象名。
"""

import json
import logging
import os
from pathlib import Path

logger = logging.getLogger("incentive_bot")

# 默认映射：抽象报表名 → 文件名前缀（不含版本号后缀）
DEFAULT_REGISTRY = {
    "新生跟进报表": "益智海外新生首通监控",
    "学管服务指标报表": "海外思维学管服务指标统计表",
    "服务池跟进质量报表": "海外思维服务SOP执行情况",
    "停课唤醒报表": "思维停课学员执行监控",
    "统合早鸟续费率报表": "统合早鸟续费率-分课包",
    "统合升舱续费率报表": "统合升舱续费率-分课包",
    "早鸟续费率报表": "早鸟续费率-（分课包）-默认一续",
    "升舱续费率报表": "升舱续费率-（分课包）-默认一续",
    "统合结课续费率报表": "统合结课续费率（近三月达成）_LP维度",
    "近三月结课续费率报表": "近三月结课续费率 --默认显示一续",
    "续费订单明细报表": "海外思维续费订单明细",
}


class ReportRegistry:
    """将抽象报表名解析为 数据源/ 目录中的实际文件名。

    匹配策略：前缀 glob + 版本号容忍
    - 在 data_dir 中查找以 prefix 开头、.xlsx 结尾的文件
    - 唯一匹配 → 直接用
    - 多个版本 → 取最新修改时间
    - 无匹配 → 抛 FileNotFoundError
    """

    def __init__(
        self,
        registry: dict[str, str] | None = None,
        data_dir: Path | None = None,
    ):
        self._registry = registry or DEFAULT_REGISTRY
        self._data_dir = data_dir

    def resolve(self, report_name: str, month: str) -> str:
        """解析抽象报表名为实际文件名。

        Args:
            report_name: data_source_mapping 中的报表名（如"新生跟进报表"）
            month: YYYY-MM 格式，预留月份特定文件名场景

        Returns:
            实际文件名（含 .xlsx 扩展名）

        Raises:
            KeyError: report_name 未在 registry 中注册
            FileNotFoundError: data_dir 中无匹配文件
        """
        prefix = self._registry.get(report_name)
        if prefix is None:
            available = list(self._registry.keys())
            raise KeyError(
                f"报表 '{report_name}' 未注册。已注册报表: {available}"
            )

        if self._data_dir is None:
            raise FileNotFoundError("未配置 data_dir，无法查找报表文件")

        # 前缀 glob 匹配
        matches = [
            f
            for f in os.listdir(self._data_dir)
            if f.startswith(prefix) and f.endswith(".xlsx")
        ]

        if len(matches) == 1:
            logger.info(f"报表 '{report_name}' → {matches[0]}")
            return matches[0]

        if len(matches) > 1:
            # 多版本：取最新修改时间
            latest = max(
                matches, key=lambda f: os.path.getmtime(self._data_dir / f)
            )
            logger.warning(
                f"报表 '{report_name}' 匹配到 {len(matches)} 个文件: {matches}，取最新: {latest}"
            )
            return latest

        # 无匹配
        all_files = [f for f in os.listdir(self._data_dir) if f.endswith(".xlsx")]
        raise FileNotFoundError(
            f"找不到报表 '{report_name}'，期望前缀 '{prefix}'。"
            f"数据源目录现有文件: {all_files}"
        )

    def resolve_path(self, report_name: str, month: str) -> Path:
        """解析报表名并返回完整文件路径。"""
        filename = self.resolve(report_name, month)
        return self._data_dir / filename

    def list_available(self) -> dict[str, str]:
        """列出所有已注册报表及其匹配结果（用于 list-rules 命令）。"""
        result = {}
        if self._data_dir is None:
            return result
        for name, prefix in self._registry.items():
            matches = [
                f
                for f in os.listdir(self._data_dir)
                if f.startswith(prefix) and f.endswith(".xlsx")
            ]
            result[name] = matches[0] if len(matches) == 1 else f"{len(matches)} 个匹配"
        return result


def load_registry_from_file(filepath: Path) -> dict[str, str]:
    """从 JSON 文件加载自定义 Registry 配置。

    JSON 格式：
    {
        "新生跟进报表": "益智海外新生首通监控",
        ...
    }
    """
    if not filepath.exists():
        raise FileNotFoundError(f"Registry 配置文件不存在: {filepath}")
    data = json.loads(filepath.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("Registry 配置必须是 dict（报表名 → 文件名前缀）")
    return data