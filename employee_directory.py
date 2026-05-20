"""employee_directory — 员工姓名 ↔ 员工编号 双向映射加载器。

数据源：数据源/海外思维LP架构表 (*).xlsx
表头在第 4 行（1-based），关键列：姓名、员工编号。

用途：CSV 输出时把 owner aggregate 模式下的"李文韬" 还原为编号 "H0019538"。
"""

import logging
from pathlib import Path

import openpyxl

logger = logging.getLogger("incentive_bot")

ARCHITECTURE_FILE_PREFIX = "海外思维LP架构表"
HEADER_ROW_INDEX = 3  # 0-based: 第 4 行是表头


class EmployeeDirectory:
    """姓名 ↔ 员工编号 双向查询。缺失返回空字符串。"""

    def __init__(self, by_name: dict[str, str], by_id: dict[str, str]):
        self._by_name = by_name
        self._by_id = by_id

    def get_id(self, name: str) -> str:
        return self._by_name.get(str(name).strip(), "")

    def get_name(self, employee_id: str) -> str:
        return self._by_id.get(str(employee_id).strip(), "")

    def resolve_pair(self, identifier: str) -> tuple[str, str]:
        """传入姓名或编号，返回 (姓名, 编号)。未匹配的位置留空字符串。"""
        if identifier is None:
            return "", ""
        key = str(identifier).strip()
        if key in self._by_name:
            return key, self._by_name[key]
        if key in self._by_id:
            return self._by_id[key], key
        return key, ""

    def __len__(self) -> int:
        return len(self._by_name)


def load_employee_directory(data_dir: Path) -> EmployeeDirectory:
    """从 LP 架构表加载 姓名↔员工编号 映射。文件不存在或读取失败返回空 directory。"""
    matches = sorted(
        data_dir.glob(f"{ARCHITECTURE_FILE_PREFIX}*.xlsx"),
        key=lambda f: f.stat().st_mtime,
        reverse=True,
    )
    matches = [f for f in matches if not f.name.startswith("~$")]
    if not matches:
        logger.warning(
            f"未找到 LP 架构表（前缀 '{ARCHITECTURE_FILE_PREFIX}'），员工编号映射为空"
        )
        return EmployeeDirectory({}, {})

    filepath = matches[0]
    logger.info(f"加载员工目录: {filepath.name}")

    wb = openpyxl.load_workbook(filepath, data_only=True)
    try:
        ws = wb.worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()

    if len(rows) <= HEADER_ROW_INDEX:
        logger.warning(f"LP 架构表行数不足，员工编号映射为空")
        return EmployeeDirectory({}, {})

    headers = [str(c).strip() if c is not None else "" for c in rows[HEADER_ROW_INDEX]]
    name_idx = _find_column(headers, "姓名", default_idx=6)
    id_idx = _find_column(headers, "员工编号", default_idx=7)

    if name_idx is None or id_idx is None:
        logger.warning(f"LP 架构表缺少 姓名/员工编号 列，headers={headers}")
        return EmployeeDirectory({}, {})

    by_name: dict[str, str] = {}
    by_id: dict[str, str] = {}
    for row in rows[HEADER_ROW_INDEX + 1 :]:
        if name_idx >= len(row) or id_idx >= len(row):
            continue
        name = row[name_idx]
        eid = row[id_idx]
        if name is None or eid is None:
            continue
        name_s = str(name).strip()
        eid_s = str(eid).strip()
        if not name_s or not eid_s:
            continue
        # 姓名重名时保留首次出现（架构表一般主管行在前），编号唯一
        by_name.setdefault(name_s, eid_s)
        by_id[eid_s] = name_s

    logger.info(f"员工目录: {len(by_name)} 个姓名 / {len(by_id)} 个编号")
    return EmployeeDirectory(by_name, by_id)


def _find_column(headers: list[str], name: str, default_idx: int) -> int | None:
    """按列名定位，找不到则用 default_idx，仍越界返回 None。"""
    if name in headers:
        return headers.index(name)
    if 0 <= default_idx < len(headers):
        return default_idx
    return None