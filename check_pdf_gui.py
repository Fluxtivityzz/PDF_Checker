# -*- coding: utf-8 -*-
"""
报货 PDF 自动汇总 GUI 工具

功能：
1. 读取 Temu/拣货单 PDF 中的 属性集、SKU ID、SKU货号、实际发货数。
2. 按自定义规则把 SKU货号 映射为你表格里的货号。
3. 支持套装拆分：一个 SKU 可拆成多个款号。
4. 支持同款多颜色拆分：白色+卡其色-L 可拆成 白色-L、卡其色-L。
5. 自动合并统计：输出款号 + 颜色尺码 相同则数量相加。
6. GUI 可维护规则，并保存到 JSON 配置。

依赖：
    pip install pymupdf openpyxl

GUI 启动：
    python baohuo_pdf_gui.py

命令行批处理：
    python baohuo_pdf_gui.py --cli TEST.pdf -o TEST_报货汇总.xlsx
    python baohuo_pdf_gui.py --cli ./pdf_folder -o ./输出目录
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import traceback
from collections import OrderedDict, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

try:
    import fitz  # PyMuPDF
except ImportError as e:
    raise SystemExit("缺少 PyMuPDF，请先运行：pip install pymupdf") from e

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
except ImportError as e:
    raise SystemExit("缺少 openpyxl，请先运行：pip install openpyxl") from e

# Tkinter 在部分精简 Python 环境中可能不存在；只有 GUI 模式才真正需要。
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except Exception:  # pragma: no cover
    tk = None
    ttk = None
    filedialog = None
    messagebox = None

APP_NAME = "报货PDF汇总工具"
CONFIG_FILE_NAME = "check_rules.json"

HAN_RE = re.compile(r"[\u4e00-\u9fff]")
# 可识别：SK-010、PG-JS-836、MGTZ-26、MG-55、SSL-14、XZ-21+XZ-20、MG-55+SSL-14
STYLE_CODE_RE = re.compile(r"^([A-Z]+(?:-[A-Z]+)*-\d+(?:\+[A-Z]+(?:-[A-Z]+)*-\d+)*)")
SIZE_ORDER = ["XS", "S", "M", "L", "XL", "XXL", "XXXL", "XXXXL"]
COLOR_ORDER = [
    "白色", "黑色", "灰色", "浅灰色", "深灰色",
    "浅蓝色", "天蓝色", "藏青色", "军绿色", "军绿", "卡其色", "杏色", "粉红色", "土黄色",
    "白色+卡其色", "白色+黑色", "黑色+藏青色", "军绿+卡其色", "军绿色+卡其色",
]

DEFAULT_CONFIG: Dict[str, Any] = {
    "version": 3,
    "general": {
        "style_prefix_replace": [
            {"from": "MGTZ-", "to": "MG-", "enabled": True},
        ],
        "split_auto_combo_styles": True,
        "merge_same_style_attr": True,
        "default_output_name": "报货汇总.xlsx",
        "default_group_name": "未分组",
        "show_group_headers": True,
        "blank_row_between_styles": True,
    },
    "color_map": {
        "White": "白色",
        "Black": "黑色",
        "Grey": "灰色",
        "Gray": "灰色",
        "LightGrey": "浅灰色",
        "LightGray": "浅灰色",
        "DarkGrey": "深灰色",
        "DarkGray": "深灰色",
        "LightBlue": "浅蓝色",
        "SkyBlue": "天蓝色",
        "NavyBlue": "藏青色",
        "DarkBlue": "藏青色",
        "ArmyGreen": "军绿色",
        "ArmyBlue": "军绿色",
        "Khaki": "卡其色",
        "Apricot": "杏色",
        "PinkRed": "粉红色",
        "EarthyYellow": "土黄色",
    },
    "rules": [
        {
            "enabled": True,
            "name": "默认：自动识别款号，组合款自动拆款",
            "priority": 1000,
            "match_type": "regex",
            "pattern": ".*",
            "output_styles": "{auto}",
            "color_mode": "keep",  # keep / split_attr_color / fixed_colors / sku_color
            "fixed_colors": "",
            "qty_multiplier": 1,
            "group_name": "自动识别",
            "note": "兜底规则。例：MG-55+SSL-14-Black 自动拆成 MG-55 与 SSL-14。",
        },
        {
            "enabled": False,
            "name": "示例：同款双色拆分 MG-55-White+Khaki",
            "priority": 10,
            "match_type": "contains",
            "pattern": "MG-55-White+Khaki",
            "output_styles": "MG-55",
            "color_mode": "sku_color",
            "fixed_colors": "",
            "qty_multiplier": 1,
            "group_name": "MG系列",
            "note": "启用后会把 MG-55-White+Khaki 拆成 MG-55 白色-尺码 与 MG-55 卡其色-尺码。",
        },
        {
            "enabled": False,
            "name": "示例：固定拆两个款式",
            "priority": 20,
            "match_type": "contains",
            "pattern": "MG-55+SSL-14-Black",
            "output_styles": "MG-55,SSL-14",
            "color_mode": "keep",
            "fixed_colors": "",
            "qty_multiplier": 1,
            "group_name": "套装拆分",
            "note": "启用后会把这个 SKU 同时计入 MG-55 与 SSL-14。",
        },
        {
            "enabled": False,
            "name": "示例：按属性集拆色",
            "priority": 30,
            "match_type": "contains",
            "pattern": "White+Black",
            "output_styles": "{auto}",
            "color_mode": "split_attr_color",
            "fixed_colors": "",
            "qty_multiplier": 1,
            "group_name": "双色拆分",
            "note": "例：白色+黑色-XL 拆成 白色-XL 与 黑色-XL。",
        },
    ],
    "group_rules": [
        {
            "enabled": True,
            "priority": 100,
            "name": "默认：MG系列",
            "match_type": "startswith",
            "pattern": "MG-",
            "group_name": "MG系列",
        },
        {
            "enabled": True,
            "priority": 110,
            "name": "默认：SSL系列",
            "match_type": "startswith",
            "pattern": "SSL-",
            "group_name": "SSL系列",
        },
        {
            "enabled": True,
            "priority": 120,
            "name": "默认：XZ系列",
            "match_type": "startswith",
            "pattern": "XZ-",
            "group_name": "XZ系列",
        },
        {
            "enabled": True,
            "priority": 130,
            "name": "默认：SK系列",
            "match_type": "startswith",
            "pattern": "SK-",
            "group_name": "SK系列",
        },
        {
            "enabled": True,
            "priority": 140,
            "name": "默认：PG系列",
            "match_type": "startswith",
            "pattern": "PG-",
            "group_name": "PG系列",
        },
    ],
}


@dataclass
class ParsedRow:
    source_file: str
    page: int
    attr: str
    sku_id: str
    sku_code: str
    qty: int


@dataclass
class OutputRow:
    source_file: str
    page: int
    original_attr: str
    sku_id: str
    sku_code: str
    output_style: str
    group_name: str
    output_attr: str
    qty: float
    rule_name: str


@dataclass
class ParseIssue:
    source_file: str
    page: Optional[int]
    issue: str
    text: str


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def default_config_path() -> Path:
    return app_dir() / CONFIG_FILE_NAME


def deep_copy_default_config() -> Dict[str, Any]:
    return json.loads(json.dumps(DEFAULT_CONFIG, ensure_ascii=False))


def load_config(path: Optional[Path] = None) -> Dict[str, Any]:
    path = path or default_config_path()
    if not path.exists():
        cfg = deep_copy_default_config()
        save_config(cfg, path)
        return cfg
    with open(path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    # 兼容旧配置缺字段
    merged = deep_copy_default_config()
    merged.update(cfg)
    merged["general"].update(cfg.get("general", {}))
    merged["color_map"].update(cfg.get("color_map", {}))
    merged["rules"] = cfg.get("rules", merged["rules"])
    merged["group_rules"] = cfg.get("group_rules", merged.get("group_rules", []))
    return merged


def save_config(cfg: Dict[str, Any], path: Optional[Path] = None) -> None:
    path = path or default_config_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def clean_text_line(s: str) -> str:
    return (s or "").strip().replace("\u3000", " ")


def is_attr_line(s: str) -> bool:
    if not s or "-" not in s:
        return False
    if not HAN_RE.search(s):
        return False
    bad_prefix = ("SKC货号", "备货", "要求", "创建", "收货", "打印", "第", "数量", "合计", "序号", "商品信息")
    return not s.startswith(bad_prefix)


def is_sku_id(s: str) -> bool:
    return bool(re.fullmatch(r"\d{8,14}", s or ""))


def is_int(s: str) -> bool:
    return bool(re.fullmatch(r"\d+", s or ""))


def detect_report_date(lines_with_page: List[Tuple[int, str]]) -> str:
    full_text = "\n".join(line for _, line in lines_with_page)
    m = re.search(r"创建时间：(\d{4})-(\d{2})-(\d{2})", full_text)
    if not m:
        m = re.search(r"要求发货时间：(\d{4})-(\d{2})-(\d{2})", full_text)
    if not m:
        return "报货汇总"
    return f"{int(m.group(2))}月{int(m.group(3))}号"


def extract_pdf_lines(pdf_path: Path) -> List[Tuple[int, str]]:
    lines: List[Tuple[int, str]] = []
    with fitz.open(pdf_path) as doc:
        for page_no, page in enumerate(doc, start=1):
            text = page.get_text("text")
            for line in text.splitlines():
                line = clean_text_line(line)
                if line:
                    lines.append((page_no, line))
    return lines


def parse_pdf_rows(pdf_path: Path) -> Tuple[List[ParsedRow], List[ParseIssue], str]:
    lines_with_page = extract_pdf_lines(pdf_path)
    lines = [line for _, line in lines_with_page]
    pages = [page for page, _ in lines_with_page]
    rows: List[ParsedRow] = []
    issues: List[ParseIssue] = []

    i = 0
    while i < len(lines):
        line = lines[i]
        if is_attr_line(line):
            attr = line
            j = i + 1

            # 处理属性集换行，例如 白色+卡其色-XX / XL -> 白色+卡其色-XXXL
            while j < len(lines) and not is_sku_id(lines[j]):
                if re.fullmatch(r"[XSML]+", lines[j]):
                    attr += lines[j]
                    j += 1
                else:
                    break

            if j < len(lines) and is_sku_id(lines[j]):
                sku_id = lines[j]
                k = j + 1
                sku_parts: List[str] = []
                while k < len(lines) and not is_int(lines[k]):
                    sku_parts.append(lines[k])
                    k += 1
                if k < len(lines) and sku_parts and is_int(lines[k]):
                    sku_code = "".join(sku_parts).strip()
                    qty = int(lines[k])
                    rows.append(ParsedRow(
                        source_file=pdf_path.name,
                        page=pages[i],
                        attr=attr,
                        sku_id=sku_id,
                        sku_code=sku_code,
                        qty=qty,
                    ))
                    i = k + 1
                    continue
                else:
                    issues.append(ParseIssue(pdf_path.name, pages[i], "疑似属性行但未读取到 SKU货号/数量", line))
        i += 1

    title = detect_report_date(lines_with_page)
    if not rows:
        issues.append(ParseIssue(pdf_path.name, None, "未解析到数据", "请确认 PDF 不是纯图片扫描件，且包含 属性集 / SKU ID / SKU货号 / 实际发货数字段。"))
    return rows, issues, title


def normalize_style(style: str, cfg: Dict[str, Any]) -> str:
    style = (style or "").strip()
    for item in cfg.get("general", {}).get("style_prefix_replace", []):
        if not item.get("enabled", True):
            continue
        old = str(item.get("from", ""))
        new = str(item.get("to", ""))
        if old and style.startswith(old):
            style = new + style[len(old):]
    return style


def get_auto_style_codes(sku_code: str, cfg: Dict[str, Any]) -> List[str]:
    code = (sku_code or "").strip().replace(" ", "-")
    code = re.sub(r"-+", "-", code)
    m = STYLE_CODE_RE.match(code)
    if not m:
        return [normalize_style(code, cfg)] if code else ["未识别款号"]
    base = m.group(1)
    split_combo = cfg.get("general", {}).get("split_auto_combo_styles", True)
    if split_combo and "+" in base:
        return [normalize_style(x, cfg) for x in base.split("+") if x]
    return [normalize_style(base, cfg)]


def split_attr(attr: str) -> Tuple[str, str]:
    attr = (attr or "").strip()
    if "-" not in attr:
        return attr, ""
    color, size = attr.rsplit("-", 1)
    return color, size


def normalize_separator_list(s: str) -> List[str]:
    if not s:
        return []
    s = s.replace("，", ",").replace("；", ",").replace(";", ",").replace("/", ",")
    parts = []
    for chunk in s.split(","):
        chunk = chunk.strip()
        if chunk:
            parts.append(chunk)
    return parts


def parse_output_styles(rule_styles: str, sku_code: str, cfg: Dict[str, Any]) -> List[str]:
    rule_styles = (rule_styles or "").strip()
    if not rule_styles or rule_styles == "{auto}":
        return get_auto_style_codes(sku_code, cfg)
    # 用逗号分隔。允许用户写 MG-55,SSL-14；不要用 + 分隔，避免和款号中的 + 混淆。
    styles = normalize_separator_list(rule_styles)
    if not styles:
        return get_auto_style_codes(sku_code, cfg)
    return [normalize_style(x, cfg) for x in styles]


def english_color_tokens_from_sku(sku_code: str) -> List[str]:
    """从 SKU 末尾颜色段里取英文颜色。例：MG-55-White+Khaki -> White, Khaki。"""
    code = (sku_code or "").strip().replace(" ", "-")
    m = STYLE_CODE_RE.match(code)
    if not m:
        return []
    rest = code[m.end():]
    rest = rest.strip("-")
    if not rest:
        return []
    # 如果 SKU 末尾还有尺码，例如 JB-003-Black-XXXL，把末尾尺码去掉。
    rest = re.sub(r"-(XS|S|M|L|XL|XXL|XXXL|XXXXL)$", "", rest, flags=re.I)
    tokens = []
    for p in re.split(r"\+|,|/", rest):
        p = p.strip("-").strip()
        if p:
            tokens.append(p)
    return tokens


def map_color_token(token: str, cfg: Dict[str, Any]) -> str:
    cmap = cfg.get("color_map", {})
    if token in cmap:
        return cmap[token]
    # 大小写兜底
    for k, v in cmap.items():
        if k.lower() == token.lower():
            return v
    return token


def output_attrs_by_color_mode(rule: Dict[str, Any], row: ParsedRow, cfg: Dict[str, Any]) -> List[str]:
    color, size = split_attr(row.attr)
    color_mode = rule.get("color_mode", "keep")

    if color_mode == "keep":
        return [row.attr]

    if color_mode == "split_attr_color":
        colors = [c.strip() for c in re.split(r"\+", color) if c.strip()]
        if not colors:
            return [row.attr]
        return [f"{c}-{size}" if size else c for c in colors]

    if color_mode == "fixed_colors":
        colors = normalize_separator_list(str(rule.get("fixed_colors", "")))
        if not colors:
            return [row.attr]
        return [f"{c}-{size}" if size else c for c in colors]

    if color_mode == "sku_color":
        tokens = english_color_tokens_from_sku(row.sku_code)
        colors = [map_color_token(t, cfg) for t in tokens]
        if not colors:
            return [row.attr]
        return [f"{c}-{size}" if size else c for c in colors]

    return [row.attr]


def rule_matches(rule: Dict[str, Any], sku_code: str) -> bool:
    if not rule.get("enabled", True):
        return False
    pattern = str(rule.get("pattern", ""))
    match_type = rule.get("match_type", "contains")
    if not pattern and match_type != "regex":
        return False
    try:
        if match_type == "exact":
            return sku_code == pattern
        if match_type == "startswith":
            return sku_code.startswith(pattern)
        if match_type == "contains":
            return pattern in sku_code
        if match_type == "regex":
            return re.search(pattern, sku_code) is not None
    except re.error:
        return False
    return False


def to_int(value: Any, default: int = 9999) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def sorted_indexed_rules(cfg: Dict[str, Any]) -> List[Tuple[int, Dict[str, Any]]]:
    return sorted(
        enumerate(cfg.get("rules", [])),
        key=lambda item: (to_int(item[1].get("priority", 9999)), item[0]),
    )


def sorted_indexed_group_rules(cfg: Dict[str, Any]) -> List[Tuple[int, Dict[str, Any]]]:
    return sorted(
        enumerate(cfg.get("group_rules", [])),
        key=lambda item: (to_int(item[1].get("priority", 9999)), item[0]),
    )


def sorted_rules(cfg: Dict[str, Any]) -> List[Dict[str, Any]]:
    return [rule for _, rule in sorted_indexed_rules(cfg)]


def sorted_group_rules(cfg: Dict[str, Any]) -> List[Dict[str, Any]]:
    return [rule for _, rule in sorted_indexed_group_rules(cfg)]


def resolve_group_name(style: str, matched_rule: Dict[str, Any], cfg: Dict[str, Any]) -> str:
    """按输出款号匹配分组规则；没有匹配时使用 SKU 规则中的分组名；仍为空则归为未分组。"""
    for gr in sorted_group_rules(cfg):
        if rule_matches(gr, style):
            name = str(gr.get("group_name", "")).strip()
            if name:
                return name
    rule_group = str(matched_rule.get("group_name", "")).strip()
    if rule_group:
        return rule_group
    return str(cfg.get("general", {}).get("default_group_name", "未分组") or "未分组")


def apply_rules_to_row(row: ParsedRow, cfg: Dict[str, Any]) -> List[OutputRow]:
    rules = sorted_rules(cfg)
    matched_rule: Optional[Dict[str, Any]] = None
    for rule in rules:
        if rule_matches(rule, row.sku_code):
            matched_rule = rule
            break

    if matched_rule is None:
        matched_rule = {
            "name": "无匹配规则，使用自动识别",
            "output_styles": "{auto}",
            "color_mode": "keep",
            "qty_multiplier": 1,
        }

    styles = parse_output_styles(str(matched_rule.get("output_styles", "{auto}")), row.sku_code, cfg)
    attrs = output_attrs_by_color_mode(matched_rule, row, cfg)
    try:
        multiplier = float(matched_rule.get("qty_multiplier", 1) or 1)
    except Exception:
        multiplier = 1

    out_rows: List[OutputRow] = []
    for style in styles:
        for attr in attrs:
            qty = row.qty * multiplier
            # 一般报货数量应为整数；若用户设置了小数倍数，保留整数四舍五入并在明细中体现。
            if abs(qty - round(qty)) < 1e-9:
                qty_value = int(round(qty))
            else:
                qty_value = qty
            out_rows.append(OutputRow(
                source_file=row.source_file,
                page=row.page,
                original_attr=row.attr,
                sku_id=row.sku_id,
                sku_code=row.sku_code,
                output_style=style,
                group_name=resolve_group_name(style, matched_rule, cfg),
                output_attr=attr,
                qty=qty_value,
                rule_name=str(matched_rule.get("name", "未命名规则")),
            ))
    return out_rows


def attr_sort_key(attr: str) -> Tuple[int, str, int, str]:
    color, size = split_attr(attr)
    try:
        color_i = COLOR_ORDER.index(color)
    except ValueError:
        color_i = 999
    try:
        size_i = SIZE_ORDER.index(size)
    except ValueError:
        size_i = 999
    return color_i, color, size_i, size


def aggregate_output_rows(out_rows: List[OutputRow]) -> Tuple[OrderedDict, Dict[Tuple[str, str], int], Dict[str, int]]:
    """汇总为：分组 -> 款号 -> 属性集 -> 数量。"""
    groups: OrderedDict[str, OrderedDict[str, OrderedDict[str, float]]] = OrderedDict()
    style_first_index: Dict[Tuple[str, str], int] = {}
    group_first_index: Dict[str, int] = {}
    for idx, r in enumerate(out_rows):
        group = (r.group_name or "未分组").strip() or "未分组"
        if group not in groups:
            groups[group] = OrderedDict()
            group_first_index[group] = idx
        if r.output_style not in groups[group]:
            groups[group][r.output_style] = OrderedDict()
            style_first_index[(group, r.output_style)] = idx
        groups[group][r.output_style][r.output_attr] = groups[group][r.output_style].get(r.output_attr, 0) + r.qty
    return groups, style_first_index, group_first_index


def group_sort_key(group: str, group_first_index: Dict[str, int]) -> Tuple[int, str]:
    return (group_first_index.get(group, 999999), group)


def style_sort_key(group: str, style: str, first_index: Dict[Tuple[str, str], int]) -> Tuple[int, str]:
    return (first_index.get((group, style), 999999), style)


def _fmt_qty(qty: float) -> Any:
    if isinstance(qty, float) and abs(qty - round(qty)) < 1e-9:
        return int(round(qty))
    return qty


def write_excel(
    groups: OrderedDict,
    first_index: Dict[Tuple[str, str], int],
    group_first_index: Dict[str, int],
    detail_rows: List[OutputRow],
    parsed_rows: List[ParsedRow],
    issues: List[ParseIssue],
    title: str,
    out_path: Path,
    cfg: Dict[str, Any],
) -> None:
    wb = Workbook()
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    group_fill = PatternFill("solid", fgColor="D9EAF7")
    total_fill = PatternFill("solid", fgColor="FFF2CC")

    ws = wb.active
    ws.title = "汇总"
    ws.merge_cells("A1:D1")
    ws["A1"] = title
    ws["A1"].font = Font(name="微软雅黑", size=18, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
    for col, width in zip(["A", "B", "C", "D"], [16, 22, 10, 10]):
        ws.column_dimensions[col].width = width

    unique_styles = set()
    total_qty = 0
    for group, styles in groups.items():
        for style, attrs in styles.items():
            unique_styles.add(style)
            total_qty += sum(attrs.values())

    ws["A2"] = "不同款号"
    ws["B2"] = f"{len(unique_styles)}款"
    ws["C2"] = "总数量"
    ws["D2"] = _fmt_qty(total_qty)
    for col in range(1, 5):
        cell = ws.cell(row=2, column=col)
        cell.font = Font(name="微软雅黑", size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = total_fill
        cell.border = border

    row_idx = 4
    show_group_headers = bool(cfg.get("general", {}).get("show_group_headers", True))
    blank_between = bool(cfg.get("general", {}).get("blank_row_between_styles", True))

    for group in sorted(groups.keys(), key=lambda g: group_sort_key(g, group_first_index)):
        styles = groups[group]
        group_styles = set(styles.keys())
        group_qty = sum(sum(attrs.values()) for attrs in styles.values())

        if show_group_headers:
            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=4)
            cell = ws.cell(row=row_idx, column=1)
            cell.value = f"分组：{group}    款数：{len(group_styles)}款    数量：{_fmt_qty(group_qty)}"
            cell.font = Font(name="微软雅黑", size=11, bold=True)
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.fill = group_fill
            cell.border = border
            for col in range(2, 5):
                ws.cell(row=row_idx, column=col).border = border
                ws.cell(row=row_idx, column=col).fill = group_fill
            row_idx += 1

        for style in sorted(styles.keys(), key=lambda st: style_sort_key(group, st, first_index)):
            attrs = styles[style]
            items = sorted(attrs.items(), key=lambda kv: attr_sort_key(kv[0]))
            style_total = sum(q for _, q in items)
            for n, (attr, qty) in enumerate(items):
                if n == 0:
                    ws.cell(row=row_idx, column=1, value=style)
                ws.cell(row=row_idx, column=2, value=attr)
                ws.cell(row=row_idx, column=3, value=_fmt_qty(qty))
                if n == len(items) - 1:
                    ws.cell(row=row_idx, column=4, value=_fmt_qty(style_total))
                for col in range(1, 5):
                    cell = ws.cell(row=row_idx, column=col)
                    cell.font = Font(name="微软雅黑", size=10, bold=(col == 1 and n == 0))
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = border
                row_idx += 1
            if blank_between:
                row_idx += 1

    ws.cell(row=row_idx, column=1, value="合计")
    ws.cell(row=row_idx, column=2, value=f"{len(unique_styles)}款")
    ws.cell(row=row_idx, column=3, value=_fmt_qty(total_qty))
    for col in range(1, 5):
        cell = ws.cell(row=row_idx, column=col)
        cell.font = Font(name="微软雅黑", size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = total_fill
        cell.border = border
    ws.freeze_panes = "A4"

    ds = wb.create_sheet("明细")
    headers = ["来源PDF", "页码", "原始SKU货号", "SKU ID", "原始属性集", "分组", "输出款号", "输出属性集", "数量", "命中规则"]
    ds.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ds.cell(row=1, column=c)
        cell.font = Font(name="微软雅黑", size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
        cell.border = border
    for r in detail_rows:
        ds.append([r.source_file, r.page, r.sku_code, r.sku_id, r.original_attr, r.group_name, r.output_style, r.output_attr, r.qty, r.rule_name])
    widths = [24, 8, 36, 16, 18, 16, 14, 18, 8, 32]
    for idx, width in enumerate(widths, start=1):
        ds.column_dimensions[chr(64 + idx)].width = width
    for row_cells in ds.iter_rows(min_row=2):
        for cell in row_cells:
            cell.font = Font(name="微软雅黑", size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
    ds.freeze_panes = "A2"

    raw = wb.create_sheet("原始解析")
    raw_headers = ["来源PDF", "页码", "属性集", "SKU ID", "SKU货号", "实际发货数"]
    raw.append(raw_headers)
    for c in range(1, len(raw_headers) + 1):
        cell = raw.cell(row=1, column=c)
        cell.font = Font(name="微软雅黑", size=10, bold=True)
        cell.fill = PatternFill("solid", fgColor="E2F0D9")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    for r in parsed_rows:
        raw.append([r.source_file, r.page, r.attr, r.sku_id, r.sku_code, r.qty])
    for idx, width in enumerate([24, 8, 18, 16, 36, 10], start=1):
        raw.column_dimensions[chr(64 + idx)].width = width
    for row_cells in raw.iter_rows(min_row=2):
        for cell in row_cells:
            cell.font = Font(name="微软雅黑", size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
    raw.freeze_panes = "A2"

    er = wb.create_sheet("异常检查")
    er_headers = ["来源PDF", "页码", "异常类型", "内容"]
    er.append(er_headers)
    for c in range(1, len(er_headers) + 1):
        cell = er.cell(row=1, column=c)
        cell.font = Font(name="微软雅黑", size=10, bold=True)
        cell.fill = PatternFill("solid", fgColor="FCE4D6")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    if issues:
        for issue in issues:
            er.append([issue.source_file, issue.page or "", issue.issue, issue.text])
    else:
        er.append(["", "", "无异常", "解析过程没有发现明显异常"])
    for idx, width in enumerate([24, 8, 24, 80], start=1):
        er.column_dimensions[chr(64 + idx)].width = width
    for row_cells in er.iter_rows(min_row=2):
        for cell in row_cells:
            cell.font = Font(name="微软雅黑", size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
    er.freeze_panes = "A2"

    rs = wb.create_sheet("规则快照")
    rs.append(["配置文件版本", cfg.get("version", "")])
    rs.append([])
    rs.append(["启用", "优先级", "规则名", "分组", "匹配方式", "匹配内容", "输出款号", "颜色模式", "固定颜色", "数量倍数", "备注"])
    for rule in sorted_rules(cfg):
        rs.append([
            "是" if rule.get("enabled", True) else "否",
            rule.get("priority", ""),
            rule.get("name", ""),
            rule.get("group_name", ""),
            rule.get("match_type", ""),
            rule.get("pattern", ""),
            rule.get("output_styles", ""),
            rule.get("color_mode", ""),
            rule.get("fixed_colors", ""),
            rule.get("qty_multiplier", ""),
            rule.get("note", ""),
        ])
    for row_cells in rs.iter_rows():
        for cell in row_cells:
            cell.font = Font(name="微软雅黑", size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
    for idx, width in enumerate([8, 8, 36, 16, 12, 32, 24, 18, 24, 10, 60], start=1):
        rs.column_dimensions[chr(64 + idx)].width = width
    rs.freeze_panes = "A4"

    grs = wb.create_sheet("分组规则快照")
    grs.append(["启用", "优先级", "规则名", "匹配方式", "匹配内容", "分组名称"])
    for gr in sorted_group_rules(cfg):
        grs.append([
            "是" if gr.get("enabled", True) else "否",
            gr.get("priority", ""),
            gr.get("name", ""),
            gr.get("match_type", ""),
            gr.get("pattern", ""),
            gr.get("group_name", ""),
        ])
    for row_cells in grs.iter_rows():
        for cell in row_cells:
            cell.font = Font(name="微软雅黑", size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
    for idx, width in enumerate([8, 8, 36, 12, 32, 18], start=1):
        grs.column_dimensions[chr(64 + idx)].width = width
    grs.freeze_panes = "A2"

    wb.save(out_path)


def collect_pdf_paths(inputs: Iterable[Path]) -> List[Path]:
    result: List[Path] = []
    for p in inputs:
        if p.is_file() and p.suffix.lower() == ".pdf":
            result.append(p)
        elif p.is_dir():
            result.extend(sorted(x for x in p.iterdir() if x.is_file() and x.suffix.lower() == ".pdf"))
    # 去重但保序
    seen = set()
    out = []
    for p in result:
        rp = str(p.resolve())
        if rp not in seen:
            seen.add(rp)
            out.append(p)
    return out


def process_pdfs(pdf_paths: List[Path], out_path: Path, cfg: Dict[str, Any]) -> Dict[str, Any]:
    all_parsed: List[ParsedRow] = []
    all_issues: List[ParseIssue] = []
    titles: List[str] = []

    for pdf in pdf_paths:
        rows, issues, title = parse_pdf_rows(pdf)
        all_parsed.extend(rows)
        all_issues.extend(issues)
        titles.append(title)

    detail_rows: List[OutputRow] = []
    for row in all_parsed:
        detail_rows.extend(apply_rules_to_row(row, cfg))

    groups, first_index, group_first_index = aggregate_output_rows(detail_rows)
    title = titles[0] if titles else "报货汇总"
    if len(set(titles)) > 1:
        title = "报货汇总"

    write_excel(groups, first_index, group_first_index, detail_rows, all_parsed, all_issues, title, out_path, cfg)

    original_qty = sum(r.qty for r in all_parsed)
    output_qty = sum(sum(sum(attrs.values()) for attrs in styles.values()) for styles in groups.values())
    style_count = len({style for styles in groups.values() for style in styles.keys()})
    summary_line_count = sum(len(attrs) for styles in groups.values() for attrs in styles.values())
    return {
        "out_path": str(out_path),
        "pdf_count": len(pdf_paths),
        "parsed_rows": len(all_parsed),
        "original_qty": original_qty,
        "output_lines": style_count,
        "summary_line_count": summary_line_count,
        "output_qty": output_qty,
        "issues": len(all_issues),
    }


# ============================== GUI ==============================

if tk is not None:
    class RuleDialog(tk.Toplevel):
        def __init__(self, master, rule: Optional[Dict[str, Any]] = None):
            super().__init__(master)
            self.title("规则设置")
            self.resizable(False, False)
            self.result: Optional[Dict[str, Any]] = None
            self.rule = rule or {
                "enabled": True,
                "name": "新规则",
                "priority": 100,
                "match_type": "contains",
                "pattern": "",
                "output_styles": "{auto}",
                "color_mode": "keep",
                "fixed_colors": "",
                "qty_multiplier": 1,
                "group_name": "",
                "note": "",
            }
            self.vars: Dict[str, tk.Variable] = {}
            self._build()
            self.grab_set()
            self.transient(master)

        def _build(self):
            pad = {"padx": 8, "pady": 5}
            frm = ttk.Frame(self)
            frm.grid(row=0, column=0, sticky="nsew")

            self.vars["enabled"] = tk.BooleanVar(value=bool(self.rule.get("enabled", True)))
            ttk.Checkbutton(frm, text="启用此规则", variable=self.vars["enabled"]).grid(row=0, column=0, columnspan=2, sticky="w", **pad)

            fields = [
                ("规则名", "name", "entry"),
                ("优先级", "priority", "entry"),
                ("匹配方式", "match_type", "combo_match"),
                ("匹配内容", "pattern", "entry"),
                ("输出款号", "output_styles", "entry"),
                ("颜色模式", "color_mode", "combo_color"),
                ("固定颜色", "fixed_colors", "entry"),
                ("数量倍数", "qty_multiplier", "entry"),
                ("默认分组", "group_name", "entry"),
                ("备注", "note", "entry"),
            ]
            row = 1
            for label, key, kind in fields:
                ttk.Label(frm, text=label).grid(row=row, column=0, sticky="e", **pad)
                if kind == "combo_match":
                    var = tk.StringVar(value=str(self.rule.get(key, "contains")))
                    self.vars[key] = var
                    cb = ttk.Combobox(frm, textvariable=var, width=38, state="readonly",
                                      values=["contains", "exact", "startswith", "regex"])
                    cb.grid(row=row, column=1, sticky="w", **pad)
                elif kind == "combo_color":
                    var = tk.StringVar(value=str(self.rule.get(key, "keep")))
                    self.vars[key] = var
                    cb = ttk.Combobox(frm, textvariable=var, width=38, state="readonly",
                                      values=["keep", "split_attr_color", "fixed_colors", "sku_color"])
                    cb.grid(row=row, column=1, sticky="w", **pad)
                else:
                    var = tk.StringVar(value=str(self.rule.get(key, "")))
                    self.vars[key] = var
                    ttk.Entry(frm, textvariable=var, width=42).grid(row=row, column=1, sticky="w", **pad)
                row += 1

            help_text = (
                "说明：\n"
                "1. 输出款号写 {auto} 表示从 SKU货号 自动识别。\n"
                "2. 拆两个款式时写：MG-55,SSL-14。\n"
                "3. keep=保留原属性集；split_attr_color=把 白色+卡其色-L 拆成 白色-L/卡其色-L。\n"
                "4. sku_color=从 SKU 英文颜色拆色，例如 White+Khaki。\n"
                "5. fixed_colors=使用你手动填写的颜色列表，例如 白色,卡其色。\n"
                "6. 默认分组只在没有命中‘分组设置’时生效；分组设置优先级更高。"
            )
            ttk.Label(frm, text=help_text, justify="left", foreground="#555").grid(row=row, column=0, columnspan=2, sticky="w", **pad)
            row += 1

            btns = ttk.Frame(frm)
            btns.grid(row=row, column=0, columnspan=2, pady=8)
            ttk.Button(btns, text="保存", command=self._save).pack(side="left", padx=6)
            ttk.Button(btns, text="取消", command=self.destroy).pack(side="left", padx=6)

        def _save(self):
            try:
                priority = int(str(self.vars["priority"].get()).strip() or "100")
            except ValueError:
                messagebox.showerror("错误", "优先级必须是整数。")
                return
            try:
                multiplier = float(str(self.vars["qty_multiplier"].get()).strip() or "1")
            except ValueError:
                messagebox.showerror("错误", "数量倍数必须是数字。")
                return
            self.result = {
                "enabled": bool(self.vars["enabled"].get()),
                "name": self.vars["name"].get().strip() or "未命名规则",
                "priority": priority,
                "match_type": self.vars["match_type"].get(),
                "pattern": self.vars["pattern"].get().strip(),
                "output_styles": self.vars["output_styles"].get().strip() or "{auto}",
                "color_mode": self.vars["color_mode"].get(),
                "fixed_colors": self.vars["fixed_colors"].get().strip(),
                "qty_multiplier": multiplier,
                "group_name": self.vars["group_name"].get().strip(),
                "note": self.vars["note"].get().strip(),
            }
            self.destroy()


    class GroupRuleDialog(tk.Toplevel):
        def __init__(self, master, rule: Optional[Dict[str, Any]] = None):
            super().__init__(master)
            self.title("分组规则设置")
            self.resizable(False, False)
            self.result: Optional[Dict[str, Any]] = None
            self.rule = rule or {
                "enabled": True,
                "priority": 100,
                "name": "新分组规则",
                "match_type": "startswith",
                "pattern": "",
                "group_name": "",
            }
            self.vars: Dict[str, tk.Variable] = {}
            self._build()
            self.grab_set()
            self.transient(master)

        def _build(self):
            pad = {"padx": 8, "pady": 5}
            frm = ttk.Frame(self)
            frm.grid(row=0, column=0, sticky="nsew")

            self.vars["enabled"] = tk.BooleanVar(value=bool(self.rule.get("enabled", True)))
            ttk.Checkbutton(frm, text="启用此分组规则", variable=self.vars["enabled"]).grid(row=0, column=0, columnspan=2, sticky="w", **pad)

            fields = [
                ("规则名", "name", "entry"),
                ("优先级", "priority", "entry"),
                ("匹配方式", "match_type", "combo_match"),
                ("匹配款号", "pattern", "entry"),
                ("分组名称", "group_name", "entry"),
            ]
            row = 1
            for label, key, kind in fields:
                ttk.Label(frm, text=label).grid(row=row, column=0, sticky="e", **pad)
                if kind == "combo_match":
                    var = tk.StringVar(value=str(self.rule.get(key, "startswith")))
                    self.vars[key] = var
                    cb = ttk.Combobox(frm, textvariable=var, width=38, state="readonly",
                                      values=["contains", "exact", "startswith", "regex"])
                    cb.grid(row=row, column=1, sticky="w", **pad)
                else:
                    var = tk.StringVar(value=str(self.rule.get(key, "")))
                    self.vars[key] = var
                    ttk.Entry(frm, textvariable=var, width=42).grid(row=row, column=1, sticky="w", **pad)
                row += 1

            help_text = (
                "说明：分组规则按‘输出款号’匹配，不按 PDF 原始 SKU 匹配。\n"
                "例：匹配方式 startswith，匹配款号 MG-，分组名称 MG系列，所有 MG- 开头的款会排在一起。\n"
                "如果同一个款号命中多个分组规则，优先级数字越小越先使用。"
            )
            ttk.Label(frm, text=help_text, justify="left", foreground="#555").grid(row=row, column=0, columnspan=2, sticky="w", **pad)
            row += 1

            btns = ttk.Frame(frm)
            btns.grid(row=row, column=0, columnspan=2, pady=8)
            ttk.Button(btns, text="保存", command=self._save).pack(side="left", padx=6)
            ttk.Button(btns, text="取消", command=self.destroy).pack(side="left", padx=6)

        def _save(self):
            try:
                priority = int(str(self.vars["priority"].get()).strip() or "100")
            except ValueError:
                messagebox.showerror("错误", "优先级必须是整数。")
                return
            group_name = self.vars["group_name"].get().strip()
            if not group_name:
                messagebox.showerror("错误", "分组名称不能为空。")
                return
            self.result = {
                "enabled": bool(self.vars["enabled"].get()),
                "name": self.vars["name"].get().strip() or "未命名分组规则",
                "priority": priority,
                "match_type": self.vars["match_type"].get(),
                "pattern": self.vars["pattern"].get().strip(),
                "group_name": group_name,
            }
            self.destroy()


    class BaohuoApp(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title(APP_NAME)
            self.geometry("1120x720")
            self.minsize(980, 640)
            self.cfg_path = default_config_path()
            self.cfg = load_config(self.cfg_path)
            self.pdf_paths: List[Path] = []
            self._build_ui()
            self.refresh_rules_tree()
            self.refresh_group_rules_tree()
            self.refresh_color_table()
            self.refresh_prefix_table()

        def _build_ui(self):
            self.nb = ttk.Notebook(self)
            self.nb.pack(fill="both", expand=True)
            self.tab_process = ttk.Frame(self.nb)
            self.tab_rules = ttk.Frame(self.nb)
            self.tab_groups = ttk.Frame(self.nb)
            self.tab_color = ttk.Frame(self.nb)
            self.nb.add(self.tab_process, text="处理PDF")
            self.nb.add(self.tab_rules, text="SKU/套装规则")
            self.nb.add(self.tab_groups, text="分组设置")
            self.nb.add(self.tab_color, text="颜色映射/款号替换")
            self._build_process_tab()
            self._build_rules_tab()
            self._build_group_tab()
            self._build_color_tab()

        def _build_process_tab(self):
            top = ttk.Frame(self.tab_process)
            top.pack(fill="x", padx=10, pady=8)
            ttk.Button(top, text="添加PDF", command=self.add_pdfs).pack(side="left", padx=4)
            ttk.Button(top, text="添加文件夹", command=self.add_folder).pack(side="left", padx=4)
            ttk.Button(top, text="移除选中", command=self.remove_selected_pdf).pack(side="left", padx=4)
            ttk.Button(top, text="清空", command=self.clear_pdfs).pack(side="left", padx=4)

            mid = ttk.Frame(self.tab_process)
            mid.pack(fill="both", expand=True, padx=10, pady=4)
            self.pdf_tree = ttk.Treeview(mid, columns=("path",), show="headings", height=12)
            self.pdf_tree.heading("path", text="待处理PDF")
            self.pdf_tree.column("path", width=900, anchor="w")
            self.pdf_tree.pack(side="left", fill="both", expand=True)
            sb = ttk.Scrollbar(mid, orient="vertical", command=self.pdf_tree.yview)
            self.pdf_tree.configure(yscrollcommand=sb.set)
            sb.pack(side="right", fill="y")

            outfrm = ttk.Frame(self.tab_process)
            outfrm.pack(fill="x", padx=10, pady=8)
            ttk.Label(outfrm, text="输出Excel：").pack(side="left")
            self.output_var = tk.StringVar(value=str((Path.cwd() / "报货汇总.xlsx").resolve()))
            ttk.Entry(outfrm, textvariable=self.output_var).pack(side="left", fill="x", expand=True, padx=6)
            ttk.Button(outfrm, text="选择位置", command=self.choose_output).pack(side="left", padx=4)
            ttk.Button(outfrm, text="开始处理", command=self.run_process).pack(side="left", padx=4)

            logfrm = ttk.LabelFrame(self.tab_process, text="日志")
            logfrm.pack(fill="both", expand=True, padx=10, pady=8)
            self.log = tk.Text(logfrm, height=10)
            self.log.pack(fill="both", expand=True)

        def _build_rules_tab(self):
            btnfrm = ttk.Frame(self.tab_rules)
            btnfrm.pack(fill="x", padx=10, pady=8)
            ttk.Button(btnfrm, text="新增规则", command=self.add_rule).pack(side="left", padx=4)
            ttk.Button(btnfrm, text="编辑选中", command=self.edit_rule).pack(side="left", padx=4)
            ttk.Button(btnfrm, text="删除选中", command=self.delete_rule).pack(side="left", padx=4)
            ttk.Button(btnfrm, text="保存规则", command=self.save_all_config).pack(side="left", padx=4)
            ttk.Button(btnfrm, text="导入配置", command=self.import_config).pack(side="left", padx=4)
            ttk.Button(btnfrm, text="导出配置", command=self.export_config).pack(side="left", padx=4)
            ttk.Button(btnfrm, text="恢复默认配置", command=self.reset_config).pack(side="left", padx=4)

            cols = ("enabled", "priority", "name", "group", "match_type", "pattern", "output_styles", "color_mode", "fixed_colors", "multiplier", "note")
            self.rules_tree = ttk.Treeview(self.tab_rules, columns=cols, show="headings")
            headings = {
                "enabled": "启用", "priority": "优先级", "name": "规则名", "group": "默认分组", "match_type": "匹配方式",
                "pattern": "匹配内容", "output_styles": "输出款号", "color_mode": "颜色模式",
                "fixed_colors": "固定颜色", "multiplier": "倍数", "note": "备注"
            }
            widths = {"enabled": 50, "priority": 60, "name": 210, "group": 120, "match_type": 90, "pattern": 160,
                      "output_styles": 150, "color_mode": 120, "fixed_colors": 140, "multiplier": 60, "note": 260}
            for col in cols:
                self.rules_tree.heading(col, text=headings[col])
                self.rules_tree.column(col, width=widths[col], anchor="center")
            self.rules_tree.pack(fill="both", expand=True, padx=10, pady=5)
            self.rules_tree.bind("<Double-1>", lambda e: self.edit_rule())

        def _build_group_tab(self):
            btnfrm = ttk.Frame(self.tab_groups)
            btnfrm.pack(fill="x", padx=10, pady=8)
            ttk.Button(btnfrm, text="新增分组规则", command=self.add_group_rule).pack(side="left", padx=4)
            ttk.Button(btnfrm, text="编辑选中", command=self.edit_group_rule).pack(side="left", padx=4)
            ttk.Button(btnfrm, text="删除选中", command=self.delete_group_rule).pack(side="left", padx=4)
            ttk.Button(btnfrm, text="保存配置", command=self.save_all_config).pack(side="left", padx=4)

            tip = ttk.Label(
                self.tab_groups,
                text="分组规则按最终输出款号匹配。例：MG- 开头归到 MG系列，SSL- 开头归到 SSL系列。输出 Excel 会先按分组归类，再按款号统计，并在每个款号之间空一行。",
                foreground="#555",
                wraplength=1050,
                justify="left",
            )
            tip.pack(fill="x", padx=10, pady=(0, 6))

            cols = ("enabled", "priority", "name", "match_type", "pattern", "group")
            self.group_tree = ttk.Treeview(self.tab_groups, columns=cols, show="headings")
            headings = {
                "enabled": "启用", "priority": "优先级", "name": "规则名", "match_type": "匹配方式",
                "pattern": "匹配款号", "group": "分组名称"
            }
            widths = {"enabled": 60, "priority": 80, "name": 260, "match_type": 120, "pattern": 260, "group": 180}
            for col in cols:
                self.group_tree.heading(col, text=headings[col])
                self.group_tree.column(col, width=widths[col], anchor="center")
            self.group_tree.pack(fill="both", expand=True, padx=10, pady=5)
            self.group_tree.bind("<Double-1>", lambda e: self.edit_group_rule())

        def _build_color_tab(self):
            main = ttk.Frame(self.tab_color)
            main.pack(fill="both", expand=True, padx=10, pady=10)

            left = ttk.LabelFrame(main, text="英文颜色 → 中文颜色")
            left.pack(side="left", fill="both", expand=True, padx=(0, 6))
            self.color_tree = ttk.Treeview(left, columns=("en", "cn"), show="headings")
            self.color_tree.heading("en", text="英文颜色")
            self.color_tree.heading("cn", text="中文颜色")
            self.color_tree.column("en", width=160, anchor="center")
            self.color_tree.column("cn", width=160, anchor="center")
            self.color_tree.pack(fill="both", expand=True, padx=6, pady=6)
            cbtn = ttk.Frame(left)
            cbtn.pack(fill="x", padx=6, pady=6)
            ttk.Button(cbtn, text="新增/修改颜色", command=self.edit_color).pack(side="left", padx=3)
            ttk.Button(cbtn, text="删除颜色", command=self.delete_color).pack(side="left", padx=3)

            right = ttk.LabelFrame(main, text="款号前缀替换")
            right.pack(side="left", fill="both", expand=True, padx=(6, 0))
            self.prefix_tree = ttk.Treeview(right, columns=("enabled", "from", "to"), show="headings")
            self.prefix_tree.heading("enabled", text="启用")
            self.prefix_tree.heading("from", text="原前缀")
            self.prefix_tree.heading("to", text="替换为")
            self.prefix_tree.column("enabled", width=60, anchor="center")
            self.prefix_tree.column("from", width=160, anchor="center")
            self.prefix_tree.column("to", width=160, anchor="center")
            self.prefix_tree.pack(fill="both", expand=True, padx=6, pady=6)
            pbtn = ttk.Frame(right)
            pbtn.pack(fill="x", padx=6, pady=6)
            ttk.Button(pbtn, text="新增/修改替换", command=self.edit_prefix).pack(side="left", padx=3)
            ttk.Button(pbtn, text="删除替换", command=self.delete_prefix).pack(side="left", padx=3)
            ttk.Button(pbtn, text="保存配置", command=self.save_all_config).pack(side="left", padx=3)

        def add_log(self, text: str):
            self.log.insert("end", text + "\n")
            self.log.see("end")
            self.update_idletasks()

        def add_pdfs(self):
            paths = filedialog.askopenfilenames(title="选择PDF", filetypes=[("PDF文件", "*.pdf")])
            self._add_pdf_paths([Path(p) for p in paths])

        def add_folder(self):
            folder = filedialog.askdirectory(title="选择PDF文件夹")
            if folder:
                self._add_pdf_paths(collect_pdf_paths([Path(folder)]))

        def _add_pdf_paths(self, paths: Iterable[Path]):
            existing = {str(p.resolve()) for p in self.pdf_paths}
            for p in paths:
                if p.exists() and p.suffix.lower() == ".pdf" and str(p.resolve()) not in existing:
                    self.pdf_paths.append(p)
                    existing.add(str(p.resolve()))
            self.refresh_pdf_tree()

        def refresh_pdf_tree(self):
            for i in self.pdf_tree.get_children():
                self.pdf_tree.delete(i)
            for p in self.pdf_paths:
                self.pdf_tree.insert("", "end", values=(str(p),))

        def remove_selected_pdf(self):
            selected = self.pdf_tree.selection()
            if not selected:
                return
            selected_paths = {self.pdf_tree.item(i, "values")[0] for i in selected}
            self.pdf_paths = [p for p in self.pdf_paths if str(p) not in selected_paths]
            self.refresh_pdf_tree()

        def clear_pdfs(self):
            self.pdf_paths = []
            self.refresh_pdf_tree()

        def choose_output(self):
            path = filedialog.asksaveasfilename(
                title="保存输出Excel",
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx")],
                initialfile="报货汇总.xlsx",
            )
            if path:
                self.output_var.set(path)

        def run_process(self):
            if not self.pdf_paths:
                messagebox.showwarning("提示", "请先添加至少一个 PDF。")
                return
            out_path = Path(self.output_var.get().strip())
            if not out_path:
                messagebox.showwarning("提示", "请设置输出 Excel 路径。")
                return
            try:
                self.save_all_config(show_msg=False)
                self.add_log("开始处理...")
                result = process_pdfs(self.pdf_paths, out_path, self.cfg)
                self.add_log(f"完成：{result['out_path']}")
                self.add_log(f"PDF数量：{result['pdf_count']}")
                self.add_log(f"解析行数：{result['parsed_rows']}")
                self.add_log(f"PDF原始数量合计：{result['original_qty']}")
                self.add_log(f"输出汇总款数：{result['output_lines']}款")
                self.add_log(f"输出汇总数量：{result['output_qty']}")
                self.add_log(f"异常数量：{result['issues']}")
                messagebox.showinfo("完成", f"已生成：\n{result['out_path']}")
            except Exception as e:
                self.add_log("处理失败：" + str(e))
                self.add_log(traceback.format_exc())
                messagebox.showerror("处理失败", str(e))

        def refresh_rules_tree(self):
            for i in self.rules_tree.get_children():
                self.rules_tree.delete(i)
            for idx, rule in sorted_indexed_rules(self.cfg):
                self.rules_tree.insert("", "end", iid=f"rule:{idx}", values=(
                    "是" if rule.get("enabled", True) else "否",
                    rule.get("priority", ""),
                    rule.get("name", ""),
                    rule.get("group_name", ""),
                    rule.get("match_type", ""),
                    rule.get("pattern", ""),
                    rule.get("output_styles", ""),
                    rule.get("color_mode", ""),
                    rule.get("fixed_colors", ""),
                    rule.get("qty_multiplier", ""),
                    rule.get("note", ""),
                ))

        def _selected_rule_actual_index(self) -> Optional[int]:
            sel = self.rules_tree.selection()
            if not sel:
                return None
            try:
                idx = int(str(sel[0]).split(":", 1)[1])
            except (IndexError, ValueError):
                return None
            return idx if 0 <= idx < len(self.cfg.get("rules", [])) else None

        def add_rule(self):
            dlg = RuleDialog(self)
            self.wait_window(dlg)
            if dlg.result:
                self.cfg.setdefault("rules", []).append(dlg.result)
                self.refresh_rules_tree()

        def edit_rule(self):
            idx = self._selected_rule_actual_index()
            if idx is None:
                messagebox.showwarning("提示", "请先选择一条规则。")
                return
            dlg = RuleDialog(self, dict(self.cfg["rules"][idx]))
            self.wait_window(dlg)
            if dlg.result:
                self.cfg["rules"][idx] = dlg.result
                self.refresh_rules_tree()

        def delete_rule(self):
            idx = self._selected_rule_actual_index()
            if idx is None:
                return
            if messagebox.askyesno("确认", "确定删除选中的规则吗？"):
                self.cfg["rules"].pop(idx)
                self.refresh_rules_tree()

        def refresh_group_rules_tree(self):
            for i in self.group_tree.get_children():
                self.group_tree.delete(i)
            for idx, rule in sorted_indexed_group_rules(self.cfg):
                self.group_tree.insert("", "end", iid=f"group:{idx}", values=(
                    "是" if rule.get("enabled", True) else "否",
                    rule.get("priority", ""),
                    rule.get("name", ""),
                    rule.get("match_type", ""),
                    rule.get("pattern", ""),
                    rule.get("group_name", ""),
                ))

        def _selected_group_rule_actual_index(self) -> Optional[int]:
            sel = self.group_tree.selection()
            if not sel:
                return None
            try:
                idx = int(str(sel[0]).split(":", 1)[1])
            except (IndexError, ValueError):
                return None
            return idx if 0 <= idx < len(self.cfg.get("group_rules", [])) else None

        def add_group_rule(self):
            dlg = GroupRuleDialog(self)
            self.wait_window(dlg)
            if dlg.result:
                self.cfg.setdefault("group_rules", []).append(dlg.result)
                self.refresh_group_rules_tree()

        def edit_group_rule(self):
            idx = self._selected_group_rule_actual_index()
            if idx is None:
                messagebox.showwarning("提示", "请先选择一条分组规则。")
                return
            dlg = GroupRuleDialog(self, dict(self.cfg["group_rules"][idx]))
            self.wait_window(dlg)
            if dlg.result:
                self.cfg["group_rules"][idx] = dlg.result
                self.refresh_group_rules_tree()

        def delete_group_rule(self):
            idx = self._selected_group_rule_actual_index()
            if idx is None:
                return
            if messagebox.askyesno("确认", "确定删除选中的分组规则吗？"):
                self.cfg["group_rules"].pop(idx)
                self.refresh_group_rules_tree()

        def refresh_color_table(self):
            for i in self.color_tree.get_children():
                self.color_tree.delete(i)
            for en, cn in sorted(self.cfg.get("color_map", {}).items()):
                self.color_tree.insert("", "end", values=(en, cn))

        def edit_color(self):
            sel = self.color_tree.selection()
            old_en = old_cn = ""
            if sel:
                vals = self.color_tree.item(sel[0], "values")
                old_en, old_cn = vals[0], vals[1]
            win = tk.Toplevel(self)
            win.title("颜色映射")
            win.resizable(False, False)
            en_var = tk.StringVar(value=old_en)
            cn_var = tk.StringVar(value=old_cn)
            ttk.Label(win, text="英文颜色").grid(row=0, column=0, padx=8, pady=6, sticky="e")
            ttk.Entry(win, textvariable=en_var, width=28).grid(row=0, column=1, padx=8, pady=6)
            ttk.Label(win, text="中文颜色").grid(row=1, column=0, padx=8, pady=6, sticky="e")
            ttk.Entry(win, textvariable=cn_var, width=28).grid(row=1, column=1, padx=8, pady=6)
            def ok():
                en = en_var.get().strip()
                cn = cn_var.get().strip()
                if not en or not cn:
                    messagebox.showerror("错误", "英文颜色和中文颜色不能为空。")
                    return
                if old_en and old_en != en:
                    self.cfg["color_map"].pop(old_en, None)
                self.cfg.setdefault("color_map", {})[en] = cn
                self.refresh_color_table()
                win.destroy()
            ttk.Button(win, text="保存", command=ok).grid(row=2, column=0, columnspan=2, pady=8)
            win.grab_set()

        def delete_color(self):
            sel = self.color_tree.selection()
            if not sel:
                return
            en = self.color_tree.item(sel[0], "values")[0]
            self.cfg.get("color_map", {}).pop(en, None)
            self.refresh_color_table()

        def refresh_prefix_table(self):
            for i in self.prefix_tree.get_children():
                self.prefix_tree.delete(i)
            for item in self.cfg.get("general", {}).get("style_prefix_replace", []):
                self.prefix_tree.insert("", "end", values=("是" if item.get("enabled", True) else "否", item.get("from", ""), item.get("to", "")))

        def edit_prefix(self):
            sel = self.prefix_tree.selection()
            actual_idx = None
            enabled = True
            old_from = old_to = ""
            if sel:
                vals = self.prefix_tree.item(sel[0], "values")
                enabled = vals[0] == "是"
                old_from, old_to = vals[1], vals[2]
                for idx, item in enumerate(self.cfg.get("general", {}).get("style_prefix_replace", [])):
                    if item.get("from", "") == old_from and item.get("to", "") == old_to:
                        actual_idx = idx
                        break
            win = tk.Toplevel(self)
            win.title("款号前缀替换")
            win.resizable(False, False)
            enabled_var = tk.BooleanVar(value=enabled)
            from_var = tk.StringVar(value=old_from)
            to_var = tk.StringVar(value=old_to)
            ttk.Checkbutton(win, text="启用", variable=enabled_var).grid(row=0, column=0, columnspan=2, padx=8, pady=6, sticky="w")
            ttk.Label(win, text="原前缀").grid(row=1, column=0, padx=8, pady=6, sticky="e")
            ttk.Entry(win, textvariable=from_var, width=28).grid(row=1, column=1, padx=8, pady=6)
            ttk.Label(win, text="替换为").grid(row=2, column=0, padx=8, pady=6, sticky="e")
            ttk.Entry(win, textvariable=to_var, width=28).grid(row=2, column=1, padx=8, pady=6)
            def ok():
                src = from_var.get().strip()
                dst = to_var.get().strip()
                if not src:
                    messagebox.showerror("错误", "原前缀不能为空。")
                    return
                item = {"enabled": bool(enabled_var.get()), "from": src, "to": dst}
                lst = self.cfg.setdefault("general", {}).setdefault("style_prefix_replace", [])
                if actual_idx is None:
                    lst.append(item)
                else:
                    lst[actual_idx] = item
                self.refresh_prefix_table()
                win.destroy()
            ttk.Button(win, text="保存", command=ok).grid(row=3, column=0, columnspan=2, pady=8)
            win.grab_set()

        def delete_prefix(self):
            sel = self.prefix_tree.selection()
            if not sel:
                return
            vals = self.prefix_tree.item(sel[0], "values")
            old_from, old_to = vals[1], vals[2]
            lst = self.cfg.get("general", {}).get("style_prefix_replace", [])
            self.cfg["general"]["style_prefix_replace"] = [x for x in lst if not (x.get("from") == old_from and x.get("to") == old_to)]
            self.refresh_prefix_table()

        def save_all_config(self, show_msg: bool = True):
            save_config(self.cfg, self.cfg_path)
            if show_msg:
                messagebox.showinfo("已保存", f"配置已保存到：\n{self.cfg_path}")

        def import_config(self):
            path = filedialog.askopenfilename(title="导入配置", filetypes=[("JSON配置", "*.json")])
            if path:
                self.cfg = load_config(Path(path))
                self.refresh_rules_tree()
                self.refresh_group_rules_tree()
                self.refresh_color_table()
                self.refresh_prefix_table()
                messagebox.showinfo("完成", "配置已导入。记得点击保存规则。")

        def export_config(self):
            path = filedialog.asksaveasfilename(title="导出配置", defaultextension=".json", filetypes=[("JSON配置", "*.json")])
            if path:
                save_config(self.cfg, Path(path))
                messagebox.showinfo("完成", f"配置已导出：\n{path}")

        def reset_config(self):
            if messagebox.askyesno("确认", "确定恢复默认配置吗？当前未保存的规则会被覆盖。"):
                self.cfg = deep_copy_default_config()
                self.refresh_rules_tree()
                self.refresh_group_rules_tree()
                self.refresh_color_table()
                self.refresh_prefix_table()


def run_cli(args: argparse.Namespace) -> None:
    cfg_path = Path(args.config) if args.config else default_config_path()
    cfg = load_config(cfg_path)
    inputs = [Path(x) for x in args.inputs]
    pdfs = collect_pdf_paths(inputs)
    if not pdfs:
        raise SystemExit("没有找到 PDF 文件。")

    out_arg = Path(args.output) if args.output else None
    if out_arg is None:
        out_path = pdfs[0].with_name(pdfs[0].stem + "_报货汇总.xlsx") if len(pdfs) == 1 else Path.cwd() / "报货汇总.xlsx"
    elif out_arg.suffix.lower() == ".xlsx":
        out_path = out_arg
    else:
        out_arg.mkdir(parents=True, exist_ok=True)
        out_path = out_arg / (pdfs[0].stem + "_报货汇总.xlsx" if len(pdfs) == 1 else "报货汇总.xlsx")

    result = process_pdfs(pdfs, out_path, cfg)
    print("完成：", result["out_path"])
    print("PDF数量：", result["pdf_count"])
    print("解析行数：", result["parsed_rows"])
    print("PDF原始数量合计：", result["original_qty"])
    print("输出汇总款数：", result["output_lines"], "款")
    print("输出汇总数量：", result["output_qty"])
    print("异常数量：", result["issues"])


def main():
    parser = argparse.ArgumentParser(description="报货 PDF 自动汇总工具")
    parser.add_argument("--cli", action="store_true", help="使用命令行模式，不打开 GUI")
    parser.add_argument("inputs", nargs="*", help="PDF 文件或包含 PDF 的文件夹")
    parser.add_argument("-o", "--output", help="输出 Excel 文件或输出目录")
    parser.add_argument("--config", help="配置 JSON 路径")
    args = parser.parse_args()

    if args.cli:
        run_cli(args)
        return

    if tk is None:
        raise SystemExit("当前 Python 环境没有 tkinter，无法启动 GUI。可以使用 --cli 命令行模式。")
    app = BaohuoApp()
    app.mainloop()


if __name__ == "__main__":
    main()
