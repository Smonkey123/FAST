import pdfplumber
import pandas as pd
import os
import PyPDF2
import re

RULE_FILE = 'J:/Engineering/ShareFolder/new_ABB_Production_Tools/Pd/document/LVD_rule.txt'
LVD_SPECIFIC_FILE = 'J:/Engineering/ShareFolder/new_ABB_Production_Tools/Pd/document/LVD_specific.txt'

# def extract_device_from_pdf(pdf_path, sheet_page_number):
#     results = []
#     with pdfplumber.open(pdf_path) as pdf:
#         for page_number, page in enumerate(pdf.pages):
#             if page_number in sheet_page_number:
#                 device_list, typical = extract_device_method(page)
#                 results.append((page_number, device_list, typical))
#     return results

def get_rule_by_drawing(cell):
    """返回 (rule_name, typical)"""
    text = str(cell).upper()
    for pattern, rule in (
        ('XZX0-', 'ZX0_rule'),
        ('XPZX0-', 'ZX0_rule'),
        ('XZS1-', 'ZS1_rule'),
        ('XZS3-', 'ZS3_rule'),
        ('XZS3.2-', 'ZS3_rule'),
        ('XZX1.5-R-', 'ZX1.5_rule'),
        ('XZX1.2-', 'ZX1.2_rule'),
        ('XZX2-', 'ZX2_rule'),
        ('XZX0.2-', 'ZX0.2_rule'),
        ('XZVC-', 'ZVC_rule'),
        ('X500R-', '500R_rule'),
        ('XZS9-', 'ZS9_rule'),
    ):
        if pattern in text:
            idx_last = text.rfind('-')
            if idx_last == -1:
                return rule, ''
            idx_prev = text.rfind('-', 0, idx_last)
            if idx_prev == -1:
                return rule, ''
            return rule, text[idx_prev + 1: idx_last]
    return None, ''

def _parse_fraction(expr):
    """把 '1/8' 或 '0.125' 转 float"""
    try:
        return float(expr)
    except ValueError:
        if '/' in expr:
            num, den = expr.split('/')
            return float(num) / float(den)
        raise

def load_rules():
    """返回 dict: {rule_name: (x0, top, x1, bottom)}；空规则返回 None"""
    rules = {}
    if not os.path.exists(RULE_FILE):
        return rules
    with open(RULE_FILE, encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or ':' not in line:
                # 无冒号 → 空规则
                rule_name = line.split(':')[0].strip()
                rules[rule_name] = None
                continue
            rule_name, params = line.split(':', 1)
            rule_name = rule_name.strip()
            if not params.strip():
                rules[rule_name] = None
                continue
            # 解析 "x0:XX,top:XX,x1:XX,bottom:XX"
            kv = {}
            for pair in params.split(','):
                k, v = pair.split(':', 1)
                kv[k.strip()] = v.strip()
            bbox = tuple(_parse_fraction(kv[k]) for k in ('x0', 'top', 'x1', 'bottom'))
            rules[rule_name] = bbox
    return rules

def load_lvd_blacklist():
    """返回 set：所有需要过滤的 -xxx 字符串"""
    blacklist = set()
    if not os.path.exists(LVD_SPECIFIC_FILE):
        return blacklist
    with open(LVD_SPECIFIC_FILE, encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line.startswith('-') and ' ' not in line:
                blacklist.add(line)
    return blacklist

_RULE_CACHE = load_rules()
_LVD_BLACKLIST = load_lvd_blacklist()

def get_rule_bbox(rule_flag, w, h):
    """返回绝对 bbox；若规则不存在或为空，返回 None"""
    bbox = _RULE_CACHE.get(rule_flag)
    if bbox is None:
        return None
    x0, top, x1, bottom = bbox
    return (x0 * w, top * h, x1 * w, bottom * h)

def apply_rule(page, rule_flag, w, h):
    bbox = get_rule_bbox(rule_flag, w, h)
    if bbox is None:
        print(f'规则 {rule_flag} 未配置或配置文件丢失，请联系管理员！')
        return []

    cropped_page = page.crop(bbox)
    result, result_position = [], []

    for ch in cropped_page.extract_words(x_tolerance=1, y_tolerance=0, vertical_ttb=False, keep_blank_chars=True):
        text = ch['text']
        if not text.startswith('-'):
            continue

        dash_cnt = text.count('-')
        if dash_cnt < 2:
            parts = [text]  # 统一用列表，便于后续过滤
        else:
            parts = list(filter(None, re.split(r'(?=-)', text)))

        for p in parts:
            if p not in _LVD_BLACKLIST:  # 关键：过滤黑名单
                result.append(p)
                result_position.append([ch['x0'], ch['top'], ch['x1'], ch['bottom']])
    # print(result)
    return result

def extract_device_from_pdf(pdf_path):
    """
    一键式接口：
    1. 读取 PDF 大纲，自动计算 sheet_page_number
    2. 提取每页的 device 列表 和 typical
    返回 list[(page_index, device_list, typical)]
    """
    # ---------- 1. 计算 sheet_page_number ----------
    sheet_page_number = []
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        outlines = pdf_reader.outline
        # 若 outlines 层级不足 4 级，请根据实际 PDF 调整索引
        if len(outlines) > 3:
            for index, item in enumerate(outlines[3]):
                title = str(item.get("/Title", ""))
                if '&C/' in title and ("门板开孔图" in title or "DOOR LAYOUT" in title.upper()):
                    sheet_page_number.append(index)

    # ---------- 2. 提取 device + typical ----------
    results = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_index, page in enumerate(pdf.pages):
            if page_index not in sheet_page_number:
                continue

            w, h = page.width, page.height
            crop_w, crop_h = w / 7, h / 11
            bbox = (w - crop_w, h - crop_h, w, h)
            cropped_page = page.crop(bbox)

            table_settings = {
                "vertical_strategy": "lines_strict",
                "horizontal_strategy": "lines_strict",
                "explicit_vertical_lines": [],
                "explicit_horizontal_lines": [],
                "min_words_vertical": 0,
            }

            page_devices = []
            page_typical = ''
            tables = cropped_page.extract_tables(table_settings)

            for table in tables:
                for row in table:
                    for cell in row:
                        if not cell:
                            continue
                        rule_name, typical = get_rule_by_drawing(cell)
                        if rule_name:
                            if not page_typical:           # 仅保留第一条
                                page_typical = typical
                            devices = apply_rule(page, rule_name, w, h)
                            page_devices.extend(devices)

            # # 去重并保持顺序（可选）
            # seen = set()
            # page_devices = [d for d in page_devices if not (d in seen or seen.add(d))]
            results.append((page_index, page_devices, page_typical))

    return results


if __name__ == "__main__":
    # 使用示例
    pdf_path = '504556032.pdf'  # 请替换为你的PDF文件路径

    all_info = extract_device_from_pdf(pdf_path)

    for page_idx, devs, typ in all_info:
        print(f'页 {page_idx}  typical={typ}  devices={devs}')