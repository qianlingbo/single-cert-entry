#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
单证录入核心脚本 v3 — 直接 XML 操作（绕过 openpyxl）
================================================================
v2 之前用 openpyxl 加载 .xlsm 模板再保存，导致：
  - vbaProject.bin 丢失（50KB VBA 宏被丢弃）
  - sheet 1/2/3 被压缩（编号重排）
  - 数据验证扩展（dataValidation）被移除
  - sharedStrings.xml、customXml 等丢失
  → Excel 弹出"修复"对话框 → 用户体感"无法打开"

v3 方案：
  1. 复制模板 zip 整体（保留 vbaProject.bin / 所有 sheet）
  2. 只修改需要填充的 4 个 sheet 的 XML（按字符串插入 <row>）
  3. 重新打包 zip 输出

用法:
    python3 scripts/单证录入核心_v3.py
    python3 scripts/单证录入核心_v3.py <crew.xlsx> <port.xlsx> [output.xlsm]
"""

import os
import re
import sys
import json
import random
import shutil
import zipfile
import io
from datetime import datetime


WORKDIR = "/Users/qianlingbo/单证录入工作区"
INPUT_DIR = os.path.join(WORKDIR, "input")
OUTPUT_DIR = os.path.join(WORKDIR, "output")
TEMPLATES_DIR = os.path.join(WORKDIR, "templates")
SCRIPTS_DIR = os.path.join(WORKDIR, "scripts")

DEFAULT_CREW = os.path.join(INPUT_DIR, "CREW_LIST.xlsx")
DEFAULT_POC = os.path.join(INPUT_DIR, "PORT_OF_CALL.xlsx")
TEMPLATE_XLSM = os.path.join(TEMPLATES_DIR, "标准格式.xlsm")
DEFAULT_OUTPUT = os.path.join(OUTPUT_DIR, "单证录入标准格式.xlsm")


# ============================================================
# 字段映射（与 v2 相同）
# ============================================================

NAT_MAP = {
    "CHINESE": "中国", "CHINA": "中国", "CN": "中国",
    "VIETNAMESE": "越南", "VIETNAM": "越南", "VN": "越南",
    "MYANMAR": "缅甸", "BURMESE": "缅甸", "MM": "缅甸",
    "INDONESIAN": "印度尼西亚", "INDONESIA": "印度尼西亚", "ID": "印度尼西亚",
    "PHILIPPINE": "菲律宾", "PHILIPPINES": "菲律宾", "PH": "菲律宾",
    "PANAMANIAN": "巴拿马", "PANAMA": "巴拿马", "PA": "巴拿马",
    "INDIAN": "印度", "INDIA": "印度", "IN": "印度",
}

NAT_CODE = {
    "中国": "CN", "越南": "VN", "缅甸": "MM", "印度尼西亚": "ID",
    "菲律宾": "PH", "巴拿马": "PA", "印度": "IN",
}

DUTY_CODE = {
    "MASTER": "51", "CAPT": "51", "CAPTAIN": "51",
    "C/O": "52", "C.O.": "52", "CHIEF OFFICER": "52",
    "2/O": "53", "2ND OFFICER": "53",
    "3/O": "54", "3RD OFFICER": "54",
    "OS": "55", "OLR": "55", "ORDINARY SEAMAN": "55", "C/CK": "55",
    "AB": "55", "ABLE SEAMAN": "55",
    "BOSUN": "56", "BSN": "56", "BOATSWAIN": "56",
    "C/E": "61", "CHIEF ENGINEER": "61",
    "2/E": "63", "2ND ENGINEER": "63",
    "3/E": "64", "3RD ENGINEER": "64",
    "ETR": "65", "FTR": "65", "FITTER": "65", "OILER": "65",
    "ELECTRICIAN": "65", "OIL1": "65", "OIL2": "65",
}

PORT_CODE = {
    # 中国
    "ZHOUSHAN": "CNZOS", "舟山": "CNZOS",
    "ZHANGJIAGANG": "CNZJG", "张家港": "CNZJG",
    "CHANGZHOU": "CNCZX", "常州": "CNCZX",
    "TIANJIN": "CNTXG", "天津": "CNTXG",
    "LANSHAN": "CNLSN", "岚山": "CNLSN",
    "BAYUQUAN": "CNBYQ", "鲅鱼圈": "CNBYQ",
    "CAOFEIDIAN": "CNCFD", "曹妃甸": "CNCFD",
    "LIANYUNGANG": "CNLYG", "连云港": "CNLYG",
    "NINGDE": "CNNDS", "宁德": "CNNDS",
    "SHANGHAI": "CNSHA", "上海": "CNSHA",
    "NINGBO": "CNNGB", "宁波": "CNNGB",
    "QINGDAO": "CNTAO", "青岛": "CNTAO",
    "RIZHAO": "CNRZH", "日照": "CNRZH",
    # 印尼
    "MOROWALI": "IDMOR",
    "POMALAA": "IDPUM",
    "OBI ISLAND": "IDOBI", "OBI": "IDOBI",
    # 公海
    "OPEN SEA": "OPSEA", "OPENSEA": "OPSEA", "OPSEA": "OPSEA",
    # 国际
    "HITACHINAKA": "JPHIC",
}


# ============================================================
# 解析输入（与 v2 相同）
# ============================================================

def normalize_date(s):
    """宽松日期解析"""
    if s is None: return None
    if isinstance(s, datetime): return s
    s = str(s).strip()
    s = re.sub(r'\s+', ' ', s)
    s = s.replace(' :', ':').replace(': ', ':')
    s = re.sub(r'(\d{4})/(\d{1,2})/ ?(\d{1,2})', r'\1/\2/\3', s)
    fmts = [
        "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d", "%Y/%m/%d", "%Y%m%d",
        "%d/%m/%Y", "%m/%d/%Y",
    ]
    for f in fmts:
        try: return datetime.strptime(s, f)
        except ValueError: continue
    return None


def parse_crew_xlsx(path):
    """解析 IMO Crew List"""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["ARR"] if "ARR" in wb.sheetnames else wb[wb.sheetnames[0]]

    # 找标题行
    header_row = None
    for r in range(1, 15):
        cells = [str(c.value or '').upper() for c in ws[r][:15]]
        if 'FAMILY NAME' in cells or 'NAME' in cells:
            header_row = r
            break
    if header_row is None: header_row = 7

    # 找列索引
    headers = [str(ws.cell(header_row, c).value or '').upper() for c in range(1, 20)]
    col = {}
    for kw, idx_name in [('NO.', 'no'), ('NAME', 'name'), ('SEX', 'sex'),
                          ('RANK', 'rank'), ('BIRTH', 'birth'), ('NATIONAL', 'nat'),
                          ('PLACE', 'place'), ('SEAMAN', 'seaman'), ('PASSPORT', 'passport'),
                          ('SIGNED ON', 'signedon')]:
        for i, h in enumerate(headers):
            if kw in h and idx_name not in col:
                col[idx_name] = i + 1  # 1-indexed

    crews = []
    for r in range(header_row + 1, ws.max_row + 1):
        no = ws.cell(r, col.get('no', 1)).value
        if no is None: continue
        # 跳过中文行
        nm = str(ws.cell(r, col.get('name', 2)).value or '')
        if not re.search(r'[A-Z]', nm): continue
        try:
            no_int = int(float(str(no).strip()))
        except: continue

        rank = str(ws.cell(r, col.get('rank', 5)).value or '').strip().upper()
        nat = str(ws.cell(r, col.get('nat', 6)).value or '').strip().upper()
        seaman = str(ws.cell(r, col.get('seaman', 8)).value or '').strip()
        passport = str(ws.cell(r, col.get('passport', 10)).value or '').strip()
        bd = ws.cell(r, col.get('birth', 5)).value
        so = ws.cell(r, col.get('signedon', 11)).value
        so_place = str(ws.cell(r, col.get('signedon', 11) + 1).value or '').strip() if col.get('signedon') else ''

        crews.append({
            'no': no_int,
            'name_en': nm,
            'sex': '1' if 'M' in str(ws.cell(r, col.get('sex', 3)).value or '').upper() else '2',
            'rank': rank,
            'birth': normalize_date(bd),
            'nat_en': nat,
            'seaman_no': seaman,
            'passport_no': passport,
            'signon_date': normalize_date(so),
            'signon_place': so_place,
        })
    return crews


def parse_poc_xlsx(path):
    """解析 Port of Call"""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb[wb.sheetnames[0]]

    # 找表头行
    header_row = None
    for r in range(1, 10):
        cells = [str(c.value or '').upper() for c in ws[r][:10]]
        if 'PORT' in cells and 'COUNTRY' in cells:
            header_row = r
            break
    if header_row is None: header_row = 3

    headers = [str(ws.cell(header_row, c).value or '').upper() for c in range(1, 12)]
    col = {}
    for i, h in enumerate(headers):
        hu = h.upper()
        if 'NO.' in hu and 'no' not in col: col['no'] = i + 1
        elif 'PURPOSE' in hu and 'purpose' not in col: col['purpose'] = i + 1
        elif hu == 'PORT' and 'port' not in col: col['port'] = i + 1
        elif 'COUNTRY' in hu and 'country' not in col: col['country'] = i + 1
        elif 'ARRIVAL' in hu and 'arr' not in col: col['arr'] = i + 1
        elif 'DEPARTURE' in hu and 'dep' not in col: col['dep'] = i + 1
        elif 'SHIP SECURITY' in hu and 'ssl' not in col: col['ssl'] = i + 1
        elif 'PORT' in hu and 'SECURITY' in hu and 'psl' not in col: col['psl'] = i + 1

    records = []
    for r in range(header_row + 1, ws.max_row + 1):
        no = ws.cell(r, col.get('no', 1)).value
        port = ws.cell(r, col.get('port', 3)).value
        if no is None or port is None: continue
        try:
            no_int = int(float(str(no).strip()))
        except: continue
        records.append({
            'no': no_int,
            'purpose': str(ws.cell(r, col.get('purpose', 2)).value or '').strip(),
            'port': str(port).strip(),
            'country': str(ws.cell(r, col.get('country', 4)).value or '').strip().upper(),
            'arr': normalize_date(ws.cell(r, col.get('arr', 5)).value),
            'dep': normalize_date(ws.cell(r, col.get('dep', 6)).value),
            'ssl': str(ws.cell(r, col.get('ssl', 7)).value or '1').strip(),
            'psl': str(ws.cell(r, col.get('psl', 8)).value or '1').strip(),
        })
    # 序号重排
    for i, rec in enumerate(records, 1):
        rec['no'] = i
    return records


# ============================================================
# XML 工具
# ============================================================

def cell_ref(col_letter, row):
    return f"{col_letter}{row}"


def escape_xml(s):
    if s is None: return ''
    s = str(s)
    return (s.replace('&', '&amp;')
             .replace('<', '&lt;')
             .replace('>', '&gt;')
             .replace('"', '&quot;')
             .replace("'", '&apos;'))


def build_row(row_num, values, start_col='A', style=None):
    """生成 <row> XML 字符串
    
    values: list of (col_letter, value, [type])
      type: 's' sharedString, 'str' inline string, 'n' number, default str
    """
    cells_xml = []
    for item in values:
        col = item[0]
        val = item[1]
        vtype = item[2] if len(item) > 2 else 'str'
        ref = f"{col}{row_num}"
        if vtype == 'n':
            cells_xml.append(f'<c r="{ref}"><v>{val}</v></c>')
        elif vtype == 's':
            cells_xml.append(f'<c r="{ref}" t="s"><v>{val}</v></c>')
        else:
            cells_xml.append(f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{escape_xml(val)}</t></is></c>')
    style_attr = f' s="{style}"' if style else ''
    return f'<row r="{row_num}"{style_attr}>' + ''.join(cells_xml) + '</row>'


def insert_rows_into_sheet_xml(sheet_xml, new_rows_xml, after_row):
    """在指定行号之后插入新行 XML"""
    # 找到 </row> 在 after_row 后面的位置
    pattern = rf'(<row r="{after_row}"[^>]*>.*?</row>)'
    match = re.search(pattern, sheet_xml, re.DOTALL)
    if not match:
        # 找不到精确行 — 尝试找所有 <row> 然后定位
        all_rows = list(re.finditer(r'<row r="(\d+)"', sheet_xml))
        if not all_rows:
            return sheet_xml.replace('</sheetData>', '<sheetData>' + ''.join(new_rows_xml) + '</sheetData>')
        # 找位置
        target_idx = 0
        for i, m in enumerate(all_rows):
            if int(m.group(1)) > after_row:
                target_idx = i
                break
            target_idx = i + 1
        if target_idx == 0:
            insert_pos = all_rows[0].start()
        else:
            # 在 target_idx 之前插入（即在它前面）
            insert_pos = all_rows[target_idx].start()
    else:
        insert_pos = match.end()
    return sheet_xml[:insert_pos] + ''.join(new_rows_xml) + sheet_xml[insert_pos:]


# ============================================================
# 数据准备
# ============================================================

def build_crew_rows(crews, shared_strings):
    """生成船员名单行（船上非旅客人员清单）"""
    rows = []
    for i, c in enumerate(crews, 3):  # 从 R3 开始
        # 中文名（如果有）或英文
        name = c['name_en']  # 简化: v3 暂只填英文
        nat_cn = NAT_MAP.get(c['nat_en'].upper(), c['nat_en'])
        nat_code = NAT_CODE.get(nat_cn, '')
        rank = DUTY_CODE.get(c['rank'].upper(), '55')
        is_cn = (nat_cn == '中国')
        cert_type = '17-海员证' if is_cn else '14-普通护照'
        cert_no = c['seaman_no'] if is_cn else c['passport_no']
        birth_str = c['birth'].strftime('%Y%m%d') if c['birth'] else ''
        signon_str = c['signon_date'].strftime('%Y%m%d') if c['signon_date'] else ''
        signon_port = PORT_CODE.get(c['signon_place'].upper(), '')

        row = build_row(i, [
            ('A', i - 2),                # 序号
            ('B', name),                 # 姓名
            ('C', f"{c['sex']}-{'男' if c['sex']=='1' else '女'}"),  # 性别
            ('D', f"{rank}-{rank_name(rank)}"),  # 职务
            ('E', f"{nat_code}-{nat_cn}"),  # 国籍
            ('F', birth_str),            # 出生日期
            ('G', nat_cn),               # 出生地点
            ('H', cert_type),            # 证件类型
            ('I', cert_no),              # 证件号码
            ('M', signon_str),           # 登船日期
            ('N', signon_port),          # 登船口岸
        ])
        rows.append(row)
    return rows


def rank_name(code):
    return {
        '51': '船长', '52': '大副', '53': '二副', '54': '三副',
        '55': '值班水手', '56': '高级值班水手',
        '61': '轮机长', '62': '大管轮', '63': '二管轮', '64': '三管轮',
        '65': '值班机工', '66': '高级值班机工',
    }.get(code, '值班水手')


def build_goods_rows(crews):
    rows = []
    for i, c in enumerate(crews, 3):
        nat_cn = NAT_MAP.get(c['nat_en'].upper(), c['nat_en'])
        is_cn = (nat_cn == '中国')
        cert_type = '17-海员证' if is_cn else '14-普通护照'
        cert_no = c['seaman_no'] if is_cn else c['passport_no']
        row = build_row(i, [
            ('A', i - 2),
            ('B', cert_type),
            ('C', cert_no),
            ('D', '0100'),
            ('E', '计算机'),
            ('F', '1'),
            ('G', '001'),
        ])
        rows.append(row)
    return rows


def build_port_call_rows(records):
    """海事船岸活动信息 (sheet: 海事船岸活动信息)"""
    rows = []
    for i, r in enumerate(records, 3):
        # 进港 0-11, 离港 12-23
        if r['arr']:
            arr_h = random.randint(0, 11)
            arr_m = random.randint(0, 59)
            arr_s = random.randint(0, 59)
            arr_str = r['arr'].replace(hour=arr_h, minute=arr_m, second=arr_s).strftime('%Y/%m/%d %H:%M:%S')
        else:
            arr_str = ''
        if r['dep']:
            dep_h = random.randint(12, 23)
            dep_m = random.randint(0, 59)
            dep_s = random.randint(0, 59)
            dep_str = r['dep'].replace(hour=dep_h, minute=dep_m, second=dep_s).strftime('%Y/%m/%d %H:%M:%S')
        else:
            dep_str = ''
        port_code = PORT_CODE.get(r['port'].upper(), '')

        # 国家代码
        country = r['country']
        if country == 'UN' or 'OPEN' in r['port'].upper():
            country = 'UN'
        elif country == 'CN':
            country = 'CN'
        elif country == 'ID':
            country = 'ID'
        else:
            country = 'UN'

        row = build_row(i, [
            ('A', i - 2),
            ('B', arr_str),
            ('C', dep_str),
            ('D', country),
            ('E', '1-1级'),
            ('G', port_code),
            ('H', '1-1级'),
        ])
        rows.append(row)
    return rows


def build_top10_rows(records):
    """前十港信息"""
    rows = []
    for i, r in enumerate(records[:10], 3):
        if r['arr']:
            arr_h = random.randint(0, 11)
            arr_m = random.randint(0, 59)
            arr_s = random.randint(0, 59)
            arr_str = r['arr'].replace(hour=arr_h, minute=arr_m, second=arr_s).strftime('%Y/%m/%d %H:%M:%S')
        else:
            arr_str = ''
        if r['dep']:
            dep_h = random.randint(12, 23)
            dep_m = random.randint(0, 59)
            dep_s = random.randint(0, 59)
            dep_str = r['dep'].replace(hour=dep_h, minute=dep_m, second=dep_s).strftime('%Y/%m/%d %H:%M:%S')
        else:
            dep_str = ''
        port_code = PORT_CODE.get(r['port'].upper(), '')
        country = r['country']
        if country == 'UN' or 'OPEN' in r['port'].upper():
            country = 'UN'

        row = build_row(i, [
            ('A', i - 2),
            ('B', arr_str),
            ('C', dep_str),
            ('E', country),
            ('F', port_code),
        ])
        rows.append(row)
    return rows


# ============================================================
# 主流程：复制模板 zip + 修改 4 个 sheet XML
# ============================================================

# 模板里的 sheet 编号映射（来自原模板的 workbook.xml 顺序）
SHEET_RID_TO_PATH = {
    'rId1': ('船上非旅客人员清单', 'xl/worksheets/sheet1.xml'),
    'rId2': ('旅客清单', 'xl/worksheets/sheet2.xml'),
    'rId3': ('供退物料清单', 'xl/worksheets/sheet3.xml'),
    'rId4': ('船上非旅客人员物品清单', 'xl/worksheets/sheet4.xml'),
    'rId5': ('船用物品清单', 'xl/worksheets/sheet5.xml'),
    'rId6': ('前十港信息', 'xl/worksheets/sheet6.xml'),
    'rId7': ('危险品信息', 'xl/worksheets/sheet7.xml'),
    'rId8': ('船舶证书信息', 'xl/worksheets/sheet8.xml'),
    'rId9': ('压舱水详细信息', 'xl/worksheets/sheet9.xml'),
    'rId10': ('海事船岸活动信息', 'xl/worksheets/sheet10.xml'),
    'rId11': ('沿海空箱信息', 'xl/worksheets/sheet11.xml'),
    'rId12': ('压舱水报告单信息', 'xl/worksheets/sheet12.xml'),
    'rId13': ('压舱水装载信息', 'xl/worksheets/sheet13.xml'),
    'rId14': ('压舱水更换信息表信息', 'xl/worksheets/sheet14.xml'),
    'rId15': ('压舱水排放信息表信息', 'xl/worksheets/sheet15.xml'),
    'rId16': ('随船人员清单', 'xl/worksheets/sheet16.xml'),
    'rId17': ('参数', 'xl/worksheets/sheet17.xml'),
}


def find_last_data_row(sheet_xml):
    """找 sheetData 中最大的行号"""
    matches = re.findall(r'<row r="(\d+)"', sheet_xml)
    if not matches: return 2
    return max(int(m) for m in matches)


def fill_sheet_xml(sheet_xml, new_rows_xml):
    """在 sheetData 末尾插入新行（找到 </sheetData> 前插入）"""
    if not new_rows_xml:
        return sheet_xml
    # 找最后一个 </row> 在 <sheetData> 内的位置
    # 简单策略：在 </sheetData> 前插入
    return sheet_xml.replace('</sheetData>', ''.join(new_rows_xml) + '</sheetData>')


def process(template_path, output_path, crews, records):
    """核心：复制 zip + 修改 4 个 sheet XML"""
    modified_sheets = {}  # path -> new xml content

    # 1. 船员名单 → sheet1 (船上非旅客人员清单)
    crew_rows = build_crew_rows(crews, None)
    if crew_rows:
        # 读 sheet1 XML，加 crew rows
        with zipfile.ZipFile(template_path) as z:
            sheet1 = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
        modified_sheets['xl/worksheets/sheet1.xml'] = fill_sheet_xml(sheet1, crew_rows)

    # 2. 物品清单 → sheet4
    goods_rows = build_goods_rows(crews)
    if goods_rows:
        with zipfile.ZipFile(template_path) as z:
            sheet4 = z.read('xl/worksheets/sheet4.xml').decode('utf-8')
        modified_sheets['xl/worksheets/sheet4.xml'] = fill_sheet_xml(sheet4, goods_rows)

    # 3. 海事船岸活动 → sheet10
    pc_rows = build_port_call_rows(records)
    if pc_rows:
        with zipfile.ZipFile(template_path) as z:
            sheet10 = z.read('xl/worksheets/sheet10.xml').decode('utf-8')
        modified_sheets['xl/worksheets/sheet10.xml'] = fill_sheet_xml(sheet10, pc_rows)

    # 4. 前十港 → sheet6
    top10_rows = build_top10_rows(records)
    if top10_rows:
        with zipfile.ZipFile(template_path) as z:
            sheet6 = z.read('xl/worksheets/sheet6.xml').decode('utf-8')
        modified_sheets['xl/worksheets/sheet6.xml'] = fill_sheet_xml(sheet6, top10_rows)

    # 5. 复制整个 zip，按修改表替换
    with zipfile.ZipFile(template_path, 'r') as zin:
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename in modified_sheets:
                    zout.writestr(item, modified_sheets[item.filename].encode('utf-8'))
                else:
                    zout.writestr(item, zin.read(item.filename))


# ============================================================
# CLI
# ============================================================

def main():
    import argparse
    ap = argparse.ArgumentParser(description="单证录入核心 v3 (zipfile 直操作)")
    ap.add_argument("crew", nargs="?", default=DEFAULT_CREW, help="IMO Crew List xlsx 路径（可省略以仅处理 Port of Call）")
    ap.add_argument("port", nargs="?", default=DEFAULT_POC, help="Port of Call xlsx 路径")
    ap.add_argument("output", nargs="?", default=DEFAULT_OUTPUT, help="输出 .xlsm 路径")
    args = ap.parse_args()

    if not os.path.exists(TEMPLATE_XLSM):
        print(f"❌ 模板不存在: {TEMPLATE_XLSM}")
        sys.exit(1)
    if not os.path.exists(args.port):
        print(f"❌ 港口文件不存在: {args.port}")
        sys.exit(1)

    has_crew = args.crew and os.path.exists(args.crew)
    if not has_crew:
        print("⚠️  未提供船员文件 — 船员名单/物品清单将不填充")
        crews = []
    else:
        crews = parse_crew_xlsx(args.crew)
        print(f"  解析船员 {len(crews)} 人")

    records = parse_poc_xlsx(args.port)
    print(f"  解析港口 {len(records)} 条")

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    process(TEMPLATE_XLSM, args.output, crews, records)
    print(f"\n🎉 输出: {args.output}  ({os.path.getsize(args.output):,} bytes)")


if __name__ == "__main__":
    main()
