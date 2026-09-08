"""
吴中结算单价补项脚本 - processId=126
输入：
  - 普通食品数据源_20260908112257_20251031115600.xlsx (sheet: data_1)
  - lmis数据源模板_20260908112308_检测项目查询统计-江苏20260812.xlsx (sheet: sheet1, 只用吴中行)
  - 吴中结算单价_20260908112311_吴中结算单价.xlsx (sheet: Sheet1)
输出：
  - 表格表解/20260908112313-p126/输出源/吴中结算单价.xlsx
"""

import re
import os
import openpyxl
from openpyxl.styles import (
    Font, Alignment, Border, Side, PatternFill, GradientFill
)
from openpyxl.utils import get_column_letter

BASE_DIR = "/workspace/表格表解/20260908112313-p126"
INPUT_DIR = os.path.join(BASE_DIR, "输入源")
OUTPUT_DIR = os.path.join(BASE_DIR, "输出源")
os.makedirs(OUTPUT_DIR, exist_ok=True)

XC_FILE = os.path.join(INPUT_DIR, "普通食品数据源_20260908112257_20251031115600.xlsx")
LMIS_FILE = os.path.join(INPUT_DIR, "lmis数据源模板_20260908112308_检测项目查询统计-江苏20260812.xlsx")
BASE_FILE = os.path.join(INPUT_DIR, "吴中结算单价_20260908112311_吴中结算单价.xlsx")
OUT_FILE = os.path.join(OUTPUT_DIR, "吴中结算单价.xlsx")


# ── 固定别名表：数据源写法 → 底表已有项目名 ──────────────────────────────
# key 是规范化后的数据源写法（已规范化处理），value 是底表已有项目
ALIAS_MAP_RAW = [
    # 微生物 n=5
    ("菌落总数(n=5)", "菌落总数"),
    ("大肠菌群(n=5)", "大肠菌群"),
    ("沙门氏菌(n=5)", "沙门氏菌"),
    ("金黄色葡萄球菌(n=5)", "金黄色葡萄球菌"),
    ("单核细胞增生李斯特氏菌(n=5)", "单核细胞增生李斯特氏菌"),
    ("铜绿假单胞菌(n=5)", "铜绿假单胞菌"),
    # 防腐剂
    ("山梨酸及其钾盐(以山梨酸计)", "山梨酸"),
    ("苯甲酸及其钠盐(以苯甲酸计)", "苯甲酸"),
    ("甜蜜素(以环己基氨基磺酸计)", "环己基氨基磺酸钠(甜蜜素)"),
    ("二氧化硫残留量", "二氧化硫"),
    ("丙酸及其盐(以丙酸计)", "丙酸钠(钙)"),
    ("丙酸及其盐含量(以丙酸计)", "丙酸钠(钙)"),
    # 酸价
    ("酸价(以脂肪计)", "酸值/酸价"),
    ("酸价(以脂肪计)(koh)", "酸值/酸价"),
    ("酸价(以koh计)", "酸值/酸价"),
    ("酸价(以脂肪计) (koh)", "酸值/酸价"),
    # 兽药
    ("氟苯尼考", "氟苯尼考（氯甲砜霉素）"),
    ("恩诺沙星", "恩诺沙星（恩诺沙星+环丙沙星总量）（不对喹诺酮打包价）"),
    ("孔雀石绿", "孔雀石绿（含结晶紫）"),
    ("磺胺类(总量)", "磺胺类（以总量计）"),
    # 黄曲霉毒素
    ("黄曲霉毒素b1", "黄曲霉毒素B1"),
    ("黄曲霉毒素m1", "黄曲霉毒素M1"),
    # 亚铁氰化钾
    ("亚铁氰化钾(以[fe(cn)6]4-计)", "亚铁氰化钾/亚铁氰化钠（以亚铁氰根计）"),
    ("亚铁氰化钾(以[fe(cn)6]计)", "亚铁氰化钾/亚铁氰化钠（以亚铁氰根计）"),
    # 二氧化碳
    ("二氧化碳气容量(20℃)", "二氧化碳"),
    # 果糖/葡萄糖 - 分列时视为已有，不新增
    ("果糖", "__SKIP__"),
    ("葡萄糖", "__SKIP__"),
    # 致泻大肠埃希氏菌
    ("致泻大肠埃希氏菌(n=5)", "大肠埃希氏菌O157：H7（不是「大肠埃希氏菌」）"),
]


def remove_html_tags(s: str) -> str:
    """去除HTML标签"""
    s = re.sub(r'<[^>]+>', '', s)
    return s


def normalize(s: str) -> str:
    """规范化名称：去HTML、trim、全角转半角、压缩空格、去括号类内容"""
    if not s:
        return ''
    s = str(s)
    # 去HTML标签
    s = remove_html_tags(s)
    # 全角转半角
    s = s.replace('（', '(').replace('）', ')')
    s = s.replace('、', ',').replace('，', ',')
    s = s.replace('．', '.').replace('。', '.')
    s = s.replace('－', '-').replace('—', '-').replace('–', '-')
    s = s.replace('／', '/')
    s = s.replace('：', ':')
    s = s.replace('；', ';')
    s = s.replace('\u201c', '"').replace('\u201d', '"')
    s = s.replace('\u2018', "'").replace('\u2019', "'")
    s = s.replace('\u3000', ' ')
    # 压缩连续空格，trim
    s = re.sub(r' +', ' ', s).strip()
    # 去掉「以XX计」类括号内容：只去掉特定的模式
    # 去掉：(以Pb计)/(以脂肪计)/(以山梨酸计)/(KOH)/(以KOH计)/(20℃)/(n=5) 等
    # 注意：这里只去掉用于匹配时的，写入时仍用原文
    s = re.sub(r'\(以[^)]*计\)', '', s)
    s = re.sub(r'\(KOH\)', '', s, flags=re.IGNORECASE)
    s = re.sub(r'\(n=\d+\)', '', s, flags=re.IGNORECASE)
    s = re.sub(r'\(20℃\)', '', s)
    # 再trim
    s = s.strip()
    # 转小写用于某些匹配（但别名表需要case-sensitive处理）
    return s


def normalize_lower(s: str) -> str:
    return normalize(s).lower()


def build_alias_map():
    """构建别名表：规范化(lower)的数据源名 → 底表项目名"""
    result = {}
    for src, tgt in ALIAS_MAP_RAW:
        key = normalize_lower(src)
        result[key] = tgt
    return result


ALIAS_MAP = build_alias_map()


def is_skip_value(v) -> bool:
    """判断是否为空/斜杠/横线，需跳过"""
    if v is None:
        return True
    s = str(v).strip()
    return s in ('', '/', '-')


def load_base_table():
    """读取底表所有数据（432行），返回列表和项目名集合"""
    wb = openpyxl.load_workbook(BASE_FILE)
    ws = wb['Sheet1']
    rows = []
    for r in range(2, ws.max_row + 1):
        row = [ws.cell(r, c).value for c in range(1, 5)]
        rows.append(row)
    print(f"底表行数: {len(rows)}，最后一行: {rows[-1]}")
    return rows


def load_xc_data():
    """读取XC数据，返回 {规范化名: (原文, 依据原文)} 字典"""
    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(XC_FILE)
    ws = wb['data_1']
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    col_item = headers.index('检验项目') + 1   # 159
    col_dep = headers.index('检验依据') + 1     # 155
    
    result = {}  # norm_key -> (原文, 依据原文列表)
    for r in range(2, ws.max_row + 1):
        item_raw = ws.cell(r, col_item).value
        dep_raw = ws.cell(r, col_dep).value
        if is_skip_value(item_raw):
            continue
        item_str = str(item_raw).strip()
        norm_key = normalize_lower(item_str)
        dep_str = '' if is_skip_value(dep_raw) else str(dep_raw).strip()
        
        if norm_key not in result:
            result[norm_key] = (item_str, [])
        if dep_str and dep_str not in ('/', '-'):
            # 归一化后去重
            deps = result[norm_key][1]
            dep_norm = re.sub(r'\s+', ' ', dep_str).strip()
            # 检查是否已存在（规范化比较）
            existing_norms = [re.sub(r'\s+', ' ', d).strip() for d in deps]
            if dep_norm not in existing_norms:
                deps.append(dep_str)
    
    print(f"XC检验项目数: {len(result)}")
    return result


def load_lmis_data():
    """读取LMIS数据（只保留委托单位含吴中的行），返回 {规范化名: (原文去HTML, 检验方法列表)}"""
    wb = openpyxl.load_workbook(LMIS_FILE)
    ws = wb['sheet1']
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    col_wt = headers.index('委托单位') + 1      # 3
    col_rpt = headers.index('报告中文名称') + 1  # 10
    col_mth = headers.index('检验方法') + 1     # 11
    
    result = {}  # norm_key -> (原文去HTML, 检验方法列表)
    wuzhong_count = 0
    for r in range(2, ws.max_row + 1):
        wt = ws.cell(r, col_wt).value
        if not wt or '吴中' not in str(wt):
            continue
        wuzhong_count += 1
        rpt_raw = ws.cell(r, col_rpt).value
        mth_raw = ws.cell(r, col_mth).value
        if is_skip_value(rpt_raw):
            continue
        rpt_str = remove_html_tags(str(rpt_raw).strip())  # 去HTML标签
        norm_key = normalize_lower(rpt_str)
        mth_str = '' if is_skip_value(mth_raw) else str(mth_raw).strip()
        
        if norm_key not in result:
            result[norm_key] = (rpt_str, [])
        if mth_str and mth_str not in ('/', '-'):
            deps = result[norm_key][1]
            mth_norm = re.sub(r'\s+', ' ', mth_str).strip()
            existing_norms = [re.sub(r'\s+', ' ', d).strip() for d in deps]
            if mth_norm not in existing_norms:
                deps.append(mth_str)
    
    print(f"LMIS吴中行数: {wuzhong_count}，独立报告中文名称数: {len(result)}")
    return result


def match_item(item_str, base_norms, alias_map):
    """
    匹配一个数据源项目名：
    返回：('alias', 底表项目名) 或 ('exact', 底表项目名) 或 ('new', None)
    """
    norm_key = normalize_lower(item_str)
    
    # 1. 先查固定别名表
    if norm_key in alias_map:
        return ('alias', alias_map[norm_key])
    
    # 2. 规范化后完全匹配底表
    if norm_key in base_norms:
        return ('exact', base_norms[norm_key])
    
    # 3. 未匹配 → 新增
    return ('new', None)


def merge_deps(xc_deps, lmis_deps):
    """
    合并XC依据和LMIS依据：去重后用；拼接
    XC依据可能是全文，LMIS依据可能是标准号
    """
    all_deps = []
    seen_norms = set()
    
    for dep in xc_deps + lmis_deps:
        dep_norm = re.sub(r'\s+', ' ', dep).strip()
        if dep_norm not in seen_norms:
            seen_norms.add(dep_norm)
            all_deps.append(dep)
    
    return '；'.join(all_deps) if all_deps else ''


# ── 12条固定新增行 ──────────────────────────────────────────────────────────
FIXED_NEW_ROWS = [
    ("喹啉黄", 150, 52.5),
    ("新红", 150, 52.5),
    ("红2G", 150, 52.5),
    ("相同色泽着色剂混合使用时各自用量占其最大使用量的比例之和", 100, 35.0),
    ("总固体", 100, 35.0),
    ("硫酸根", 100, 35.0),
    ("氯离子", 100, 35.0),
    ("耗氧量（以 O₂计）", 80, 28.0),
    ("烯唑醇", 150, 52.5),
    ("氟唑菌酰胺", 150, 52.5),
    ("噻唑膦", 150, 52.5),
    ("氟胺氰菊酯", 150, 52.5),
]


def get_dep_for_new_item(item_name, xc_data, lmis_data):
    """从XC和LMIS中提取新增项目的检验依据"""
    norm_key = normalize_lower(item_name)
    
    xc_deps = []
    lmis_deps = []
    
    # 从XC匹配
    if norm_key in xc_data:
        xc_deps = xc_data[norm_key][1]
    
    # 从LMIS匹配
    if norm_key in lmis_data:
        lmis_deps = lmis_data[norm_key][1]
    
    return merge_deps(xc_deps, lmis_deps)


def make_border():
    side = Side(style='thin', color='000000')
    return Border(left=side, right=side, top=side, bottom=side)


def make_alignment():
    return Alignment(horizontal='center', vertical='center', wrap_text=True)


def make_font():
    return Font(name='宋体', size=12, color='000000')


def make_yellow_fill():
    return PatternFill(fill_type='solid', fgColor='FFFF00')


def main():
    print("=== 开始处理 ===")
    
    # 1. 读取底表
    base_rows = load_base_table()  # list of [序号, 检验项目, 最高限价, 投标报价]
    
    # 构建底表规范化名称集合
    base_norms = {}  # normalize_lower(项目名) -> 底表原文
    for row in base_rows:
        item = row[1]
        if item:
            norm = normalize_lower(str(item))
            base_norms[norm] = str(item)
    print(f"底表规范化项目数: {len(base_norms)}")
    
    # 2. 读取XC数据
    xc_data = load_xc_data()  # norm_lower -> (原文, [依据])
    
    # 3. 读取LMIS数据（仅吴中）
    lmis_data = load_lmis_data()  # norm_lower -> (原文去HTML, [方法])
    
    # 4. 验证12条固定新增行确实不在底表中，打印匹配情况
    print("\n=== 验证固定新增行 ===")
    for item_name, price_max, price_bid in FIXED_NEW_ROWS:
        norm_key = normalize_lower(item_name)
        in_alias = norm_key in ALIAS_MAP
        in_base = norm_key in base_norms
        # 检验依据
        dep = get_dep_for_new_item(item_name, xc_data, lmis_data)
        print(f"  [{item_name}] 别名={in_alias}, 底表={in_base}, 依据前50={dep[:50] if dep else '(空)'}")
    
    # 5. 创建输出文件（参考模板样式）
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = 'Sheet1'
    
    # 写表头
    headers = ['序号', '检验项目', '最高限价（元/次）', '投标报价（元/次）', '检验依据']
    for c, h in enumerate(headers, 1):
        cell = ws_out.cell(1, c, h)
        cell.font = Font(name='宋体', size=12, bold=True, color='000000')
        cell.alignment = make_alignment()
        cell.border = make_border()
    
    # 6. 写旧的432行
    for i, row in enumerate(base_rows, 1):
        r = i + 1  # 第2行开始
        for c in range(1, 5):
            cell = ws_out.cell(r, c, row[c-1])
            cell.font = make_font()
            cell.alignment = make_alignment()
            cell.border = make_border()
        # E列（检验依据）留空
        cell_e = ws_out.cell(r, 5, None)
        cell_e.font = make_font()
        cell_e.alignment = make_alignment()
        cell_e.border = make_border()
    
    # 7. 写12条新增行
    yellow_fill = make_yellow_fill()
    
    for idx, (item_name, price_max, price_bid) in enumerate(FIXED_NEW_ROWS):
        seq_no = 432 + idx + 1
        r = 432 + idx + 2  # 底表432行 + 表头1行 + 新增行起始
        dep = get_dep_for_new_item(item_name, xc_data, lmis_data)
        
        row_data = [seq_no, item_name, price_max, price_bid, dep]
        for c, val in enumerate(row_data, 1):
            cell = ws_out.cell(r, c, val)
            cell.font = make_font()
            cell.alignment = make_alignment()
            cell.border = make_border()
            # B列（检验项目）黄底
            if c == 2:
                cell.fill = yellow_fill
    
    # 8. 设置列宽
    col_widths = [8, 50, 18, 18, 60]
    for c, w in enumerate(col_widths, 1):
        ws_out.column_dimensions[get_column_letter(c)].width = w
    
    # 9. 保存
    wb_out.save(OUT_FILE)
    print(f"\n✓ 输出文件已保存: {OUT_FILE}")
    
    # 10. 验证输出
    wb_check = openpyxl.load_workbook(OUT_FILE)
    ws_check = wb_check['Sheet1']
    print(f"  输出行数（含表头）: {ws_check.max_row}")
    print(f"  第2行: {[ws_check.cell(2, c).value for c in range(1, 6)]}")
    print(f"  第433行: {[ws_check.cell(433, c).value for c in range(1, 6)]}")
    print(f"  第434行: {[ws_check.cell(434, c).value for c in range(1, 6)]}")
    print(f"  最后行(445): {[ws_check.cell(445, c).value for c in range(1, 6)]}")
    # 验证第433行B列黄底
    fill433 = ws_check.cell(433, 2).fill
    print(f"  第433行B列填充类型={fill433.fill_type}, fgColor={fill433.fgColor.rgb if fill433.fill_type=='solid' else 'N/A'}")
    # 验证第2行投标报价
    print(f"  第2行投标报价: {ws_check.cell(2, 4).value}")
    
    # 写入manifest
    import json
    manifest = {"files": ["吴中结算单价.xlsx"]}
    manifest_path = os.path.join(OUTPUT_DIR, "manifest.json")
    with open(manifest_path, 'w', encoding='utf-8') as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    print(f"✓ manifest.json 已写入: {manifest_path}")


if __name__ == '__main__':
    main()
