#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
资质匹配模块 - 根据招标公告资质要求匹配合适的施工单位
"""

import re
import json
from pathlib import Path

# 资质等级顺序（数字越小越高）
LEVEL_ORDER = {
    '特级': 0,
    '一级': 1,
    '二级': 2,
    '三级': 3,
    '四级': 4,
    '五级': 5,
    '不分等级': 99,
}

# 常见施工总承包/专业承包类别关键词
# 关键词格式有两种：
#   1. 标准格式：XXX工程施工总承包（如"建筑工程施工总承包"）
#   2. 源文本格式：施工总承包不分行业XXX（如"施工总承包不分行业建筑工程"）
# 添加"XXX工程"形式的关键词以覆盖第2种格式
CATEGORY_KEYWORDS = {
    '市政公用工程施工总承包': [
        '市政公用工程施工总承包', '市政公用工程', '市政公用',
    ],
    '建筑工程施工总承包': [
        '建筑工程施工总承包', '房屋建筑工程施工总承包',
        '房屋建筑工程', '工业与民用建筑工程',  # 标准格式
        '建筑工程', '房屋建筑',                 # 中间型（来源文本格式）
    ],
    '建筑装修装饰工程专业承包': [
        '装修装饰工程专业承包', '装修装饰', '建筑装修装饰',
        '建筑装修装饰工程',
    ],
    '水利水电工程施工总承包': [
        '水利水电工程施工总承包', '水利水电工程', '水利水电',
    ],
    '公路工程施工总承包': [
        '公路工程施工总承包', '公路工程', '道路工程', '道路',
    ],
    '公路路面工程专业承包': [
        '公路路面工程专业承包', '公路路面工程', '路面工程',
    ],
    '公路路基工程专业承包': [
        '公路路基工程专业承包', '公路路基工程', '路基工程',
    ],
    '机电工程施工总承包': [
        '机电工程施工总承包', '机电工程', '机电',
    ],
    '电力工程施工总承包': [
        '电力工程施工总承包', '电力工程', '电力',
    ],
    '地基基础工程专业承包': [
        '地基基础工程专业承包', '地基基础工程', '地基基础',
    ],
    '钢结构工程专业承包': [
        '钢结构工程专业承包', '钢结构工程', '钢结构',
    ],
    '消防设施工程专业承包': [
        '消防设施工程专业承包', '消防设施工程', '消防设施',
    ],
    '电子与智能化工程专业承包': [
        '电子与智能化工程专业承包', '电子与智能化工程', '智能化工程', '电子与智能化',
    ],
    '建筑幕墙工程专业承包': [
        '建筑幕墙工程专业承包', '建筑幕墙工程', '幕墙工程', '幕墙',
    ],
    '园林绿化工程施工': [
        '园林绿化工程施工', '园林绿化工程', '园林绿化',
    ],
    '输变电工程专业承包': [
        '输变电工程专业承包', '输变电工程', '输变电',
    ],
    '建筑机电安装工程专业承包': [
        '建筑机电安装工程专业承包', '建筑机电安装工程', '建筑机电安装',
    ],
    '环保工程专业承包': [
        '环保工程专业承包', '环保工程', '环保施工',
    ],
    '特种工程专业承包': [
        '特种工程专业承包', '特种工程',
    ],
}


def normalize_text(text):
    """统一文本，去除多余空白"""
    if not text:
        return ''
    return re.sub(r'\s+', '', str(text))


def parse_level(text):
    """
    从资质要求文本中解析等级。
    返回: ("等级字符串", has_above) 或 None

    "以上"判断：搜索"XXX级"后，看后面紧跟的是否有"以上"或"及以上"
    """
    text = normalize_text(text)
    # 找所有"XXX级"出现位置
    for m in re.finditer(r'([一二三四特]级)', text):
        lvl = m.group(1)
        lvl_end = m.end()
        # 取等级词后6个字符，看是否有"以上"或"及以上"
        after = text[lvl_end:min(len(text), lvl_end+6)]
        if '以上' in after or '及以上' in after:
            return lvl, True
        # 也可能是"一级"在"及以上"里，要往前看
        lvl_start = m.start()
        before = text[max(0, lvl_start-2):lvl_start]
        after2 = text[lvl_end:min(len(text), lvl_end+6)]
        if '以上' in (before + after2):
            return lvl, True
        # 默认：如果后面紧跟"级"（重复匹配到同一个），继续找下一个
    # 不分等级
    if any(kw in text for kw in ['不分等级', '不分专业', '不分行业']):
        return '不分等级', False
    return None


def find_categories(text):
    """
    从文本中识别所有可能的资质类别。
    返回: [(category_name, matched_keyword, start_pos), ...]
    """
    text = normalize_text(text)
    matches = []
    for cat_name, keywords in CATEGORY_KEYWORDS.items():
        for kw in keywords:
            idx = text.find(kw)
            if idx >= 0:
                matches.append((cat_name, kw, idx))
    # 按关键词长度降序（长词优先）
    matches.sort(key=lambda x: -len(x[1]))
    return matches


def parse_qual_req(text):
    """
    解析完整的资质要求文本。
    返回: [{"category": "...", "min_level": "三级", "has_above": True}, ...]

    策略：分割后逐段解析，优先尝试长类别名匹配。
    """
    if not text:
        return []

    text = normalize_text(text)
    # 按中英文分号、逗号、顿号分割
    segments = re.split(r'[;；，,]', text)
    requirements = []

    for seg in segments:
        seg = seg.strip()
        if len(seg) < 5:
            continue

        # 去除前缀
        seg_clean = re.sub(r'^(具备|具有|同时具备|须具备|必须具备)+', '', seg).strip()
        # 去除无用括号和"资质"后缀，但保留"及以上"（level_text 需要它）
        seg_for_cat = re.sub(r'[）\]\)资质的?]$', '', seg_clean).strip()
        seg_for_cat = re.sub(r'^(的)?资质要求?$', '', seg_for_cat).strip()
        seg_for_cat = re.sub(r'^(和|以及|或|\|).*$', '', seg_for_cat).strip()

        if len(seg_for_cat) < 5:
            continue

        # 找类别
        cats = find_categories(seg_for_cat)
        if not cats:
            continue

        best_cat, best_kw, _ = cats[0]

        # 在段内找等级（以best_kw之后的位置为参考）
        kw_end = seg_for_cat.find(best_kw) + len(best_kw)
        level_text = seg_for_cat[kw_end:kw_end+15]  # 等级在类别名之后

        parsed_level = parse_level(level_text)
        if not parsed_level:
            # 等级可能在类别名之前（如"三级建筑工程施工总承包"）
            lvl_candidates = re.findall(r'([一二三四特]级)', seg_for_cat)
            if lvl_candidates:
                parsed_level = parse_level(lvl_candidates[0])
            else:
                # 不分等级的情况
                if any(kw in seg_for_cat for kw in ['不分等级', '不分专业', '不分行业']):
                    parsed_level = ('不分等级', False)

        if parsed_level:
            min_level, has_above = parsed_level
            requirements.append({
                'category': best_cat,
                'min_level': min_level,
                'has_above': has_above,
            })

    # 去重：同一类别保留最严格的等级要求
    seen = {}
    for r in requirements:
        cat = r['category']
        lvl_rank = LEVEL_ORDER.get(r['min_level'], 99)
        if cat not in seen or lvl_rank < LEVEL_ORDER.get(seen[cat]['min_level'], 99):
            seen[cat] = r
    return list(seen.values())


def level_satisfies(company_level, required_level, has_above=False):
    """
    判断公司资质等级是否满足要求。
    """
    if company_level == required_level:
        return True
    if company_level not in LEVEL_ORDER or required_level not in LEVEL_ORDER:
        return False
    if has_above:
        return LEVEL_ORDER[company_level] < LEVEL_ORDER[required_level]
    return False


def company_matches(company_qualifications, qual_requirements):
    """
    判断公司是否满足招标资质要求。所有要求（AND 逻辑）必须同时满足。
    """
    if not qual_requirements:
        return False

    for req in qual_requirements:
        req_cat = req['category']
        req_min = req['min_level']
        req_has_above = req.get('has_above', True)

        matched = False
        for cq in company_qualifications:
            cq_cat = cq.get('category', '')
            cq_lvl = cq.get('level', '')

            # 类别精确匹配（标准化名称）
            if req_cat != cq_cat:
                continue
            if level_satisfies(cq_lvl, req_min, req_has_above):
                matched = True
                break
        if not matched:
            return False
    return True


def match_companies_to_bid(companies, bid_qual_req):
    """
    将招标资质要求与公司列表匹配，返回符合条件公司列表。
    """
    if not companies or not bid_qual_req:
        return []

    requirements = parse_qual_req(bid_qual_req)
    if not requirements:
        return []

    matched = []
    for company in companies:
        quals = company.get('qualifications', [])
        if company_matches(quals, requirements):
            matched.append(company)
    return matched


# ── 工具函数 ──────────────────────────────────────────────────────────────

COMPANY_DB_PATH = Path(__file__).parent / 'companies.json'


def load_companies(db_path=None):
    path = Path(db_path) if db_path else COMPANY_DB_PATH
    if not path.exists():
        return []
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data.get('companies', [])


def save_companies(companies, db_path=None):
    path = Path(db_path) if db_path else COMPANY_DB_PATH
    data = {
        'meta': {'version': '1.0', 'updated': '', 'description': '施工单位资质数据库'},
        'companies': companies,
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ── 测试 ──────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    test_cases = [
        "建筑工程施工总承包三级（含）以上资质，并在人员、设备、资金等方面具有相应的施工能力；",
        "市政公用工程施工总承包三级（含）以上资质；",
        "施工总承包不分专业建筑工程三级（含）以上资质；",
        "具备【施工总承包不分行业建筑工程三级(含)以上资质】；",
        "市政公用工程施工总承包一级及以上；",
        "具有房屋建筑工程施工总承包三级（含）以上资质；",
        "同时具备建筑工程施工总承包二级（含）以上资质，和建筑装修装饰工程专业承包一级（含）以上资质；",
    ]

    print("=== 资质解析测试 ===")
    for tc in test_cases:
        result = parse_qual_req(tc)
        print(f"  原文: {tc[:60]}")
        print(f"  解析: {result}")
        print()

    companies = [
        {
            "name": "江苏八达建设集团有限公司",
            "email": "bada@example.com",
            "qualifications": [
                {"category": "建筑工程施工总承包", "level": "二级"},
                {"category": "市政公用工程施工总承包", "level": "三级"},
            ]
        },
        {
            "name": "常州华新建筑安装工程有限公司",
            "email": "huaxin@example.com",
            "qualifications": [
                {"category": "建筑工程施工总承包", "level": "三级"},
            ]
        },
        {
            "name": "江苏利安建设工程有限公司",
            "email": "liian@example.com",
            "qualifications": [
                {"category": "机电工程施工总承包", "level": "二级"},
            ]
        },
        {
            "name": "江苏中如有市政工程有限公司",
            "email": "zhongrou@example.com",
            "qualifications": [
                {"category": "市政公用工程施工总承包", "level": "一级"},
            ]
        },
    ]

    print("=== 匹配测试 ===")
    for tc in test_cases:
        matched = match_companies_to_bid(companies, tc)
        print(f"  原文: {tc[:60]}")
        print(f"  匹配: {[c['name'] for c in matched]}")
        print()
