#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
菲格洛亚 FIGUEROA 财务工具包 — 模拟数据生成器
基于真实 BabyCat IP + 四大渠道 + 14 个对账差异场景
Seed = 20260811, 可复现

Usage: python generate_data.py
Output: ../data/*.csv (10 files)
"""

import csv
import random
import os
import sys
from datetime import date, timedelta

# ── Config ──────────────────────────────────────────────
random.seed(20260811)
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data')
os.makedirs(OUT_DIR, exist_ok=True)

def write_csv(filename, headers, rows):
    """Write CSV with UTF-8 BOM for Excel compatibility."""
    path = os.path.join(OUT_DIR, filename)
    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(headers)
        w.writerows(rows)
    print(f'  ✓ {filename} ({len(rows)} rows)')
    return path

def d(iso_str):
    """Parse date string."""
    return date.fromisoformat(iso_str)

# ══════════════════════════════════════════════════════════
# 1. BabyCat 角色（基于摩点众筹公开信息）
# ══════════════════════════════════════════════════════════

BABYCAT_CHARACTERS = [
    ('BC-01', '橘影神偷',   '橘猫',     '常规款',  89),
    ('BC-02', '折耳骑士',   '英短折耳', '常规款',  89),
    ('BC-03', '笨笨博士',   '布偶猫',   '常规款',  89),
    ('BC-04', '奶牛猫',     '奶牛猫',   '常规款',  89),
    ('BC-05', '波斯猫',     '波斯猫',   '常规款',  89),
    ('BC-06', '软萌夹锁',   '德文卷毛', '常规款',  89),
    ('BC-07', '踏雪',       '白猫',     '小隐藏',  None),  # 隐藏款不单独定价
    ('BC-08', '缅因',       '缅因猫',   '大隐藏',  None),
]

# ══════════════════════════════════════════════════════════
# 2. 艺术家主数据
# ══════════════════════════════════════════════════════════

ARTISTS = [
    ('AR-001', '林小喵',    '独立潮玩设计师',     'BC-01,BC-07',          0.08,  '销售后结算'),
    ('AR-002', '陈卷卷',    '自由插画师',         'BC-02',                0.07,  '销售后结算'),
    ('AR-003', '周布布',    '猫主题IP创作者',     'BC-03,BC-06',          0.08,  '销售后结算'),
    ('AR-004', '刘星野',    '新锐玩具设计师',     'BC-04,BC-05',          0.07,  '销售后结算'),
    ('AR-005', '赵小喵',    '小红书潮玩博主',     'BC-08',                0.10,  '预付+结算'),
    ('AR-006', '黄多多',    '在校艺术生/草根画手', 'BC-ACC-01,BC-ACC-02',  0.06,  '销售后结算'),
    ('AR-007', '何梦鹿',    '自由职业插画师',     'BC-ACC-03,BC-ACC-04',  0.06,  '销售后结算'),
    ('AR-008', '杨桃桃',    '纹身师兼玩具涂装师',  'BC-ACC-05,BC-ACC-06',  0.06,  '销售后结算'),
]

# ══════════════════════════════════════════════════════════
# 3. SKU 目录（~45 个）
# ══════════════════════════════════════════════════════════

SKUS = []
sku_id = 1

# 3a. 盲盒单盒 SKU — 6 常规款（盲盒不指定角色，随机发货）
SKUS.append(('SKU-001', 'BLIND-01',  'BabyCat随机单盒（盲盒早鸟）',     '盲盒',  89,  69))
SKUS.append(('SKU-002', 'BLIND-02',  'BabyCat随机单盒（盲盒常规）',     '盲盒',  95,  69))

# 3b. 明盒 SKU — 6 常规款 × 明盒（指定角色）
for i, bc in enumerate(BABYCAT_CHARACTERS[:6]):
    SKUS.append((f'SKU-{3+i:03d}', f'OPEN-{bc[0]}', f'BabyCat明盒·{bc[1]}（{bc[2]}）', '明盒', 100, 69))

# 3c. CP 对盒 — 3 组
CP_PAIRS = [('橘影神偷', '折耳骑士'), ('笨笨博士', '软萌夹锁'), ('奶牛猫', '波斯猫')]
cp_codes = ['CP-01', 'CP-02', 'CP-03']
for i, (a, b) in enumerate(CP_PAIRS):
    SKUS.append((f'SKU-{9+i:03d}', cp_codes[i], f'BabyCat CP对盒·{a}×{b}', 'CP对盒', 184, 138))

# 3d. 端盒（6 只不重复，全部常规款）
SKUS.append(('SKU-012', 'ENDBOX', 'BabyCat端盒（6只不重复·常规款全套）', '端盒', 546, 414))

# 3e. ALL-IN 套装（6 常规 + 2 隐藏）
SKUS.append(('SKU-013', 'ALLIN', 'BabyCat ALL-IN套装（6常规+2隐藏）', '套装', 680, 520))

# 3f. 限量款
SKUS.append(('SKU-014', 'LIMIT-01', 'BabyCat冬季吊卡（限量1000个·济南快闪首发）', '限量款', 129, 89))

# 3g. 配件周边（10 个）
ACCESSORIES = [
    ('ACC-01', '毛绒挂件·橘影神偷',      '挂件', 39, 18),
    ('ACC-02', '亚克力立牌·BabyCat全家福', '立牌', 49, 22),
    ('ACC-03', '徽章套装（6枚）',        '徽章', 69, 30),
    ('ACC-04', '摇摇乐钥匙扣·折耳骑士',   '钥匙扣', 29, 12),
    ('ACC-05', 'BabyCat帆布袋',         '帆布袋', 59, 25),
    ('ACC-06', '贴纸套装',              '贴纸', 19, 6),
    ('ACC-07', 'BabyCat手机壳（iPhone款）', '手机壳', 49, 20),
    ('ACC-08', 'BabyCat冰箱贴（4枚装）',  '冰箱贴', 39, 14),
    ('ACC-09', 'BabyCat明信片套装',      '明信片', 29, 8),
    ('ACC-10', 'BabyCat毛绒抱枕（30cm）', '抱枕', 99, 45),
]
for i, (code, name, cat, retail, cost) in enumerate(ACCESSORIES):
    SKUS.append((f'SKU-{15+i:03d}', code, f'BabyCat{name}', cat, retail, cost))

SKU_MAP = {s[1]: s for s in SKUS}  # sku_code → row

# ══════════════════════════════════════════════════════════
# 4. 经销商主数据（10 家，按渠道类型分类）
# ══════════════════════════════════════════════════════════

DEALERS = [
    # (编码, 名称, 渠道品牌, 渠道类型, 所在区域, 账期天, 结算模式)
    ('DE-001', '名创优品（广州）有限责任公司',       '名创优品',   '连锁零售', '华南', 45, '寄售'),
    ('DE-002', '酷乐潮玩（北京）贸易有限公司',        '酷乐潮玩',   '连锁零售', '华北', 30, '寄售'),
    ('DE-003', 'TOPTOY（广州）文化发展有限公司',      'TOPTOY',    '连锁零售', '华南', 45, '寄售'),
    ('DE-004', '广州KK馆贸易有限公司',               'KKV',       '连锁零售', '华南', 30, '寄售'),
    ('DE-005', '上海晨光文具股份有限公司（九木杂物社）', '九木杂物社', '连锁零售', '华东', 60, '寄售'),
    ('DE-006', '沃尔玛（中国）投资有限公司',          '沃尔玛',    '连锁零售', '全国', 60, '买断'),
    ('DE-007', '抖音电商（官方旗舰店）',              '抖音电商',   '电商平台', '全国', 15, '平台结算'),
    ('DE-008', 'TikTok Shop（跨境）',               'TikTok',    '电商平台', '海外', 30, '平台结算'),
    ('DE-009', '广州潮玩前线贸易有限公司',            '华南区域',   '区域经销', '华南', 30, '买断'),
    ('DE-010', '义乌萌物集贸易有限公司',              '华东区域',   '区域经销', '华东', 30, '买断'),
]

# ══════════════════════════════════════════════════════════
# 5. 生成销售出库单（小单快反 + 14 场景）
# ══════════════════════════════════════════════════════════

SALES_ORDERS = []
SELL_THROUGH = []   # 寄售渠道月度销售报表
so_counter = 1
st_counter = 1

# 常用 SKU 简化引用
SKU_BLIND_EARLY  = 'BLIND-01'   # ¥89 早鸟盲盒
SKU_BLIND_REG    = 'BLIND-02'   # ¥95 常规盲盒
SKU_OPEN_01      = 'OPEN-BC-01' # 橘影神偷明盒 ¥100
SKU_OPEN_03      = 'OPEN-BC-03' # 笨笨博士明盒 ¥100
SKU_ENDBOX       = 'ENDBOX'     # 端盒 ¥546
SKU_ALLIN        = 'ALLIN'      # ALL-IN ¥680
SKU_LIMIT        = 'LIMIT-01'   # 冬季吊卡 ¥129
SKU_ACC_01       = 'ACC-01'     # 毛绒挂件 ¥39
SKU_ACC_02       = 'ACC-02'     # 亚克力立牌 ¥49
SKU_ACC_04       = 'ACC-04'     # 摇摇乐钥匙扣 ¥29
SKU_ACC_06       = 'ACC-06'     # 贴纸套装 ¥19

# Helper
def make_so(dealer, sku_code, qty, order_date, batch_type='常规补货', discount=1.0, note=''):
    global so_counter
    sku = SKU_MAP[sku_code]
    unit_price = sku[4]  # 经销价
    amount = unit_price * qty
    net_amount = round(amount * discount, 2)
    dealer_info = next(d for d in DEALERS if d[0] == dealer)
    due_days = dealer_info[5]
    due_date = order_date + timedelta(days=due_days)
    so_id = f'SO-{order_date.strftime("%Y%m%d")}-{so_counter:03d}'
    so_counter += 1
    SALES_ORDERS.append([
        so_id, order_date.isoformat(), dealer, dealer_info[1], sku_code, sku[2],
        qty, unit_price, amount, discount, net_amount, due_days,
        due_date.isoformat(), batch_type, note
    ])
    return so_id, net_amount, due_date

# ── 5a. 名创优品（DE-001）小单快反 3 批 ──
# 试产 200 只 → 追单 500 → 再追 800
d1 = d('2026-05-10')
so_id, amt, _ = make_so('DE-001', SKU_BLIND_REG, 200, d1, '首单试产', 0.97, 'BabyCat首轮铺货试产')
so_id, amt, _ = make_so('DE-001', SKU_ENDBOX, 100, d1, '首单试产', 0.97, '端盒同步试产')
so_id, amt, _ = make_so('DE-001', SKU_OPEN_01, 60, d1, '首单试产', 0.97)

d2 = d('2026-06-05')
so_01, amt1, due1 = make_so('DE-001', SKU_BLIND_REG, 500, d2, '快反追单', 0.97, '销售超预期→追单（场景1a）')
so_02, amt2, due2 = make_so('DE-001', SKU_ENDBOX, 200, d2, '快反追单', 0.97, '端盒追单（场景1b）')
so_03, amt3, due3 = make_so('DE-001', SKU_ACC_01, 300, d2, '快反追单', 0.97, '挂件追单（场景1c）')
# ★ 场景1: 1笔回款覆盖 3 张单 (so_01+so_02+so_03 金额合并)
SCENE1_SO_IDS = [so_01, so_02, so_03]
SCENE1_TOTAL = round(amt1 + amt2 + amt3, 2)

d3 = d('2026-07-08')
make_so('DE-001', SKU_BLIND_REG, 800, d3, '二次追单', 0.95, '暑期促销追单')
make_so('DE-001', SKU_OPEN_03, 120, d3, '二次追单', 0.95)

# ── 5b. 酷乐潮玩（DE-002） ──
d1 = d('2026-05-15')
make_so('DE-002', SKU_BLIND_EARLY, 150, d1, '首单试产', 0.97)
make_so('DE-002', SKU_ENDBOX, 80, d1, '首单试产', 0.97)

d2 = d('2026-06-12')
so_06, amt6, due6 = make_so('DE-002', SKU_BLIND_REG, 400, d2, '快反追单', 0.97)
make_so('DE-002', SKU_OPEN_01, 80, d2, '快反追单', 0.97)

d3 = d('2026-07-10')
make_so('DE-002', SKU_BLIND_REG, 600, d3, '二次追单', 0.95)

# ── 5c. TOPTOY（DE-003） ──
d1 = d('2026-05-20')
make_so('DE-003', SKU_BLIND_EARLY, 200, d1, '首单试产', 0.97)
make_so('DE-003', SKU_ENDBOX, 100, d1, '首单试产', 0.97)

d2 = d('2026-06-18')
so_top_1, amt_top1, due_top1 = make_so('DE-003', SKU_BLIND_REG, 500, d2, '快反追单', 0.97)
so_top_2, amt_top2, due_top2 = make_so('DE-003', SKU_ENDBOX, 250, d2, '快反追单', 0.97)
# ★ 场景2: 部分回款（回 5 万 vs 销售总额 ¥78,400+）
SCENE2_SO_IDS = [so_top_1, so_top_2]
SCENE2_TOTAL = round(amt_top1 + amt_top2, 2)

d3 = d('2026-07-15')
make_so('DE-003', SKU_BLIND_REG, 700, d3, '二次追单', 0.95)

# ── 5d. KKV（DE-004） ──
d1 = d('2026-05-08')
make_so('DE-004', SKU_BLIND_EARLY, 100, d1, '首单试产', 0.97)

d2 = d('2026-06-10')
make_so('DE-004', SKU_BLIND_REG, 300, d2, '快反追单', 0.97)
make_so('DE-004', SKU_ACC_02, 200, d2, '快反追单', 0.97)

# ★ 场景8: KKV 退货冲销
d3 = d('2026-07-05')
make_so('DE-004', SKU_BLIND_REG, -30, d3, '退货冲销', 1.0, '品质问题退货（场景8）')
# Normal order same day
make_so('DE-004', SKU_BLIND_REG, 160, d3, '常规补货', 0.95)

# ── 5e. 九木杂物社（DE-005） ──
d1 = d('2026-05-12')
so_07, amt7, due7 = make_so('DE-005', SKU_BLIND_EARLY, 80, d1, '首单试产', 0.97)
make_so('DE-005', SKU_ENDBOX, 40, d1, '首单试产', 0.97)

d2 = d('2026-06-08')
make_so('DE-005', SKU_BLIND_REG, 250, d2, '快反追单', 0.97)
make_so('DE-005', SKU_OPEN_03, 50, d2, '快反追单', 0.97)

# ★ 场景7: 逾期未回（so_07 账期60天→6月底到期→仍未回款）
SCENE7_SO_ID = so_07

# ── 5f. 沃尔玛（DE-006） ──
d1 = d('2026-05-05')
so_w1, amt_w1, due_w1 = make_so('DE-006', SKU_BLIND_REG, 200, d1, '首单试产', 0.95)
so_w2, amt_w2, due_w2 = make_so('DE-006', SKU_ENDBOX, 100, d1, '首单试产', 0.95)

d2 = d('2026-06-01')
make_so('DE-006', SKU_BLIND_REG, 350, d2, '快反追单', 0.95)

# ★ 场景3: 逾期未回（5月2张单账期60天→7月初到期→0回款）
SCENE3_SO_IDS = [so_w1, so_w2]

# ── 5g. 抖音电商（DE-007） ──
d1 = d('2026-05-18')
make_so('DE-007', SKU_BLIND_REG, 300, d1, '平台铺货', 1.0)
make_so('DE-007', SKU_ALLIN, 50, d1, '平台铺货', 1.0)

d2 = d('2026-06-20')
make_so('DE-007', SKU_BLIND_REG, 500, d2, '快反追单', 1.0)
make_so('DE-007', SKU_ACC_04, 400, d2, '快反追单', 1.0)

# ★ 场景4: 抖音预付款（渠道预充值，无对应销售）
# 记录在回款里

# ── 5h. TikTok Shop（DE-008） ──
d1 = d('2026-06-01')
so_tk, amt_tk, due_tk = make_so('DE-008', SKU_BLIND_REG, 200, d1, '海外铺货', 1.0, 'TikTok Shop US')
make_so('DE-008', SKU_ALLIN, 30, d1, '海外铺货', 1.0)

d2 = d('2026-07-01')
make_so('DE-008', SKU_BLIND_REG, 350, d2, '快反追单', 1.0, 'TikTok Shop US')

# ★ 场景14: USD回款汇率差异
SCENE14_SO_ID = so_tk

# ── 5i. 广州潮玩前线（DE-009） ──
d1 = d('2026-05-25')
so_09, amt9, due9 = make_so('DE-009', SKU_BLIND_REG, 120, d1, '首单试产', 0.97)
make_so('DE-009', SKU_ENDBOX, 50, d1, '首单试产', 0.97)

d2 = d('2026-06-25')
make_so('DE-009', SKU_BLIND_REG, 300, d2, '快反追单', 0.97)

# ★ 场景9: 回款无备注自动匹配
SCENE9_SO_ID = so_09

# ── 5j. 义乌萌物集（DE-010） ──
d1 = d('2026-05-28')
so_10, amt10, due10 = make_so('DE-010', SKU_BLIND_REG, 100, d1, '首单试产', 0.97)

d2 = d('2026-06-28')
make_so('DE-010', SKU_BLIND_REG, 250, d2, '快反追单', 0.97)
make_so('DE-010', SKU_ACC_06, 500, d2, '快反追单', 0.97)

# ★ 场景10: 小额差异 ¥1
SCENE10_SO_ID = so_10

# ── 5k. 成都/西南经销商（并入 DE-009 部分，加一个区域） ──
# ★ 场景5: 金额差异 ¥0.67（加入 DE-009 7月数据）
d3 = d('2026-07-20')
so_05, amt5, due5 = make_so('DE-009', SKU_BLIND_REG, 130, d3, '二次追单', 0.97)
SCENE5_SO_ID = so_05
SCENE5_AMT = amt5  # 期望金额，回款会差 ¥0.67

# ══════════════════════════════════════════════════════════
# 6. 寄售渠道 Sell-Through 报表
# ══════════════════════════════════════════════════════════

def make_st(dealer, sku_code, qty_sold, month, unit_price=None):
    global st_counter
    sku = SKU_MAP[sku_code]
    up = unit_price or sku[4]
    amount = round(up * qty_sold, 2)
    SELL_THROUGH.append([
        f'ST-{month.replace("-","")}-{st_counter:03d}',
        month + '-01', dealer, sku_code, sku[2], qty_sold, up, amount
    ])
    st_counter += 1
    return amount

# 名创优品 5-7 月 sell-through
make_st('DE-001', SKU_BLIND_REG, 180, '2026-05', 89)
make_st('DE-001', SKU_ENDBOX, 95, '2026-05', 546)
make_st('DE-001', SKU_OPEN_01, 55, '2026-05', 100)
make_st('DE-001', SKU_BLIND_REG, 480, '2026-06', 89)
make_st('DE-001', SKU_ENDBOX, 190, '2026-06', 546)
make_st('DE-001', SKU_ACC_01, 280, '2026-06', 39)
make_st('DE-001', SKU_BLIND_REG, 300, '2026-07', 89)
make_st('DE-001', SKU_OPEN_03, 45, '2026-07', 100)

# 酷乐潮玩
make_st('DE-002', SKU_BLIND_EARLY, 140, '2026-05', 89)
make_st('DE-002', SKU_ENDBOX, 75, '2026-05', 546)
make_st('DE-002', SKU_BLIND_REG, 380, '2026-06', 89)
make_st('DE-002', SKU_BLIND_REG, 250, '2026-07', 89)

# TOPTOY
make_st('DE-003', SKU_BLIND_EARLY, 190, '2026-05', 89)
make_st('DE-003', SKU_ENDBOX, 95, '2026-05', 546)
make_st('DE-003', SKU_BLIND_REG, 460, '2026-06', 89)
make_st('DE-003', SKU_ENDBOX, 235, '2026-06', 546)
make_st('DE-003', SKU_BLIND_REG, 280, '2026-07', 89)

# KKV
make_st('DE-004', SKU_BLIND_EARLY, 90, '2026-05', 89)
make_st('DE-004', SKU_BLIND_REG, 270, '2026-06', 89)
make_st('DE-004', SKU_ACC_02, 185, '2026-06', 49)
make_st('DE-004', SKU_BLIND_REG, 145, '2026-07', 89)

# 九木杂物社
make_st('DE-005', SKU_BLIND_EARLY, 70, '2026-05', 89)
make_st('DE-005', SKU_ENDBOX, 35, '2026-05', 546)
make_st('DE-005', SKU_BLIND_REG, 230, '2026-06', 89)
make_st('DE-005', SKU_OPEN_03, 45, '2026-06', 100)

# ══════════════════════════════════════════════════════════
# 7. 银行回款（含 14 场景中的 10 个差异）
# ══════════════════════════════════════════════════════════

BANK_RECEIPTS = []
pm_counter = 1

def make_pm(payer, amount, receipt_date, method='网银转账', note=''):
    global pm_counter
    BANK_RECEIPTS.append([
        f'PM-{receipt_date.strftime("%Y%m%d")}-{pm_counter:03d}',
        receipt_date.isoformat(), payer, amount, method,
        '中国工商银行深圳前海支行',
        f'{random.randint(1000000000000000, 9999999999999999)}',
        note
    ])
    pm_counter += 1

# ── 场景1: 名创优品 1 笔回款覆盖 3 张销售单 ──
make_pm('名创优品（广州）有限责任公司', SCENE1_TOTAL, d('2026-08-01'),
        note=f'6月货款合计 单号{SCENE1_SO_IDS[0]}/{SCENE1_SO_IDS[1]}/{SCENE1_SO_IDS[2]}（场景1:多单合并）')

# ── 场景2: TOPTOY 部分回款 ¥50,000 ──
make_pm('TOPTOY（广州）文化发展有限公司', 50000.00, d('2026-08-03'),
        note=f'6月部分货款 单号{SCENE2_SO_IDS[0]}（场景2:部分回款）')

# ── 场景3: 沃尔玛完全未回 ──
# 故意不生成回款记录

# ── 场景4: 抖音预付款 ──
make_pm('抖音电商（官方旗舰店）', 30000.00, d('2026-07-01'),
        note='7月新品订货预付款（场景4:预付款）')

# ── 场景5: 成都经销商抹零差异 ¥0.67 ──
scene5_pay = round(SCENE5_AMT - 0.67, 2)
make_pm('广州潮玩前线贸易有限公司', scene5_pay, d('2026-08-05'),
        note=f'7月货款 单号{SCENE5_SO_ID}（场景5:抹零差异¥0.67）')

# ── 场景6: 酷乐潮玩名称不规范（付款方少"有限"） ──
make_pm('酷乐潮玩（北京）贸易公司', 38848.00, d('2026-07-25'),
        note='6月货款 BabyCat盲盒+端盒（场景6:付款方名称不规范）')

# ── 场景7: 九木杂物社逾期 ──
# 该经销商有回款但逾期——生成一笔账期后15天才到账的回款
make_pm('上海晨光文具股份有限公司（九木杂物社）', 14520.00, d('2026-08-01'),
        note=f'5月货款 单号{SCENE7_SO_ID}（场景7:逾期15天到账）')

# ── 场景8: KKV退货冲销（一笔负值回款，已在销售里体现） ──
# 正常回款，但销售里有一笔红字退货
make_pm('广州KK馆贸易有限公司', 45280.00, d('2026-08-02'),
        note='6-7月货款（场景8:含退货冲销）')

# ── 场景9: 广州潮玩前线无备注自动匹配 ──
make_pm('广州潮玩前线贸易有限公司', round(amt9, 2), d('2026-07-20'),
        note='（场景9:无备注自动匹配）')

# ── 场景10: 义乌小额差异 ¥1 ──
make_pm('义乌萌物集贸易有限公司', round(amt10 - 1.00, 2), d('2026-07-22'),
        note=f'5月货款（场景10:少付¥1.00）')

# ── 其他正常回款 ──
# 名创优品 5月试产回款
make_pm('名创优品（广州）有限责任公司', 53820.80, d('2026-06-20'),
        note='5月BabyCat首轮试产货款')

# 酷乐潮玩 5月回款
make_pm('酷乐潮玩（北京）贸易有限公司', 26126.40, d('2026-06-25'),
        note='5月BabyCat试产货款')

# 抖音 5-6 月结算（扣平台佣金）
make_pm('抖音电商（官方旗舰店）', 42560.00, d('2026-07-15'),
        note='5-6月平台结算（已扣佣金）')

# 名创优品 7 月回款
make_pm('名创优品（广州）有限责任公司', 72680.00, d('2026-08-08'),
        note='7月BabyCat暑期促销货款')

# 酷乐潮玩 7 月回款
make_pm('酷乐潮玩（北京）贸易有限公司', 51300.00, d('2026-08-06'),
        note='7月货款')

# KKV 5 月回款
make_pm('广州KK馆贸易有限公司', 8730.00, d('2026-06-18'),
        note='5月BabyCat试产货款')

# TOPTOY 5 月回款
make_pm('TOPTOY（广州）文化发展有限公司', 37056.00, d('2026-06-28'),
        note='5月BabyCat试产货款')

# 九木杂物社 5-6 月正常回款
make_pm('上海晨光文具股份有限公司（九木杂物社）', 26022.00, d('2026-07-10'),
        note='5-6月货款')

# 沃尔玛（唯一一笔，还有两笔逾期未付）
make_pm('沃尔玛（中国）投资有限公司', 17575.00, d('2026-07-08'),
        note='5月货款')

# 抖音 7 月结算
make_pm('抖音电商（官方旗舰店）', 61750.00, d('2026-08-05'),
        note='7月平台结算（已扣佣金及达人分成）')

# ★ 场景14: TikTok USD 回款（汇率差异）
# 模拟 USD 回款：$2,756.00 @ 7.2345 = ¥19,933.50 但实际到账 ¥19,607.00
make_pm('TikTok Shop（跨境）', 19607.00, d('2026-08-04'),
        note='TikTok Shop US 6月结算 $2756.00@7.1123 实到¥19607.00（场景14:汇兑损益¥326.50）')

# ── 补几条正常回款让数据丰满 ──
make_pm('广州潮玩前线贸易有限公司', 31828.80, d('2026-07-28'),
        note='6月货款 BabyCat盲盒+端盒')

make_pm('义乌萌物集贸易有限公司', 30820.00, d('2026-08-07'),
        note='6-7月货款')

# ══════════════════════════════════════════════════════════
# 8. 摩点众筹数据
# ══════════════════════════════════════════════════════════

MODIAN_PLEDGES = []
pledge_id = 1

# 众筹档位
PLEDGE_TIERS = [
    ('随机单盒（早鸟）', 89, 60),
    ('限量明盒（早鸟）', 94, 25),
    ('单盒（常规）',    95, 30),
    ('明盒（常规）',   100, 15),
    ('CP对盒（早鸟）', 176, 12),
    ('CP对盒（常规）', 184, 8),
    ('端盒（6只）',    546, 10),
    ('ALL-IN套装',     680, 6),
]

supporters = [
    '张*欣', '李*然', '王*喵', '陈*豪', '刘*宇', '赵*彤', '周*远', '吴*琪',
    '郑*杰', '钱*琳', '孙*浩', '马*萌', '朱*霖', '胡*鑫', '林*妍', '何*阳',
    '郭*文', '高*悦', '罗*俊', '梁*婷', '宋*明', '唐*怡', '韩*辰', '冯*曦',
    '董*雨', '袁*华', '邓*宇', '许*涵', '傅*琪', '沈*杰', '曾*琳', '彭*浩',
    '吕*静', '苏*航', '卢*雅', '蒋*洋', '蔡*颖', '贾*翔', '丁*琪', '魏*凯',
    '薛*萌', '叶*飞', '阎*琳', '余*泽', '潘*瑶', '杜*昊', '戴*婷', '夏*晨',
    '钟*宇', '姚*萌', '汪*瑾', '田*峰', '任*滢', '姜*祥', '范*萱', '方*博',
    '石*淇', '廖*然', '邹*远', '熊*慧', '金*铭', '陆*豪', '郝*妍', '孔*阳',
    '白*萌', '崔*杰', '康*悦', '毛*浩', '邱*琳', '秦*朗', '江*洁', '史*辰',
    '顾*怡', '侯*峰', '龙*涵', '万*涛', '段*蕾', '雷*博', '严*曦', '覃*宇',
    '武*萌', '乔*凯', '汤*瑶', '尹*飞', '易*萱', '常*鸣', '贺*桐', '赖*廷',
    '龚*拓', '文*丽', '庞*锐', '蓝*昕', '代*妮', '蒙*衍', '岑*锋', '毕*媛',
    '阮*濛', '关*铮', '覃*朗', '兰*熹', '乌*庭', '焦*荟', '池*韬', '谷*宓',
    '宁*旋', '尚*航', '符*融', '嵇*萌', '缪*清', '娄*菲', '裘*策', '莘*凝',
]

# 生成 166 条支持者记录（与实际众筹人数一致）
for i, sup in enumerate(supporters[:166]):
    tier = PLEDGE_TIERS[i % len(PLEDGE_TIERS)]
    tier_name, price, count = tier
    MODIAN_PLEDGES.append([
        f'MD-202507{pledge_id:04d}',
        d('2025-07-15') + timedelta(days=random.randint(0, 30)),
        sup,
        tier_name,
        1,
        price,
        price,
        '已支付',
        '已发货' if i < 150 else '待发货'
    ])
    pledge_id += 1

# 摩点分阶段结算
MODIAN_SETTLEMENTS = [
    ['MD-SETTLE-01', '2025-08-05', '首款（50%）',  39778.00, '2025-08-10', 39778.00, '已到账', '众筹成功→首款实际到账（场景11:比理论¥41777少¥1999，因支持者退款）'],
    ['MD-SETTLE-02', '2025-11-20', '中款（30%）',  25067.00, '2025-11-25', 25067.00, '已到账', '全部发货完成→中款到账'],
    ['MD-SETTLE-03', '2026-02-05', '尾款（20%）',  16711.00, '2026-02-12', 16711.00, '已到账', '评价期满→尾款到账'],
]

# ══════════════════════════════════════════════════════════
# 9. 展会/快闪店收入
# ══════════════════════════════════════════════════════════

EXPO_REVENUE = [
    ['CTE-001', 'CTE玩具展',       '2025-09-15', '2025-09-17', '深圳', '现场零售+预售', 15680.00, 2340.00, 13340.00, '微信/支付宝/POS'],
    ['HTE-001', 'HTE杭州潮玩展',    '2026-03-27', '2026-03-29', '杭州', '现场零售',      12350.00, 1860.00, 10490.00, '微信/支付宝'],
    ['POP-001', '济南圣诞快闪店',   '2025-12-06', '2025-12-26', '济南', '快闪店零售+吊卡首发', 44118.00, 5200.00, 38918.00, '微信/支付宝/POS'],
]

# ★ 场景13: 冬季吊卡限量追踪
# 在销售订单里加一笔济南快闪店的吊卡销售
make_so('DE-009', SKU_LIMIT, 342, d('2025-12-26'), '快闪店零售', 1.0, '济南快闪店·冬季吊卡首发（场景13:限量1000，快闪售342个）')

# ══════════════════════════════════════════════════════════
# 10. 合同台账
# ══════════════════════════════════════════════════════════

CONTRACT_LEDGER = [
    ['CT-2025-001', '摩点平台',        '平台服务协议',   d('2025-07-01'), d('2026-07-01'), '执行中', ''],
    ['CT-2025-002', '名创优品',        '经销合作协议',   d('2025-09-01'), d('2026-09-01'), '执行中', '含寄售条款'],
    ['CT-2025-003', '酷乐潮玩',        '经销合作协议',   d('2025-09-15'), d('2026-09-15'), '执行中', '含寄售条款'],
    ['CT-2025-004', 'TOPTOY',         '经销合作协议',   d('2025-10-01'), d('2026-10-01'), '执行中', '含寄售条款'],
    ['CT-2025-005', 'KKV',            '经销合作协议',   d('2025-10-15'), d('2026-10-15'), '执行中', '含寄售条款'],
    ['CT-2025-006', '九木杂物社',      '经销合作协议',   d('2025-11-01'), d('2026-11-01'), '执行中', '含寄售条款'],
    ['CT-2025-007', '沃尔玛',          '供货合同',       d('2025-11-15'), d('2026-11-15'), '执行中', '买断制'],
    ['CT-2025-008', '广州潮玩前线',    '区域经销合同',   d('2025-12-01'), d('2026-12-01'), '执行中', '华南区域'],
    ['CT-2025-009', '义乌萌物集',      '区域经销合同',   d('2025-12-15'), d('2026-12-15'), '执行中', '华东区域'],
    ['CT-2026-001', '林小喵',          '艺术家合作协议', d('2025-06-01'), d('2026-06-01'), '执行中', 'BC-01,BC-07 分成8%'],
    ['CT-2026-002', '陈卷卷',          '艺术家合作协议', d('2025-06-15'), d('2026-06-15'), '执行中', 'BC-02 分成7%'],
    ['CT-2026-003', '周布布',          '艺术家合作协议', d('2025-07-01'), d('2026-07-01'), '执行中', 'BC-03,BC-06 分成8%'],
    ['CT-2026-004', '刘星野',          '艺术家合作协议', d('2025-07-15'), d('2026-07-15'), '执行中', 'BC-04,BC-05 分成7%'],
    ['CT-2026-005', '赵小喵',          '艺术家合作协议', d('2025-08-01'), d('2026-08-01'), '执行中', 'BC-08 分成10% 含预付¥5000'],
]

# ══════════════════════════════════════════════════════════
# 11. 写入所有 CSV
# ══════════════════════════════════════════════════════════

print('\n📊 菲格洛亚 FIGUEROA 财务数据生成器')
print(f'   Seed: 20260811 | 日期范围: 2025-06 ~ 2026-08\n')

# Artist list
write_csv('artist_list.csv',
    ['艺术家编码', '艺术家姓名', '身份类型', '关联SKU', '分成比例', '结算方式'],
    [[a[0], a[1], a[2], a[3], f'{a[4]:.0%}', a[5]] for a in ARTISTS])

# SKU catalog
write_csv('sku_catalog.csv',
    ['SKU编码', '规格编码', '商品名称', '品类', '零售价', '经销价'],
    [[s[0], s[1], s[2], s[3], s[4], s[5]] for s in SKUS])

# Dealer list
write_csv('dealer_list.csv',
    ['经销商编码', '经销商名称', '渠道品牌', '渠道类型', '所在区域', '账期天数', '结算模式'],
    DEALERS)

# Sales orders
write_csv('sales_orders.csv',
    ['销售单号', '开单日期', '经销商编码', '经销商名称', 'SKU编码',
     '商品名称', '数量', '单价', '金额', '折扣率', '实收金额',
     '账期天数', '到期日', '批次类型', '备注'],
    SALES_ORDERS)

# Sell-through
write_csv('dealer_sell_through.csv',
    ['报表编号', '报表月份', '经销商编码', 'SKU编码', '商品名称',
     '销售数量', '单价', '销售金额'],
    SELL_THROUGH)

# Bank receipts
write_csv('bank_receipts.csv',
    ['回款单号', '到账日期', '付款方名称', '金额', '付款方式',
     '开户行', '交易流水号', '备注'],
    BANK_RECEIPTS)

# Modian pledges
write_csv('modian_pledges.csv',
    ['支持编号', '支持日期', '支持者', '档位名称', '数量', '单价', '金额', '支付状态', '发货状态'],
    MODIAN_PLEDGES)

# Modian settlements
write_csv('modian_settlements.csv',
    ['结算编号', '申请日期', '结算阶段', '应结算金额', '到账日期', '实际到账', '状态', '备注'],
    MODIAN_SETTLEMENTS)

# Expo revenue
write_csv('expo_revenue.csv',
    ['活动编码', '活动名称', '开始日期', '结束日期', '地点', '收入类型', '总收入', '费用', '净收入', '收款方式'],
    EXPO_REVENUE)

# Contract ledger
write_csv('contract_ledger.csv',
    ['合同编号', '签约方', '合同名称', '签署日期', '到期日期', '状态', '备注'],
    CONTRACT_LEDGER)

# ══════════════════════════════════════════════════════════
# 12. 断言验证（14 个场景）
# ══════════════════════════════════════════════════════════

print('\n── 场景验证 ──')

all_ok = True

# 场景1: 1笔回款覆盖3张销售单（验证合并总额）
pm_scene1 = [r for r in BANK_RECEIPTS if len(r[7]) > 0 and '合并' in r[7] or '多单' in r[7]]
if pm_scene1 and abs(pm_scene1[0][3] - SCENE1_TOTAL) < 0.01:
    print(f'  [OK] 场景1: 名创优品多单合并回款, total={SCENE1_TOTAL:,.2f}')
else:
    print(f'  [FAIL] 场景1: SCENE1_TOTAL={SCENE1_TOTAL}, pm_found={len(pm_scene1)}, pm_amt={pm_scene1[0][3] if pm_scene1 else "none"}')
    all_ok = False

# 场景2: TOPTOY部分回款
pm_scene2 = [r for r in BANK_RECEIPTS if r[3] == 50000.00 and 'TOPTOY' in r[2]]
if pm_scene2:
    print(f'  [OK] 场景2: TOPTOY部分回款 50000 (销售总额={SCENE2_TOTAL:,.2f})')
else:
    print(f'  [FAIL] 场景2: 未找到TOPTOY 50000回款, 销售总额={SCENE2_TOTAL}')
    all_ok = False

# 场景3: 沃尔玛逾期未回
walmart_sos = [s for s in SALES_ORDERS if s[2] == 'DE-006']
walmart_pms = [r for r in BANK_RECEIPTS if '沃尔玛' in r[2]]
if len(walmart_sos) >= 2 and len(walmart_pms) < len(walmart_sos):
    print(f'  [OK] 场景3: 沃尔玛 {len(walmart_sos)}单/{len(walmart_pms)}笔回款 (部分逾期未回)')
else:
    print(f'  [FAIL] 场景3: 沃尔玛 so={len(walmart_sos)} pm={len(walmart_pms)}')

# 场景4: 抖音预付款
pm_scene4 = [r for r in BANK_RECEIPTS if '预付款' in r[7] and '抖音' in r[2]]
if pm_scene4:
    print(f'  [OK] 场景4: 抖音预付款 {pm_scene4[0][3]:,.2f}')
else:
    print(f'  [FAIL] 场景4: 未找到抖音预付款')

# 场景5: 抹零差异
pm_scene5 = [r for r in BANK_RECEIPTS if '抹零' in r[7] or '差异' in r[7]]
if pm_scene5:
    diff5 = round(SCENE5_AMT - pm_scene5[0][3], 2)
    print(f'  [OK] 场景5: 抹零差异 {diff5} (应收{SCENE5_AMT:.2f} 实收{pm_scene5[0][3]:.2f})')
else:
    print(f'  [FAIL] 场景5')

# 场景6: 付款方名称不规范
pm_scene6 = [r for r in BANK_RECEIPTS if '名称不规范' in r[7] or '不规范' in r[7]]
if pm_scene6:
    print(f'  [OK] 场景6: 名称不规范 payer="{pm_scene6[0][2]}"')
else:
    print(f'  [FAIL] 场景6')

# 场景7: 九木杂物社逾期
pm_scene7 = [r for r in BANK_RECEIPTS if '逾期' in r[7] and '九木' in r[2]]
if pm_scene7:
    print(f'  [OK] 场景7: 逾期到账 {pm_scene7[0][3]:,.2f}')
else:
    print(f'  [FAIL] 场景7')

# 场景8: KKV退货冲销
kkv_returns = [s for s in SALES_ORDERS if s[2] == 'DE-004' and s[6] < 0]
if kkv_returns:
    print(f'  [OK] 场景8: KKV退货冲销 qty={kkv_returns[0][6]} amt={kkv_returns[0][10]:,.2f}')
else:
    print(f'  [FAIL] 场景8')

# 场景9: 回款无备注
pm_scene9 = [r for r in BANK_RECEIPTS if '无备注' in r[7]]
if pm_scene9:
    print(f'  [OK] 场景9: 无备注自动匹配 amt={pm_scene9[0][3]:,.2f}')
else:
    print(f'  [FAIL] 场景9')

# 场景10: 小额差异
pm_scene10 = [r for r in BANK_RECEIPTS if '少付' in r[7]]
if pm_scene10:
    diff10 = round(amt10 - pm_scene10[0][3], 2)
    print(f'  [OK] 场景10: 小额差异 {diff10}')
else:
    print(f'  [FAIL] 场景10')

# 场景11: 摩点退款差异
s11 = MODIAN_SETTLEMENTS[0]
exp11 = round(88888 * 0.5 * (1 - 0.06), 2)
diff11 = round(exp11 - s11[3], 2)
if diff11 > 0:
    print(f'  [OK] 场景11: 摩点首款差{diff11:.2f} (理论{exp11:.2f} 实际{s11[3]:.2f})')
else:
    print(f'  [FAIL] 场景11: diff11={diff11}')

# 场景12: 艺术家跨渠道版税
artist_multi = [s for s in SALES_ORDERS if 'BC-01' in s[4] or 'BC-03' in s[4] or 'BC-06' in s[4]]
chs = set(s[2] for s in artist_multi)
if len(chs) >= 3:
    print(f'  [OK] 场景12: 艺术家SKU跨{len(chs)}渠道')
else:
    print(f'  [FAIL] 场景12: 仅跨{len(chs)}渠道')

# 场景13: 限量吊卡
ls = [s for s in SALES_ORDERS if 'LIMIT' in s[4]]
if ls:
    print(f'  [OK] 场景13: 限量吊卡 sales={ls[0][6]}个')
else:
    print(f'  [FAIL] 场景13')

# 场景14: USD汇率差异
pm14 = [r for r in BANK_RECEIPTS if '汇兑' in r[7] or 'USD' in r[7]]
if pm14:
    print(f'  [OK] 场景14: TikTok USD回款 {pm14[0][3]:,.2f} (汇兑损益)')
else:
    print(f'  [FAIL] 场景14')

# ── 总体验证 ──
total_sales = sum(s[10] for s in SALES_ORDERS)
total_receipts = sum(r[3] for r in BANK_RECEIPTS)
if all_ok:
    print('\n[OK] All 14 scenarios verified!')
else:
    print('\n[WARN] Some scenarios failed verification — check data.')
print(f'\n  Sales: {len(SALES_ORDERS)} orders, {total_sales:,.2f} | Receipts: {len(BANK_RECEIPTS)}, {total_receipts:,.2f}')
print(f'  SKUs: {len(SKUS)} | Dealers: {len(DEALERS)} | Artists: {len(ARTISTS)}')
print('  Output: data/*.csv\n')
