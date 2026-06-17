import pandas as pd, sys, io, json, os

DEFAULT_REGION = '张瑛'
OUT_DIR = r'C:\Users\zhangying\WorkBuddy\2026-06-09-14-21-08'

# ========== 解析门店数据 ==========
df_store = pd.read_excel(r'C:/Users/zhangying/Downloads/workbuddy-小程序月度业绩达成看板-截至6.16.xlsx', header=1, engine='openpyxl')

stores = []
for _, row in df_store.iterrows():
    store_no = str(row.get('门店编号', '') or '').strip()
    store_name = str(row.get('门店名称', '') or '').strip()
    if not store_no or store_no == 'nan' or store_name in ('nan', '合计', '小计', ''):
        continue
    region = str(row.get('区经', '') or '').strip()
    if not region or region == 'nan':
        region = DEFAULT_REGION
    store_type = str(row.get('正价/奥莱', '') or '').strip()
    if store_type == 'nan':
        store_type = ''

    def safe_float(v):
        try:
            return round(float(v), 2) if v is not None and str(v) not in ('nan', '') else 0.0
        except:
            return 0.0

    target = safe_float(row.get('6月目标'))
    achieved = safe_float(row.get('6月达成'))
    actual_receipt = safe_float(row.get('实收金额'))
    after_sales = safe_float(row.get('售后金额'))

    stores.append({
        'store_no': store_no,
        'store_name': store_name,
        'store_type': store_type,
        'region': region,
        'target': target,
        'achieved': achieved,
        'actual_receipt': actual_receipt if actual_receipt else None,
        'after_sales': after_sales if after_sales else None,
    })

print(f"门店数据: {len(stores)} 条")

# ========== 解析导购数据 ==========
df_guide = pd.read_excel(r'C:/Users/zhangying/Downloads/workbuddy-导购实收金额汇总表-0616.xlsx', engine='openpyxl', header=None)
guides = []
current_store = None
for _, row in df_guide.iterrows():
    col1 = str(row.iloc[1] if len(row) > 1 else '').strip()
    col2 = str(row.iloc[2] if len(row) > 2 else '').strip()
    col3 = row.iloc[3] if len(row) > 3 else None

    if col1 in ('导购门店编号', 'nan', '') and col2 in ('导购名称', 'nan', ''):
        continue

    if col1 and col1 != 'nan' and col2 in ('nan', ''):
        current_store = col1
        continue

    if col2 and col2 != 'nan' and current_store:
        try:
            amount = round(float(col3), 2) if col3 and str(col3) not in ('nan', '') else 0.0
        except:
            amount = 0.0
        if amount > 0:
            guides.append({
                'storeCode': current_store,
                'guideName': col2,
                'amount': amount
            })

print(f"导购数据: {len(guides)} 条")

# ========== 生成门店 JS 代码 ==========
store_lines = []
for s in stores:
    ar = str(s['actual_receipt']) if s['actual_receipt'] is not None else 'null'
    af = str(s['after_sales']) if s['after_sales'] is not None else 'null'
    sname = s['store_name'].replace('\\', '\\\\').replace('"', '\\"')
    line = '{ ' + (
        "store_no:'" + s['store_no'] + "', " +
        'store_name:"' + sname + '", ' +
        "store_type:'" + s['store_type'] + "', " +
        "region:'" + s['region'] + "', " +
        "target:" + str(s['target']) + ", " +
        "achieved:" + str(s['achieved']) + ", " +
        "actual_receipt:" + ar + ", " +
        "after_sales:" + af
    ) + ' }'
    store_lines.append(line)

store_js = 'storeData = [\n' + ',\n'.join(['  ' + l for l in store_lines]) + '\n];\n  saveToStorage();'

# ========== 生成导购 JS 代码 ==========
guide_lines = []
for g in guides:
    gname = g['guideName'].replace('\\', '\\\\').replace("'", "\\'")
    line = "  { storeCode:'" + g['storeCode'] + "', guideName:'" + gname + "', amount:" + str(g['amount']) + " }"
    guide_lines.append(line)

guide_js = 'guideRawData = [\n' + ',\n'.join(guide_lines) + '\n];\n  buildGuideMap();\n  renderGuideSection();\n  saveGuideToStorage();'

store_out = os.path.join(OUT_DIR, 'store_js.txt')
guide_out = os.path.join(OUT_DIR, 'guide_js.txt')

with open(store_out, 'w', encoding='utf-8') as f:
    f.write(store_js)
with open(guide_out, 'w', encoding='utf-8') as f:
    f.write(guide_js)

print(f"JS代码已写入:\n  {store_out}\n  {guide_out}")
print("门店第一条:", stores[0] if stores else '无')
print("导购第一条:", guides[0] if guides else '无')
