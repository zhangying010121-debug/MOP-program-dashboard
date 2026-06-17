import re, sys, io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 读取当前HTML
with open(r'C:\Users\zhangying\WorkBuddy\2026-06-09-14-21-08\index.html', 'r', encoding='utf-8') as f:
    html = f.read()

# 读取新的数据JS
with open(r'C:\Users\zhangying\WorkBuddy\2026-06-09-14-21-08\store_js.txt', 'r', encoding='utf-8') as f:
    new_store_js = f.read()

with open(r'C:\Users\zhangying\WorkBuddy\2026-06-09-14-21-08\guide_js.txt', 'r', encoding='utf-8') as f:
    new_guide_js = f.read()

# ---------- 替换 loadSampleData 函数体内的数据 ----------
# 找到函数开始和结束，替换函数内的数组内容
store_func_pattern = re.compile(
    r'(function loadSampleData\(\) \{)[^}]*(\];[\s\n]*saveToStorage\(\);[\s\n]*refreshAll\(\);[\s\n]*showToast\([^\)]+\);[\s\n]*\})',
    re.DOTALL
)

new_store_func = (
    'function loadSampleData() {\n  ' +
    new_store_js +
    '\n  refreshAll();\n  showToast(\'✅ 截至6.16门店数据已载入（48家门店）\');\n}'
)

if store_func_pattern.search(html):
    html = store_func_pattern.sub(new_store_func, html)
    print("✅ loadSampleData 替换成功")
else:
    print("❌ loadSampleData 未找到匹配模式，尝试手动替换...")
    # 手动找到函数范围
    start_marker = 'function loadSampleData() {'
    end_marker_after = 'showToast'
    si = html.index(start_marker)
    # 找到函数结束：最近一个以 } 为单独一行的位置
    func_region = html[si:si+10000]
    # 找 toast 调用后的 }
    toast_idx = func_region.index("showToast('✅ 截至6.11")
    brace_idx = func_region.index('\n}', toast_idx)
    end_idx = si + brace_idx + 2
    html = html[:si] + new_store_func + html[end_idx:]
    print("✅ loadSampleData 手动替换成功")

# ---------- 替换 loadPresetGuideData 函数体内的数据 ----------
guide_func_pattern = re.compile(
    r'(function loadPresetGuideData\(\) \{).*?(\})',
    re.DOTALL
)

new_guide_func = (
    'function loadPresetGuideData() {\n  ' +
    new_guide_js +
    '\n}'
)

if guide_func_pattern.search(html):
    html = guide_func_pattern.sub(new_guide_func, html, count=1)
    print("✅ loadPresetGuideData 替换成功")
else:
    print("❌ loadPresetGuideData 未找到匹配模式")

# ---------- 版本号升至 v2.4 ----------
html = html.replace("version: 'v2.3'", "version: 'v2.4'", 10)
html = html.replace("version !== 'v2.3'", "version !== 'v2.4'", 10)
html = html.replace("version: \"v2.3\"", "version: \"v2.4\"", 10)
html = html.replace('v2.3-截至6.11', 'v2.4-截至6.16')
html = html.replace('截至6.11门店数据已载入', '截至6.16门店数据已载入')
print("✅ 版本号已升至 v2.4")

# ---------- 写回文件 ----------
with open(r'C:\Users\zhangying\WorkBuddy\2026-06-09-14-21-08\index.html', 'w', encoding='utf-8') as f:
    f.write(html)

print(f"✅ index.html 已更新，总字符数: {len(html)}")

# 验证
store_count = html.count("store_no:'95S")
guide_count = html.count("storeCode:'95S")
print(f"验证：HTML中门店条目数={store_count}，导购条目数={guide_count}")
