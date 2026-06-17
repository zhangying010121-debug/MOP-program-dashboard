import re, sys, io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

HTML_PATH = r'C:\Users\zhangying\WorkBuddy\2026-06-09-14-21-08\index.html'

with open(HTML_PATH, 'r', encoding='utf-8') as f:
    html = f.read()

print(f"原始大小: {len(html)} 字符")

# 删除旧的 localStorage 注入脚本块（整个 <script> 块 7-16行）
# 识别: 开头包含 "// 嵌入业绩数据" 的 script 块
old_inject_pattern = re.compile(
    r'<script>\s*// 嵌入业绩数据.*?</script>\s*',
    re.DOTALL
)

if old_inject_pattern.search(html):
    html = old_inject_pattern.sub('', html, count=1)
    print("✅ 已删除旧的 localStorage 注入脚本块")
else:
    print("⚠️ 未找到旧注入脚本，跳过")

print(f"更新后大小: {len(html)} 字符")

with open(HTML_PATH, 'w', encoding='utf-8') as f:
    f.write(html)

print("✅ index.html 已保存")

# 验证
if 'localStorage.setItem' in html[:500]:
    print("⚠️ 警告: 文件头部仍有 localStorage.setItem!")
else:
    print("✅ 验证通过：旧注入脚本已清除")
    
# 检查版本和数据
store_count = html.count("store_no:'95S")
guide_count = html.count("storeCode:'95S")
print(f"内置门店条目: {store_count}，导购条目: {guide_count}")
