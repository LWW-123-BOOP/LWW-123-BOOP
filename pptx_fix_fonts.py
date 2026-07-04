# -*- coding: utf-8 -*-
"""
PPTX 字体修复脚本（保持可编辑）
解决：pptx 发到别人电脑乱码 / 手机微信打开弹退。
原因：原文件用了 PMingLiU(新细明体)、微軟正黑體 等地域限定字体，对方设备没装就乱码。
本脚本把这类字体替换成 Windows 自带的微软雅黑，保持 pptx 可编辑且不乱码。

用法:
    python pptx_fix_fonts.py 输入.pptx [输出.pptx]
    python pptx_fix_fonts.py            # 不带参数则弹出文件选择框
"""
import sys, os, zipfile, shutil, re

# 需要替换掉的"易缺字体" -> 通用字体（Win7+ 自带）
FONT_MAP = {
    "PMingLiU":  "微软雅黑",   # 新细明体（繁体）
    "MingLiU":   "微软雅黑",   # 细明体（繁体）
    "微軟正黑體":"微软雅黑",   # Microsoft JhengHei
    "微软正黑体":"微软雅黑",
    "Microsoft JhengHei": "微软雅黑",
    "新細明體":  "微软雅黑",
    "細明體":    "微软雅黑",
    "DFKai-SB":  "微软雅黑",   # 标楷体（繁体）
    "楷体_GB2312":"微软雅黑",
}

def pick_file():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw()
        return filedialog.askopenfilename(
            title="选择要修复字体的 PPTX",
            filetypes=[("PowerPoint", "*.pptx"), ("所有文件", "*.*")])
    except Exception:
        return None

def main():
    src = sys.argv[1] if len(sys.argv) > 1 else pick_file()
    if not src or not os.path.exists(src):
        print("未提供有效文件"); sys.exit(1)
    dst = sys.argv[2] if len(sys.argv) > 2 else os.path.splitext(src)[0] + "_修复.pptx"

    zin = zipfile.ZipFile(src, 'r')
    zout = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)

    replaced_count = 0
    touched_files = []

    for item in zin.infolist():
        raw = zin.read(item.filename)
        # 只处理 xml 文本
        if item.filename.lower().endswith(('.xml', '.rels')):
            try:
                text = raw.decode('utf-8')
            except UnicodeDecodeError:
                zout.writestr(item, raw); continue
            original = text
            # 替换 typeface="..." 中的字体名
            def repl(m):
                name = m.group(1)
                return 'typeface="%s"' % FONT_MAP.get(name, name)
            text = re.sub(r'typeface="([^"]+)"', repl, text)
            if text != original:
                # 统计实际发生了多少替换
                replaced_count += sum(
                    1 for k in FONT_MAP
                    if ('typeface="%s"' % k) in original
                )
                touched_files.append(item.filename)
            raw = text.encode('utf-8')
        zout.writestr(item, raw)

    zin.close(); zout.close()

    print(f"✅ 字体修复完成: {dst}")
    print(f"   替换次数: {replaced_count}")
    print(f"   涉及文件: {len(touched_files)} 个")
    for f in touched_files:
        print(f"     - {f}")
    print(f"   替换为: 微软雅黑 (Win7+ 自带，几乎所有人都有)")
    print(f"\n现在把这个 _修复.pptx 发给别人，对方能编辑且不乱码。")

if __name__ == "__main__":
    main()
