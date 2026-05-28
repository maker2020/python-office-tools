#!/usr/bin/env python3
"""
🤖 AI 办公助手 - 一个脚本搞定所有文件处理
============================================
功能：Excel合并拆分 / PDF处理 / 图片压缩 / 数据清洗 / 文件重命名 / 二维码生成
用法：python ai_office.py
依赖：pip install pandas openpyxl PyPDF2 Pillow qrcode reportlab
作者：一个想请你喝杯咖啡的程序员
"""

import os, sys, glob, math, re, io, json
from pathlib import Path
from datetime import datetime

# ===== 颜色输出 =====
def c(text, color):
    colors = {'red':'\033[91m','green':'\033[92m','yellow':'\033[93m',
              'blue':'\033[94m','cyan':'\033[96m','bold':'\033[1m','end':'\033[0m'}
    return f"{colors.get(color,'')}{text}{colors['end']}"

def banner():
    print(c("""
╔══════════════════════════════════════════════╗
║  🤖  AI 办公助手 v1.0                       ║
║  一个脚本搞定所有文件处理                    ║
╚══════════════════════════════════════════════╝""", 'cyan'))

def menu():
    print(c("\n📋 功能列表:", 'bold'))
    tools = [
        ("1", "📊 Excel 合并", "多个Excel文件合并为一个"),
        ("2", "📊 Excel 拆分", "大Excel按行数拆分"),
        ("3", "📄 PDF 合并", "多个PDF合并为一个"),
        ("4", "📄 PDF 加水印", "批量添加文字水印"),
        ("5", "🖼️  图片压缩", "批量压缩图片"),
        ("6", "🖼️  图片缩放", "批量调整尺寸"),
        ("7", "📁 批量重命名", "按规则重命名文件"),
        ("8", "🔲 二维码生成", "文本/WiFi/名片二维码"),
        ("9", "🧹 数据清洗", "CSV/Excel去重填空"),
        ("0", "📝 文字提取", "PDF/图片OCR提取文字"),
    ]
    for num, name, desc in tools:
        print(f"  {c(num, 'yellow')}. {name}  {c(desc, 'cyan')}")
    print(f"\n  {c('q', 'red')}. 退出")
    return input(c("\n请选择功能 (0-9): ", 'bold')).strip()

# ===== 工具函数 =====
def pick_folder():
    folder = input("📁 文件夹路径: ").strip().strip('"').strip("'")
    if not os.path.isdir(folder):
        print(c("❌ 文件夹不存在", 'red')); return None
    return folder

def pick_file(ext_filter=None):
    filepath = input("📄 文件路径: ").strip().strip('"').strip("'")
    if not os.path.isfile(filepath):
        print(c("❌ 文件不存在", 'red')); return None
    return filepath

def find_files(folder, exts):
    files = []
    for ext in exts:
        files.extend(glob.glob(os.path.join(folder, ext)))
        files.extend(glob.glob(os.path.join(folder, ext.upper())))
    return sorted(set(files))

def confirm(msg="确认执行？"):
    return input(f"\n{msg} (y/N): ").strip().lower() == 'y'

def done_msg(output, count=None):
    print(c(f"\n✅ 完成！", 'green'))
    if count: print(f"   📊 处理: {count} 个文件")
    print(f"   📁 输出: {output}")

# ===== 1. Excel 合并 =====
def excel_merge():
    print(c("\n📊 Excel 合并", 'bold'))
    folder = pick_folder()
    if not folder: return

    files = find_files(folder, ['*.xlsx', '*.xls'])
    if not files:
        print(c("❌ 没有找到Excel文件", 'red')); return

    print(f"✅ 找到 {len(files)} 个文件")
    for f in files: print(f"   📄 {os.path.basename(f)}")

    if not confirm(f"合并 {len(files)} 个文件？"): return

    import pandas as pd
    dfs = []
    for f in files:
        try:
            df = pd.read_excel(f)
            df['_来源'] = os.path.basename(f)
            dfs.append(df)
            print(f"   ✅ {os.path.basename(f)} ({len(df)}行)")
        except Exception as e:
            print(f"   ⚠️ 跳过: {e}")

    if dfs:
        merged = pd.concat(dfs, ignore_index=True)
        output = os.path.join(folder, "合并结果.xlsx")
        merged.to_excel(output, index=False, engine='openpyxl')
        done_msg(output, len(files))

# ===== 2. Excel 拆分 =====
def excel_split():
    print(c("\n📊 Excel 拆分", 'bold'))
    filepath = pick_file()
    if not filepath: return

    import pandas as pd
    rows_per = int(input("每个文件行数 (默认1000): ").strip() or "1000")
    df = pd.read_excel(filepath)
    num_files = math.ceil(len(df) / rows_per)

    output_dir = os.path.splitext(filepath)[0] + "_拆分"
    os.makedirs(output_dir, exist_ok=True)

    for i in range(num_files):
        start = i * rows_per
        end = min((i + 1) * rows_per, len(df))
        chunk = df.iloc[start:end]
        out_path = os.path.join(output_dir, f"part_{i+1:03d}.xlsx")
        chunk.to_excel(out_path, index=False, engine='openpyxl')
        print(f"   ✅ part_{i+1:03d}.xlsx ({len(chunk)}行)")

    done_msg(output_dir, num_files)

# ===== 3. PDF 合并 =====
def pdf_merge():
    print(c("\n📄 PDF 合并", 'bold'))
    folder = pick_folder()
    if not folder: return

    from PyPDF2 import PdfMerger
    files = sorted(find_files(folder, ['*.pdf']))
    if not files:
        print(c("❌ 没有找到PDF文件", 'red')); return

    print(f"✅ 找到 {len(files)} 个文件")
    for i, f in enumerate(files, 1): print(f"   {i}. {os.path.basename(f)}")

    if not confirm(f"合并 {len(files)} 个文件？"): return

    merger = PdfMerger()
    for f in files:
        merger.append(f)
        print(f"   ✅ {os.path.basename(f)}")

    output = os.path.join(folder, "合并结果.pdf")
    merger.write(output)
    merger.close()
    done_msg(output, len(files))

# ===== 4. PDF 加水印 =====
def pdf_watermark():
    print(c("\n📄 PDF 加水印", 'bold'))
    folder = pick_folder()
    if not folder: return

    from PyPDF2 import PdfReader, PdfWriter
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.colors import Color

    text = input("水印文字 (默认: 仅供参考): ").strip() or "仅供参考"
    opacity = float(input("透明度 0.01-1 (默认0.1): ").strip() or "0.1")

    # 创建水印
    packet = io.BytesIO()
    c_pdf = rl_canvas.Canvas(packet, pagesize=A4)
    c_pdf.setFont("Helvetica", 50)
    c_pdf.setFillColor(Color(0.5, 0.5, 0.5, alpha=opacity))
    c_pdf.saveState()
    c_pdf.translate(A4[0]/2, A4[1]/2)
    c_pdf.rotate(45)
    c_pdf.drawCentredString(0, 0, text)
    c_pdf.restoreState()
    c_pdf.save()
    packet.seek(0)
    wm_page = PdfReader(packet).pages[0]

    files = find_files(folder, ['*.pdf'])
    output_dir = os.path.join(folder, "水印输出")
    os.makedirs(output_dir, exist_ok=True)

    for f in files:
        try:
            reader = PdfReader(f)
            writer = PdfWriter()
            for page in reader.pages:
                page.merge_page(wm_page)
                writer.add_page(page)
            out_path = os.path.join(output_dir, os.path.basename(f))
            with open(out_path, "wb") as fp:
                writer.write(fp)
            print(f"   ✅ {os.path.basename(f)}")
        except Exception as e:
            print(f"   ❌ {os.path.basename(f)}: {e}")

    done_msg(output_dir, len(files))

# ===== 5. 图片压缩 =====
def image_compress():
    print(c("\n🖼️  图片压缩", 'bold'))
    folder = pick_folder()
    if not folder: return

    from PIL import Image
    files = find_files(folder, ['*.jpg','*.jpeg','*.png','*.bmp','*.webp'])
    if not files:
        print(c("❌ 没有找到图片", 'red')); return

    quality = int(input("压缩质量 1-100 (默认70): ").strip() or "70")
    output_dir = os.path.join(folder, "压缩输出")
    os.makedirs(output_dir, exist_ok=True)

    total_before, total_after = 0, 0
    for f in files:
        try:
            img = Image.open(f)
            orig = os.path.getsize(f)
            total_before += orig
            if img.mode in ('RGBA', 'P'): img = img.convert('RGB')
            out_path = os.path.join(output_dir, os.path.splitext(os.path.basename(f))[0] + ".jpg")
            img.save(out_path, "JPEG", quality=quality, optimize=True)
            new = os.path.getsize(out_path)
            total_after += new
            print(f"   ✅ {os.path.basename(f)}: {orig//1024}KB → {new//1024}KB (↓{(1-new/orig)*100:.0f}%)")
        except Exception as e:
            print(f"   ❌ {os.path.basename(f)}: {e}")

    if total_before:
        print(f"\n📊 总计: {total_before//1024//1024}MB → {total_after//1024//1024}MB (↓{(1-total_after/total_before)*100:.0f}%)")
    done_msg(output_dir, len(files))

# ===== 6. 图片缩放 =====
def image_resize():
    print(c("\n🖼️  图片缩放", 'bold'))
    folder = pick_folder()
    if not folder: return

    from PIL import Image
    files = find_files(folder, ['*.jpg','*.jpeg','*.png','*.bmp','*.webp'])
    if not files:
        print(c("❌ 没有找到图片", 'red')); return

    max_w = int(input("最大宽度 (默认1920): ").strip() or "1920")
    max_h = int(input("最大高度 (默认1080): ").strip() or "1080")
    output_dir = os.path.join(folder, "缩放输出")
    os.makedirs(output_dir, exist_ok=True)

    for f in files:
        try:
            img = Image.open(f)
            orig_size = img.size
            img.thumbnail((max_w, max_h), Image.Resampling.LANCZOS)
            out_path = os.path.join(output_dir, os.path.basename(f))
            img.save(out_path, quality=95)
            print(f"   ✅ {os.path.basename(f)}: {orig_size[0]}x{orig_size[1]} → {img.size[0]}x{img.size[1]}")
        except Exception as e:
            print(f"   ❌ {os.path.basename(f)}: {e}")

    done_msg(output_dir, len(files))

# ===== 7. 批量重命名 =====
def file_rename():
    print(c("\n📁 批量重命名", 'bold'))
    folder = pick_folder()
    if not folder: return

    pattern = input("匹配模式 (默认*.*): ").strip() or "*.*"
    files = sorted([f for f in glob.glob(os.path.join(folder, pattern)) if os.path.isfile(f)])
    if not files:
        print(c("❌ 没有找到文件", 'red')); return

    print(f"✅ 找到 {len(files)} 个文件")
    print("  1.添加前缀  2.添加后缀  3.替换文字  4.序号重命名")
    mode = input("模式 (1-4): ").strip()

    new_names = []
    for f in files:
        name = os.path.basename(f)
        base, ext = os.path.splitext(name)
        if mode == "1":
            prefix = input("前缀: ").strip()
            new_names.append(prefix + name)
        elif mode == "2":
            suffix = input("后缀: ").strip()
            new_names.append(base + suffix + ext)
        elif mode == "3":
            old_str = input("查找: ").strip()
            new_str = input("替换为: ").strip()
            new_names.append(name.replace(old_str, new_str))
        elif mode == "4":
            prefix = input("前缀 (可空): ").strip()
            idx_start = int(input("起始序号 (默认1): ").strip() or "1")
            new_names.append(f"{prefix}{str(files.index(f)+idx_start).zfill(3)}{ext}")

    # 预览
    print(c("\n📋 预览:", 'yellow'))
    for old, new in zip(files, new_names):
        print(f"   {os.path.basename(old)} → {new}")

    if not confirm("确认重命名？"): return

    for old_path, new_name in zip(files, new_names):
        try:
            os.rename(old_path, os.path.join(folder, new_name))
        except: pass
    done_msg(folder, len(files))

# ===== 8. 二维码生成 =====
def qr_gen():
    print(c("\n🔲 二维码生成", 'bold'))
    import qrcode

    print("  1.文本/网址  2.WiFi  3.名片")
    mode = input("类型 (1-3): ").strip() or "1"

    if mode == "1":
        data = input("内容: ").strip()
    elif mode == "2":
        ssid = input("WiFi名称: ").strip()
        pwd = input("密码: ").strip()
        data = f"WIFI:T:WPA;S:{ssid};P:{pwd};;"
    else:
        name = input("姓名: ").strip()
        phone = input("电话: ").strip()
        data = f"BEGIN:VCARD\nVERSION:3.0\nFN:{name}\nTEL:{phone}\nEND:VCARD"

    qr = qrcode.QRCode(version=1, box_size=10, border=4)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")

    output = input("输出文件 (默认qrcode.png): ").strip() or "qrcode.png"
    img.save(output)
    done_msg(os.path.abspath(output))

# ===== 9. 数据清洗 =====
def data_clean():
    print(c("\n🧹 数据清洗", 'bold'))
    filepath = pick_file()
    if not filepath: return

    import pandas as pd
    ext = os.path.splitext(filepath)[1].lower()

    try:
        if ext == '.csv':
            enc = input("编码 (默认utf-8): ").strip() or "utf-8"
            df = pd.read_csv(filepath, encoding=enc)
        else:
            df = pd.read_excel(filepath)
    except Exception as e:
        print(c(f"❌ 读取失败: {e}", 'red')); return

    print(f"\n📊 {len(df)}行 x {len(df.columns)}列 | 空值: {df.isnull().sum().sum()}")
    print("  1.删除重复行  2.删除空值行  3.去除文本空格  4.全部执行")
    ops = input("操作 (可多选，逗号分隔): ").strip()

    if "4" in ops: ops = "1,2,3"
    log = []

    if "1" in ops:
        before = len(df)
        df = df.drop_duplicates()
        log.append(f"删除重复: {before - len(df)}行")

    if "2" in ops:
        before = len(df)
        df = df.dropna()
        log.append(f"删除空值: {before - len(df)}行")

    if "3" in ops:
        str_cols = df.select_dtypes(include=['object']).columns
        for col in str_cols:
            df[col] = df[col].astype(str).str.strip()
        log.append(f"去空格: {len(str_cols)}列")

    base, ext = os.path.splitext(os.path.basename(filepath))
    output = os.path.join(os.path.dirname(filepath), f"cleaned_{base}{ext}")
    if ext == '.csv':
        df.to_csv(output, index=False, encoding='utf-8-sig')
    else:
        df.to_excel(output, index=False, engine='openpyxl')

    for l in log: print(f"   • {l}")
    done_msg(output, f"{len(df)}行x{len(df.columns)}列")

# ===== 10. 文字提取 =====
def text_extract():
    print(c("\n📝 文字提取", 'bold'))
    print("  1.PDF提取文字  2.图片OCR")
    mode = input("模式 (1-2): ").strip()

    if mode == "1":
        from PyPDF2 import PdfReader
        filepath = pick_file()
        if not filepath: return
        reader = PdfReader(filepath)
        text = ""
        for i, page in enumerate(reader.pages):
            t = page.extract_text()
            if t: text += f"\n--- 第{i+1}页 ---\n{t}"
        output = os.path.splitext(filepath)[0] + ".txt"
        with open(output, "w", encoding="utf-8") as f:
            f.write(text)
        done_msg(output, f"{len(reader.pages)}页")
    else:
        try:
            import pytesseract
            from PIL import Image
        except ImportError:
            print(c("❌ 需要: pip install pytesseract", 'red')); return
        folder = pick_folder()
        if not folder: return
        files = find_files(folder, ['*.jpg','*.jpeg','*.png','*.bmp'])
        all_text = ""
        for f in files:
            img = Image.open(f)
            text = pytesseract.image_to_string(img, lang='chi_sim+eng')
            all_text += f"\n=== {os.path.basename(f)} ===\n{text}"
        output = os.path.join(folder, "ocr_result.txt")
        with open(output, "w", encoding="utf-8") as f:
            f.write(all_text)
        done_msg(output, len(files))

# ===== 主程序 =====
HANDLERS = {
    '1': excel_merge, '2': excel_split, '3': pdf_merge, '4': pdf_watermark,
    '5': image_compress, '6': image_resize, '7': file_rename, '8': qr_gen,
    '9': data_clean, '0': text_extract,
}

def main():
    banner()
    while True:
        choice = menu()
        if choice in ('q', 'Q', 'quit', 'exit'):
            print(c("\n👋 感谢使用！如果帮到了你，请我喝杯咖啡吧 ☕", 'yellow'))
            print(c("   微信扫码赞赏: [见同目录 qrcode_wechat.jpg]", 'cyan'))
            break
        handler = HANDLERS.get(choice)
        if handler:
            try:
                handler()
            except Exception as e:
                print(c(f"\n❌ 出错了: {e}", 'red'))
        else:
            print(c("❌ 无效选择", 'red'))
        input(c("\n按回车继续...", 'cyan'))

if __name__ == "__main__":
    main()
