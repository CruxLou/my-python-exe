# -*- coding: utf-8 -*-
import os
import sys
import traceback
from pathlib import Path

import win32com.client as win32
from win32com.client import constants


# =========================================================
# Step 1: Word -> PDF + content edits
# =========================================================
def replace_in_header_range(header_range, find_text, replace_text):
    """在页眉 Range 内做查找替换（全部替换）。"""
    find = header_range.Find
    find.ClearFormatting()
    find.Replacement.ClearFormatting()

    find.Text = find_text
    find.Replacement.Text = replace_text

    find.Forward = True
    find.Wrap = constants.wdFindContinue
    find.MatchCase = False
    find.MatchWholeWord = False
    find.MatchWildcards = False
    find.MatchSoundsLike = False
    find.MatchAllWordForms = False
    find.Format = False

    find.Execute(Replace=constants.wdReplaceAll)


def word_to_pdf(doc, pdf_path: Path):
    """将当前 Word 文档保存为 PDF。"""
    doc.SaveAs(str(pdf_path), FileFormat=constants.wdFormatPDF)


def delete_after_second_table(doc):
    """查找第2个表格，并删除该表格之后的所有内容。"""
    tables = doc.Tables
    if tables.Count >= 2:
        tbl2 = tables.Item(2)
        rng = doc.Range(Start=tbl2.Range.End, End=doc.Content.End)
        rng.Delete()


def get_base_dir() -> Path:
    """
    扫描目录：
    - 打包 EXE：以 exe 所在目录为准（双击最直观）
    - 直接跑 py：以脚本所在目录为准
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


# =========================================================
# Step 2: Signature on PDF (search by all names in /signature)
# =========================================================
def load_signature_folder(base_dir: Path):
    """
    读取 base_dir/signature 下的签名资源。
    约定：
      - 名字 = 文件名 stem（不含扩展名）
      - 签名图片 = signature/名字.jpg 或 png
    返回：
      sig_names: [name1, name2, ...]（全部stem去重）
      img_map: {name_lower: image_path}（仅图片）
    """
    sig_dir = base_dir / "signature"
    sig_names = []
    img_map = {}

    if not sig_dir.exists() or not sig_dir.is_dir():
        print(f"[WARN] signature folder not found: {sig_dir}")
        return sig_names, img_map

    for f in sig_dir.iterdir():
        if not f.is_file():
            continue
        sig_names.append(f.stem)
        if f.suffix.lower() in (".jpg", ".jpeg", ".png"):
            img_map[f.stem.lower()] = f

    sig_names = sorted(set(sig_names))
    print(f"[INFO] signature names(stem): {sig_names}")
    return sig_names, img_map


def build_word_rect_index_first_page(page):
    """
    用 PyMuPDF 在第一页提取 word-level 文本与坐标。
    返回：dict[word_lower] -> list[fitz.Rect]（按阅读顺序）
    """
    import fitz  # PyMuPDF

    # get_text("words") 返回词级别数据及坐标 (x0,y0,x1,y1,word,...)
    # 这能让我们精准定位“名字”在 PDF 上的位置。[1](https://pymupdftest.readthedocs.io/en/stable/index.html)
    words = page.get_text("words")
    words.sort(key=lambda w: (w[1], w[0]))  # 按阅读顺序：先y后x

    idx = {}
    for w in words:
        x0, y0, x1, y1, txt = w[0], w[1], w[2], w[3], w[4]
        if not txt:
            continue
        key = txt.strip().lower()
        if not key:
            continue
        idx.setdefault(key, []).append(fitz.Rect(x0, y0, x1, y1))
    return idx


def sign_pdf_search_all_names(pdf_in: Path, pdf_out: Path, sig_names, img_map,
                             max_sign=2, img_width_pt=120, x_offset_pt=0, y_offset_pt=-5):
    """
    按你的最新规则：
      1) 搜索 signature 文件夹里“所有名字”
      2) 选择在第一页出现次数最多的名字 best_name
      3) 插入 min(best_hits, max_sign) 个签名（0/1/2）
      4) 仅当 best_hits < 2（即所有名字出现次数都<2）时，返回 need_warn=True
    返回：
      inserted(int), best_name(str|None), best_hits(int), need_warn(bool)
    """
    try:
        import fitz  # PyMuPDF
    except ImportError:
        print("[FATAL] PyMuPDF (fitz) not installed. Please 'pip install pymupdf'.")
        return 0, None, 0, True

    doc = fitz.open(str(pdf_in))
    try:
        if doc.page_count < 1:
            doc.save(str(pdf_out))
            return 0, None, 0, True

        page = doc[0]  # 只查第一页
        word_rects = build_word_rect_index_first_page(page)

        # 只保留 signature 中“有对应图片”的名字
        valid_names = [n for n in sig_names if n and (n.strip().lower() in img_map)]
        if not valid_names:
            doc.save(str(pdf_out))
            return 0, None, 0, True

        # 统计所有名字出现次数，取最大者
        best_name = None
        best_rects = []
        for n in valid_names:
            rects = word_rects.get(n.strip().lower(), [])
            if len(rects) > len(best_rects):
                best_name = n
                best_rects = rects

        best_hits = len(best_rects)

        # 是否需要提示：只有当“所有名字都<2次”才提示
        # 等价于 best_hits < 2
        need_warn = (best_hits < 2)

        # 命中0次：不插入
        if best_hits == 0 or best_name is None:
            doc.save(str(pdf_out))
            return 0, best_name, best_hits, need_warn

        img = img_map.get(best_name.strip().lower())
        if not img:
            doc.save(str(pdf_out))
            return 0, best_name, best_hits, need_warn

        insert_count = min(best_hits, max_sign)  # 1次插1，>=2次插2
        inserted = 0

        for r in best_rects[:insert_count]:
            # 计算图片高度（保持比例）
            try:
                pix = fitz.Pixmap(str(img))
                w_px, h_px = pix.width, pix.height
                pix = None
            except Exception:
                w_px, h_px = 300, 120  # 兜底

            aspect = h_px / max(w_px, 1)
            img_h_pt = img_width_pt * aspect

            # 以“名字词框左上角”为锚点，可用偏移微调
            x0 = r.x0 + x_offset_pt
            y0 = r.y0 + y_offset_pt
            rect_img = fitz.Rect(x0, y0, x0 + img_width_pt, y0 + img_h_pt)

            # insert_image: 在指定矩形插入图片（overlay=True 覆盖显示）
            page.insert_image(rect_img, filename=str(img), keep_proportion=True, overlay=True)
            inserted += 1

        doc.save(str(pdf_out))
        return inserted, best_name, best_hits, need_warn

    finally:
        doc.close()


# =========================================================
# Pipeline per document
# =========================================================
def process_document(app, file_path: str, sig_names, sig_img_map):
    """
    Step 1: Word处理并导出PDF
      - xxx.pdf (原始)
      - xxx_Summary.pdf (修改后)
    Step 2: 在 Summary PDF 首页查找 signature 名字并贴签名
      - xxx_Signed.pdf
    """
    p = Path(file_path)
    folder = p.parent
    stem = p.stem

    doc = None
    summary_pdf = None

    try:
        doc = app.Documents.Open(str(p))

        # 1) 原始 PDF
        pdf_path = folder / f"{stem}.pdf"
        word_to_pdf(doc, pdf_path)
        print(f"[OK] Export original PDF: {pdf_path}")

        # 2) 删除第二个表格之后内容
        delete_after_second_table(doc)

        # 3) 页眉替换
        for sec in doc.Sections:
            header = sec.Headers.Item(constants.wdHeaderFooterPrimary)
            header_range = header.Range
            replace_in_header_range(header_range, "Test Plan", "Test Report")
            replace_in_header_range(header_range, "Attachment", "")

        # 4) Summary PDF
        summary_pdf = folder / f"{stem}_Summary.pdf"
        word_to_pdf(doc, summary_pdf)
        print(f"[OK] Export summary PDF: {summary_pdf}")

    except Exception as e:
        print(f"[Error] processing Word: {file_path}\n{e}")
        traceback.print_exc()
        return False

    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass

    # Step 2: PDF 签名
    try:
        signed_pdf = folder / f"{stem}_Signed.pdf"

        inserted, best_name, best_hits, need_warn = sign_pdf_search_all_names(
            pdf_in=summary_pdf,
            pdf_out=signed_pdf,
            sig_names=sig_names,
            img_map=sig_img_map,
            max_sign=2,
            img_width_pt=120,
            x_offset_pt=0,
            y_offset_pt=-5
        )

        print(f"[INFO] Best name on page1: {best_name}, hits={best_hits}, inserted={inserted}")

        # ✅ 你的规则：所有名字都<2次才提示
        if need_warn:
            print("No found signature")

        if inserted > 0:
            print(f"[OK] Signed PDF saved: {signed_pdf}")

        return True

    except Exception as e:
        print(f"[Error] processing PDF signature: {summary_pdf}\n{e}")
        traceback.print_exc()
        return False


# =========================================================
# Main
# =========================================================
def main():
    word = None
    try:
        base_dir = get_base_dir()
        print(f"[INFO] Scan folder: {base_dir}")

        # signature 资源
        sig_names, sig_img_map = load_signature_folder(base_dir)

        # 扫描 Word 文件（exe 同目录）
        word_files = [
            p for p in base_dir.iterdir()
            if p.is_file() and p.suffix.lower() in (".doc", ".docx")
        ]

        if not word_files:
            print("No found docx doc")
            return

        # 启动 Word COM
        word = win32.gencache.EnsureDispatch("Word.Application")
        word.Visible = False
        try:
            word.DisplayAlerts = 0
        except Exception:
            pass

        processed, success = 0, 0
        for p in word_files:
            processed += 1
            print(f"[INFO] Processing: {p.name}")
            if process_document(word, str(p), sig_names, sig_img_map):
                success += 1

        print("PDF transfer done")
        print(f"[INFO] Done. Total: {processed}, Success: {success}")

    except Exception as e:
        print(f"[FATAL] {e}")
        traceback.print_exc()

    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass

        # 不自动关闭窗口
        print("\nPress Enter to exit...")
        try:
            input()
        except EOFError:
            pass


if __name__ == "__main__":
    main()
