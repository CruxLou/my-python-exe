
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
    - 打包 EXE：以 exe 所在目录为准
    - 直接跑 py：以脚本所在目录为准
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


# =========================================================
# Step 2: Signature on PDF
# =========================================================
def load_signature_folder(base_dir: Path):
    """
    读取 base_dir/signature 下的签名资源。
    约定：名字 = 文件名 stem（不含扩展名），图片= jpg/jpeg/png
    返回：
      img_map: {name_lower: image_path}
    """
    sig_dir = base_dir / "signature"
    img_map = {}

    if not sig_dir.exists() or not sig_dir.is_dir():
        print(f"[WARN] signature folder not found: {sig_dir}")
        return img_map

    names = []
    for f in sig_dir.iterdir():
        if not f.is_file():
            continue
        names.append(f.stem)
        if f.suffix.lower() in (".jpg", ".jpeg", ".png"):
            img_map[f.stem.lower()] = f

    names = sorted(set(names))
    print(f"[INFO] signature names(stem): {names}")
    return img_map


def insert_signatures_on_first_page(pdf_in: Path, pdf_out: Path, img_map: dict,
                                    target_count=2,
                                    img_width_pt=120,
                                    height_scale=0.5,
                                    x_offset_pt=0,
                                    y_offset_pt=-5,
                                    background=True):
    """
    在 PDF 第1页按阅读顺序扫描所有 words：
      - 若 word.lower() 在 img_map 中，则插入对应签名图片
      - 直到插满 target_count 或扫描结束
    返回 inserted_count

    关键点：
    - 用 page.get_text("words") 获取词与坐标用于定位。[2](https://github.com/pyinstaller/pyinstaller/blob/develop/doc/spec-files.rst)
    - background=True 时 overlay=False -> 图片在背景层不遮挡文字。[1](https://github.com/pymupdf/PyMuPDF/issues/1976)
    """
    try:
        import fitz  # PyMuPDF
    except ImportError:
        print("[FATAL] PyMuPDF (fitz) not installed. Please 'pip install pymupdf'.")
        return 0

    if not pdf_in.exists():
        print(f"[WARN] PDF not found: {pdf_in}")
        return 0

    doc = fitz.open(str(pdf_in))
    try:
        if doc.page_count < 1:
            doc.save(str(pdf_out))
            return 0

        page = doc[0]

        # 词级别提取 (x0, y0, x1, y1, word, block, line, word_no)
        words = page.get_text("words")  # [2](https://github.com/pyinstaller/pyinstaller/blob/develop/doc/spec-files.rst)
        words.sort(key=lambda w: (w[1], w[0]))  # 阅读顺序

        inserted = 0
        for w in words:
            if inserted >= target_count:
                break

            x0, y0, x1, y1, txt = w[0], w[1], w[2], w[3], w[4]
            if not txt:
                continue

            key = txt.strip().lower()
            if not key:
                continue

            img_path = img_map.get(key)
            if not img_path:
                continue

            # 计算显示尺寸：宽固定，高度按比例并 * height_scale（你要求高度减半）
            try:
                pix = fitz.Pixmap(str(img_path))
                w_px, h_px = pix.width, pix.height
                pix = None
            except Exception:
                w_px, h_px = 300, 120

            aspect = h_px / max(w_px, 1)
            img_h_pt = img_width_pt * aspect * height_scale

            rx0 = x0 + x_offset_pt
            ry0 = y0 + y_offset_pt
            rect_img = fitz.Rect(rx0, ry0, rx0 + img_width_pt, ry0 + img_h_pt)

            # overlay=False => 背景层（不遮挡文字）[1](https://github.com/pymupdf/PyMuPDF/issues/1976)
            overlay_flag = False if background else True

            page.insert_image(
                rect_img,
                filename=str(img_path),
                keep_proportion=True,
                overlay=overlay_flag
            )

            inserted += 1
            print(f"[INFO] Signature inserted for '{key}' at {rect_img}")

        doc.save(str(pdf_out))
        return inserted

    finally:
        doc.close()


# =========================================================
# Pipeline per document
# =========================================================
def process_document(app, file_path: str, img_map: dict):
    """
    最终只保留两份 PDF：
      1) {stem}_Signed.pdf
      2) {stem}_Summary_Signed.pdf   （你要求 Summary 加签名后用这个名字）

    中间文件会删除（但仅当签名输出成功后才删除）：
      - {stem}.pdf
      - {stem}_Summary.pdf
    """
    p = Path(file_path)
    folder = p.parent
    stem = p.stem

    # 中间未签名文件
    raw_pdf = folder / f"{stem}.pdf"
    summary_pdf = folder / f"{stem}_Summary.pdf"

    # 最终保留文件
    raw_signed_pdf = folder / f"{stem}_Signed.pdf"
    summary_signed_pdf = folder / f"{stem}_Summary_Signed.pdf"

    doc = None
    try:
        doc = app.Documents.Open(str(p))

        # 1) 原始 PDF（未修改）
        word_to_pdf(doc, raw_pdf)
        print(f"[OK] Export original PDF: {raw_pdf}")

        # 2) 修改内容
        delete_after_second_table(doc)

        # 3) 页眉替换
        for sec in doc.Sections:
            header = sec.Headers.Item(constants.wdHeaderFooterPrimary)
            header_range = header.Range
            replace_in_header_range(header_range, "Test Plan", "Test Report")
            replace_in_header_range(header_range, "Attachment", "")

        # 4) Summary PDF（未签名）
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

    # ---------- Step 2: PDF 插入签名 ----------
    try:
        if not img_map:
            print("No found signature")
            # 不删中间文件，避免丢失
            return True

        # A) 原始 PDF -> Signed
        inserted_raw = insert_signatures_on_first_page(
            pdf_in=raw_pdf,
            pdf_out=raw_signed_pdf,
            img_map=img_map,
            target_count=2,
            img_width_pt=120,
            height_scale=0.5,      # ✅ 高度减半
            x_offset_pt=0,
            y_offset_pt=-5,
            background=True        # ✅ 背景层不遮挡文字 [1](https://github.com/pymupdf/PyMuPDF/issues/1976)
        )

        # B) Summary PDF -> Summary_Signed（你要求的命名）
        inserted_sum = insert_signatures_on_first_page(
            pdf_in=summary_pdf,
            pdf_out=summary_signed_pdf,
            img_map=img_map,
            target_count=2,
            img_width_pt=120,
            height_scale=0.5,
            x_offset_pt=0,
            y_offset_pt=-5,
            background=True
        )

        # 你的规则：如果“所有名字都出现<2次”，等价于插入总数 < 2 时提示
        # 这里分别对 raw 和 summary 判断，你也可以只判断 summary（看你更关注哪份）
        if inserted_raw < 2 or inserted_sum < 2:
            print("No found signature")

        # 只有当最终文件生成成功才删除中间文件，避免丢文件
        if raw_signed_pdf.exists() and summary_signed_pdf.exists():
            for f in (raw_pdf, summary_pdf):
                try:
                    if f.exists():
                        f.unlink()
                except Exception:
                    pass

        print(f"[OK] Keep: {raw_signed_pdf} (inserted={inserted_raw})")
        print(f"[OK] Keep: {summary_signed_pdf} (inserted={inserted_sum})")
        return True

    except Exception as e:
        print(f"[Error] processing PDF signature\n{e}")
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

        img_map = load_signature_folder(base_dir)

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
            if process_document(word, str(p), img_map):
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

        print("\nPress Enter to exit...")
        try:
            input()
        except EOFError:
            pass


if __name__ == "__main__":
    main()
