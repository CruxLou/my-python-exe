
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
      sig_names: 所有stem去重列表
      img_map: {name_lower: image_path}
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
    - 使用 page.get_text("words") 获取词和坐标（用于定位）。[3](https://deepwiki.com/pyinstaller/pyinstaller/4.1-understanding-and-writing-hooks)
    - background=True 时使用 overlay=False，把图片放到背景层不遮挡文字。[1](https://github.com/pymupdf/PyMuPDF/discussions/3717)[2](https://developer.aliyun.com/article/1559433)
    """
    try:
        import fitz  # PyMuPDF
    except ImportError:
        print("[FATAL] PyMuPDF (fitz) not installed. Please 'pip install pymupdf'.")
        return 0

    doc = fitz.open(str(pdf_in))
    try:
        if doc.page_count < 1:
            doc.save(str(pdf_out))
            return 0

        page = doc[0]

        # 读取第一页所有单词及坐标 (x0, y0, x1, y1, word, block, line, word_no)
        words = page.get_text("words")  # word-level + coordinates [3](https://deepwiki.com/pyinstaller/pyinstaller/4.1-understanding-and-writing-hooks)
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

            # 计算签名图片显示尺寸：宽固定，高度=按比例 * height_scale（你要求高度减半）
            try:
                pix = fitz.Pixmap(str(img_path))
                w_px, h_px = pix.width, pix.height
                pix = None
            except Exception:
                w_px, h_px = 300, 120

            aspect = h_px / max(w_px, 1)
            img_h_pt = img_width_pt * aspect * height_scale

            # 以单词左上角为锚点，可微调偏移
            rx0 = x0 + x_offset_pt
            ry0 = y0 + y_offset_pt
            rect_img = fitz.Rect(rx0, ry0, rx0 + img_width_pt, ry0 + img_h_pt)

            # overlay=False => 放到背景层（不遮挡文字）[1](https://github.com/pymupdf/PyMuPDF/discussions/3717)[2](https://developer.aliyun.com/article/1559433)
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
    Step 1:
      - 导出原始 PDF：xxx.pdf
      - 删除第2表格后内容 + 替换页眉
      - 导出 Summary PDF：xxx_Summary.pdf

    Step 2:
      - 对 Summary PDF 本身插入签名（覆盖写回）
      - 额外生成 Signed 备份：xxx_Signed.pdf
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
        if not img_map:
            print("No found signature")
            return True

        # (A) Summary PDF 本身也需要插签名 —— 覆盖写回
        tmp_summary = folder / f"{stem}_Summary_tmp.pdf"
        inserted_summary = insert_signatures_on_first_page(
            pdf_in=summary_pdf,
            pdf_out=tmp_summary,
            img_map=img_map,
            target_count=2,
            img_width_pt=120,
            height_scale=0.5,   # ✅ 高度减半
            x_offset_pt=0,
            y_offset_pt=-5,
            background=True     # ✅ 背景层，不遮挡文字 [1](https://github.com/pymupdf/PyMuPDF/discussions/3717)[2](https://developer.aliyun.com/article/1559433)
        )

        # 覆盖写回 Summary
        if tmp_summary.exists():
            try:
                summary_pdf.unlink(missing_ok=True)
            except Exception:
                pass
            tmp_summary.replace(summary_pdf)

        # (B) 额外生成一个 Signed 备份文件（方便你对比/归档）
        signed_pdf = folder / f"{stem}_Signed.pdf"
        inserted_signed = insert_signatures_on_first_page(
            pdf_in=summary_pdf,
            pdf_out=signed_pdf,
            img_map=img_map,
            target_count=2,
            img_width_pt=120,
            height_scale=0.5,
            x_offset_pt=0,
            y_offset_pt=-5,
            background=True
        )

        # 提示逻辑：你要求“如果所有名字都出现 <2 次 才提示 No found signature”
        # 这里用 inserted_summary 来判断：插入次数 < 2 说明第一页合计命中不足 2
        if inserted_summary < 2:
            print("No found signature")

        print(f"[OK] Summary signed (in-place): {summary_pdf} (inserted={inserted_summary})")
        print(f"[OK] Signed PDF saved: {signed_pdf} (inserted={inserted_signed})")
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
            if process_document(word, str(p), sig_img_map):
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
