# -*- coding: utf-8 -*-
import os
import sys
import traceback
from pathlib import Path

import win32com.client as win32
from win32com.client import constants


# ---------------------------
# Word helper functions
# ---------------------------
def replace_in_header_range(header_range, find_text, replace_text):
    """在指定 Range 内做查找替换（全部替换）。"""
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
    """查找第 2 个表格，并删除该表格之后的所有内容。"""
    tables = doc.Tables
    if tables.Count >= 2:
        tbl2 = tables.Item(2)
        rng = doc.Range(Start=tbl2.Range.End, End=doc.Content.End)
        rng.Delete()


# ---------------------------
# Runtime paths
# ---------------------------
def get_base_dir() -> Path:
    """EXE：exe 所在目录；py：脚本所在目录。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


# ---------------------------
# Signature (PDF) helpers
# ---------------------------
def load_signature_folder(base_dir: Path):
    """
    读取 base_dir/signature 下的签名资源。
    约定：
      - 名字 = 文件名 stem（不含扩展名）
      - 签名图片 = signature/名字.jpg 或 png
    返回：
      names: [名字1, 名字2...]
      img_map: {名字lower: Path_to_image}
    """
    sig_dir = base_dir / "signature"
    names = []
    img_map = {}

    if not sig_dir.exists() or not sig_dir.is_dir():
        print(f"[WARN] signature folder not found: {sig_dir}")
        return names, img_map

    for f in sig_dir.iterdir():
        if not f.is_file():
            continue

        # 记录所有“文档名称/名字”（你要求的记录）
        names.append(f.stem)

        # 只把图片加入 map
        if f.suffix.lower() in (".jpg", ".jpeg", ".png"):
            img_map[f.stem.lower()] = f

    print(f"[INFO] signature names(stem): {sorted(set(names))}")
    return sorted(set(names)), img_map


def sign_pdf_by_names_first_page(pdf_in: Path, pdf_out: Path, names, img_map,
                                 need_count=2, img_width_pt=120, y_offset_pt=-5):
    """
    在 PDF 首页搜索 names（名字列表），若命中则在该文字位置插入对应图片。
    - need_count：需要插入的签名数量（你要求 2 个位置）
    - img_width_pt：图片宽度（PDF point）
    - y_offset_pt：插入时 y 方向微调（负值表示略往上）
    返回：插入数量 inserted
    """
    try:
        import fitz  # PyMuPDF
    except ImportError:
        print("[FATAL] PyMuPDF (fitz) not installed. Please 'pip install pymupdf'.")
        return 0

    inserted = 0

    doc = fitz.open(str(pdf_in))
    try:
        if doc.page_count < 1:
            return 0

        page = doc[0]  # 只处理第一页

        # 依次用 signature 文件夹里的“名字”去搜
        for name in names:
            key = name.lower()
            img = img_map.get(key)
            if not img:
                continue

            # 搜索名字出现的位置（返回矩形列表）
            rects = page.search_for(name)
            if not rects:
                continue

            for r in rects:
                # 计算插入图片矩形：以文字矩形左边为锚，图片宽度固定，高度按图片比例缩放
                # 先拿图片尺寸（像素）换算比例，保持不变形
                try:
                    pix = fitz.Pixmap(str(img))
                    w_px, h_px = pix.width, pix.height
                    pix = None
                except Exception:
                    w_px, h_px = 300, 120  # 兜底

                aspect = h_px / max(w_px, 1)
                img_h_pt = img_width_pt * aspect

                # 把签名贴在“文字位置”附近：默认贴在文字矩形的左上角区域
                x0 = r.x0
                y0 = r.y0 + y_offset_pt
                rect_img = fitz.Rect(x0, y0, x0 + img_width_pt, y0 + img_h_pt)

                page.insert_image(rect_img, filename=str(img), keep_proportion=True, overlay=True)

                inserted += 1
                if inserted >= need_count:
                    break

            if inserted >= need_count:
                break

        doc.save(str(pdf_out))
        return inserted

    finally:
        doc.close()


# ---------------------------
# Document processing
# ---------------------------
def process_document(app, file_path: str, sig_names, sig_img_map):
    """
    Step 1: Word 处理 + 导出 PDF
    Step 2: 对 Summary PDF 进行签名贴图 -> Signed PDF
    """
    p = Path(file_path)
    folder = p.parent
    stem = p.stem

    doc = None
    try:
        doc = app.Documents.Open(str(p))

        # 1) 原始 PDF
        pdf_path = folder / f"{stem}.pdf"
        word_to_pdf(doc, pdf_path)
        print(f"[OK] Export original PDF: {pdf_path}")

        # 2) 修改内容
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

    # ----------------------
    # Step 2: PDF 签名贴图
    # ----------------------
    try:
        signed_pdf = folder / f"{stem}_Signed.pdf"
        inserted = sign_pdf_by_names_first_page(
            pdf_in=summary_pdf,
            pdf_out=signed_pdf,
            names=sig_names,
            img_map=sig_img_map,
            need_count=2,
            img_width_pt=120,
            y_offset_pt=-5
        )

        if inserted < 2:
            print("No found signature")
        else:
            print(f"[OK] Signature inserted: {inserted} -> {signed_pdf}")

        return True

    except Exception as e:
        print(f"[Error] processing PDF signature: {summary_pdf}\n{e}")
        traceback.print_exc()
        return False


def main():
    word = None
    try:
        base_dir = get_base_dir()
        print(f"[INFO] Scan folder: {base_dir}")

        # 读取 signature 文件夹
        sig_names, sig_img_map = load_signature_folder(base_dir)

        # 扫描 Word 文件
        word_files = [p for p in base_dir.iterdir()
                      if p.is_file() and p.suffix.lower() in (".doc", ".docx")]

        if not word_files:
            print("No found docx doc")
            return

        # 启动 Word
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

        print("\nPress Enter to exit...")
        try:
            input()
        except EOFError:
            pass


if __name__ == "__main__":
    main()
