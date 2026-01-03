
# -*- coding: utf-8 -*-
import os
import sys
import traceback
from pathlib import Path

import win32com.client as win32
from win32com.client import constants


def replace_in_header_range(header_range, find_text, replace_text):
    """在指定 Range 内做查找替换（全部替换）。"""
    find = header_range.Find

    # ✅ Word COM：清格式应作用于 Find 和 Replacement
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
    find.Format = False  # ✅ 不按格式匹配，更稳定

    find.Execute(Replace=constants.wdReplaceAll)


def word_to_pdf(doc, pdf_path: Path):
    """将当前 Word 文档保存为 PDF。"""
    doc.SaveAs(str(pdf_path), FileFormat=constants.wdFormatPDF)


def delete_after_second_table(doc):
    """查找第 2 个表格，并删除该表格之后的所有内容。"""
    tables = doc.Tables
    if tables.Count >= 2:
        tbl2 = tables.Item(2)
        start = tbl2.Range.End
        end = doc.Content.End
        rng = doc.Range(Start=start, End=end)
        rng.Delete()


def process_document(app, file_path: str):
    """
    处理单个文档：
      1) 原始保存 PDF
      2) 删除第 2 个表格之后内容
      3) 替换页眉文本
      4) 保存 Summary PDF
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

        # 2) 删除第二表格后内容
        delete_after_second_table(doc)

        # 3) 页眉替换（每个 section）
        for sec in doc.Sections:
            header = sec.Headers.Item(constants.wdHeaderFooterPrimary)

            # ✅ 有些文档 section 的 header 可能不存在
            try:
                if hasattr(header, "Exists") and not header.Exists:
                    continue
            except Exception:
                pass

            header_range = header.Range

            replace_in_header_range(header_range, "Test Plan", "Test Report")
            replace_in_header_range(header_range, "Attachment", "")

        # 4) Summary PDF
        summary_pdf_path = folder / f"{stem}_Summary.pdf"
        word_to_pdf(doc, summary_pdf_path)
        print(f"[OK] Export summary PDF: {summary_pdf_path}")

        return True

    except Exception as e:
        print(f"[Error] processing: {file_path}\n{e}")
        traceback.print_exc()
        return False

    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass


def get_scan_dir() -> Path:
    """
    扫描目录：
    - EXE：exe 所在目录（最符合双击习惯）
    - .py：脚本所在目录
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def main():
    word = None
    try:
        scan_dir = get_scan_dir()
        print(f"[INFO] Scan folder: {scan_dir}")

        word_files = [
            p for p in scan_dir.iterdir()
            if p.is_file() and p.suffix.lower() in (".doc", ".docx")
        ]

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

        processed = 0
        success = 0

        for p in word_files:
            processed += 1
            print(f"[INFO] Processing: {p.name}")
            if process_document(word, str(p)):
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

        # ✅ 不自动关闭窗口
        print("\nPress Enter to exit...")
        try:
            input()
        except EOFError:
            pass


if __name__ == "__main__":
    main()
