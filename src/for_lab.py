# -*- coding: utf-8 -*-
import os
import sys
import traceback
from pathlib import Path

import win32com.client as win32
from win32com.client import constants


def replace_in_header_range(header_range, find_text, replace_text):
    """在页眉范围内执行查找替换（全部替换）。"""
    find = header_range.Find
    find.ClearFormatting()
    header_range.ClearFormatting()

    find.Text = find_text
    find.Replacement.Text = replace_text
    find.Forward = True
    find.Wrap = constants.wdFindContinue
    find.MatchCase = False
    find.MatchWholeWord = False
    find.MatchWildcards = False
    find.MatchSoundsLike = False
    find.MatchAllWordForms = False

    find.Execute(Replace=constants.wdReplaceAll)


def word_to_pdf(doc, pdf_path: Path):
    """将当前 Word 文档保存为 PDF。"""
    doc.SaveAs(str(pdf_path), FileFormat=constants.wdFormatPDF)


def delete_after_second_table(doc):
    """查找第 2 个表格，并删除该表格之后的所有内容。"""
    tables = doc.Tables
    # ✅ 必须是 >=，不能出现 &gt;= 或 &amp;gt;=
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

        pdf_path = folder / f"{stem}.pdf"
        word_to_pdf(doc, pdf_path)
        print(f"[OK] Export original PDF: {pdf_path}")

        delete_after_second_table(doc)

        for sec in doc.Sections:
            header = sec.Headers.Item(constants.wdHeaderFooterPrimary)
            header_range = header.Range
            replace_in_header_range(header_range, "Test Plan", "Test Report")
            replace_in_header_range(header_range, "Attachment", "")

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
    获取扫描目录：
    - 打包成 EXE：用 exe 所在目录（最符合双击使用习惯）
    - 直接跑 .py：用脚本所在目录
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def main():
    """
    扫描“当前目录”下的 doc/docx 并批处理。
    - 没找到：打印 'No found docx doc'
    - 全部转换完成：打印 'PDF transfer done'
    - 不自动关闭窗口：等待用户按回车
    """
    word = None
    try:
        script_dir = get_scan_dir()
        print(f"[INFO] Scan folder: {script_dir}")

        word_files = [
            p for p in script_dir.iterdir()
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

        # ✅ 关键：不自动关闭窗口
        print("\nPress Enter to exit...")
        try:
            input()
        except EOFError:
            pass


if __name__ == "__main__":
    main()
