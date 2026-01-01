# -*- coding: utf-8 -*-
import os
import sys
import traceback
from pathlib import Path

import win32com.client as win32
from win32com.client import constants


def replace_in_header_range(header_range, find_text, replace_text):
    """
    在页眉范围内执行查找替换（全部替换）。
    等价于 VBS: headerRange.Find.Execute ..., wdReplaceAll
    """
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

    # 全部替换
    find.Execute(Replace=constants.wdReplaceAll)


def word_to_pdf(doc, pdf_path: Path):
    """
    将当前 Word 文档保存为 PDF。
    """
    doc.SaveAs(str(pdf_path), FileFormat=constants.wdFormatPDF)


def delete_after_second_table(doc):
    """
    查找第 2 个表格，并删除该表格之后的所有内容。
    VBS 逻辑：
      tbl = objDoc.Tables(2)
      deleteRange = objDoc.Range(tbl.Range.End, objDoc.Content.End)
      deleteRange.Delete
    """
    tables = doc.Tables
    # ✅ 修复：原来是 tables.Count &gt;= 2
    if tables.Count >= 2:
        tbl2 = tables.Item(2)
        start = tbl2.Range.End
        end = doc.Content.End
        rng = doc.Range(Start=start, End=end)
        rng.Delete()


def process_document(app, file_path: str):
    """
    打开 Word 文档，先保存原始 PDF，然后按要求：
      - 删除第 2 个表格之后的所有内容
      - 替换页眉中的文本
      - 保存为 *_Summary.pdf
    最后关闭文档（不保存 .doc/.docx）
    """
    p = Path(file_path)
    folder = p.parent
    stem = p.stem

    doc = None
    try:
        # 打开文档
        doc = app.Documents.Open(str(p))

        # 1) 保存原始 PDF（未修改版本）
        pdf_path = folder / f"{stem}.pdf"
        word_to_pdf(doc, pdf_path)
        print(f"[OK] Export original PDF: {pdf_path}")

        # 2) 删除第二个表格之后的所有内容
        delete_after_second_table(doc)

        # 3) 修改所有节的主页眉（Primary Header）
        for sec in doc.Sections:
            header = sec.Headers.Item(constants.wdHeaderFooterPrimary)
            header_range = header.Range

            # “Test Plan” -> “Test Report”
            replace_in_header_range(header_range, "Test Plan", "Test Report")
            # “Attachment” -> ""（删除）
            replace_in_header_range(header_range, "Attachment", "")

        # 4) 保存为 Summary PDF（修改后的版本）
        summary_pdf_path = folder / f"{stem}_Summary.pdf"
        word_to_pdf(doc, summary_pdf_path)
        print(f"[OK] Export summary PDF: {summary_pdf_path}")

    except Exception as e:
        print(f"[Error] processing: {file_path}\n{e}")
        traceback.print_exc()

    finally:
        # 关闭文档（不保存 .doc/.docx 的更改）
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass


def main():
    # 当前脚本所在文件夹
    try:
        script_dir = Path(os.path.abspath(os.path.dirname(__file__)))
    except NameError:
        script_dir = Path(os.getcwd())

    # 启动 Word
    # 使用 EnsureDispatch 更稳（如果环境未生成缓存，会自动生成）
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False

    # 可选：减少弹窗（如“是否保存 Normal.dotm”等）
    # 0 = 不显示警告（wdAlertsNone）
    try:
        word.DisplayAlerts = 0
    except Exception:
        pass

    try:
        # 遍历当前目录的 doc/docx 文件
        for p in script_dir.iterdir():
            if p.is_file() and p.suffix.lower() in (".doc", ".docx"):
                print(f"[INFO] Processing: {p}")
                process_document(word, str(p))

    finally:
        # 退出 Word
        try:
            word.Quit()
        except Exception:
            pass

    print("Done.")


if __name__ == "__main__":
    main()
