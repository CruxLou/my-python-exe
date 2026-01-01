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
    等价于 VBS 里的：headerRange.Find.Execute "text", , , , , , , , , "replacement", wdReplaceAll
    """
    find = header_range.Find
    # 设置查找选项
    find.ClearFormatting()
    header_range.ClearFormatting()

    find.Text = find_text
    find.Replacement.Text = replace_text
    find.Forward = True
    find.Wrap = constants.wdFindContinue       # 循环查找
    find.MatchCase = False
    find.MatchWholeWord = False
    find.MatchWildcards = False
    find.MatchSoundsLike = False
    find.MatchAllWordForms = False

    # 2 对应 wdReplaceAll
    find.Execute(Replace=constants.wdReplaceAll)


def word_to_pdf(doc, pdf_path):
    """
    将当前文档保存为 PDF。
    17 对应 wdFormatPDF。
    """
    doc.SaveAs(str(pdf_path), FileFormat=constants.wdFormatPDF)


def delete_after_second_table(doc):
    """
    查找第 2 个表格，并删除该表格之后的所有内容。
    VBS 等价逻辑：
      Set tbl = objDoc.Tables(2)
      Set deleteRange = objDoc.Range(tbl.Range.End, objDoc.Content.End)
      deleteRange.Delete
    """
    tables = doc.Tables
    if tables.Count >= 2:
        tbl2 = tables.Item(2)  # 第二个表格
        # 以该表格的结束位置到文档尾部为删除范围
        start = tbl2.Range.End
        end = doc.Content.End
        rng = doc.Range(Start=start, End=end)
        rng.Delete()


def process_document(app, file_path):
    """
    打开 Word 文档，先保存原始 PDF，然后按要求：
      - 删除第 2 个表格之后的所有内容
      - 替换页眉中的文本
      - 保存为 *_Summary.pdf
    最后关闭文档（不保存 .doc/.docx）
    """
    folder = Path(file_path).parent
    stem = Path(file_path).stem

    # 打开文档（与 VBS 的 Documents.Open 相同）
    doc = app.Documents.Open(str(file_path))

    try:
        # 1) 保存原始 PDF（未修改版本）
        pdf_path = folder / f"{stem}.pdf"
        word_to_pdf(doc, pdf_path)

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

    except Exception as e:
        print(f"[Error] processing: {file_path}\n{e}")
        traceback.print_exc()
    finally:
        # 关闭文档（不保存 .doc/.docx 的更改）
        doc.Close(SaveChanges=False)


def main():
    # 当前脚本所在文件夹（等价于 VBS 的 GetParentFolderName(WScript.ScriptFullName)）
    try:
        script_dir = Path(os.path.abspath(os.path.dirname(__file__)))
    except NameError:
        # 如果在交互式环境中 __file__ 不存在，则使用当前工作目录
        script_dir = Path(os.getcwd())

    # 启动 Word（等价于 CreateObject("Word.Application")）
    word = win32.Dispatch("Word.Application")
    word.Visible = False  # 与 VBS 保持一致：不显示 Word 窗口

    try:
        # 遍历当前目录的 doc/docx 文件
        for p in script_dir.iterdir():
            if p.is_file():
                ext = p.suffix.lower()
                if ext in (".doc", ".docx"):
                    process_document(word, str(p))
    finally:
        # 退出 Word
        word.Quit()

    print("Done.")

if __name__ == "__main__":
    main()
``
