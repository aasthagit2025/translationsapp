import ast
import html
import pathlib
import re
from difflib import SequenceMatcher

import openpyxl
import pandas as pd
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph


APP = pathlib.Path(r"D:\Projects 2026\automation\transapp.py")
XLSX = r"C:\Users\admin\Downloads\Translated_Output (10).xlsx"
DOCX = r"D:\Projects 2026\automation\SAT-13 - Caregiver Screener & QNR_PH v1.2.docx"

tree = ast.parse(APP.read_text(encoding="utf-8"))
nodes = [node for node in tree.body if isinstance(node, ast.FunctionDef)]
ns = {
    "html": html,
    "pd": pd,
    "re": re,
    "Document": Document,
    "CT_Tbl": CT_Tbl,
    "CT_P": CT_P,
    "Table": Table,
    "Paragraph": Paragraph,
    "fuzz": None,
    "process": None,
}
exec(compile(ast.Module(body=nodes, type_ignores=[]), str(APP), "exec"), ns)
memory = ns["build_translation_memory"](DOCX)
keys = list(memory.keys())

wb = openpyxl.load_workbook(XLSX, data_only=False)
ws = wb["Translated"]


def is_yellow(cell):
    fill = cell.fill
    value = str(fill.fgColor.rgb or fill.fgColor.indexed or fill.fgColor.theme or "").upper()
    return bool(fill.fill_type and ("FFFF00" in value or "FFEB9C" in value))


for row in range(2, ws.max_row + 1):
    if not any(is_yellow(ws.cell(row, col)) for col in range(1, ws.max_column + 1)):
        continue
    source = ws.cell(row, 2).value or ""
    target = ws.cell(row, 3).value or ""
    if str(target).strip():
        continue
    norm = ns["normalize"](source)
    best = sorted(
        ((SequenceMatcher(None, norm, key).ratio(), key) for key in keys),
        reverse=True,
    )[:3]
    print("\nROW", row, "ID", ws.cell(row, 1).value)
    print("SRC", str(source)[:260].replace("\n", "\\n"))
    print("NORM", norm[:260])
    for score, key in best:
        print("BEST", round(score, 3), key[:180], "=>", memory[key]["target"][:160])