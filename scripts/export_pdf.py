"""Open each report in Microsoft Word, refresh the contents pages and all fields, then export a PDF.

Usage (from the repository root, Windows with Word installed):
    python scripts/export_pdf.py                 # both reports
    python scripts/export_pdf.py report/X.docx   # one file
"""
import sys
from pathlib import Path

import win32com.client

ROOT = Path(__file__).resolve().parents[1]
DEFAULT = [ROOT / "report" / "Student_Depression_Report.docx",
           ROOT / "report" / "Student_Depression_Research_Paper.docx"]
WD_FORMAT_PDF = 17

files = [Path(f).resolve() for f in sys.argv[1:]] or DEFAULT
word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
try:
    for docx in files:
        doc = word.Documents.Open(str(docx))
        # Update twice: the first pass fills the contents, the second fixes page numbers it shifted.
        for _ in range(2):
            doc.Fields.Update()
            for toc in doc.TablesOfContents:
                toc.Update()
        doc.Save()
        pdf = docx.with_suffix(".pdf")
        doc.ExportAsFixedFormat(str(pdf), WD_FORMAT_PDF)
        pages = doc.ComputeStatistics(2)  # wdStatisticPages
        doc.Close(False)
        print(f"Saved {pdf.relative_to(ROOT)} ({pages} pages)")
finally:
    word.Quit()
