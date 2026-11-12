from pathlib import Path

import win32com.client

RAIZ = Path(__file__).resolve().parents[1]
DOCX = RAIZ / "Informe Final Scouting David Cuello.docx"
PDF = DOCX.with_suffix(".pdf")

word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
try:
    doc = word.Documents.Open(str(DOCX))
    doc.Fields.Update()
    for i in range(1, doc.TablesOfContents.Count + 1):
        doc.TablesOfContents(i).Update()
    doc.Save()
    doc.ExportAsFixedFormat(str(PDF), ExportFormat=17)
    paginas = doc.ComputeStatistics(2)
    doc.Close(False)
    print(f"PDF generado: {PDF} ({paginas} páginas)")
finally:
    word.Quit()
