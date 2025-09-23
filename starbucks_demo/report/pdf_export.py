from pathlib import Path


def export_excel_to_pdf(xlsx_path: Path, pdf_path: Path):
    try:
        import win32com.client as win32
        import pythoncom
    except ImportError:
        print("pywin32 not installed; skipping PDF export.")
        return

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(str(xlsx_path.resolve()))
        # 0 = xlTypePDF, 0 = xlQualityStandard
        wb.ExportAsFixedFormat(
            Type=0,
            Filename=str(pdf_path.resolve()),
            Quality=0,
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )
        wb.Close(False)
        print("PDF exported:", pdf_path)
    except Exception as e:
        print("PDF export failed:", e)
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()
