from pathlib import Path  # Standard Python Library
import xlwings as xw  # pip install xlwings
from PyPDF2 import PdfFileMerger, PdfFileReader  # pip install PyPDF2


# ---Documentations:
# PyPDF2: https://pythonhosted.org/PyPDF2/
# xlwings: https://docs.xlwings.org/en/stable/


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]

    merger = PdfFileMerger()
    sheet.range("status").clear_contents()
    source_dir = sheet.range("source_dir").value
    output_name = sheet.range("output_name").value + ".pdf"
    pdf_files = list(Path(source_dir).glob("*.pdf"))

    for pdf_file in pdf_files:
        merger.append(PdfFileReader(str(pdf_file), "rb"))

    output_path = str(Path(__file__).parent / output_name)
    merger.write(output_path)
    sheet.range("status").value = f"The file have been saved here: {output_path}"


if __name__ == "__main__":
    xw.Book("pdfmerger.xlsm").set_mock_caller()
    main()
