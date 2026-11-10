from contextlib import contextmanager
from copy import deepcopy
from io import BytesIO
from pathlib import Path
from typing import Annotated, Literal
import re
from zipfile import BadZipFile, ZipFile

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.oxml.ns import qn
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, Response
from lxml.etree import XMLSyntaxError
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

ROOT = Path(__file__).resolve().parent
TEMPLATE = ROOT / "Student Progress Template.docx"
UPLOAD_LIMIT = 10 * 1024 * 1024
TERMS = ("first", "second", "summer")
COURSE_TYPES = ("lecture", "laboratory", "rle")
EXAMS = {"PE": "PrelimExam", "ME": "MidtermExam", "FE": "FinalExam"}
SCORES = {
    **EXAMS,
    "MG": "MidtermGrade",
    **{f"{prefix}{i}": f"{prefix}{i}" for prefix in ("MQ", "FQ", "MA", "FA") for i in range(1, 8)},
}
app = FastAPI(title="Student Progress Maker")


REQUIRED_HEADERS = {"StudentName", *SCORES.values()}


@contextmanager
def open_class_record(contents):
    try:
        with ZipFile(BytesIO(contents)) as archive:
            if sum(member.file_size for member in archive.infolist()) > 50 * 1024 * 1024:
                raise ValueError("The expanded workbook exceeds 50 MB.")
        book = load_workbook(BytesIO(contents), read_only=True, data_only=True)
        try:
            yield book
        finally:
            book.close()
    except ValueError as exc:
        raise HTTPException(400, str(exc)) from exc
    except (BadZipFile, InvalidFileException, KeyError, OSError, XMLSyntaxError) as exc:
        raise HTTPException(400, "Excel could not read this class record. Re-save it and try again.") from exc


def worksheet_columns(sheet):
    header = next(sheet.iter_rows(max_row=1, max_col=256, values_only=True), ())
    names = [str(value).strip() for value in header]
    if any(names.count(name) != 1 for name in REQUIRED_HEADERS):
        return None
    return {name: names.index(name) for name in REQUIRED_HEADERS}


def read_students(contents, worksheet=None):
    with open_class_record(contents) as book:
        if worksheet is None:
            if len(book.sheetnames) != 1:
                raise ValueError("Select a worksheet from the uploaded class record.")
            worksheet = book.sheetnames[0]
        sheets = {sheet.title: sheet for sheet in book.worksheets}
        if worksheet not in sheets:
            raise ValueError("The selected worksheet is missing from the class record.")
        sheet = sheets[worksheet]
        columns = worksheet_columns(sheet)
        if columns is None:
            raise ValueError("The selected worksheet needs one of each required class-record header in row 1.")
        students = []
        for row_number, row in enumerate(
            sheet.iter_rows(min_row=2, max_col=max(columns.values()) + 1), start=2
        ):
            name = row[columns["StudentName"]].value
            if not isinstance(name, str) or not name.strip():
                continue
            student = {"StudentName": name.strip()}
            for column in SCORES.values():
                cell = row[columns[column]]
                if cell.data_type == "e":
                    raise ValueError(f"Excel error in {column}, row {row_number}. Correct it and upload again.")
                student[column] = cell.value
            students.append(student)
            if len(students) > 100:
                raise ValueError("Upload a class record with at most 100 students.")
        if not students:
            raise ValueError("The class record contains no student names.")
        return students


def display(value):
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return f"{value:.2f}".rstrip("0").rstrip(".")
    return str(value)


def fill_sheet(template, student, fields, term, course_type):
    doc = Document(BytesIO(template))
    for paragraph in doc.paragraphs:
        # Edit text nodes directly so tabs and the original checkbox symbols survive.
        for node in paragraph._p.xpath("./w:r/w:t"):
            original = node.text or ""
            if original in fields:
                node.text = fields[original]
            elif original == ": AcadYear":
                node.text = ": " + fields["AcadYear"]
        symbols = paragraph._p.xpath("./w:r/w:sym")
        if paragraph.text.startswith("NAME: "):
            first_tab = True
            for run in paragraph.runs:
                if run.text == "\t":
                    run.text = "  " if first_tab else ""
                    first_tab = False
                else:
                    first_tab = True
            for node in paragraph._p.xpath("./w:r/w:t"):
                if node.text in (" FIRST ", " SECOND ", " SUMMER"):
                    node.text = "\u00a0" + node.text[1:]
            symbols[TERMS.index(term)].set(qn("w:char"), "F0FE")
        elif paragraph.text.startswith("COURSE CODE: "):
            symbols[COURSE_TYPES.index(course_type)].set(qn("w:char"), "F0FE")

    for table_index, table in enumerate(doc.tables):
        for row_index, row in enumerate(table.rows):
            if table_index == 0 and row_index == 0:
                continue  # The first exam row labels the columns.
            for cell in row.cells:
                column = SCORES.get(cell.text)
                if column:
                    runs = cell.paragraphs[0].runs
                    runs[0].text = display(student[column])
                    for run in runs[1:]:
                        run.text = ""

    return doc


def build_document(students, fields, term, course_type):
    template = TEMPLATE.read_bytes()
    document = fill_sheet(template, students[0], {**fields, "StudentName": students[0]["StudentName"]}, term, course_type)
    for student in students[1:]:
        sheet = fill_sheet(template, student, {**fields, "StudentName": student["StudentName"]}, term, course_type)
        section = document.add_section(WD_SECTION_START.NEW_PAGE)
        for element in sheet.element.body:
            if element.tag != qn("w:sectPr"):
                section._sectPr.addprevious(deepcopy(element))

    output = BytesIO()
    document.save(output)
    return output.getvalue()


def build_student_zip(students, fields, term, course_type):
    template = TEMPLATE.read_bytes()
    output = BytesIO()
    with ZipFile(output, "w") as archive:
        for number, student in enumerate(students, start=1):
            name = student["StudentName"]
            filename = re.sub(r"[^\w .-]", "_", name)[:100].strip(" .")
            sheet = fill_sheet(template, student, {**fields, "StudentName": name}, term, course_type)
            document = BytesIO()
            sheet.save(document)
            archive.writestr(f"{number:03d} - {filename or 'Student'}.docx", document.getvalue())
    return output.getvalue()


@app.get("/", response_class=HTMLResponse)
def index():
    return (ROOT / "index.html").read_text(encoding="utf-8")


@app.get("/class-record-template")
def download_class_record_template():
    workbook = ROOT / "with_macros_-_2026_blank_grading_sheet_with_autotransmuter.xlsm"
    return FileResponse(
        workbook,
        filename=workbook.name,
        media_type="application/vnd.ms-excel.sheet.macroEnabled.12",
    )


def read_upload(workbook):
    if Path(workbook.filename or "").suffix.lower() not in {".xlsx", ".xlsm"}:
        raise HTTPException(400, "Choose an Excel .xlsx or .xlsm class record.")
    contents = workbook.file.read(UPLOAD_LIMIT + 1)
    if len(contents) > UPLOAD_LIMIT:
        raise HTTPException(413, "The class record exceeds 10 MB.")
    return contents


@app.post("/worksheets")
def list_worksheets(workbook: Annotated[UploadFile, File()]):
    with open_class_record(read_upload(workbook)) as book:
        sheets = {sheet.title: sheet for sheet in book.worksheets}
        return [
            {"name": name, "compatible": name in sheets and worksheet_columns(sheets[name]) is not None}
            for name in book.sheetnames
        ]


@app.post("/generate")
def generate(
    workbook: Annotated[UploadFile, File()],
    section: Annotated[str, Form(max_length=60)],
    course_code: Annotated[str, Form(max_length=30)],
    academic_year: Annotated[str, Form(max_length=20)],
    course_description: Annotated[str, Form(max_length=100)],
    term: Annotated[Literal["first", "second", "summer"], Form()],
    course_type: Annotated[Literal["lecture", "laboratory", "rle"], Form()],
    output_format: Annotated[Literal["combined", "individual"], Form()] = "combined",
    worksheet: Annotated[str | None, Form()] = None,
):
    contents = read_upload(workbook)
    fields = {
        "Section": section.strip(),
        "CourseCode": course_code.strip(),
        "AcadYear": academic_year.strip(),
        "CourseDescription": course_description.strip(),
    }
    if any(not value or any(ord(char) < 32 for char in value) for value in fields.values()):
        raise HTTPException(400, "Enter all class details on a single line per field.")
    students = read_students(contents, worksheet)

    if output_format == "individual":
        return Response(
            build_student_zip(students, fields, term, course_type),
            media_type="application/zip",
            headers={
                "Content-Disposition": 'attachment; filename="student-progress-sheets.zip"',
                "Cache-Control": "no-store",
            },
        )
    return Response(
        build_document(students, fields, term, course_type),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": 'attachment; filename="student-progress-sheets.docx"',
            "Cache-Control": "no-store",
        },
    )
