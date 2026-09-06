from flask import Flask, render_template, request, redirect, url_for
import os
import re
import sqlite3
import hashlib
import zipfile
from datetime import datetime, date
from difflib import SequenceMatcher
from werkzeug.utils import secure_filename

# Document readers
from pypdf import PdfReader
from docx import Document
from openpyxl import load_workbook
from pptx import Presentation


# ============================================================
# FLASK CONFIGURATION
# ============================================================

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

UPLOAD_FOLDER = os.path.join(
    BASE_DIR,
    "uploads"
)

DATABASE = os.path.join(
    BASE_DIR,
    "facultyflow.db"
)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# Maximum upload size = 50 MB
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

os.makedirs(
    UPLOAD_FOLDER,
    exist_ok=True
)


# ============================================================
# SUPPORTED DOCUMENT TYPES
# ============================================================

ALLOWED_EXTENSIONS = {
    "pdf",
    "doc",
    "docx",
    "xls",
    "xlsx",
    "ppt",
    "pptx",
    "zip",
    "jpg",
    "jpeg",
    "png",
    "txt"
}


def allowed_file(filename):
    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower()
        in ALLOWED_EXTENSIONS
    )


# ============================================================
# DATABASE CONNECTION
# ============================================================

def get_db():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn


# ============================================================
# DATABASE INITIALIZATION
# ============================================================

def init_db():

    conn = get_db()
    cursor = conn.cursor()

    # --------------------------------------------------------
    # FACULTY RECORDS
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            faculty_name TEXT,
            department TEXT,
            activity TEXT,
            date TEXT,
            duration TEXT,
            document_type TEXT,
            status TEXT,
            created_at TEXT
        )
    """)

    # --------------------------------------------------------
    # TASKS
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            deadline TEXT,
            importance TEXT,
            created_at TEXT
        )
    """)

    # --------------------------------------------------------
    # UPLOADED DOCUMENTS
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT,
            file_path TEXT,
            file_hash TEXT,
            normalized_text TEXT,
            category TEXT,
            uploaded_at TEXT
        )
    """)

    # --------------------------------------------------------
    # FACULTY CLASSES
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS faculty_classes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            subject_name TEXT,
            subject_code TEXT,
            section TEXT,
            semester TEXT,
            academic_year TEXT,
            student_count INTEGER,
            classes_scheduled INTEGER,
            classes_completed INTEGER,
            syllabus_completion REAL,
            created_at TEXT
        )
    """)

    # --------------------------------------------------------
    # STUDENTS
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_name TEXT,
            usn TEXT,
            subject TEXT,
            previous_sem REAL,
            cia REAL,
            mid_sem REAL,
            end_sem REAL,
            project_marks REAL,
            assignment_marks REAL,
            attendance REAL,
            remarks TEXT,
            created_at TEXT
        )
    """)

    # --------------------------------------------------------
    # STUDENT PROJECTS
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_title TEXT,
            student_team TEXT,
            subject TEXT,
            guide TEXT,
            review1 REAL,
            review2 REAL,
            review3 REAL,
            final_evaluation REAL,
            status TEXT,
            remarks TEXT,
            created_at TEXT
        )
    """)

    # --------------------------------------------------------
    # CLUB ACTIVITIES
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS club_activities (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            club_name TEXT,
            activity TEXT,
            faculty_role TEXT,
            activity_date TEXT,
            students_involved INTEGER,
            status TEXT,
            remarks TEXT,
            created_at TEXT
        )
    """)

    # --------------------------------------------------------
    # FACULTY PROFESSIONAL ACTIVITIES
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS faculty_activities (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            activity_type TEXT,
            activity_name TEXT,
            organization TEXT,
            activity_date TEXT,
            role TEXT,
            status TEXT,
            remarks TEXT,
            created_at TEXT
        )
    """)

    # --------------------------------------------------------
    # FACULTY PROFILE
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS faculty_profile (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            faculty_name TEXT,
            employee_id TEXT,
            department TEXT,
            designation TEXT,
            email TEXT,
            phone TEXT,
            specialization TEXT,
            joining_date TEXT,
            created_at TEXT
        )
    """)

    # --------------------------------------------------------
    # FACULTY ATTENDANCE
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS faculty_attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            faculty_name TEXT,
            attendance_date TEXT,
            status TEXT,
            leave_type TEXT,
            remarks TEXT,
            created_at TEXT
        )
    """)

    conn.commit()
    conn.close()


init_db()


# ============================================================
# DATABASE HELPERS
# ============================================================

def get_all_records():

    conn = get_db()

    records = conn.execute("""
        SELECT *
        FROM records
        ORDER BY id DESC
    """).fetchall()

    conn.close()

    return records


def get_all_tasks():

    conn = get_db()

    tasks = conn.execute("""
        SELECT *
        FROM tasks
        ORDER BY id DESC
    """).fetchall()

    conn.close()

    return tasks


def get_all_documents():

    conn = get_db()

    documents = conn.execute("""
        SELECT *
        FROM documents
        ORDER BY id DESC
    """).fetchall()

    conn.close()

    return documents


def get_all_classes():

    conn = get_db()

    classes = conn.execute("""
        SELECT *
        FROM faculty_classes
        ORDER BY id DESC
    """).fetchall()

    conn.close()

    return classes


def get_all_students():

    conn = get_db()

    students = conn.execute("""
        SELECT *
        FROM students
        ORDER BY id DESC
    """).fetchall()

    conn.close()

    return students


def get_all_projects():

    conn = get_db()

    projects = conn.execute("""
        SELECT *
        FROM projects
        ORDER BY id DESC
    """).fetchall()

    conn.close()

    return projects


def get_all_club_activities():

    conn = get_db()

    activities = conn.execute("""
        SELECT *
        FROM club_activities
        ORDER BY id DESC
    """).fetchall()

    conn.close()

    return activities


def get_all_faculty_activities():

    conn = get_db()

    activities = conn.execute("""
        SELECT *
        FROM faculty_activities
        ORDER BY id DESC
    """).fetchall()

    conn.close()

    return activities


def get_faculty_profile():

    conn = get_db()

    profile = conn.execute("""
        SELECT *
        FROM faculty_profile
        ORDER BY id DESC
        LIMIT 1
    """).fetchone()

    conn.close()

    return profile


def get_all_faculty_attendance():

    conn = get_db()

    attendance = conn.execute("""
        SELECT *
        FROM faculty_attendance
        ORDER BY attendance_date DESC, id DESC
    """).fetchall()

    conn.close()

    return attendance


# ============================================================
# ADD RECORD
# ============================================================

def add_record(
    faculty_name,
    department,
    activity,
    record_date,
    duration,
    document_type,
    status
):

    conn = get_db()

    conn.execute("""
        INSERT INTO records
        (
            faculty_name,
            department,
            activity,
            date,
            duration,
            document_type,
            status,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        faculty_name,
        department,
        activity,
        record_date,
        duration,
        document_type,
        status,
        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ))

    conn.commit()
    conn.close()


# ============================================================
# ADD TASK
# ============================================================

def add_task(
    name,
    deadline,
    importance
):

    conn = get_db()

    conn.execute("""
        INSERT INTO tasks
        (
            name,
            deadline,
            importance,
            created_at
        )
        VALUES (?, ?, ?, ?)
    """, (
        name,
        deadline,
        importance,
        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ))

    conn.commit()
    conn.close()


def delete_task(task_id):

    conn = get_db()

    conn.execute("""
        DELETE FROM tasks
        WHERE id = ?
    """, (task_id,))

    conn.commit()
    conn.close()


# ============================================================
# ADD DOCUMENT
# ============================================================

def add_document(
    filename,
    file_path,
    file_hash,
    normalized_text,
    category
):

    conn = get_db()

    conn.execute("""
        INSERT INTO documents
        (
            filename,
            file_path,
            file_hash,
            normalized_text,
            category,
            uploaded_at
        )
        VALUES (?, ?, ?, ?, ?, ?)
    """, (
        filename,
        file_path,
        file_hash,
        normalized_text,
        category,
        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ))

    conn.commit()
    conn.close()


# ============================================================
# TEXT CLEANING
# ============================================================

def clean_text(text):

    if not text:
        return ""

    text = text.replace("\x00", " ")

    text = re.sub(
        r"[ \t]+",
        " ",
        text
    )

    text = re.sub(
        r"\n\s*\n+",
        "\n",
        text
    )

    return text.strip()


# ============================================================
# TXT EXTRACTION
# ============================================================

def extract_txt(file_path):

    try:

        with open(
            file_path,
            "r",
            encoding="utf-8",
            errors="ignore"
        ) as file:

            return clean_text(
                file.read()
            )

    except Exception as e:

        return f"Unable to read TXT file: {e}"


# ============================================================
# PDF EXTRACTION
# ============================================================

def extract_pdf(file_path):

    text_parts = []

    try:

        reader = PdfReader(file_path)

        for page in reader.pages:

            try:

                page_text = page.extract_text()

                if page_text:
                    text_parts.append(page_text)

            except Exception:
                continue

        return clean_text(
            "\n".join(text_parts)
        )

    except Exception as e:

        return f"Unable to read PDF file: {e}"


# ============================================================
# DOCX EXTRACTION
# ============================================================

def extract_docx(file_path):

    try:

        document = Document(file_path)

        parts = []

        for paragraph in document.paragraphs:

            if paragraph.text.strip():

                parts.append(
                    paragraph.text.strip()
                )

        for table in document.tables:

            for row in table.rows:

                row_values = []

                for cell in row.cells:

                    row_values.append(
                        cell.text.strip()
                    )

                if any(row_values):

                    parts.append(
                        " | ".join(row_values)
                    )

        return clean_text(
            "\n".join(parts)
        )

    except Exception as e:

        return f"Unable to read DOCX file: {e}"


# ============================================================
# XLSX EXTRACTION
# ============================================================

def extract_xlsx(file_path):

    try:

        workbook = load_workbook(
            file_path,
            data_only=True
        )

        parts = []

        for sheet in workbook.worksheets:

            parts.append(
                f"Sheet: {sheet.title}"
            )

            for row in sheet.iter_rows(
                values_only=True
            ):

                values = []

                for value in row:

                    if value is not None:

                        values.append(
                            str(value)
                        )

                if values:

                    parts.append(
                        " | ".join(values)
                    )

        return clean_text(
            "\n".join(parts)
        )

    except Exception as e:

        return f"Unable to read XLSX file: {e}"


# ============================================================
# PPTX EXTRACTION
# ============================================================

def extract_pptx(file_path):

    try:

        presentation = Presentation(
            file_path
        )

        parts = []

        for slide_number, slide in enumerate(
            presentation.slides,
            start=1
        ):

            parts.append(
                f"Slide {slide_number}:"
            )

            for shape in slide.shapes:

                if hasattr(shape, "text"):

                    if shape.text.strip():

                        parts.append(
                            shape.text.strip()
                        )

        return clean_text(
            "\n".join(parts)
        )

    except Exception as e:

        return f"Unable to read PPTX file: {e}"


# ============================================================
# ZIP INSPECTION
# ============================================================

def inspect_zip(file_path):

    try:

        parts = []

        with zipfile.ZipFile(
            file_path,
            "r"
        ) as archive:

            parts.append(
                "ZIP contains:"
            )

            for filename in archive.namelist()[:100]:

                parts.append(filename)

        return clean_text(
            "\n".join(parts)
        )

    except Exception as e:

        return f"Unable to inspect ZIP file: {e}"


# ============================================================
# UNIVERSAL DOCUMENT EXTRACTION
# ============================================================

def extract_document_content(
    file_path,
    extension
):

    extension = extension.lower()

    if extension == "txt":

        return extract_txt(file_path)

    elif extension == "pdf":

        return extract_pdf(file_path)

    elif extension == "docx":

        return extract_docx(file_path)

    elif extension == "xlsx":

        return extract_xlsx(file_path)

    elif extension == "pptx":

        return extract_pptx(file_path)

    elif extension == "zip":

        return inspect_zip(file_path)

    elif extension in {
        "doc",
        "xls",
        "ppt"
    }:

        return (
            "Legacy Office document uploaded "
            "successfully."
        )

    elif extension in {
        "jpg",
        "jpeg",
        "png"
    }:

        return (
            "Image document uploaded successfully. "
            "OCR can be added for automatic text extraction."
        )

    return ""


# ============================================================
# DOCUMENT CATEGORY DETECTION
# ============================================================

def detect_document_category(
    text,
    filename
):

    combined = (
        text + " " + filename
    ).lower()

    if any(word in combined for word in [
        "marks",
        "result",
        "grade",
        "score",
        "gpa",
        "cgpa"
    ]):

        return "Marks / Academic Result"

    if any(word in combined for word in [
        "attendance",
        "present",
        "absent",
        "attendance percentage"
    ]):

        return "Attendance Record"

    if any(word in combined for word in [
        "research paper",
        "journal",
        "publication",
        "abstract",
        "methodology",
        "references",
        "conference paper"
    ]):

        return "Research / Academic Paper"

    if any(word in combined for word in [
        "faculty development",
        "fdp",
        "training",
        "workshop",
        "seminar",
        "professional development"
    ]):

        return "Faculty Development / Training"

    if any(word in combined for word in [
        "meeting",
        "minutes",
        "agenda",
        "committee meeting"
    ]):

        return "Meeting / Minutes"

    if any(word in combined for word in [
        "notice",
        "circular",
        "announcement",
        "notification"
    ]):

        return "Notice / Circular"

    if any(word in combined for word in [
        "project",
        "project report",
        "implementation",
        "final report"
    ]):

        return "Project / Report"

    if any(word in combined for word in [
        "assignment",
        "coursework",
        "submission"
    ]):

        return "Assignment / Submission"

    if any(word in combined for word in [
        "timetable",
        "schedule",
        "class schedule"
    ]):

        return "Timetable / Schedule"

    return "General Faculty Document"


# ============================================================
# KEYWORD EXTRACTION
# ============================================================

def extract_keywords(text):

    if not text:
        return []

    words = re.findall(
        r"\b[A-Za-z][A-Za-z0-9-]{3,}\b",
        text.lower()
    )

    stop_words = {
        "this",
        "that",
        "with",
        "from",
        "have",
        "were",
        "been",
        "will",
        "into",
        "about",
        "their",
        "there",
        "which",
        "where",
        "when",
        "faculty",
        "department",
        "document"
    }

    frequency = {}

    for word in words:

        if word in stop_words:
            continue

        frequency[word] = (
            frequency.get(word, 0) + 1
        )

    sorted_words = sorted(
        frequency.items(),
        key=lambda x: x[1],
        reverse=True
    )

    return [
        word
        for word, count
        in sorted_words[:10]
    ]


# ============================================================
# AI SUMMARY
# ============================================================

def generate_summary(
    text,
    category
):

    if not text:

        return (
            f"FacultyFlow AI identified this "
            f"document as {category}. "
            f"No readable text was detected."
        )

    sentences = re.split(
        r"(?<=[.!?])\s+",
        text
    )

    sentences = [
        sentence.strip()
        for sentence in sentences
        if sentence.strip()
    ]

    if sentences:

        preview = " ".join(
            sentences[:3]
        )

    else:

        preview = text[:500]

    if len(preview) > 600:

        preview = (
            preview[:600]
            + "..."
        )

    return (
        f"AI classified this document as "
        f"{category}. "
        f"Content summary: {preview}"
    )


# ============================================================
# FIELD EXTRACTION
# ============================================================

def extract_field(
    text,
    field_name
):

    pattern = (
        rf"(?im)^\s*"
        rf"{re.escape(field_name)}"
        rf"\s*:\s*(.*?)\s*$"
    )

    match = re.search(
        pattern,
        text
    )

    if match:

        value = match.group(1).strip()

        if value:
            return value

    return "Not specified"


# ============================================================
# FACULTY INFORMATION EXTRACTION
# ============================================================

def extract_possible_faculty_information(
    text,
    category
):

    fields = {}

    field_names = [
        "Faculty Name",
        "Department",
        "Activity",
        "Description",
        "Date",
        "Duration",
        "Document Type",
        "Status"
    ]

    for field in field_names:

        fields[field] = extract_field(
            text,
            field
        )

    # --------------------------------------------------------
    # DATE DETECTION
    # --------------------------------------------------------

    if fields["Date"] == "Not specified":

        date_patterns = [

            r"\b\d{1,2}\s+"
            r"(January|February|March|April|May|June|July|"
            r"August|September|October|November|December)"
            r"\s+\d{4}\b",

            r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b",

            r"\b\d{4}[/-]\d{1,2}[/-]\d{1,2}\b"
        ]

        for pattern in date_patterns:

            match = re.search(
                pattern,
                text,
                re.IGNORECASE
            )

            if match:

                fields["Date"] = match.group(0)

                break

    # --------------------------------------------------------
    # FALLBACKS
    # --------------------------------------------------------

    if fields["Activity"] == "Not specified":

        fields["Activity"] = category

    if fields["Document Type"] == "Not specified":

        fields["Document Type"] = category

    if fields["Status"] == "Not specified":

        fields["Status"] = "Uploaded"

    return fields


# ============================================================
# FILE HASH
# ============================================================

def calculate_file_hash(file_path):

    sha256 = hashlib.sha256()

    with open(
        file_path,
        "rb"
    ) as file:

        while True:

            chunk = file.read(8192)

            if not chunk:
                break

            sha256.update(chunk)

    return sha256.hexdigest()


# ============================================================
# DUPLICATE DOCUMENT DETECTOR
# ============================================================

def check_duplicate_document(
    file_hash,
    normalized_text
):

    conn = get_db()

    documents = conn.execute("""
        SELECT *
        FROM documents
    """).fetchall()

    conn.close()

    # Exact duplicate

    for document in documents:

        if document["file_hash"] == file_hash:

            return {
                "duplicate": True,
                "reason":
                    "Exact duplicate file detected."
            }

    # Similar content

    if normalized_text:

        new_text = normalized_text.lower()

        for document in documents:

            old_text = (
                document["normalized_text"]
                or ""
            ).lower()

            if (
                len(new_text) > 50
                and len(old_text) > 50
            ):

                similarity = SequenceMatcher(
                    None,
                    new_text[:10000],
                    old_text[:10000]
                ).ratio()

                if similarity >= 0.85:

                    return {
                        "duplicate": True,
                        "reason":
                            "A highly similar document "
                            "already exists."
                    }

    return {
        "duplicate": False,
        "reason": ""
    }


# ============================================================
# UNIVERSAL DOCUMENT ANALYZER
# ============================================================

def analyze_document(
    filename,
    file_path,
    extension
):

    text = extract_document_content(
        file_path,
        extension
    )

    category = detect_document_category(
        text,
        filename
    )

    keywords = extract_keywords(
        text
    )

    summary = generate_summary(
        text,
        category
    )

    faculty_info = (
        extract_possible_faculty_information(
            text,
            category
        )
    )

    record_data = {

        "faculty_name":
            faculty_info.get(
                "Faculty Name",
                "Not specified"
            ),

        "department":
            faculty_info.get(
                "Department",
                "Not specified"
            ),

        "activity":
            faculty_info.get(
                "Activity",
                category
            ),

        "date":
            faculty_info.get(
                "Date",
                "Not specified"
            ),

        "duration":
            faculty_info.get(
                "Duration",
                "Not specified"
            ),

        "document_type":
            faculty_info.get(
                "Document Type",
                category
            ),

        "status":
            faculty_info.get(
                "Status",
                "Uploaded"
            )
    }

    warnings = []

    if (
        not text
        or len(text.strip()) < 10
    ):

        warnings.append(
            "Limited readable content was detected. "
            "The file can still be saved."
        )

    if extension.lower() in {
        "jpg",
        "jpeg",
        "png"
    }:

        warnings.append(
            "Image uploaded successfully. "
            "OCR can be added for automatic text extraction."
        )

    if extension.lower() in {
        "doc",
        "xls",
        "ppt"
    }:

        warnings.append(
            "Legacy Office format detected. "
            "The document was uploaded successfully."
        )

    return {

        "filename": filename,

        "file_type": extension.upper(),

        "category": category,

        "summary": summary,

        "keywords": keywords,

        "metadata": {

            "File Name": filename,

            "File Type": extension.upper(),

            "Detected Category": category,

            "Content Length": len(text)
        },

        "faculty_info": faculty_info,

        "record_data": record_data,

        "warnings": warnings,

        "can_save_record": True
    }


# ============================================================
# LOGIN
# ============================================================

@app.route("/")
def home():

    return render_template(
        "login.html"
    )


# ============================================================
# DASHBOARD
# ============================================================

@app.route("/dashboard")
def dashboard():

    records = get_all_records()
    tasks = get_all_tasks()
    documents = get_all_documents()

    total_tasks = len(tasks)

    pending_tasks = len(tasks)

    high_priority = 0

    for task in tasks:

        priority = calculate_priority(
            task["deadline"],
            task["importance"]
        )

        if priority == "High":

            high_priority += 1

    completed_records = len(records)

    return render_template(
        "dashboard.html",
        total_tasks=total_tasks,
        pending_tasks=pending_tasks,
        high_priority=high_priority,
        completed_records=completed_records,
        records=records,
        tasks=tasks,
        documents=documents
    )


# ============================================================
# DOCUMENTS
# ============================================================

@app.route(
    "/documents",
    methods=["GET"]
)
def documents():

    documents = get_all_documents()

    return render_template(
        "documents.html",
        documents=documents
    )


# ============================================================
# DOCUMENT UPLOAD
# ============================================================

@app.route(
    "/upload",
    methods=["POST"]
)
def upload():

    if "document" not in request.files:

        return "No document selected."

    file = request.files["document"]

    if file.filename == "":

        return "No document selected."

    if not allowed_file(file.filename):

        return "Unsupported file type."

    filename = secure_filename(
        file.filename
    )

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    name, extension = os.path.splitext(
        filename
    )

    unique_filename = (
        f"{name}_{timestamp}{extension}"
    )

    file_path = os.path.join(
        app.config["UPLOAD_FOLDER"],
        unique_filename
    )

    file.save(file_path)

    extension_without_dot = (
        extension
        .replace(".", "")
        .lower()
    )

    # HASH

    file_hash = calculate_file_hash(
        file_path
    )

    # EXTRACT CONTENT

    extracted_text = (
        extract_document_content(
            file_path,
            extension_without_dot
        )
    )

    normalized_text = re.sub(
        r"\s+",
        " ",
        extracted_text
    ).strip()

    # DUPLICATE CHECK

    duplicate_result = (
        check_duplicate_document(
            file_hash,
            normalized_text
        )
    )

    # AI ANALYSIS

    analysis = analyze_document(
        unique_filename,
        file_path,
        extension_without_dot
    )

    if duplicate_result["duplicate"]:

        analysis["warnings"].append(
            "⚠️ "
            + duplicate_result["reason"]
        )

    # STORE DOCUMENT

    add_document(
        unique_filename,
        file_path,
        file_hash,
        normalized_text,
        analysis["category"]
    )

    return render_template(
        "analysis.html",
        analysis=analysis
    )


# ============================================================
# SAVE TO FACULTY RECORDS
# ============================================================

@app.route(
    "/save_record",
    methods=["POST"]
)
def save_record():

    faculty_name = request.form.get(
        "faculty_name",
        ""
    ).strip()

    department = request.form.get(
        "department",
        ""
    ).strip()

    activity = request.form.get(
        "activity",
        ""
    ).strip()

    record_date = request.form.get(
        "date",
        ""
    ).strip()

    duration = request.form.get(
        "duration",
        ""
    ).strip()

    document_type = request.form.get(
        "document_type",
        ""
    ).strip()

    status = request.form.get(
        "status",
        ""
    ).strip()

    faculty_name = (
        faculty_name
        or "Not specified"
    )

    department = (
        department
        or "Not specified"
    )

    activity = (
        activity
        or "General Faculty Activity"
    )

    record_date = (
        record_date
        or "Not specified"
    )

    duration = (
        duration
        or "Not specified"
    )

    document_type = (
        document_type
        or "General Faculty Document"
    )

    status = (
        status
        or "Uploaded"
    )

    records = get_all_records()

    for record in records:

        if (
            record["faculty_name"]
            == faculty_name
            and record["activity"]
            == activity
            and record["date"]
            == record_date
        ):

            return render_template(
                "duplicate.html",
                message=(
                    "This record already exists "
                    "in Faculty Records."
                )
            )

    add_record(
        faculty_name,
        department,
        activity,
        record_date,
        duration,
        document_type,
        status
    )

    return redirect(
        url_for("records")
    )


# ============================================================
# RECORDS
# ============================================================

@app.route("/records")
def records():

    records = get_all_records()

    return render_template(
        "records.html",
        records=records
    )


# ============================================================
# TASK PRIORITY ENGINE
# ============================================================

def calculate_priority(
    deadline,
    importance
):

    try:

        deadline_date = datetime.strptime(
            deadline,
            "%Y-%m-%d"
        ).date()

        today = date.today()

        days_left = (
            deadline_date - today
        ).days

    except Exception:

        return "Medium"

    importance = (
        importance or ""
    ).lower()

    if "critical" in importance:

        return "High"

    if days_left <= 1:

        return "High"

    if days_left <= 3:

        return "Medium"

    return "Low"


# ============================================================
# TASK MANAGER
# ============================================================

@app.route("/tasks")
def tasks():

    tasks = get_all_tasks()

    task_data = []

    for task in tasks:

        priority = calculate_priority(
            task["deadline"],
            task["importance"]
        )

        task_data.append({

            "id": task["id"],

            "name": task["name"],

            "deadline":
                task["deadline"],

            "importance":
                task["importance"],

            "priority":
                priority
        })

    return render_template(
        "tasks.html",
        tasks=task_data
    )


@app.route(
    "/add_task",
    methods=["POST"]
)
def add_new_task():

    task_name = request.form.get(
        "task",
        ""
    ).strip()

    deadline = request.form.get(
        "deadline",
        ""
    ).strip()

    importance = request.form.get(
        "importance",
        ""
    ).strip()

    if task_name and deadline:

        add_task(
            task_name,
            deadline,
            importance
        )

    return redirect(
        url_for("tasks")
    )


@app.route(
    "/complete_task/<int:task_id>",
    methods=["POST"]
)
def complete_task(task_id):

    delete_task(task_id)

    return redirect(
        url_for("tasks")
    )


# ============================================================
# AI ASSISTANT
# ============================================================

@app.route("/ai-assistant")
def ai_assistant():

    return render_template(
        "ai_assistant.html",
        response=(
            "Hello! 👋 I am FacultyFlow AI. "
            "I can help you manage documents, "
            "tasks, students, classes, attendance, "
            "projects and faculty records."
        )
    )


@app.route(
    "/ask_ai",
    methods=["POST"]
)
def ask_ai():

    question = request.form.get(
        "question",
        ""
    ).strip().lower()

    tasks = get_all_tasks()
    records = get_all_records()
    students = get_all_students()
    classes = get_all_classes()
    projects = get_all_projects()
    documents = get_all_documents()
    faculty_attendance = get_all_faculty_attendance()

    response = ""

    # --------------------------------------------------------
    # PRIORITY
    # --------------------------------------------------------

    if (
        "priority" in question
        or "urgent" in question
        or "important" in question
    ):

        high_tasks = []

        for task in tasks:

            priority = calculate_priority(
                task["deadline"],
                task["importance"]
            )

            if priority == "High":

                high_tasks.append(task)

        if high_tasks:

            response = (
                "<strong>"
                "🔴 High Priority Tasks"
                "</strong><br><br>"
            )

            for task in high_tasks:

                response += (
                    f"• {task['name']} "
                    f"(Deadline: "
                    f"{task['deadline']})"
                    "<br>"
                )

        else:

            response = (
                "🎉 You currently have "
                "no high-priority tasks."
            )

    # --------------------------------------------------------
    # FACULTY ATTENDANCE
    # --------------------------------------------------------

    elif (
        "faculty attendance" in question
        or "teacher attendance" in question
        or "faculty present" in question
        or "teacher present" in question
        or "faculty absent" in question
        or "teacher absent" in question
        or "faculty leave" in question
        or "teacher leave" in question
    ):

        if faculty_attendance:

            total = len(
                faculty_attendance
            )

            present = len([
                row
                for row in faculty_attendance
                if row["status"].lower()
                == "present"
            ])

            absent = len([
                row
                for row in faculty_attendance
                if row["status"].lower()
                == "absent"
            ])

            leave = len([
                row
                for row in faculty_attendance
                if row["status"].lower()
                == "leave"
            ])

            percentage = round(
                (present / total) * 100,
                1
            )

            response = (
                "<strong>"
                "👩‍🏫 Faculty Attendance Intelligence"
                "</strong><br><br>"
                f"Attendance records: {total}<br>"
                f"Days present: {present}<br>"
                f"Days absent: {absent}<br>"
                f"Leave days: {leave}<br>"
                f"Attendance percentage: "
                f"{percentage}%<br><br>"
                "💡 Faculty attendance is "
                "being tracked using the "
                "FacultyFlow attendance records."
            )

        else:

            response = (
                "No faculty attendance records "
                "have been entered yet."
            )

    # --------------------------------------------------------
    # TASKS
    # --------------------------------------------------------

    elif (
        "pending" in question
        or "tasks" in question
        or "task" in question
    ):

        if tasks:

            response = (
                "<strong>"
                "📋 Your Current Tasks"
                "</strong><br><br>"
            )

            for task in tasks:

                priority = calculate_priority(
                    task["deadline"],
                    task["importance"]
                )

                response += (
                    f"• {task['name']} — "
                    f"{priority} Priority — "
                    f"Due {task['deadline']}"
                    "<br>"
                )

        else:

            response = (
                "✅ You have no pending tasks."
            )

    # --------------------------------------------------------
    # WORKLOAD
    # --------------------------------------------------------

    elif (
        "workload" in question
        or "manage" in question
        or "overview" in question
    ):

        high_count = 0

        for task in tasks:

            if calculate_priority(
                task["deadline"],
                task["importance"]
            ) == "High":

                high_count += 1

        response = (
            "<strong>"
            "📊 Faculty Workload Intelligence"
            "</strong><br><br>"
            f"Pending tasks: {len(tasks)}<br>"
            f"High-priority tasks: "
            f"{high_count}<br>"
            f"Uploaded documents: "
            f"{len(documents)}<br>"
            f"Faculty records: "
            f"{len(records)}<br>"
            f"Classes: "
            f"{len(classes)}<br>"
            f"Students: "
            f"{len(students)}<br>"
            f"Projects: "
            f"{len(projects)}<br>"
            f"Faculty attendance records: "
            f"{len(faculty_attendance)}<br><br>"
            "💡 Recommendation: "
            "Complete high-priority tasks first "
            "and review students and faculty "
            "attendance requiring attention."
        )

    # --------------------------------------------------------
    # STUDENTS
    # --------------------------------------------------------

    elif (
        "student" in question
        or "students" in question
    ):

        if students:

            low_attendance = [

                student

                for student in students

                if float(
                    student["attendance"] or 0
                ) < 75
            ]

            response = (
                "<strong>"
                "👨‍🎓 Student Overview"
                "</strong><br><br>"
                f"Students recorded: "
                f"{len(students)}<br>"
                f"Students below 75% attendance: "
                f"{len(low_attendance)}<br>"
            )

            if low_attendance:

                response += (
                    "<br><strong>"
                    "⚠️ Students needing "
                    "attendance attention:"
                    "</strong><br>"
                )

                for student in low_attendance:

                    response += (
                        f"• "
                        f"{student['student_name']} "
                        f"— "
                        f"{student['attendance']}%"
                        "<br>"
                    )

        else:

            response = (
                "No student data has been entered yet."
            )

    # --------------------------------------------------------
    # STUDENT ATTENDANCE
    # --------------------------------------------------------

    elif (
        "attendance" in question
        or "absent" in question
        or "present" in question
    ):

        if students:

            attendance_values = [

                float(
                    student["attendance"] or 0
                )

                for student in students
            ]

            average = round(
                sum(attendance_values)
                / len(attendance_values),
                1
            )

            low_count = len([

                student

                for student in students

                if float(
                    student["attendance"] or 0
                ) < 75
            ])

            response = (
                "<strong>"
                "📅 Student Attendance Intelligence"
                "</strong><br><br>"
                f"Average student attendance: "
                f"{average}%<br>"
                f"Students below 75%: "
                f"{low_count}<br><br>"
                "Students below the attendance "
                "threshold should be reviewed."
            )

        else:

            response = (
                "No student attendance data "
                "has been entered."
            )

    # --------------------------------------------------------
    # MARKS
    # --------------------------------------------------------

    elif (
        "marks" in question
        or "cia" in question
        or "mid sem" in question
        or "end sem" in question
    ):

        if students:

            cia_values = [

                float(
                    student["cia"] or 0
                )

                for student in students
            ]

            average_cia = round(
                sum(cia_values)
                / len(cia_values),
                1
            )

            response = (
                "<strong>"
                "📊 Academic Performance"
                "</strong><br><br>"
                f"Students recorded: "
                f"{len(students)}<br>"
                f"Average CIA marks: "
                f"{average_cia}<br><br>"
                "FacultyFlow AI can compare "
                "previous semester, CIA, mid-semester, "
                "end-semester and project performance."
            )

        else:

            response = (
                "No student marks have been entered yet."
            )

    # --------------------------------------------------------
    # CLASSES
    # --------------------------------------------------------

    elif (
        "class" in question
        or "classes" in question
        or "subject" in question
    ):

        if classes:

            response = (
                "<strong>"
                "📚 Teaching Overview"
                "</strong><br><br>"
            )

            for classroom in classes:

                response += (
                    f"• "
                    f"{classroom['subject_name']} "
                    f"({classroom['subject_code']}) "
                    f"— Section "
                    f"{classroom['section']} "
                    f"— "
                    f"{classroom['student_count']} students"
                    "<br>"
                )

        else:

            response = (
                "No classes have been entered yet."
            )

    # --------------------------------------------------------
    # PROJECTS
    # --------------------------------------------------------

    elif (
        "project" in question
        or "projects" in question
    ):

        response = (
            "<strong>"
            "🧪 Project Overview"
            "</strong><br><br>"
            f"Projects recorded: "
            f"{len(projects)}<br><br>"
            "Project information includes "
            "students, guides, reviews, evaluations "
            "and status."
        )

    # --------------------------------------------------------
    # DOCUMENTS
    # --------------------------------------------------------

    elif (
        "document" in question
        or "record" in question
    ):

        response = (
            "<strong>"
            "📁 Document & Record Overview"
            "</strong><br><br>"
            f"Uploaded documents: "
            f"{len(documents)}<br>"
            f"Faculty records: "
            f"{len(records)}<br><br>"
            "Documents can be uploaded, analyzed "
            "and saved to Faculty Records."
        )

    # --------------------------------------------------------
    # DEFAULT
    # --------------------------------------------------------

    else:

        response = (
            "<strong>"
            "🤖 FacultyFlow AI can help with:"
            "</strong><br><br>"
            "• Classes and subjects<br>"
            "• Student information<br>"
            "• Student attendance analysis<br>"
            "• Faculty attendance analysis<br>"
            "• CIA / Mid-Sem / End-Sem marks<br>"
            "• Projects<br>"
            "• Club activities<br>"
            "• Faculty activities<br>"
            "• High-priority tasks<br>"
            "• Documents and records<br>"
            "• Faculty workload"
        )

    return render_template(
        "ai_assistant.html",
        response=response
    )


# ============================================================
# FACULTY MANAGEMENT
# ============================================================

@app.route(
    "/faculty-management"
)
def faculty_management():

    classes = get_all_classes()

    students = get_all_students()

    projects = get_all_projects()

    club_activities = (
        get_all_club_activities()
    )

    faculty_activities = (
        get_all_faculty_activities()
    )

    faculty_attendance = (
        get_all_faculty_attendance()
    )

    faculty_profile = (
        get_faculty_profile()
    )

    # --------------------------------------------------------
    # BASIC STATISTICS
    # --------------------------------------------------------

    total_classes = len(classes)

    total_students = len(students)

    total_projects = len(projects)

    total_club_activities = (
        len(club_activities)
    )

    # --------------------------------------------------------
    # STUDENT ATTENDANCE
    # --------------------------------------------------------

    if students:

        attendance_values = [

            float(
                student["attendance"] or 0
            )

            for student in students
        ]

        average_attendance = round(
            sum(attendance_values)
            / len(attendance_values),
            1
        )

    else:

        average_attendance = 0

    low_attendance = [

        student

        for student in students

        if float(
            student["attendance"] or 0
        ) < 75
    ]

    # --------------------------------------------------------
    # AVERAGE CIA
    # --------------------------------------------------------

    if students:

        cia_values = [

            float(
                student["cia"] or 0
            )

            for student in students
        ]

        average_cia = round(
            sum(cia_values)
            / len(cia_values),
            1
        )

    else:

        average_cia = 0

    # --------------------------------------------------------
    # SYLLABUS COMPLETION
    # --------------------------------------------------------

    if classes:

        syllabus_values = [

            float(
                classroom[
                    "syllabus_completion"
                ] or 0
            )

            for classroom in classes
        ]

        average_syllabus = round(
            sum(syllabus_values)
            / len(syllabus_values),
            1
        )

    else:

        average_syllabus = 0

    # --------------------------------------------------------
    # FACULTY ATTENDANCE INTELLIGENCE
    # --------------------------------------------------------

    faculty_present = len([

        row

        for row in faculty_attendance

        if row["status"].lower()
        == "present"
    ])

    faculty_absent = len([

        row

        for row in faculty_attendance

        if row["status"].lower()
        == "absent"
    ])

    faculty_leave = len([

        row

        for row in faculty_attendance

        if row["status"].lower()
        == "leave"
    ])

    total_faculty_attendance = (
        len(faculty_attendance)
    )

    if total_faculty_attendance > 0:

        faculty_attendance_percentage = round(
            (
                faculty_present
                / total_faculty_attendance
            ) * 100,
            1
        )

    else:

        faculty_attendance_percentage = 0

    return render_template(

        "faculty_management.html",

        classes=classes,

        students=students,

        projects=projects,

        club_activities=
            club_activities,

        faculty_activities=
            faculty_activities,

        faculty_attendance=
            faculty_attendance,

        faculty_profile=
            faculty_profile,

        total_classes=
            total_classes,

        total_students=
            total_students,

        total_projects=
            total_projects,

        total_club_activities=
            total_club_activities,

        average_attendance=
            average_attendance,

        average_cia=
            average_cia,

        average_syllabus=
            average_syllabus,

        low_attendance=
            low_attendance,

        faculty_present=
            faculty_present,

        faculty_absent=
            faculty_absent,

        faculty_leave=
            faculty_leave,

        faculty_attendance_percentage=
            faculty_attendance_percentage
    )


# ============================================================
# ADD FACULTY PROFILE
# ============================================================

@app.route(
    "/add_faculty_profile",
    methods=["POST"]
)
def add_faculty_profile():

    faculty_name = request.form.get(
        "faculty_name",
        ""
    ).strip()

    employee_id = request.form.get(
        "employee_id",
        ""
    ).strip()

    department = request.form.get(
        "department",
        ""
    ).strip()

    designation = request.form.get(
        "designation",
        ""
    ).strip()

    email = request.form.get(
        "email",
        ""
    ).strip()

    phone = request.form.get(
        "phone",
        ""
    ).strip()

    specialization = request.form.get(
        "specialization",
        ""
    ).strip()

    joining_date = request.form.get(
        "joining_date",
        ""
    ).strip()

    conn = get_db()

    conn.execute("""
        INSERT INTO faculty_profile
        (
            faculty_name,
            employee_id,
            department,
            designation,
            email,
            phone,
            specialization,
            joining_date,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (

        faculty_name,
        employee_id,
        department,
        designation,
        email,
        phone,
        specialization,
        joining_date,

        datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    ))

    conn.commit()
    conn.close()

    return redirect(
        url_for(
            "faculty_management"
        )
    )


# ============================================================
# ADD FACULTY ATTENDANCE
# ============================================================

@app.route(
    "/add_faculty_attendance",
    methods=["POST"]
)
def add_faculty_attendance():

    faculty_name = request.form.get(
        "faculty_name",
        ""
    ).strip()

    attendance_date = request.form.get(
        "attendance_date",
        ""
    ).strip()

    status = request.form.get(
        "status",
        ""
    ).strip()

    leave_type = request.form.get(
        "leave_type",
        ""
    ).strip()

    remarks = request.form.get(
        "remarks",
        ""
    ).strip()

    if not faculty_name:

        profile = get_faculty_profile()

        if profile:

            faculty_name = (
                profile["faculty_name"]
            )

        else:

            faculty_name = "Faculty"

    conn = get_db()

    # Prevent duplicate attendance
    existing = conn.execute("""
        SELECT id
        FROM faculty_attendance
        WHERE faculty_name = ?
        AND attendance_date = ?
    """, (
        faculty_name,
        attendance_date
    )).fetchone()

    if existing:

        conn.close()

        return redirect(
            url_for(
                "faculty_management"
            )
        )

    conn.execute("""
        INSERT INTO faculty_attendance
        (
            faculty_name,
            attendance_date,
            status,
            leave_type,
            remarks,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?)
    """, (

        faculty_name,
        attendance_date,
        status,
        leave_type,
        remarks,

        datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    ))

    conn.commit()
    conn.close()

    return redirect(
        url_for(
            "faculty_management"
        )
    )


# ============================================================
# ADD CLASS
# ============================================================

@app.route(
    "/add_class",
    methods=["POST"]
)
def add_class():

    subject_name = request.form.get(
        "subject_name",
        ""
    ).strip()

    subject_code = request.form.get(
        "subject_code",
        ""
    ).strip()

    section = request.form.get(
        "section",
        ""
    ).strip()

    semester = request.form.get(
        "semester",
        ""
    ).strip()

    academic_year = request.form.get(
        "academic_year",
        ""
    ).strip()

    student_count = request.form.get(
        "student_count",
        0
    )

    classes_scheduled = request.form.get(
        "classes_scheduled",
        0
    )

    classes_completed = request.form.get(
        "classes_completed",
        0
    )

    syllabus_completion = request.form.get(
        "syllabus_completion",
        0
    )

    conn = get_db()

    conn.execute("""
        INSERT INTO faculty_classes
        (
            subject_name,
            subject_code,
            section,
            semester,
            academic_year,
            student_count,
            classes_scheduled,
            classes_completed,
            syllabus_completion,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (

        subject_name,
        subject_code,
        section,
        semester,
        academic_year,
        student_count,
        classes_scheduled,
        classes_completed,
        syllabus_completion,

        datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    ))

    conn.commit()
    conn.close()

    return redirect(
        url_for(
            "faculty_management"
        )
    )


# ============================================================
# ADD STUDENT
# ============================================================

@app.route(
    "/add_student",
    methods=["POST"]
)
def add_student():

    student_name = request.form.get(
        "student_name",
        ""
    ).strip()

    usn = request.form.get(
        "usn",
        ""
    ).strip()

    subject = request.form.get(
        "subject",
        ""
    ).strip()

    previous_sem = request.form.get(
        "previous_sem",
        0
    )

    cia = request.form.get(
        "cia",
        0
    )

    mid_sem = request.form.get(
        "mid_sem",
        0
    )

    end_sem = request.form.get(
        "end_sem",
        0
    )

    project_marks = request.form.get(
        "project_marks",
        0
    )

    assignment_marks = request.form.get(
        "assignment_marks",
        0
    )

    attendance = request.form.get(
        "attendance",
        0
    )

    remarks = request.form.get(
        "remarks",
        ""
    ).strip()

    conn = get_db()

    conn.execute("""
        INSERT INTO students
        (
            student_name,
            usn,
            subject,
            previous_sem,
            cia,
            mid_sem,
            end_sem,
            project_marks,
            assignment_marks,
            attendance,
            remarks,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (

        student_name,
        usn,
        subject,
        previous_sem,
        cia,
        mid_sem,
        end_sem,
        project_marks,
        assignment_marks,
        attendance,
        remarks,

        datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    ))

    conn.commit()
    conn.close()

    return redirect(
        url_for(
            "faculty_management"
        )
    )


# ============================================================
# ADD PROJECT
# ============================================================

@app.route(
    "/add_project",
    methods=["POST"]
)
def add_project():

    project_title = request.form.get(
        "project_title",
        ""
    ).strip()

    student_team = request.form.get(
        "student_team",
        ""
    ).strip()

    subject = request.form.get(
        "subject",
        ""
    ).strip()

    guide = request.form.get(
        "guide",
        ""
    ).strip()

    review1 = request.form.get(
        "review1",
        0
    )

    review2 = request.form.get(
        "review2",
        0
    )

    review3 = request.form.get(
        "review3",
        0
    )

    final_evaluation = request.form.get(
        "final_evaluation",
        0
    )

    status = request.form.get(
        "status",
        ""
    ).strip()

    remarks = request.form.get(
        "remarks",
        ""
    ).strip()

    conn = get_db()

    conn.execute("""
        INSERT INTO projects
        (
            project_title,
            student_team,
            subject,
            guide,
            review1,
            review2,
            review3,
            final_evaluation,
            status,
            remarks,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (

        project_title,
        student_team,
        subject,
        guide,
        review1,
        review2,
        review3,
        final_evaluation,
        status,
        remarks,

        datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    ))

    conn.commit()
    conn.close()

    return redirect(
        url_for(
            "faculty_management"
        )
    )


# ============================================================
# ADD CLUB ACTIVITY
# ============================================================

@app.route(
    "/add_club_activity",
    methods=["POST"]
)
def add_club_activity():

    club_name = request.form.get(
        "club_name",
        ""
    ).strip()

    activity = request.form.get(
        "activity",
        ""
    ).strip()

    faculty_role = request.form.get(
        "faculty_role",
        ""
    ).strip()

    activity_date = request.form.get(
        "activity_date",
        ""
    ).strip()

    students_involved = request.form.get(
        "students_involved",
        0
    )

    status = request.form.get(
        "status",
        ""
    ).strip()

    remarks = request.form.get(
        "remarks",
        ""
    ).strip()

    conn = get_db()

    conn.execute("""
        INSERT INTO club_activities
        (
            club_name,
            activity,
            faculty_role,
            activity_date,
            students_involved,
            status,
            remarks,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (

        club_name,
        activity,
        faculty_role,
        activity_date,
        students_involved,
        status,
        remarks,

        datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    ))

    conn.commit()
    conn.close()

    return redirect(
        url_for(
            "faculty_management"
        )
    )


# ============================================================
# ADD FACULTY ACTIVITY
# ============================================================

@app.route(
    "/add_faculty_activity",
    methods=["POST"]
)
def add_faculty_activity():

    activity_type = request.form.get(
        "activity_type",
        ""
    ).strip()

    activity_name = request.form.get(
        "activity_name",
        ""
    ).strip()

    organization = request.form.get(
        "organization",
        ""
    ).strip()

    activity_date = request.form.get(
        "activity_date",
        ""
    ).strip()

    role = request.form.get(
        "role",
        ""
    ).strip()

    status = request.form.get(
        "status",
        ""
    ).strip()

    remarks = request.form.get(
        "remarks",
        ""
    ).strip()

    conn = get_db()

    conn.execute("""
        INSERT INTO faculty_activities
        (
            activity_type,
            activity_name,
            organization,
            activity_date,
            role,
            status,
            remarks,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (

        activity_type,
        activity_name,
        organization,
        activity_date,
        role,
        status,
        remarks,

        datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    ))

    conn.commit()
    conn.close()

    return redirect(
        url_for(
            "faculty_management"
        )
    )


# ============================================================
# RUN APPLICATION
# ============================================================

if __name__ == "__main__":

    app.run(
        debug=True
    )