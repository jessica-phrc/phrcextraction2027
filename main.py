import pymupdf
import re
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# SETTINGS
# ============================================================

# Put all of your PHRC PDFs in a folder called "pdfs"
PDF_FOLDER = Path("your/path/here")

# Name of the Excel file that will be created
OUTPUT_FILE = "PHRC_metric_scores.xlsx"


# ============================================================
# PDF TEXT EXTRACTION
# ============================================================

def extract_pages(pdf_path):
    """
    Extract text from a PDF page-by-page.
    """

    doc = pymupdf.open(pdf_path)

    pages = []

    for page_number, page in enumerate(doc, start=1):

        text = page.get_text("text")

        # Clean a few common PDF characters
        text = (
            text
            .replace("\u00a0", " ")
            .replace("ﬁ", "fi")
            .replace("ﬂ", "fl")
        )

        pages.append({
            "page": page_number,
            "text": text
        })

    doc.close()

    return pages


# ============================================================
# EXTRACT INSTITUTION + DISCIPLINE
# ============================================================

def extract_institution_info(pdf_path):
    """
    Extracts institution and discipline from the first page.

    Example:

        Planetary Health Report Card
        (Medicine) 2026:
        Dalhousie University
    """

    pages = extract_pages(pdf_path)

    first_page = pages[0]["text"]

    # -----------------------------
    # Discipline
    # -----------------------------

    discipline_match = re.search(
        r"\(([^()]+)\)\s*2026",
        first_page,
        flags=re.IGNORECASE
    )

    if discipline_match:
        discipline = discipline_match.group(1).strip()
    else:
        discipline = ""

    # -----------------------------
    # Institution
    # -----------------------------

    institution_match = re.search(
        r"\)\s*2026\s*:\s*\n?\s*([^\n]+)",
        first_page,
        flags=re.IGNORECASE
    )

    if institution_match:
        institution = institution_match.group(1).strip()

    else:
        # Fallback: use filename
        institution = pdf_path.stem

    return {
        "institution": institution,
        "discipline": discipline
    }


# ============================================================
# EXTRACT METRIC SCORES
# ============================================================

def extract_metric_scores(pdf_path):
    """
    Extracts all metrics with an explicit:

        Score Assigned: X

    Example:

        1.7. Does your medical school curriculum...
        ...
        Score Assigned: 3

    Returns a dataframe like:

        metric    score_assigned    page
        1.1       2                 8
        1.2       3                 9
        ...
    """

    pages = extract_pages(pdf_path)

    records = []

    current_metric = None

    # Metric headings:
    #
    # 1.1. Does...
    # 1.10. Does...
    # 2.4. Does...
    #
    metric_pattern = re.compile(
        r"(?m)^\s*(\d+\.\d+)\.?\s+(?=[A-Z])"
    )

    # Handles:
    #
    # Score Assigned: 2
    #
    # or
    #
    # Score Assigned:
    # 2
    #
    score_pattern = re.compile(
        r"Score\s*Assigned\s*:\s*"
        r"([0-9]+(?:\.[0-9]+)?)",
        flags=re.IGNORECASE
    )

    for page_data in pages:

        page_number = page_data["page"]
        text = page_data["text"]

        events = []

        # Find metric headings
        for match in metric_pattern.finditer(text):

            events.append({
                "position": match.start(),
                "type": "metric",
                "value": match.group(1)
            })

        # Find assigned scores
        for match in score_pattern.finditer(text):

            events.append({
                "position": match.start(),
                "type": "score",
                "value": float(match.group(1))
            })

        # Process in the order they occur on the page
        events.sort(
            key=lambda x: x["position"]
        )

        for event in events:

            if event["type"] == "metric":

                current_metric = event["value"]

            elif event["type"] == "score":

                if current_metric is None:
                    continue

                score = event["value"]

                # 2.0 -> 2
                if score.is_integer():
                    score = int(score)

                records.append({
                    "metric": current_metric,
                    "score_assigned": score,
                    "page": page_number
                })

    df = pd.DataFrame(records)

    # Remove duplicate metric entries if extraction
    # happens to see one twice.
    if not df.empty:

        df = (
            df
            .drop_duplicates(
                subset="metric",
                keep="first"
            )
            .reset_index(drop=True)
        )

    return df


# ============================================================
# EXTRACT FINAL GRADING TABLE
# ============================================================

def extract_grading_table(pdf_path):
    """
    Extracts the table in the Grading section near the end
    of the PDF.

    Extracts:

        Curriculum raw score / % / grade
        Interdisciplinary Research raw score / % / grade
        Community Outreach raw score / % / grade
        Student-Led Initiatives raw score / % / grade
        Campus Sustainability raw score / % / grade
        Institutional % / grade
    """

    pages = extract_pages(pdf_path)

    # Grading section should be near the end
    ending_text = "\n".join(
        page["text"]
        for page in pages[-5:]
    )

    grading_match = re.search(
        r"(?m)^\s*Grading\s*$",
        ending_text,
        flags=re.IGNORECASE
    )

    if not grading_match:
        return {}

    grading_text = ending_text[
        grading_match.start():
    ]

    result = {}

    section_patterns = {

        "Curriculum":
            r"Planetary Health Curriculum",

        "Interdisciplinary Research":
            r"Interdisciplinary Research",

        "Community Outreach and Advocacy":
            r"Community Outreach\s+and\s+Advocacy",

        "Student-Led Initiatives":
            r"Support for Student[- ]led Planetary\s+Health Initiatives",

        "Campus Sustainability":
            r"Campus Sustainability"
    }

    # ========================================================
    # SECTION GRADES
    # ========================================================

    for section_name, section_pattern in section_patterns.items():

        heading = re.search(
            section_pattern,
            grading_text,
            flags=re.IGNORECASE
        )

        if not heading:
            continue

        # Examine text immediately following heading
        block = grading_text[
            heading.end():
            heading.end() + 350
        ]

        # Looks for:
        #
        # (43/69) x 100 = 63.77% B-
        #
        score_match = re.search(
            r"\(\s*(\d+)\s*/\s*(\d+)\s*\)"
            r".{0,100}?"
            r"([0-9]+(?:\.[0-9]+)?)\s*%"
            r"\s*([ABCDF][+-]?)",
            block,
            flags=re.IGNORECASE | re.DOTALL
        )

        if score_match:

            numerator = int(
                score_match.group(1)
            )

            denominator = int(
                score_match.group(2)
            )

            percentage = float(
                score_match.group(3)
            )

            grade = (
                score_match
                .group(4)
                .upper()
            )

            result[
                f"{section_name} Raw Score"
            ] = f"{numerator}/{denominator}"

            result[
                f"{section_name} %"
            ] = percentage

            result[
                f"{section_name} Grade"
            ] = grade

    # ========================================================
    # INSTITUTIONAL GRADE
    # ========================================================

    institutional_match = re.search(
        r"Institutional\s+Grade"
        r".{0,500}?"
        r"([0-9]+(?:\.[0-9]+)?)\s*%"
        r"\s*([ABCDF][+-]?)",
        grading_text,
        flags=re.IGNORECASE | re.DOTALL
    )

    if institutional_match:

        result["Institutional %"] = float(
            institutional_match.group(1)
        )

        result["Institutional Grade"] = (
            institutional_match
            .group(2)
            .upper()
        )

    return result


# ============================================================
# EXTRACT ONE PHRC PDF
# ============================================================

def extract_phrc(pdf_path):
    """
    Extract everything needed from one PHRC PDF.
    """

    info = extract_institution_info(
        pdf_path
    )

    metrics = extract_metric_scores(
        pdf_path
    )

    grading = extract_grading_table(
        pdf_path
    )

    return {
        "institution": info["institution"],
        "discipline": info["discipline"],
        "metrics": metrics,
        "grading": grading
    }


# ============================================================
# METRIC SORTING
# ============================================================

def metric_sort_key(metric):
    """
    Makes columns sort as:

    1.1
    1.2
    1.3
    ...
    1.9
    1.10
    1.11

    rather than alphabetical sorting.
    """

    try:

        section, number = str(metric).split(".")

        return (
            int(section),
            int(number)
        )

    except Exception:

        return (
            999,
            999
        )


# ============================================================
# PROCESS ALL PDF FILES
# ============================================================

metric_rows = []
grading_rows = []

pdf_files = sorted(
    PDF_FOLDER.glob("*.pdf")
)

print(
    f"Found {len(pdf_files)} PDF files."
)

for pdf_path in pdf_files:

    print(
        f"Processing: {pdf_path.name}"
    )

    try:

        result = extract_phrc(
            pdf_path
        )

        institution = result[
            "institution"
        ]

        discipline = result[
            "discipline"
        ]

        metrics = result[
            "metrics"
        ]

        grading = result[
            "grading"
        ]

        # ====================================================
        # METRIC SCORES ROW
        # ====================================================

        metric_row = {

            "Institution":
                institution,

            "Discipline":
                discipline,

            "File":
                pdf_path.name
        }

        for _, row in metrics.iterrows():

            metric = str(
                row["metric"]
            )

            score = row[
                "score_assigned"
            ]

            metric_row[
                metric
            ] = score

        metric_rows.append(
            metric_row
        )

        # ====================================================
        # GRADING SUMMARY ROW
        # ====================================================

        grading_row = {

            "Institution":
                institution,

            "Discipline":
                discipline,

            "File":
                pdf_path.name
        }

        grading_row.update(
            grading
        )

        grading_rows.append(
            grading_row
        )

    except Exception as e:

        print(
            f"ERROR processing "
            f"{pdf_path.name}: {e}"
        )


# ============================================================
# CREATE HORIZONTAL METRIC DATAFRAME
# ============================================================

metric_df = pd.DataFrame(
    metric_rows
)

metadata_columns = [
    "Institution",
    "Discipline",
    "File"
]

metric_columns = [

    column

    for column in metric_df.columns

    if column
    not in metadata_columns
]

metric_columns = sorted(
    metric_columns,
    key=metric_sort_key
)

metric_df = metric_df[
    metadata_columns
    +
    metric_columns
]


# ============================================================
# CREATE GRADING SUMMARY DATAFRAME
# ============================================================

grading_df = pd.DataFrame(
    grading_rows
)

grading_columns = [

    "Institution",
    "Discipline",
    "File",

    "Curriculum Raw Score",
    "Curriculum %",
    "Curriculum Grade",

    "Interdisciplinary Research Raw Score",
    "Interdisciplinary Research %",
    "Interdisciplinary Research Grade",

    "Community Outreach and Advocacy Raw Score",
    "Community Outreach and Advocacy %",
    "Community Outreach and Advocacy Grade",

    "Student-Led Initiatives Raw Score",
    "Student-Led Initiatives %",
    "Student-Led Initiatives Grade",

    "Campus Sustainability Raw Score",
    "Campus Sustainability %",
    "Campus Sustainability Grade",

    "Institutional %",
    "Institutional Grade"
]

# Some PDFs may be missing fields.
# Only retain columns that actually exist.
grading_columns = [

    column

    for column in grading_columns

    if column in grading_df.columns
]

grading_df = grading_df[
    grading_columns
]


# ============================================================
# WRITE TO EXCEL
# ============================================================

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    # Individual metric scores
    metric_df.to_excel(
        writer,
        sheet_name="Metric Scores",
        index=False
    )

    # Grading table
    grading_df.to_excel(
        writer,
        sheet_name="Grading Summary",
        index=False
    )


# ============================================================
# FORMAT EXCEL WORKBOOK
# ============================================================

wb = load_workbook(
    OUTPUT_FILE
)


for ws in wb.worksheets:

    # Freeze top row
    ws.freeze_panes = "A2"

    # ----------------------------------------
    # Header formatting
    # ----------------------------------------

    for cell in ws[1]:

        cell.font = Font(
            bold=True,
            color="FFFFFF"
        )

        cell.fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78"
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    # ----------------------------------------
    # Filters
    # ----------------------------------------

    ws.auto_filter.ref = (
        ws.dimensions
    )

    # ----------------------------------------
    # Column widths
    # ----------------------------------------

    for column_cells in ws.columns:

        max_length = 0

        column_letter = (
            get_column_letter(
                column_cells[0].column
            )
        )

        for cell in column_cells:

            if cell.value is None:
                continue

            value_length = len(
                str(cell.value)
            )

            max_length = max(
                max_length,
                value_length
            )

        width = min(
            max(
                max_length + 2,
                10
            ),
            35
        )

        ws.column_dimensions[
            column_letter
        ].width = width


# ============================================================
# FORMAT GRADING PERCENTAGES
# ============================================================

if "Grading Summary" in wb.sheetnames:

    ws = wb[
        "Grading Summary"
    ]

    header_lookup = {

        cell.value:
            cell.column

        for cell in ws[1]
    }

    percentage_columns = [

        "Curriculum %",

        "Interdisciplinary Research %",

        "Community Outreach and Advocacy %",

        "Student-Led Initiatives %",

        "Campus Sustainability %",

        "Institutional %"
    ]

    for header in percentage_columns:

        column_number = (
            header_lookup.get(
                header
            )
        )

        if column_number is None:
            continue

        for row_number in range(
            2,
            ws.max_row + 1
        ):

            ws.cell(
                row=row_number,
                column=column_number
            ).number_format = (
                '0.00"%"'
            )


# ============================================================
# SAVE
# ============================================================

wb.save(
    OUTPUT_FILE
)

print()
print(
    "Finished!"
)

print(
    f"Excel file saved as: {OUTPUT_FILE}"
)