from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.styles.colors import Color

wb = Workbook()

# Colors
DARK_BLUE = "1F4E79"
MID_BLUE = "2E75B6"
LIGHT_BLUE = "D6E4F0"
WHITE = "FFFFFF"
LIGHT_GRAY = "F2F2F2"

def header_font():
    return Font(name="Arial", bold=True, color=WHITE, size=11)

def title_font():
    return Font(name="Arial", bold=True, color=DARK_BLUE, size=14)

def normal_font():
    return Font(name="Arial", size=10)

def header_fill():
    return PatternFill("solid", start_color=DARK_BLUE)

def alt_fill():
    return PatternFill("solid", start_color=LIGHT_BLUE)

def white_fill():
    return PatternFill("solid", start_color=WHITE)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

def apply_row(ws, row, cols, fill, font=None):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        if font:
            cell.font = font
        cell.border = border

def add_hyperlink(cell, url, display):
    cell.hyperlink = url
    cell.value = display
    cell.font = Font(name="Arial", size=10, color="0563C1", underline="single")

# ── SHEET 1: o9 OMP Roles ──────────────────────────────────────────────────
ws1 = wb.active
ws1.title = "o9 OMP Roles"

ws1.row_dimensions[1].height = 30
ws1.merge_cells("A1:G1")
ws1["A1"] = "o9 / OMP Consultant Roles — US Job Listings"
ws1["A1"].font = title_font()
ws1["A1"].alignment = center()
ws1["A1"].fill = PatternFill("solid", start_color="DEEAF1")

ws1.row_dimensions[2].height = 8  # spacer

ws1.row_dimensions[3].height = 22
headers1 = ["#", "Job Title", "Company", "Location", "Type", "Link", "Recruiter / HM", "Notes"]
for c, h in enumerate(headers1, 1):
    cell = ws1.cell(row=3, column=c, value=h)
    cell.font = header_font()
    cell.fill = header_fill()
    cell.alignment = center()
    cell.border = border

data1 = [
    (1, "Principal O9 Consultant", "Infosys", "Raleigh, NC", "On-site · Full-time",
     "https://www.linkedin.com/jobs/view/4388512492/", "Managed off LinkedIn", "Reach out via Infosys careers portal"),
    (2, "Lead o9 Consultant ⭐", "Infosys", "Richardson, TX", "On-site · Full-time",
     "https://www.linkedin.com/jobs/view/4378154088/", "Managed off LinkedIn", "HIGH MATCH — Reach out via Infosys careers portal"),
    (3, "O9 Solutions Consultant", "InfoVision Inc.", "New York, US", "Remote · Contract",
     "https://www.linkedin.com/jobs/view/4383145463/", "Not listed", "Easy Apply · Actively reviewing applicants"),
]

for i, row_data in enumerate(data1):
    r = i + 4
    ws1.row_dimensions[r].height = 20
    fill = alt_fill() if i % 2 == 0 else white_fill()
    for c, val in enumerate(row_data, 1):
        cell = ws1.cell(row=r, column=c)
        cell.fill = fill
        cell.border = border
        cell.font = normal_font()
        if c == 1:
            cell.alignment = center()
        else:
            cell.alignment = left()
        if c == 6:
            add_hyperlink(cell, val, "View Job →")
            cell.fill = fill
        else:
            cell.value = val

ws1.freeze_panes = "A4"
col_widths1 = [5, 30, 18, 18, 18, 14, 30, 42]
for i, w in enumerate(col_widths1, 1):
    ws1.column_dimensions[get_column_letter(i)].width = w

# ── SHEET 2: Supply Chain Architect Roles ─────────────────────────────────
ws2 = wb.create_sheet("Supply Chain Architect Roles")

ws2.row_dimensions[1].height = 30
ws2.merge_cells("A1:I1")
ws2["A1"] = "Supply Chain Architect Roles — US Job Listings"
ws2["A1"].font = title_font()
ws2["A1"].alignment = center()
ws2["A1"].fill = PatternFill("solid", start_color="DEEAF1")

ws2.row_dimensions[2].height = 8  # spacer

ws2.row_dimensions[3].height = 22
headers2 = ["#", "Job Title", "Company", "Location", "Type", "Salary", "Link", "Recruiter / HM", "Notes"]
for c, h in enumerate(headers2, 1):
    cell = ws2.cell(row=3, column=c, value=h)
    cell.font = header_font()
    cell.fill = header_fill()
    cell.alignment = center()
    cell.border = border

data2 = [
    (4, "Solutions Architect - Kinaxis ⭐", "Genpact", "USA", "Remote · Full-time", "$150K – $180K",
     "https://www.linkedin.com/jobs/view/4330227649/",
     "Vipul Gupta\nGlobal Talent Acquisition @ Genpact",
     "Easy Apply · Actively reviewing · HIGH MATCH"),
    (5, "Senior Solution Designer ⭐", "Körber Supply Chain", "USA", "Remote · Full-time", "Not listed",
     "https://www.linkedin.com/jobs/view/4382197450/",
     "Michael Unser\nDirector Solution Design, Körber NA",
     "HIGH MATCH"),
    (6, "Sr Functional Solution Architect ⭐", "Blue Yonder", "Coppell, TX", "Remote · Full-time", "Not listed",
     "https://www.linkedin.com/jobs/view/4394225134/",
     "Not listed publicly",
     "HIGH MATCH · Be an early applicant"),
    (7, "Supply Chain Solutions Architect", "Texas Instruments", "Dallas, TX", "On-site · Full-time", "Not listed",
     "https://www.linkedin.com/jobs/view/4382493831/",
     "Not listed publicly",
     ""),
]

for i, row_data in enumerate(data2):
    r = i + 4
    ws2.row_dimensions[r].height = 36
    fill = alt_fill() if i % 2 == 0 else white_fill()
    for c, val in enumerate(row_data, 1):
        cell = ws2.cell(row=r, column=c)
        cell.fill = fill
        cell.border = border
        cell.font = normal_font()
        if c == 1:
            cell.alignment = center()
        else:
            cell.alignment = left()
        if c == 7:
            add_hyperlink(cell, val, "View Job →")
            cell.fill = fill
        else:
            cell.value = val

ws2.freeze_panes = "A4"
col_widths2 = [5, 32, 20, 14, 18, 16, 14, 32, 38]
for i, w in enumerate(col_widths2, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w

wb.save("C:/Users/amolp/Prometheus/job_listings.xlsx")
print("Done: job_listings.xlsx created")
