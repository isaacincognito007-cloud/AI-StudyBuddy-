import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
import win32com.client
import shutil

SRC_DOCX = r"d:\AI-StudyBuddy\LearnMate - AI assistant for smarter learning and study support.docx"
OUT_DOCX_D = r"d:\AI-StudyBuddy\Study_Buddy_AI_Report.docx"
OUT_PDF_D = r"d:\AI-StudyBuddy\Study_Buddy_AI_Report.pdf"
OUT_DOCX_DL = r"c:\Users\acer\Downloads\Study_Buddy_AI_Report.docx"
OUT_PDF_DL = r"c:\Users\acer\Downloads\Study_Buddy_AI_Report.pdf"

doc = docx.Document(SRC_DOCX)

# 1. Update first heading in body
doc.paragraphs[0].text = "Study Buddy AI (LearnMate) – AI Assistant for Smarter Learning and Study Support"

# Change paragraph 1 ("Description") to "Introduction"
for p in doc.paragraphs[:10]:
    if p.text.strip() == "Description":
        p.text = "Introduction"
        break

# Update API Testing section headers in the body so they match the TOC
for i, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if txt == "API Testing":
        p.text = "Authentication & Material API Testing"
    elif txt == "4. Generate AI Summary":
        p_ai = p.insert_paragraph_before()
        p_ai.style = doc.styles['Heading 2']
        p_ai.text = "AI API Testing"
        p_ai.paragraph_format.space_before = Pt(12)
        p_ai.paragraph_format.space_after = Pt(6)
        break

# Add formal Conclusion section before links at the end
p_concl_h = doc.add_paragraph()
p_concl_h.style = doc.styles['Heading 2']
p_concl_h.text = "Conclusion"
p_concl_h.paragraph_format.space_before = Pt(14)
p_concl_h.paragraph_format.space_after = Pt(6)

p_concl_body = doc.add_paragraph()
p_concl_body.paragraph_format.space_after = Pt(10)
p_concl_body.paragraph_format.line_spacing = 1.15
p_concl_body.add_run(
    "Study Buddy AI (LearnMate) successfully delivers an intelligent, secure, and scalable backend platform "
    "for automated educational resource generation. By uniting Node.js and Express.js RESTful architecture with "
    "the Google Gemini 2.5 Flash Generative AI model and MongoDB, the system transforms raw study materials into "
    "structured revision assets—including concise summaries, active-recall flashcards, four-option multiple choice "
    "quizzes, and day-by-day exam preparation schedules. Comprehensive defense-in-depth measures, including bcryptjs "
    "salted hashing, stateless JWT authentication with role-based access control (RBAC), custom Express 5 NoSQL-injection "
    "sanitization, and two-tier rate limiting, ensure that student data remains isolated and protected against malicious inputs. "
    "All RESTful endpoints have been comprehensively asserted through Postman test suites, establishing an enterprise-ready "
    "backend ready for seamless frontend client integration and future feature expansion."
)

# Add GitHub repo link (Updated to https://github.com/ILAZAHAMEDSK/AI-STUDYBUDDY)
p_repo = doc.add_paragraph()
p_repo.paragraph_format.space_before = Pt(8)
r1 = p_repo.add_run("GitHub Repository Link: ")
r1.bold = True
r1.font.size = Pt(11)
r2 = p_repo.add_run("https://github.com/ILAZAHAMEDSK/AI-STUDYBUDDY")
r2.font.size = Pt(11)
r2.font.color.rgb = RGBColor(37, 99, 235)

# Insert Cover Page (Page 1) and Table of Contents (Page 2) BEFORE paragraph 0
first_para = doc.paragraphs[0]

def add_p_before(target, text="", font_size=12, bold=False, color=(30, 41, 59), space_before=0, space_after=6, align=WD_ALIGN_PARAGRAPH.CENTER):
    p = target.insert_paragraph_before()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.25
    if text:
        r = p.add_run(text)
        r.font.name = "Calibri"
        r.font.size = Pt(font_size)
        r.bold = bold
        r.font.color.rgb = RGBColor(*color)
    return p

def add_label_val_before(target, label, val, font_size=13, space_before=2, space_after=5):
    p = target.insert_paragraph_before()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.25
    
    r1 = p.add_run(label)
    r1.font.name = "Calibri"
    r1.font.size = Pt(font_size)
    r1.bold = True
    r1.font.color.rgb = RGBColor(51, 65, 85)
    
    r2 = p.add_run(val)
    r2.font.name = "Calibri"
    r2.font.size = Pt(font_size)
    r2.bold = False
    r2.font.color.rgb = RGBColor(15, 23, 42)
    return p

# --- COVER PAGE ---
add_p_before(first_para, "TN Skills", 30, True, (15, 23, 42), space_before=150, space_after=14)
add_p_before(first_para, "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", 11, True, (37, 99, 235), space_before=0, space_after=18)
add_p_before(first_para, "PROJECT NAME", 11, True, (100, 116, 139), space_before=4, space_after=2)
add_p_before(first_para, "Study Buddy AI", 24, True, (30, 58, 138), space_before=0, space_after=18)
add_label_val_before(first_para, "Subject Name: ", "AI-Augmented Backend Development - A&S", 14, space_before=4, space_after=6)
add_label_val_before(first_para, "Subject Code: ", "[Subject Code]", 14, space_before=2, space_after=24)

add_p_before(first_para, "TEAM LEADER", 12, True, (30, 58, 138), space_before=8, space_after=4)
add_label_val_before(first_para, "ILAZ AHAMED S K", "  —  [Reg. No]", 13, space_before=2, space_after=18)

add_p_before(first_para, "TEAMMATES", 12, True, (30, 58, 138), space_before=6, space_after=6)
teammates = [
    ("BILL CLINTON J", "  —  [Reg. No]"),
    ("MANIKANDAN T", "  —  [Reg. No]"),
    ("DAVISH SHYAM S", "  —  [Reg. No]"),
    ("SAN MATHEW M", "  —  [Reg. No]"),
]
for name, reg in teammates:
    add_label_val_before(first_para, name, reg, 13, space_before=2, space_after=5)

# Page Break after Cover Page
p_pb1 = first_para.insert_paragraph_before()
p_pb1.add_run().add_break(WD_BREAK.PAGE)

# --- TABLE OF CONTENT PAGE ---
p_toc_heading = first_para.insert_paragraph_before()
p_toc_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_toc_heading.paragraph_format.space_before = Pt(30)
p_toc_heading.paragraph_format.space_after = Pt(20)
r_th = p_toc_heading.add_run("TABLE OF CONTENT")
r_th.font.name = "Calibri"
r_th.font.size = Pt(18)
r_th.bold = True
r_th.font.color.rgb = RGBColor(15, 23, 42)

# Insert Table of Content table before first_para
toc_rows = [
    ("1.", "Introduction", "1-2"),
    ("2.", "Scenario-Based Case Study", "2-3"),
    ("3.", "Usage", "3-4"),
    ("4.", "Outcome", "4-5"),
    ("5.", "Epic 1: Project Architecture", "5-6"),
    ("6.", "Components Definition", "6-7"),
    ("7.", "Entity Relationship (ER) Diagram Description", "7-9"),
    ("8.", "Key Features", "9-10"),
    ("9.", "Roles and Responsibilities", "10"),
    ("10.", "User Flow", "10-11"),
    ("11.", "MVC Architecture Pattern", "11-12"),
    ("12.", "Epic 2: Project Setup and Configuration", "12-14"),
    ("13.", "Epic 3: Backend Development", "14-18"),
    ("14.", "Epic 4: Database Configuration", "18-20"),
    ("15.", "Epic 5: Project Execution", "20-21"),
    ("16.", "Authentication & Material API Testing", "21-23"),
    ("17.", "AI API Testing", "23-25"),
    ("18.", "Conclusion", "25"),
]

table = doc.add_table(rows=len(toc_rows) + 1, cols=3)
table.alignment = WD_TABLE_ALIGNMENT.CENTER

tblPr = table._tbl.tblPr
tblBorders = parse_xml(r'''
    <w:tblBorders %s>
        <w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>
        <w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>
        <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>
        <w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>
        <w:insideH w:val="single" w:sz="6" w:space="0" w:color="000000"/>
        <w:insideV w:val="single" w:sz="6" w:space="0" w:color="000000"/>
    </w:tblBorders>
''' % nsdecls('w'))
tblPr.append(tblBorders)

col_widths = [Inches(1.0), Inches(4.3), Inches(1.3)]
for row in table.rows:
    for i, w in enumerate(col_widths):
        row.cells[i].width = w

hdr_cells = table.rows[0].cells
headers = ["S.No", "Table of content", "Page no"]
for i, name in enumerate(headers):
    hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = hdr_cells[i].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(name)
    run.font.name = "Calibri"
    run.font.size = Pt(11)
    run.bold = True

for r_idx, (sno, title, pno) in enumerate(toc_rows):
    row_cells = table.rows[r_idx + 1].cells
    for c_idx, val in enumerate([sno, title, pno]):
        cell = row_cells[c_idx]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(2.5)
        p.paragraph_format.space_after = Pt(2.5)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(val)
        run.font.name = "Calibri"
        run.font.size = Pt(10.5)

first_para._p.addprevious(table._tbl)

p_pb2 = first_para.insert_paragraph_before()
p_pb2.add_run().add_break(WD_BREAK.PAGE)

# Save DOCX
doc.save(OUT_DOCX_D)
print("Saved updated DOCX to:", OUT_DOCX_D)

# Convert to PDF
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    wdoc = word.Documents.Open(OUT_DOCX_D)
    wdoc.SaveAs(OUT_PDF_D, FileFormat=17) # wdFormatPDF
    wdoc.Close(False)
    print("Saved updated PDF to:", OUT_PDF_D)
finally:
    word.Quit()

# Sync to Downloads
try:
    shutil.copyfile(OUT_DOCX_D, OUT_DOCX_DL)
    shutil.copyfile(OUT_PDF_D, OUT_PDF_DL)
    print("Successfully synced updated DOCX and PDF to Downloads!")
except Exception as e:
    print("Downloads sync notice:", e)
