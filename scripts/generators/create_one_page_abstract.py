import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
import win32com.client
import fitz # PyMuPDF

OUT_DOCX = r"c:\Users\acer\Downloads\Study_Buddy_AI_Abstract.docx"
OUT_PDF = r"c:\Users\acer\Downloads\Study_Buddy_AI_Abstract.pdf"
SCRATCH_PDF = r"C:\Users\acer\.gemini\antigravity\brain\3df468b6-62f3-4271-a7b4-01a98becf01a\scratch\Study_Buddy_AI_Abstract.pdf"
PREVIEW_IMG = r"C:\Users\acer\.gemini\antigravity\brain\3df468b6-62f3-4271-a7b4-01a98becf01a\scratch\abstract_page_1.png"

doc = docx.Document()

# Page Setup: Standard A4 with balanced margins
section = doc.sections[0]
section.page_width = Inches(8.27)
section.page_height = Inches(11.69)
section.top_margin = Inches(0.55)
section.bottom_margin = Inches(0.55)
section.left_margin = Inches(0.65)
section.right_margin = Inches(0.65)

# Styling Constants
COLOR_NAVY = RGBColor(15, 30, 60)       # #0F1E3C
COLOR_TEAL = RGBColor(15, 118, 110)     # #0F766E
COLOR_CHARCOAL = RGBColor(23, 32, 51)   # #172033
COLOR_MUTED = RGBColor(100, 116, 139)   # #64748B

def set_cell_border(cell, **kwargs):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(r'<w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = f'w:{edge}'
            element = parse_xml(
                f'<{tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                f'w:val="{edge_data.get("val", "single")}" '
                f'w:sz="{edge_data.get("sz", "4")}" '
                f'w:space="0" '
                f'w:color="{edge_data.get("color", "auto")}"/>'
            )
            tcBorders.append(element)
    tcPr.append(tcBorders)

def set_cell_shading(cell, color_hex):
    shading_xml = f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="{color_hex}"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))

# -------------------------------------------------------------
# 1. INSTITUTION & PROJECT HEADER BLOCK
# -------------------------------------------------------------
p_inst = doc.add_paragraph()
p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_inst.paragraph_format.space_before = Pt(0)
p_inst.paragraph_format.space_after = Pt(2)
p_inst.paragraph_format.line_spacing = 1.08

r_inst1 = p_inst.add_run("PATRICIAN COLLEGE OF ARTS AND SCIENCE\n")
r_inst1.font.name = "Calibri"
r_inst1.font.size = Pt(13)
r_inst1.font.bold = True
r_inst1.font.color.rgb = COLOR_NAVY

r_inst2 = p_inst.add_run("Affiliated to the University of Madras  •  Adyar, Chennai – 600 020\n")
r_inst2.font.name = "Calibri"
r_inst2.font.size = Pt(9.5)
r_inst2.font.color.rgb = COLOR_MUTED

r_inst3 = p_inst.add_run("DEPARTMENT OF COMPUTER APPLICATIONS  |  TN SKILLS PROJECT RECORD")
r_inst3.font.name = "Calibri"
r_inst3.font.size = Pt(10)
r_inst3.font.bold = True
r_inst3.font.color.rgb = COLOR_TEAL

# -------------------------------------------------------------
# 2. PROJECT TITLE & METADATA (As clean bold points, no highlight box)
# -------------------------------------------------------------
meta_items = [
    ("Project Title", "STUDY BUDDY AI (LearnMate) — AI Assistant for Smarter Learning and Study Support"),
    ("Subject & Degree", "AI-Augmented Backend Development - A&S (NMU88)  |  BCA (III Year – V Semester)"),
    ("Team Leader", "ISAAC SHILOAH HERBERT V (212404414)"),
    ("Team Members", "MOHAMMED IRFAN A (212404420), BHUVANESWARI A (212404441), BENSAM GILBERT E (212404405), JOEL EDWIN P (212404416)")
]

for idx, (lbl, val) in enumerate(meta_items):
    p_item = doc.add_paragraph()
    p_item.paragraph_format.space_before = Pt(1)
    p_item.paragraph_format.space_after = Pt(2)
    p_item.paragraph_format.line_spacing = 1.12
    p_item.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    r_dot = p_item.add_run("• ")
    r_dot.font.name = "Calibri"
    r_dot.font.size = Pt(9.5)
    r_dot.font.bold = True
    r_dot.font.color.rgb = COLOR_TEAL
    
    r_lbl = p_item.add_run(lbl + ": ")
    r_lbl.font.name = "Calibri"
    r_lbl.font.size = Pt(9.5)
    r_lbl.font.bold = True
    r_lbl.font.color.rgb = COLOR_NAVY
    
    r_val = p_item.add_run(val)
    r_val.font.name = "Calibri"
    r_val.font.size = Pt(9)
    r_val.font.bold = (idx == 0 or idx == 2)
    r_val.font.color.rgb = COLOR_CHARCOAL

# -------------------------------------------------------------
# 2. ABSTRACT TITLE
# -------------------------------------------------------------
p_ab_title = doc.add_paragraph()
p_ab_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_ab_title.paragraph_format.space_before = Pt(8)
p_ab_title.paragraph_format.space_after = Pt(5)

r_at = p_ab_title.add_run("PROJECT ABSTRACT")
r_at.font.name = "Calibri"
r_at.font.size = Pt(12)
r_at.font.bold = True
r_at.font.color.rgb = COLOR_NAVY

# -------------------------------------------------------------
# 3. STRUCTURED ABSTRACT BODY PARAGRAPHS
# -------------------------------------------------------------
def add_abstract_para(doc, heading, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = 1.12
    
    r_h = p.add_run(heading + ": ")
    r_h.font.name = "Calibri"
    r_h.font.size = Pt(10)
    r_h.font.bold = True
    r_h.font.color.rgb = COLOR_TEAL
    
    r_b = p.add_run(text)
    r_b.font.name = "Calibri"
    r_b.font.size = Pt(9.5)
    r_b.font.color.rgb = COLOR_CHARCOAL
    return p

add_abstract_para(
    doc,
    "Context & Problem Statement",
    "Modern academic curricula require learners to digest substantial volumes of study content within constrained examination timelines. However, students frequently face severe operational friction: educational resources remain scattered across disparate files and formats, synthesizing comprehensive revision notes is heavily labor-intensive, and manually constructing active-recall flashcards or self-assessment quizzes is repetitive and inefficient. Furthermore, the lack of structured, adaptive study timelines synchronized with upcoming exam schedules leaves students unprepared and unable to objectively gauge their subject readiness."
)

add_abstract_para(
    doc,
    "Proposed System Architecture",
    "To resolve these foundational learning challenges, Study Buddy AI (LearnMate) introduces an intelligent, secure, and scalable educational backend platform that seamlessly unifies web application engineering with Generative Artificial Intelligence. Engineered using Node.js and Express.js under the Model-View-Controller (MVC) architectural pattern, the platform interfaces with MongoDB via Mongoose ODM for resilient data persistence. Incoming learning documents (.txt, .md, .pdf) are ingested through dedicated Multer multipart upload middleware and routed to an AI service layer powered by the Google Gemini 2.5 Flash API (@google/genai), leveraging dynamic prompt engineering and strict JSON schema guarantees to yield highly deterministic educational outputs."
)

add_abstract_para(
    doc,
    "Core Functional Capabilities",
    "Study Buddy AI provides four core student-centric capabilities: (1) AI-Driven Conceptual Summaries that condense lengthy textbooks and lecture materials into concise, high-yield takeaways; (2) Active-Recall Flashcards generating structured question-and-answer pairs designed for optimal spaced repetition; (3) Objective Multiple-Choice Quizzes generating four-option MCQs accompanied by verified correct answers to facilitate quantitative self-assessment; and (4) Personalized Study Planners and Contextual Q&A that formulate structured, day-by-day revision milestones aligned with user availability. Crucially, all AI-generated assets are automatically cataloged in MongoDB under user-scoped collections, enabling students to re-access, review, and filter their learning artifacts at any point."
)

add_abstract_para(
    doc,
    "Security & Administrative Governance",
    "The architecture implements a rigorous multi-tier defense-in-depth security model. User authentication is managed through cryptographically salted password hashing via bcryptjs (10 rounds) and stateless JSON Web Tokens (JWT) featuring 7-day expiration. Server-side middleware strictly enforces Role-Based Access Control (RBAC), segregating Student privileges from Administrative capabilities while preventing privilege escalation. Query-level data isolation ({ userId: req.user._id }) guarantees horizontal privacy across all database operations. Application defenses are further bolstered by custom Express 5 NoSQL-injection sanitizers, Helmet HTTP security headers, and two-tier API rate limiting. Administrators are provided dedicated telemetry endpoints for user oversight and platform AI quota monitoring."
)

add_abstract_para(
    doc,
    "Validation & Experimental Outcome",
    "The platform was comprehensively verified using automated Postman test suites encompassing authentication lifecycles, multipart document uploads, and AI inference endpoints, uniformly demonstrating authenticated HTTP 200/201 status codes and strict schema compliance. Study Buddy AI successfully bridges the gap between raw study content and structured academic preparation—dramatically reducing manual revision overhead, accelerating subject comprehension, and delivering an autonomous, AI-augmented study companion."
)

# -------------------------------------------------------------
# 4. KEYWORDS
# -------------------------------------------------------------
p_kw = doc.add_paragraph()
p_kw.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p_kw.paragraph_format.space_before = Pt(4)
p_kw.paragraph_format.space_after = Pt(6)
p_kw.paragraph_format.line_spacing = 1.08

r_kwh = p_kw.add_run("Keywords: ")
r_kwh.font.name = "Calibri"
r_kwh.font.size = Pt(9.5)
r_kwh.font.bold = True
r_kwh.font.color.rgb = COLOR_NAVY

r_kwb = p_kw.add_run("Artificial Intelligence, Generative AI, Google Gemini 2.5 Flash, Node.js, Express.js REST API, MongoDB, Mongoose ODM, JWT Authentication, Role-Based Access Control (RBAC), Active Recall, EdTech.")
r_kwb.font.name = "Calibri"
r_kwb.font.size = Pt(9)
r_kwb.font.italic = True
r_kwb.font.color.rgb = COLOR_CHARCOAL

# -------------------------------------------------------------
# 5. FOOTER / PROJECT LINKS TABLE
# -------------------------------------------------------------
table_foot = doc.add_table(rows=1, cols=1)
table_foot.alignment = WD_TABLE_ALIGNMENT.CENTER
table_foot.autofit = False
cell_f = table_foot.rows[0].cells[0]
cell_f.width = Inches(6.97)
set_cell_shading(cell_f, "F1F5F9")
set_cell_border(cell_f, 
                top=dict(val='single', sz=4, color='CBD5E1'),
                bottom=dict(val='none'),
                left=dict(val='none'),
                right=dict(val='none'))

p_f = cell_f.paragraphs[0]
p_f.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_f.paragraph_format.space_before = Pt(3)
p_f.paragraph_format.space_after = Pt(3)
p_f.paragraph_format.line_spacing = 1.08

r_fl1 = p_f.add_run("GitHub Repository: ")
r_fl1.font.name = "Calibri"
r_fl1.font.size = Pt(8.5)
r_fl1.font.bold = True
r_fl1.font.color.rgb = COLOR_TEAL

r_fv1 = p_f.add_run("https://github.com/isaacincognito007-cloud/AI-StudyBuddy-    |    ")
r_fv1.font.name = "Calibri"
r_fv1.font.size = Pt(8.5)
r_fv1.font.color.rgb = COLOR_CHARCOAL

r_fl2 = p_f.add_run("Demo Video: ")
r_fl2.font.name = "Calibri"
r_fl2.font.size = Pt(8.5)
r_fl2.font.bold = True
r_fl2.font.color.rgb = COLOR_TEAL

r_fv2 = p_f.add_run("https://youtu.be/N5lNC06rn7w")
r_fv2.font.name = "Calibri"
r_fv2.font.size = Pt(8.5)
r_fv2.font.color.rgb = COLOR_CHARCOAL

doc.save(OUT_DOCX)
print("Saved DOCX:", OUT_DOCX)

# Convert to PDF via Word COM
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc_word = word.Documents.Open(OUT_DOCX)
doc_word.SaveAs(OUT_PDF, FileFormat=17) # 17 = wdFormatPDF
doc_word.SaveAs(SCRATCH_PDF, FileFormat=17)
doc_word.Close()
word.Quit()
print("Saved PDF:", OUT_PDF)

# Verify page count with PyMuPDF
pdf_doc = fitz.open(OUT_PDF)
page_count = len(pdf_doc)
print(f"Total Pages in PDF: {page_count}")
page1 = pdf_doc[0]
pix = page1.get_pixmap(dpi=150)
pix.save(PREVIEW_IMG)
print("Saved Preview Image:", PREVIEW_IMG)
if page_count > 1:
    print("WARNING: Document spilled over to page 2!")
    for p in range(1, page_count):
        print(f"Page {p+1} text length:", len(pdf_doc[p].get_text()))
pdf_doc.close()
