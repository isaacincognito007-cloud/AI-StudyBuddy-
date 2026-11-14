import os
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

SRC_DOCX = r"d:\AI-StudyBuddy\LearnMate - AI assistant for smarter learning and study support.docx"
OUT_DOCX_D = r"d:\AI-StudyBuddy\Study_Buddy_AI_Report.docx"
OUT_PDF_D = r"d:\AI-StudyBuddy\Study_Buddy_AI_Report.pdf"
OUT_DOCX_DL = r"c:\Users\acer\Downloads\Study_Buddy_AI_Report.docx"
OUT_PDF_DL = r"c:\Users\acer\Downloads\Study_Buddy_AI_Report.pdf"

doc = docx.Document(SRC_DOCX)

# First, update the first heading in the body to explicitly include Study Buddy AI
first_body_para = doc.paragraphs[0]
first_body_para.text = "Study Buddy AI (LearnMate) – AI Assistant for Smarter Learning and Study Support"

def insert_centered_para_before(target_para, text="", font_size=12, bold=False, color_rgb=(30, 41, 59), space_before=0, space_after=6, font_name="Calibri"):
    p = target_para.insert_paragraph_before()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.25
    if text:
        run = p.add_run(text)
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.bold = bold
        run.font.color.rgb = RGBColor(*color_rgb)
    return p

def insert_label_value_para_before(target_para, label, value, font_size=13, space_before=4, space_after=6, font_name="Calibri"):
    p = target_para.insert_paragraph_before()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.25
    
    r_label = p.add_run(label)
    r_label.font.name = font_name
    r_label.font.size = Pt(font_size)
    r_label.bold = True
    r_label.font.color.rgb = RGBColor(51, 65, 85)
    
    r_val = p.add_run(value)
    r_val.font.name = font_name
    r_val.font.size = Pt(font_size)
    r_val.bold = False
    r_val.font.color.rgb = RGBColor(15, 23, 42)
    return p

# Insert Cover Page (First Page) before first_body_para
# Position "TN Skills" in the middle/center of the first page, and all other details falling below it
insert_centered_para_before(
    first_body_para,
    text="TN Skills",
    font_size=30,
    bold=True,
    color_rgb=(15, 23, 42),
    space_before=150,  # pushes TN Skills down into the middle-center of the page
    space_after=14
)

# Subtle divider line / accent
insert_centered_para_before(
    first_body_para,
    text="━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    font_size=11,
    bold=True,
    color_rgb=(37, 99, 235),
    space_before=0,
    space_after=18
)

# Project Name
insert_centered_para_before(
    first_body_para,
    text="PROJECT NAME",
    font_size=11,
    bold=True,
    color_rgb=(100, 116, 139),
    space_before=4,
    space_after=2
)
insert_centered_para_before(
    first_body_para,
    text="Study Buddy AI",
    font_size=24,
    bold=True,
    color_rgb=(30, 58, 138),
    space_before=0,
    space_after=18
)

# Subject Name & Subject Code
insert_label_value_para_before(
    first_body_para,
    label="Subject Name: ",
    value="AI-Augmented Backend Development - A&S",
    font_size=14,
    space_before=4,
    space_after=6
)
insert_label_value_para_before(
    first_body_para,
    label="Subject Code: ",
    value="[Subject Code]",
    font_size=14,
    space_before=2,
    space_after=24
)

# Team Leader section
insert_centered_para_before(
    first_body_para,
    text="TEAM LEADER",
    font_size=12,
    bold=True,
    color_rgb=(30, 58, 138),
    space_before=8,
    space_after=4
)
insert_label_value_para_before(
    first_body_para,
    label="ISAAC SHILOAH HERBERT V",
    value="  —  [Reg. No]",
    font_size=13,
    space_before=2,
    space_after=18
)

# Teammates section
insert_centered_para_before(
    first_body_para,
    text="TEAMMATES",
    font_size=12,
    bold=True,
    color_rgb=(30, 58, 138),
    space_before=6,
    space_after=6
)

teammates = [
    ("MOHAMMED IRFAN A", "  —  [Reg. No]"),
    ("BHUVANESWARI", "  —  [Reg. No]"),
    ("BENSAM GILBERT E", "  —  [Reg. No]"),
    ("JOEL EDWIN P", "  —  [Reg. No]"),
]

for name, reg_placeholder in teammates:
    insert_label_value_para_before(
        first_body_para,
        label=name,
        value=reg_placeholder,
        font_size=13,
        space_before=2,
        space_after=5
    )

# Add Page Break before the main report body starts
p_break = first_body_para.insert_paragraph_before()
p_break.add_run().add_break(docx.enum.text.WD_BREAK.PAGE)

# Also enrich the end of the document with the GitHub repository link and README production highlights
p_repo = doc.add_paragraph()
p_repo.paragraph_format.space_before = Pt(8)
r1 = p_repo.add_run("GitHub Repository Link: ")
r1.bold = True
r1.font.size = Pt(12)
r2 = p_repo.add_run("https://github.com/isaacincognito007-cloud/AI-StudyBuddy-")
r2.font.size = Pt(12)
r2.font.color.rgb = RGBColor(37, 99, 235)

doc.save(OUT_DOCX_D)
doc.save(OUT_DOCX_DL)
print("Saved DOCX to:", OUT_DOCX_D, "and", OUT_DOCX_DL)

# Convert DOCX to PDF via Microsoft Word COM automation
import win32com.client
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    wdoc = word.Documents.Open(OUT_DOCX_D)
    wdoc.SaveAs(OUT_PDF_D, FileFormat=17)  # 17 = wdFormatPDF
    wdoc.Close(False)
    print("Saved PDF to:", OUT_PDF_D)
finally:
    word.Quit()

import shutil
shutil.copyfile(OUT_PDF_D, OUT_PDF_DL)
print("Copied PDF to:", OUT_PDF_DL)
