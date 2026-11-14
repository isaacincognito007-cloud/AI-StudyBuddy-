import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import win32com.client
import fitz # PyMuPDF
import shutil

MEDIA_DIR = r"C:\Users\acer\.gemini\antigravity\brain\3df468b6-62f3-4271-a7b4-01a98becf01a\scratch\media"
CROPPED_DIR = os.path.join(MEDIA_DIR, "cropped")

OUT_DIR_D = r"d:\AI-StudyBuddy"
os.makedirs(OUT_DIR_D, exist_ok=True)

OUT_PPTX_D = os.path.join(OUT_DIR_D, "Study_Buddy_AI_Presentation.pptx")
OUT_PPTX_ROOT = r"d:\Study_Buddy_AI_Presentation.pptx"
OUT_PPTX_DL = r"c:\Users\acer\Downloads\Study_Buddy_AI_Presentation.pptx"

OUT_PDF_DL = r"c:\Users\acer\Downloads\Study_Buddy_AI_Presentation.pdf"
SCRATCH_PDF = r"C:\Users\acer\.gemini\antigravity\brain\3df468b6-62f3-4271-a7b4-01a98becf01a\scratch\clean8_preview.pdf"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

# ==============================================================================
# MINIMALIST DESIGN THEME
# ==============================================================================
BG_PAGE       = RGBColor(248, 250, 252) # #F8FAFC - Soft light grey
BG_WHITE      = RGBColor(255, 255, 255) # #FFFFFF - Clean white
CHARCOAL_DARK = RGBColor(23, 32, 51)    # #172033 - Primary text
CHARCOAL_MID  = RGBColor(51, 65, 85)    # #334155 - Subtitle text
TEAL_PRIMARY  = RGBColor(15, 118, 110)  # #0F766E - Teal accent
TEAL_LIGHT    = RGBColor(240, 253, 250) # #F0FDFA - Subtle teal tint
TEAL_BORDER   = RGBColor(153, 246, 228) # #99F6E4 - Crisp teal line
GREY_MUTED    = RGBColor(100, 116, 139) # #64748B - Secondary / meta text
BORDER_LIGHT  = RGBColor(226, 232, 240) # #E2E8F0 - Subtle divider / border
RED_ALERT     = RGBColor(220, 38, 38)   # #DC2626 - Problem accent
RED_LIGHT     = RGBColor(254, 242, 242) # #FEF2F2 - Red tint
RED_BORDER    = RGBColor(254, 202, 202) # #FECACA - Red border
BLUE_ACCENT   = RGBColor(2, 132, 199)   # #0284C7 - Info blue
WHITE         = RGBColor(255, 255, 255)

FONT_HEADING = "Calibri"
FONT_BODY    = "Calibri"

def set_slide_bg(slide, color=BG_PAGE):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_header(slide, title_text, category_text="AI-AUGMENTED BACKEND DEVELOPMENT", subtitle_text=None, slide_num=None):
    set_slide_bg(slide, BG_PAGE)

    # Category Pill / Label at top left
    tb_cat = slide.shapes.add_textbox(Inches(0.8), Inches(0.38), Inches(9.5), Inches(0.28))
    tf_cat = tb_cat.text_frame
    tf_cat.word_wrap = True
    p_cat = tf_cat.paragraphs[0]
    p_cat.text = category_text.upper()
    p_cat.font.name = FONT_HEADING
    p_cat.font.size = Pt(10)
    p_cat.font.bold = True
    p_cat.font.color.rgb = TEAL_PRIMARY

    # Slide Number at top right
    if slide_num:
        tb_num = slide.shapes.add_textbox(Inches(10.5), Inches(0.38), Inches(2.0), Inches(0.28))
        tf_num = tb_num.text_frame
        p_num = tf_num.paragraphs[0]
        p_num.alignment = PP_ALIGN.RIGHT
        p_num.text = f"SLIDE {slide_num} OF 8"
        p_num.font.name = FONT_HEADING
        p_num.font.size = Pt(10)
        p_num.font.bold = True
        p_num.font.color.rgb = GREY_MUTED

    # Main Slide Title
    tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.62), Inches(11.7), Inches(0.58))
    tf_title = tb_title.text_frame
    tf_title.word_wrap = True
    p_title = tf_title.paragraphs[0]
    p_title.text = title_text
    p_title.font.name = FONT_HEADING
    p_title.font.size = Pt(25)
    p_title.font.bold = True
    p_title.font.color.rgb = CHARCOAL_DARK

    if subtitle_text:
        tb_sub = slide.shapes.add_textbox(Inches(0.8), Inches(1.22), Inches(11.7), Inches(0.35))
        tf_sub = tb_sub.text_frame
        p_sub = tf_sub.paragraphs[0]
        p_sub.text = subtitle_text
        p_sub.font.name = FONT_BODY
        p_sub.font.size = Pt(12.5)
        p_sub.font.color.rgb = GREY_MUTED

    # Subtle bottom header divider rule
    divider = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.62), Inches(11.733), Inches(0.015))
    divider.fill.solid()
    divider.fill.fore_color.rgb = BORDER_LIGHT
    divider.line.fill.background()

def add_clean_card(slide, left, top, width, height, bg_color=BG_WHITE, border_color=BORDER_LIGHT, border_width=1):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.adjustments[0] = 0.02
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    if border_color:
        card.line.color.rgb = border_color
        card.line.width = Pt(border_width)
    else:
        card.line.fill.background()
    return card

# ==============================================================================
# SLIDE 1: TITLE SLIDE (Points format, bold, no highlighting)
# ==============================================================================
s1 = prs.slides.add_slide(blank_layout)
set_slide_bg(s1, BG_WHITE)

# Category at top
tb_c1 = s1.shapes.add_textbox(Inches(0.85), Inches(0.70), Inches(8.0), Inches(0.35))
tf_c1 = tb_c1.text_frame
p_c1 = tf_c1.paragraphs[0]
p_c1.text = "AI-AUGMENTED BACKEND DEVELOPMENT  |  TN SKILLS PROJECT RECORD"
p_c1.font.name = FONT_HEADING
p_c1.font.size = Pt(11)
p_c1.font.bold = True
p_c1.font.color.rgb = TEAL_PRIMARY

# Project Title - Bold, Not highlighted, No background box
tb_t1 = s1.shapes.add_textbox(Inches(0.85), Inches(1.05), Inches(8.5), Inches(0.8))
tf_t1 = tb_t1.text_frame
tf_t1.word_wrap = True
p_t1 = tf_t1.paragraphs[0]
p_t1.text = "Study Buddy AI"
p_t1.font.name = FONT_HEADING
p_t1.font.size = Pt(44)
p_t1.font.bold = True
p_t1.font.color.rgb = CHARCOAL_DARK

# Subtitle
p_sub1 = tf_t1.add_paragraph()
p_sub1.space_before = Pt(2)
p_sub1.text = "LearnMate — AI Assistant for Smarter Learning and Study Support"
p_sub1.font.name = FONT_BODY
p_sub1.font.size = Pt(17)
p_sub1.font.bold = True
p_sub1.font.color.rgb = TEAL_PRIMARY

# Executive Project Overview - Formatted as Bold Bullet Points (NOT paragraph, NO highlight)
tb_points = s1.shapes.add_textbox(Inches(0.85), Inches(2.35), Inches(8.2), Inches(1.8))
tf_pts = tb_points.text_frame
tf_pts.word_wrap = True
tf_pts.margin_left = Inches(0.0)
tf_pts.margin_right = Inches(0.0)

overview_points = [
    ("Project Title", "Study Buddy AI (LearnMate) — Educational Backend Architecture"),
    ("Core Function", "Converts raw study files into summaries, flashcards, quizzes & revision timelines"),
    ("AI Integration", "Powered by Google Gemini 2.5 Flash API with deterministic JSON schema output"),
    ("Data Isolation", "RESTful Node.js + Express + MongoDB architecture with query-level user isolation")
]

for idx, (label, val) in enumerate(overview_points):
    p_pt = tf_pts.paragraphs[0] if idx == 0 else tf_pts.add_paragraph()
    if idx > 0:
        p_pt.space_before = Pt(5)
    p_pt.line_spacing = 1.18
    
    r_dot = p_pt.add_run()
    r_dot.text = "• "
    r_dot.font.name = FONT_BODY
    r_dot.font.size = Pt(12)
    r_dot.font.bold = True
    r_dot.font.color.rgb = TEAL_PRIMARY
    
    r_lbl = p_pt.add_run()
    r_lbl.text = f"{label}: "
    r_lbl.font.name = FONT_BODY
    r_lbl.font.size = Pt(12)
    r_lbl.font.bold = True
    r_lbl.font.color.rgb = CHARCOAL_DARK
    
    r_val = p_pt.add_run()
    r_val.text = val
    r_val.font.name = FONT_BODY
    r_val.font.size = Pt(12)
    r_val.font.color.rgb = CHARCOAL_MID

# Abstract Connected AI Node Graphic on the Right
nodes = [
    (10.8, 2.1, 0.55, TEAL_PRIMARY, True),
    (9.5, 1.4, 0.38, TEAL_PRIMARY, False),
    (12.0, 1.5, 0.34, GREY_MUTED, False),
    (9.3, 3.0, 0.40, GREY_MUTED, False),
    (12.1, 2.9, 0.42, TEAL_PRIMARY, False),
    (10.9, 3.8, 0.36, TEAL_PRIMARY, False),
    (12.4, 3.9, 0.30, GREY_MUTED, False),
]
connections = [
    (0, 1), (0, 2), (0, 3), (0, 4), (0, 5),
    (1, 3), (2, 4), (4, 6), (5, 6)
]
for i1, i2 in connections:
    x1, y1, s1_w, _, _ = nodes[i1]
    x2, y2, s2_w, _, _ = nodes[i2]
    conn = s1.shapes.add_connector(1, Inches(x1 + s1_w/2), Inches(y1 + s1_w/2), Inches(x2 + s2_w/2), Inches(y2 + s2_w/2))
    conn.line.color.rgb = BORDER_LIGHT
    conn.line.width = Pt(1.2)

for x, y, size, col, is_center in nodes:
    node_circle = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(size), Inches(size))
    node_circle.fill.solid()
    node_circle.fill.fore_color.rgb = col if is_center else BG_WHITE
    node_circle.line.color.rgb = col
    node_circle.line.width = Pt(2 if is_center else 1.5)

# Horizontal Divider Line
hdiv1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.85), Inches(4.35), Inches(11.6), Inches(0.015))
hdiv1.fill.solid()
hdiv1.fill.fore_color.rgb = BORDER_LIGHT
hdiv1.line.fill.background()

# Bottom Two Balanced Cards: Tech Stack (Left) & Team Information (Right)
add_clean_card(s1, 0.85, 4.60, 5.7, 2.35, bg_color=BG_PAGE, border_color=BORDER_LIGHT)
tb_ts = s1.shapes.add_textbox(Inches(1.10), Inches(4.75), Inches(5.2), Inches(2.05))
tf_ts = tb_ts.text_frame
tf_ts.word_wrap = True
p_tsh = tf_ts.paragraphs[0]
p_tsh.text = "CORE TECHNOLOGY STACK"
p_tsh.font.name = FONT_HEADING
p_tsh.font.size = Pt(11)
p_tsh.font.bold = True
p_tsh.font.color.rgb = TEAL_PRIMARY

p_tsb = tf_ts.add_paragraph()
p_tsb.space_before = Pt(6)
p_tsb.text = "Node.js  •  Express.js  •  MongoDB  •  Mongoose ODM\nJWT & RBAC Security  •  Google Gemini 2.5 Flash API"
p_tsb.font.name = FONT_BODY
p_tsb.font.size = Pt(13)
p_tsb.font.bold = True
p_tsb.font.color.rgb = CHARCOAL_DARK
p_tsb.line_spacing = 1.3

p_tss = tf_ts.add_paragraph()
p_tss.space_before = Pt(6)
p_tss.text = "Production RESTful architecture enforcing query-level data isolation."
p_tss.font.name = FONT_BODY
p_tss.font.size = Pt(11)
p_tss.font.color.rgb = GREY_MUTED

# Right: Team Information
add_clean_card(s1, 6.75, 4.60, 5.7, 2.35, bg_color=BG_PAGE, border_color=BORDER_LIGHT)
tb_tm1 = s1.shapes.add_textbox(Inches(7.00), Inches(4.75), Inches(5.2), Inches(2.05))
tf_tm1 = tb_tm1.text_frame
tf_tm1.word_wrap = True
p_tmh = tf_tm1.paragraphs[0]
p_tmh.text = "PROJECT TEAM (PATRICIAN COLLEGE / TN SKILLS)"
p_tmh.font.name = FONT_HEADING
p_tmh.font.size = Pt(11)
p_tmh.font.bold = True
p_tmh.font.color.rgb = TEAL_PRIMARY

p_tl = tf_tm1.add_paragraph()
p_tl.space_before = Pt(5)
r_tll = p_tl.add_run()
r_tll.text = "Team Leader: "
r_tll.font.name = FONT_BODY
r_tll.font.size = Pt(12)
r_tll.font.bold = True
r_tll.font.color.rgb = CHARCOAL_DARK
r_tln = p_tl.add_run()
r_tln.text = "Isaac Shiloah Herbert V (212404414)"
r_tln.font.name = FONT_BODY
r_tln.font.size = Pt(12)
r_tln.font.bold = True
r_tln.font.color.rgb = TEAL_PRIMARY

p_tmb = tf_tm1.add_paragraph()
p_tmb.space_before = Pt(4)
r_tml = p_tmb.add_run()
r_tml.text = "Team Members:\n"
r_tml.font.name = FONT_BODY
r_tml.font.size = Pt(11.5)
r_tml.font.bold = True
r_tml.font.color.rgb = CHARCOAL_DARK

r_tmn = p_tmb.add_run()
r_tmn.text = "Mohammed Irfan A (212404420)  •  Bhuvaneswari A (212404441)\nBensam Gilbert E (212404405)  •  Joel Edwin P (212404416)"
r_tmn.font.name = FONT_BODY
r_tmn.font.size = Pt(11)
r_tmn.font.color.rgb = CHARCOAL_MID
p_tmb.line_spacing = 1.25

# ==============================================================================
# SLIDE 2: PROBLEM STATEMENT & CASE STUDY
# ==============================================================================
s2 = prs.slides.add_slide(blank_layout)
add_header(s2, "Academic Bottlenecks & Practical Motivation", "PROBLEM STATEMENT & CASE STUDY", 
           "Analysis of student revision friction: scattered materials, manual overhead, and absence of active recall.", 2)

# Case Study Context Card (Top Horizontal)
add_clean_card(s2, 0.8, 1.85, 11.733, 1.05, bg_color=BG_WHITE, border_color=BORDER_LIGHT)
tb_cs = s2.shapes.add_textbox(Inches(1.0), Inches(1.92), Inches(11.3), Inches(0.9))
tf_cs = tb_cs.text_frame
tf_cs.word_wrap = True
p_csh = tf_cs.paragraphs[0]
r_ch1 = p_csh.add_run()
r_ch1.text = "Case Study Scenario (The Student Dilemma): "
r_ch1.font.name = FONT_HEADING
r_ch1.font.size = Pt(12)
r_ch1.font.bold = True
r_ch1.font.color.rgb = RED_ALERT

r_ch2 = p_csh.add_run()
r_ch2.text =  "Students facing tight semester examination schedules must digest voluminous lecture slides, textbooks, and notes. " "Traditional revision forces students to spend the majority of their study hours manually retyping notes, copying definitions, " "and fabricating practice questions—leaving minimal time for deep conceptual learning and diagnostic self-assessment." 
r_ch2.font.name = FONT_BODY
r_ch2.font.size = Pt(11)
r_ch2.font.color.rgb = CHARCOAL_DARK
p_csh.line_spacing = 1.2

# 3 Detailed Problem Cards (Left, Middle, Right)
prob_cards = [
    ("Scattered & Disorganized Notes", 
     "Operational Friction",
     "• Study resources are fragmented across physical notebooks, loose PDFs, LMS portals, and shared links.\n"
     "• No unified search or central repository to quickly find relevant topics before an exam.\n"
     "• Inconsistent formats prevent automated analysis and fast keyword retrieval."),
    ("Labor-Intensive Note Synthesis", 
     "Productivity Bottleneck",
     "• Manually summarizing 50+ page chapters takes hours of passive writing with low retention yield.\n"
     "• Extracting key concepts and formulas requires repetitive clerical effort rather than cognitive processing.\n"
     "• Drafting active-recall question/answer pairs is too time-consuming for full syllabus coverage."),
    ("Absence of Structured Workflows", 
     "Assessment & Timeline Deficit",
     "• Students suffer from the 'illusion of competence' due to passive reading without objective quizzes.\n"
     "• No adaptive revision timelines mapped to remaining days before examination dates.\n"
     "• Lack of on-demand contextual Q&A to resolve doubts immediately during late-night study sessions.")
]

for idx, (title, tag, body) in enumerate(prob_cards):
    cx = 0.8 + idx * 4.0
    add_clean_card(s2, cx, 3.10, 3.733, 3.85, bg_color=RED_LIGHT, border_color=RED_BORDER)
    
    tb_p = s2.shapes.add_textbox(Inches(cx + 0.2), Inches(3.25), Inches(3.333), Inches(3.5))
    tf_p = tb_p.text_frame
    tf_p.word_wrap = True
    
    p_tag = tf_p.paragraphs[0]
    p_tag.text = tag.upper()
    p_tag.font.name = FONT_HEADING
    p_tag.font.size = Pt(10)
    p_tag.font.bold = True
    p_tag.font.color.rgb = RED_ALERT
    
    p_t = tf_p.add_paragraph()
    p_t.space_before = Pt(3)
    p_t.text = title
    p_t.font.name = FONT_HEADING
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = CHARCOAL_DARK
    
    p_b = tf_p.add_paragraph()
    p_b.space_before = Pt(8)
    p_b.line_spacing = 1.25
    p_b.text = body
    p_b.font.name = FONT_BODY
    p_b.font.size = Pt(10.5)
    p_b.font.color.rgb = CHARCOAL_MID

# ==============================================================================
# SLIDE 3: PROPOSED SOLUTION & CORE VALUE PROPOSITION
# ==============================================================================
s3 = prs.slides.add_slide(blank_layout)
add_header(s3, "Smarter Study Preparation, One Platform", "PROPOSED SOLUTION & PLATFORM VALUE", 
           "Transforming raw study files into structured, persistent revision assets via Generative AI.", 3)

sol_cards = [
    ("1. Centralized Material Repository", 
     "STUDY MATERIAL MANAGEMENT",
     "• Ingests raw notes and documents (.txt, .md, .pdf) through secure multipart form-data upload.\n"
     "• Enforces strict file validation (MIME-type checks and size limits via Multer middleware).\n"
     "• Associates every file with an authenticated user ID in MongoDB for instant anytime retrieval.",
     TEAL_LIGHT, TEAL_BORDER),
    ("2. AI-Driven Active Recall Engine", 
     "GENERATIVE LEARNING UTILITIES",
     "• Dynamic Summaries: Synthesizes high-yield conceptual takeaways from uploaded notes.\n"
     "• Active-Recall Flashcards: Generates structured Q&A pairs via guaranteed JSON schema for spaced repetition.\n"
     "• 4-Option MCQ Quizzes: Constructs objective practice tests with verified answer keys to gauge readiness.",
     BG_WHITE, BORDER_LIGHT),
    ("3. Adaptive Timelines & Direct Q&A", 
     "PERSONALIZED REVISION SUPPORT",
     "• Tailored Study Plans: Builds realistic day-by-day revision schedules matched to exam dates.\n"
     "• Contextual Q&A: Students query uploaded materials directly; AI extracts answers strictly from context.\n"
     "• Eliminates hallucination risk by grounding model prompts in actual student materials.",
     BG_WHITE, BORDER_LIGHT),
    ("4. Persistent Asset Cataloging", 
     "LONG-TERM REVISION RETENTION",
     "• Every generated summary, flashcard set, and quiz is automatically persisted to MongoDB.\n"
     "• Students can re-access, filter, and review all study assets across subsequent revision sessions.\n"
     "• Eliminates token waste by ensuring AI outputs are permanently preserved rather than ephemeral.",
     TEAL_LIGHT, TEAL_BORDER)
]

for idx, (title, tag, body, bg_col, border_col) in enumerate(sol_cards):
    row = idx // 2
    col = idx % 2
    cx = 0.8 + col * 5.95
    cy = 1.85 + row * 2.60
    
    add_clean_card(s3, cx, cy, 5.75, 2.40, bg_color=bg_col, border_color=border_col)
    
    tb_s = s3.shapes.add_textbox(Inches(cx + 0.2), Inches(cy + 0.15), Inches(5.35), Inches(2.10))
    tf_s = tb_s.text_frame
    tf_s.word_wrap = True
    
    p_tag = tf_s.paragraphs[0]
    p_tag.text = tag
    p_tag.font.name = FONT_HEADING
    p_tag.font.size = Pt(10)
    p_tag.font.bold = True
    p_tag.font.color.rgb = TEAL_PRIMARY
    
    p_t = tf_s.add_paragraph()
    p_t.space_before = Pt(2)
    p_t.text = title
    p_t.font.name = FONT_HEADING
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = CHARCOAL_DARK
    
    p_b = tf_s.add_paragraph()
    p_b.space_before = Pt(5)
    p_b.line_spacing = 1.2
    p_b.text = body
    p_b.font.name = FONT_BODY
    p_b.font.size = Pt(10)
    p_b.font.color.rgb = CHARCOAL_MID

# ==============================================================================
# SLIDE 4: SYSTEM ARCHITECTURE & DATA FLOW
# ==============================================================================
s4 = prs.slides.add_slide(blank_layout)
add_header(s4, "End-to-End System Architecture", "SYSTEM ARCHITECTURE & DATA FLOW", 
           "Multi-tier MVC request lifecycle, defense-in-depth security boundary, and dual backend services.", 4)

def draw_arch_card(slide, x, y, w, h, title, subtitle, bullets, is_teal=False, is_accent=False):
    bg = TEAL_LIGHT if is_teal else (RGBColor(241, 245, 249) if is_accent else BG_WHITE)
    border = TEAL_PRIMARY if is_teal else (TEAL_BORDER if is_accent else BORDER_LIGHT)
    add_clean_card(slide, x, y, w, h, bg_color=bg, border_color=border, border_width=1.5 if is_teal else 1)
    
    tb = slide.shapes.add_textbox(Inches(x + 0.12), Inches(y + 0.08), Inches(w - 0.24), Inches(h - 0.16))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.04)
    tf.margin_bottom = Inches(0.04)
    
    p0 = tf.paragraphs[0]
    p0.text = title
    p0.font.name = FONT_HEADING
    p0.font.size = Pt(12)
    p0.font.bold = True
    p0.font.color.rgb = TEAL_PRIMARY if is_teal else CHARCOAL_DARK
    
    p1 = tf.add_paragraph()
    p1.space_before = Pt(2)
    p1.text = subtitle
    p1.font.name = FONT_BODY
    p1.font.size = Pt(10)
    p1.font.bold = True
    p1.font.color.rgb = GREY_MUTED
    
    for b in bullets:
        pb = tf.add_paragraph()
        pb.space_before = Pt(2.5)
        pb.line_spacing = 1.12
        pb.text = f"• {b}"
        pb.font.name = FONT_BODY
        pb.font.size = Pt(9.5)
        pb.font.color.rgb = CHARCOAL_MID

# Box 1: Client Layer
draw_arch_card(s4, 0.8, 1.95, 2.3, 3.65, "Client Interface", "Web & Postman Clients",
               ["Student & Admin clients", "HTTPS requests + Bearer JWT", "Uploads text & PDF materials", "Triggers AI generation", "Views saved revision assets"])

# Arrow 1: Client -> Gateway
a1 = s4.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(3.20), Inches(3.65), Inches(0.35), Inches(0.25))
a1.fill.solid()
a1.fill.fore_color.rgb = TEAL_PRIMARY
a1.line.fill.background()

# Box 2: API Gateway
draw_arch_card(s4, 3.65, 1.95, 2.5, 3.65, "REST API Gateway", "Express.js Router Layer",
               ["Centralized endpoint routing", "Helmet HTTP security headers", "CORS & Morgan request logger", "Two-tier rate limiting", "Multer multipart file parser", "Express 5 NoSQL-injection guard"])

# Arrow 2: Gateway -> Controllers
a2 = s4.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(6.25), Inches(3.65), Inches(0.35), Inches(0.25))
a2.fill.solid()
a2.fill.fore_color.rgb = TEAL_PRIMARY
a2.line.fill.background()

# Box 3: Auth & Controllers
draw_arch_card(s4, 6.7, 1.95, 2.6, 3.65, "Auth & Controllers", "Middleware & Business Logic",
               ["Verifies 7-day JWT Bearer tokens", "Role check (Student / Admin)", "Scoped queries { userId }", "Input validator schema rules", "Global error-handling middleware", "Dispatches AI & DB services"])

# Arrow 3A: Controllers -> Gemini
a3a = s4.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(9.40), Inches(2.70), Inches(0.4), Inches(0.22))
a3a.fill.solid()
a3a.fill.fore_color.rgb = TEAL_PRIMARY
a3a.line.fill.background()

# Box 4A: Gemini Service
draw_arch_card(s4, 9.9, 1.95, 2.65, 1.72, "Google Gemini 2.5 Flash", "@google/genai Service",
               ["Material-grounded prompts", "Summaries & Study Plans (prose)", "Flashcards & Quizzes (JSON schema)"], is_teal=True)

# Arrow 3B: Controllers -> MongoDB
a3b = s4.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(9.40), Inches(4.65), Inches(0.4), Inches(0.22))
a3b.fill.solid()
a3b.fill.fore_color.rgb = TEAL_PRIMARY
a3b.line.fill.background()

# Box 4B: MongoDB Persistence
draw_arch_card(s4, 9.9, 3.88, 2.65, 1.72, "MongoDB Atlas & Mongoose", "Data Persistence Layer",
               ["User accounts & password hashes", "Uploaded notes & metadata", "Saves summaries, cards & quizzes", "Enables retrieval for later revision"], is_accent=True)

# Bottom Horizontal Persistent Retrieval Banner
bot_bar = add_clean_card(s4, 0.8, 5.82, 11.75, 1.10, bg_color=BG_WHITE, border_color=BORDER_LIGHT)
tb_bb = s4.shapes.add_textbox(Inches(1.0), Inches(5.86), Inches(11.35), Inches(0.98))
tf_bb = tb_bb.text_frame
tf_bb.word_wrap = True
tf_bb.margin_left = Inches(0.06)
tf_bb.margin_right = Inches(0.06)
tf_bb.margin_top = Inches(0.04)
tf_bb.margin_bottom = Inches(0.04)
p_bb = tf_bb.paragraphs[0]
r_bbh = p_bb.add_run()
r_bbh.text = "Persistent Resource Retrieval: "
r_bbh.font.name = FONT_HEADING
r_bbh.font.size = Pt(11.5)
r_bbh.font.bold = True
r_bbh.font.color.rgb = TEAL_PRIMARY

r_bbb = p_bb.add_run()
r_bbb.text =  "Authorized students can re-access, search, filter, and review all their generated resources at any time. " "When an AI feature executes, its structured response is persisted directly to MongoDB associated with the user's account " "({ userId: req.user._id })—guaranteeing that study assets are preserved for exam revision rather than lost after generation." 
r_bbb.font.name = FONT_BODY
r_bbb.font.size = Pt(10.5)
r_bbb.font.color.rgb = CHARCOAL_MID
p_bb.line_spacing = 1.25

# ==============================================================================
# SLIDE 5: TECHNICAL STACK & ENGINEERING PROGRAMS
# ==============================================================================
s5 = prs.slides.add_slide(blank_layout)
add_header(s5, "Technology Stack & Engineering Programs", "TECH STACK & PROGRAMS USED", 
           "Exact packages, runtime environments, AI SDKs, and developer tools powering the backend.", 5)

tech_quadrants = [
    ("Runtime & Backend Framework", "CORE COMPUTING LAYER", [
        ("Node.js (v18+ LTS)", "High-performance asynchronous JavaScript runtime engine with non-blocking I/O event loop."),
        ("Express.js (v5.0)", "Fast, unopinionated web framework structuring RESTful routes, middleware chaining, and HTTP controllers."),
        ("Mongoose ODM (v8.0)", "Object Data Modeling library providing strict schema validation, type casting, and MongoDB query construction.")
    ]),
    ("Generative AI & Inference", "AI COGNITIVE ENGINE", [
        ("@google/genai SDK", "Official Google client library providing direct, low-latency integration with Gemini 2.5 Flash."),
        ("Gemini 2.5 Flash", "Advanced multimodal LLM delivering rapid token inference, strong context comprehension, and cost efficiency."),
        ("Structured JSON Mode", "Deterministic schema enforcement (`responseSchema`) guaranteeing valid JSON arrays for flashcards and quizzes.")
    ]),
    ("Security & Defense Utilities", "DATA PROTECTION & GOVERNANCE", [
        ("bcryptjs (10 rounds)", "Adaptive cryptographic password hashing with 10 salt rounds to eliminate plaintext exposure."),
        ("jsonwebtoken (JWT)", "Stateless authentication issuing 7-day HMAC-SHA256 Bearer tokens carrying decoded user ID and role."),
        ("helmet & express-rate-limit", "Protective HTTP headers mitigation (XSS, clickjacking) plus two-tier rate limiting against API abuse.")
    ]),
    ("File Handling & Developer Tooling", "FILE INGESTION & VALIDATION", [
        ("Multer Middleware", "Multipart/form-data handler managing incoming file uploads (.txt, .md, .pdf) with disk & memory buffers."),
        ("dotenv & cors", "Zero-dependency environment variable isolation and Cross-Origin Resource Sharing governance."),
        ("Postman Test Suites", "Automated API assertion framework verifying HTTP status codes, payload structures, and token security.")
    ])
]

for idx, (title, category, items) in enumerate(tech_quadrants):
    col = idx % 2
    row = idx // 2
    cx = 0.8 + col * 5.95
    cy = 1.85 + row * 2.60
    
    add_clean_card(s5, cx, cy, 5.75, 2.40, bg_color=BG_WHITE, border_color=BORDER_LIGHT)
    
    tb_t = s5.shapes.add_textbox(Inches(cx + 0.2), Inches(cy + 0.15), Inches(5.35), Inches(2.10))
    tf_t = tb_t.text_frame
    tf_t.word_wrap = True
    
    p_cat = tf_t.paragraphs[0]
    p_cat.text = category
    p_cat.font.name = FONT_HEADING
    p_cat.font.size = Pt(10)
    p_cat.font.bold = True
    p_cat.font.color.rgb = TEAL_PRIMARY
    
    p_t = tf_t.add_paragraph()
    p_t.space_before = Pt(2)
    p_t.text = title
    p_t.font.name = FONT_HEADING
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = CHARCOAL_DARK
    
    for item_name, item_desc in items:
        p_i = tf_t.add_paragraph()
        p_i.space_before = Pt(4)
        p_i.line_spacing = 1.15
        
        r_in = p_i.add_run()
        r_in.text = f"• {item_name}: "
        r_in.font.name = FONT_BODY
        r_in.font.size = Pt(10)
        r_in.font.bold = True
        r_in.font.color.rgb = CHARCOAL_DARK
        
        r_id = p_i.add_run()
        r_id.text = item_desc
        r_id.font.name = FONT_BODY
        r_id.font.size = Pt(9.5)
        r_id.font.color.rgb = CHARCOAL_MID

# ==============================================================================
# SLIDE 6: DATABASE DESIGN & DATA PERSISTENCE
# ==============================================================================
s6 = prs.slides.add_slide(blank_layout)
add_header(s6, "Database Design & Data Persistence", "DATABASE CONFIGURATION & SCHEMAS", 
           "Relational modeling in MongoDB NoSQL via Mongoose schemas, indexes, and document references.", 6)

schemas_2x2 = [
    ("User Schema (models/User.js)", "AUTHENTICATION & ROLES", [
        ("Fields", "name (String), email (String, unique, indexed), password (bcrypt-hashed), role (enum: ['student', 'admin']), createdAt (Date)."),
        ("Pre-save Hook", "Triggers bcrypt.hash(password, 10) before persisting modified credentials to the database."),
        ("Instance Methods", "matchPassword(enteredPassword) performs timing-safe cryptographic comparisons for login validation.")
    ], TEAL_LIGHT, TEAL_BORDER),
    
    ("Study Material Schema (models/Material.js)", "LEARNING ASSETS REPOSITORY", [
        ("Fields", "userId (ObjectId ref: 'User'), title (String), content (String/Text), fileType (String), fileSize (Number), uploadedAt (Date)."),
        ("Relationships", "Primary input anchor for all AI features; linked via 1-to-Many parent relationship with AI assets."),
        ("Indexing", "Compound index on { userId: 1, uploadedAt: -1 } for fast chronologically sorted user queries.")
    ], BG_WHITE, BORDER_LIGHT),
    
    ("Active-Recall Learning Schemas", "FLASHCARDS & QUIZZES", [
        ("Flashcard Schema", "userId, materialId, cards: [{ question: String, answer: String }], cardCount (Number)."),
        ("Quiz Schema", "userId, materialId, questions: [{ question: String, options: [String], correctAnswer: Number, explanation: String }]."),
        ("JSON Validation", "Enforces strict JSON schema guarantees (`responseSchema`) for deterministic array parsing.")
    ], BG_WHITE, BORDER_LIGHT),
    
    ("Synthesis & Planning Schemas", "SUMMARIES & REVISION TIMELINES", [
        ("Summary Schema", "userId (ref User), materialId (ref Material), summaryText (String), keyTakeaways ([String]), generatedAt."),
        ("Study Plan Schema", "userId, targetExam: String, examDate: Date, dailyPlan: [{ day: Number, topics: [String], hours: Number }]."),
        ("Scoped Retrieval", "Linked to authenticated student account; queried on demand for anytime exam review.")
    ], TEAL_LIGHT, TEAL_BORDER)
]

for idx, (title, category, items, bg_col, border_col) in enumerate(schemas_2x2):
    col = idx % 2
    row = idx // 2
    cx = 0.8 + col * 5.95
    cy = 1.85 + row * 2.60
    
    add_clean_card(s6, cx, cy, 5.75, 2.40, bg_color=bg_col, border_color=border_col)
    
    tb_sc = s6.shapes.add_textbox(Inches(cx + 0.2), Inches(cy + 0.15), Inches(5.35), Inches(2.10))
    tf_sc = tb_sc.text_frame
    tf_sc.word_wrap = True
    
    p_c = tf_sc.paragraphs[0]
    p_c.text = category
    p_c.font.name = FONT_HEADING
    p_c.font.size = Pt(10)
    p_c.font.bold = True
    p_c.font.color.rgb = TEAL_PRIMARY
    
    p_tit = tf_sc.add_paragraph()
    p_tit.space_before = Pt(2)
    p_tit.text = title
    p_tit.font.name = FONT_HEADING
    p_tit.font.size = Pt(13)
    p_tit.font.bold = True
    p_tit.font.color.rgb = CHARCOAL_DARK
    
    for iname, idesc in items:
        p_item = tf_sc.add_paragraph()
        p_item.space_before = Pt(4)
        p_item.line_spacing = 1.15
        
        r_in = p_item.add_run()
        r_in.text = f"• {iname}: "
        r_in.font.name = FONT_BODY
        r_in.font.size = Pt(10)
        r_in.font.bold = True
        r_in.font.color.rgb = CHARCOAL_DARK
        
        r_id = p_item.add_run()
        r_id.text = idesc
        r_id.font.name = FONT_BODY
        r_id.font.size = Pt(9.5)
        r_id.font.color.rgb = CHARCOAL_MID


# ==============================================================================
# SLIDE 7: DEFENSE-IN-DEPTH SECURITY ARCHITECTURE
# ==============================================================================
s7 = prs.slides.add_slide(blank_layout)
add_header(s7, "Security Engineering & Access Governance", "DEFENSE-IN-DEPTH SECURITY", 
           "Multi-tier security architecture enforcing identity protection, privilege isolation, and abuse prevention.", 7)

sec_pillars = [
    ("Bcrypt Password Hashing", "CRYPTOGRAPHIC STORAGE",
     "• Passwords salted with 10 rounds using bcryptjs before database commit.\n"
     "• Plaintext credentials never stored or exposed in application logs.\n"
     "• Resistant to precomputed rainbow-table and brute-force attacks.",
     BG_WHITE),
    ("Stateless JWT Authentication", "IDENTITY VERIFICATION",
     "• 7-day HMAC-SHA256 signed JSON Web Tokens issued upon successful login.\n"
     "• Bearer tokens extracted and cryptographically verified by authMiddleware.\n"
     "• Decoded user object attached to req.user for downstream handler access.",
     TEAL_LIGHT),
    ("Role-Based Access Control (RBAC)", "PRIVILEGE SEGREGATION",
     "• Strict separation of Student and Administrator roles.\n"
     "• Role validation middleware blocks privilege escalation.\n"
     "• Admin endpoints guarded: user management and usage analytics.",
     BG_WHITE),
    ("User-Scoped Query Isolation", "HORIZONTAL PRIVACY",
     "• Every database query strictly enforces { userId: req.user._id }.\n"
     "• Prevents insecure direct object reference (IDOR) vulnerabilities.\n"
     "• Guarantees students cannot view, modify, or delete peer materials.",
     BG_WHITE),
    ("Input Guard & Rate Limiting", "ABUSE PREVENTION & HYGIENE",
     "• Custom Express 5 NoSQL-injection guard sanitizes request payloads.\n"
     "• express-validator enforces strict field types, lengths, and schemas.\n"
     "• Two-tier rate limiting: 100 req/15min general + 15 req/15min on AI quotas.",
     TEAL_LIGHT),
    ("Audit Logs & Error Masking", "TELEMETRY & RESILIENCE",
     "• Morgan logger records HTTP request methods, status codes, and latency.\n"
     "• Centralized error-handling middleware normalizes response payloads.\n"
     "• Internal database error stacks masked in production to prevent leaks.",
     BG_WHITE)
]

for idx, (title, category, body, bg_col) in enumerate(sec_pillars):
    col = idx % 3
    row = idx // 3
    cx = 0.8 + col * 3.95
    cy = 1.85 + row * 2.60
    
    border_col = TEAL_BORDER if bg_col == TEAL_LIGHT else BORDER_LIGHT
    add_clean_card(s7, cx, cy, 3.80, 2.40, bg_color=bg_col, border_color=border_col)
    
    tb_sec = s7.shapes.add_textbox(Inches(cx + 0.15), Inches(cy + 0.12), Inches(3.50), Inches(2.15))
    tf_sec = tb_sec.text_frame
    tf_sec.word_wrap = True
    
    p_cat = tf_sec.paragraphs[0]
    p_cat.text = category
    p_cat.font.name = FONT_HEADING
    p_cat.font.size = Pt(9.5)
    p_cat.font.bold = True
    p_cat.font.color.rgb = TEAL_PRIMARY
    
    p_t = tf_sec.add_paragraph()
    p_t.space_before = Pt(2)
    p_t.text = title
    p_t.font.name = FONT_HEADING
    p_t.font.size = Pt(12)
    p_t.font.bold = True
    p_t.font.color.rgb = CHARCOAL_DARK
    
    p_b = tf_sec.add_paragraph()
    p_b.space_before = Pt(5)
    p_b.line_spacing = 1.2
    p_b.text = body
    p_b.font.name = FONT_BODY
    p_b.font.size = Pt(9.5)
    p_b.font.color.rgb = CHARCOAL_MID

# ==============================================================================
# SLIDE 8: IMPLEMENTATION, POSTMAN VALIDATION & DELIVERABLES
# ==============================================================================
s8 = prs.slides.add_slide(blank_layout)
add_header(s8, "Implementation, API Validation & Deliverables", "IMPLEMENTATION, VALIDATION & LINKS", 
           "Verified endpoint testing in Postman, live demonstration evidence, and project repositories.", 8)

shots = [
    ("User Authentication (201 Created)", "postman_auth.png", "Bcrypt password hashing & 7-day JWT Bearer token issuance."),
    ("Material Upload (201 Created)", "postman_upload.png", "Multipart form-data upload (.txt) securely parsed and saved to MongoDB."),
    ("AI Flashcards (200 OK)", "postman_flashcards.png", "Google Gemini 2.5 Flash generating structured active-recall Q&A arrays.")
]

card_w = 3.65
card_h = 3.30
y_top = 1.85

for idx, (caption, filename, note) in enumerate(shots):
    x_pos = 0.8 + idx * 4.05
    add_clean_card(s8, x_pos, y_top, card_w, card_h, bg_color=BG_WHITE, border_color=BORDER_LIGHT)
    
    # Title
    tb_cap = s8.shapes.add_textbox(Inches(x_pos + 0.1), Inches(y_top + 0.08), Inches(card_w - 0.2), Inches(0.35))
    p_c = tb_cap.text_frame.paragraphs[0]
    p_c.text = caption
    p_c.font.name = FONT_HEADING
    p_c.font.size = Pt(11)
    p_c.font.bold = True
    p_c.font.color.rgb = TEAL_PRIMARY
    
    # Cropped Image
    img_path = os.path.join(CROPPED_DIR, filename)
    if os.path.exists(img_path):
        s8.shapes.add_picture(img_path, Inches(x_pos + 0.12), Inches(y_top + 0.42), Inches(card_w - 0.24), Inches(2.15))
        
    # Note / Annotation below image
    tb_n = s8.shapes.add_textbox(Inches(x_pos + 0.1), Inches(y_top + 2.62), Inches(card_w - 0.2), Inches(0.6))
    tf_n = tb_n.text_frame
    tf_n.word_wrap = True
    p_n = tf_n.paragraphs[0]
    p_n.text = note
    p_n.font.name = FONT_BODY
    p_n.font.size = Pt(10)
    p_n.font.color.rgb = GREY_MUTED
    p_n.line_spacing = 1.15

# Testing Statement Box
stmt_box = add_clean_card(s8, 0.8, 5.30, 11.75, 0.58, bg_color=BG_PAGE, border_color=BORDER_LIGHT)
tb_stmt = s8.shapes.add_textbox(Inches(0.95), Inches(5.32), Inches(11.45), Inches(0.52))
tf_stmt = tb_stmt.text_frame
tf_stmt.word_wrap = True
p_stmt = tf_stmt.paragraphs[0]
r_st1 = p_stmt.add_run()
r_st1.text = "API Verification: "
r_st1.font.name = FONT_HEADING
r_st1.font.size = Pt(11)
r_st1.font.bold = True
r_st1.font.color.rgb = TEAL_PRIMARY

r_st2 = p_stmt.add_run()
r_st2.text =  "The backend RESTful API was comprehensively tested using Postman, verifying route protection, role authorization, " "multipart file uploads, and Gemini AI JSON schema validation across all endpoints." 
r_st2.font.name = FONT_BODY
r_st2.font.size = Pt(10.5)
r_st2.font.color.rgb = CHARCOAL_DARK

# Project Links & Closing Statement Banner
bot_box = add_clean_card(s8, 0.8, 6.02, 11.75, 1.10, bg_color=TEAL_LIGHT, border_color=TEAL_BORDER)

tb_links = s8.shapes.add_textbox(Inches(1.0), Inches(6.06), Inches(11.35), Inches(1.00))
tf_l = tb_links.text_frame
tf_l.word_wrap = True

# Links row
p_l1 = tf_l.paragraphs[0]
r_gh_lbl = p_l1.add_run()
r_gh_lbl.text = "GitHub Repository: "
r_gh_lbl.font.name = FONT_HEADING
r_gh_lbl.font.size = Pt(11)
r_gh_lbl.font.bold = True
r_gh_lbl.font.color.rgb = TEAL_PRIMARY

r_gh_val = p_l1.add_run()
r_gh_val.text = "https://github.com/isaacincognito007-cloud/AI-StudyBuddy-    "
r_gh_val.font.name = FONT_BODY
r_gh_val.font.size = Pt(11)
r_gh_val.font.bold = True
r_gh_val.font.color.rgb = CHARCOAL_DARK

r_yt_lbl = p_l1.add_run()
r_yt_lbl.text = "|   Demo Video: "
r_yt_lbl.font.name = FONT_HEADING
r_yt_lbl.font.size = Pt(11)
r_yt_lbl.font.bold = True
r_yt_lbl.font.color.rgb = TEAL_PRIMARY

r_yt_val = p_l1.add_run()
r_yt_val.text = "https://youtu.be/N5lNC06rn7w"
r_yt_val.font.name = FONT_BODY
r_yt_val.font.size = Pt(11)
r_yt_val.font.bold = True
r_yt_val.font.color.rgb = CHARCOAL_DARK

# Closing statement
p_l2 = tf_l.add_paragraph()
p_l2.space_before = Pt(4)
r_close = p_l2.add_run()
r_close.text = "“Study Buddy AI — Turning Study Materials into Structured Learning Resources.”"
r_close.font.name = FONT_HEADING
r_close.font.size = Pt(12)
r_close.font.bold = True
r_close.font.color.rgb = TEAL_PRIMARY

prs.save(OUT_PPTX_D)
print("Saved PPTX to:", OUT_PPTX_D)

shutil.copyfile(OUT_PPTX_D, OUT_PPTX_ROOT)
print("Saved PPTX to:", OUT_PPTX_ROOT)

shutil.copyfile(OUT_PPTX_D, OUT_PPTX_DL)
print("Saved PPTX to:", OUT_PPTX_DL)

# Convert to PDF
powerpoint = win32com.client.Dispatch("PowerPoint.Application")
deck = powerpoint.Presentations.Open(OUT_PPTX_D, WithWindow=False)
deck.SaveAs(SCRATCH_PDF, 32)
deck.SaveAs(OUT_PDF_DL, 32)
deck.Close()
powerpoint.Quit()
print("Converted to PDF:", SCRATCH_PDF)

# Render all 8 slides to PNG
doc = fitz.open(SCRATCH_PDF)
print(f"Total slides in PDF: {len(doc)}")
for i, page in enumerate(doc):
    pix = page.get_pixmap(dpi=150)
    out_img = rf"C:\Users\acer\.gemini\antigravity\brain\3df468b6-62f3-4271-a7b4-01a98becf01a\scratch\clean8_slide_{i+1}.png"
    pix.save(out_img)
    print(f"Saved slide {i+1} to {out_img}")
doc.close()
