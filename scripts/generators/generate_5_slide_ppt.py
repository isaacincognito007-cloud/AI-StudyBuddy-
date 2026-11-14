import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image

MEDIA_DIR = r"C:\Users\acer\.gemini\antigravity\brain\3df468b6-62f3-4271-a7b4-01a98becf01a\scratch\media"
OUT_PPTX_D = r"d:\AI-StudyBuddy\Study_Buddy_AI_Presentation.pptx"
OUT_PPTX_DL = r"c:\Users\acer\Downloads\Study_Buddy_AI_Presentation.pptx"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

# NEW THEME: "Cyber Emerald & Obsidian AI" (Modern Deep Forest / Emerald / Mint / Ice Teal)
OBSIDIAN_DEEP = RGBColor(11, 19, 27)    # #0B131B - Deepest Obsidian Slate
EMERALD_DARK  = RGBColor(6, 78, 59)     # #064E3B - Deep Rich Emerald
EMERALD_MED   = RGBColor(5, 150, 105)   # #059669 - Vibrant Emerald Green
MINT_NEON     = RGBColor(52, 211, 153)  # #34D399 - Glowing Mint
TEAL_CYAN     = RGBColor(20, 184, 166)  # #14B8A6 - Tech Teal
ICE_BLUE      = RGBColor(204, 251, 241) # #CCFBF1 - Soft Ice Mint
SLATE_DARK    = RGBColor(15, 23, 42)    # #0F172A - Deep Text
SLATE_MUTED   = RGBColor(71, 85, 105)   # #475569 - Muted Text
BG_PEARL      = RGBColor(246, 250, 248) # #F6FAF8 - Ultra-clean soft ambient background
CARD_WHITE    = RGBColor(255, 255, 255) # #FFFFFF - Crisp white cards
CARD_BORDER   = RGBColor(209, 231, 221) # #D1E7DD - Subtle emerald-tinted border
WHITE         = RGBColor(255, 255, 255)

def set_slide_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_header(slide, title_text, category_text="STUDY BUDDY AI  |  BACKEND ARCHITECTURE", slide_num=None):
    set_slide_bg(slide, BG_PEARL)
    
    # Top banner bar - Obsidian Dark
    banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.05))
    banner.fill.solid()
    banner.fill.fore_color.rgb = OBSIDIAN_DEEP
    banner.line.fill.background()

    # Glowing Mint & Emerald Accent stripe
    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.01), Inches(13.333), Inches(0.05))
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = MINT_NEON
    stripe.line.fill.background()

    # Category label
    tb_cat = slide.shapes.add_textbox(Inches(0.55), Inches(0.12), Inches(10), Inches(0.28))
    tf_cat = tb_cat.text_frame
    tf_cat.word_wrap = True
    p_cat = tf_cat.paragraphs[0]
    p_cat.text = category_text.upper()
    p_cat.font.name = "Calibri"
    p_cat.font.size = Pt(10)
    p_cat.font.bold = True
    p_cat.font.color.rgb = MINT_NEON

    # Main Slide Title
    tb_title = slide.shapes.add_textbox(Inches(0.55), Inches(0.36), Inches(11.2), Inches(0.55))
    tf_title = tb_title.text_frame
    tf_title.word_wrap = True
    p_title = tf_title.paragraphs[0]
    p_title.text = title_text
    p_title.font.name = "Calibri"
    p_title.font.size = Pt(22)
    p_title.font.bold = True
    p_title.font.color.rgb = WHITE

    # Footer
    if slide_num is not None:
        footer = slide.shapes.add_textbox(Inches(0.55), Inches(7.12), Inches(12.2), Inches(0.3))
        tf_f = footer.text_frame
        p_f = tf_f.paragraphs[0]
        p_f.text = f"Study Buddy AI (LearnMate)  •  AI-Augmented Backend Development                                                                                                  Slide {slide_num} of 5"
        p_f.font.name = "Calibri"
        p_f.font.size = Pt(9.5)
        p_f.font.color.rgb = SLATE_MUTED

def add_card(slide, left, top, width, height, title=None, bg_color=CARD_WHITE, border_color=CARD_BORDER, header_bg=None):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.adjustments[0] = 0.02
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    if border_color:
        card.line.color.rgb = border_color
        card.line.width = Pt(1.2)
    else:
        card.line.fill.background()

    content_top = top + 0.12
    if title:
        if header_bg:
            hbar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(0.46))
            hbar.adjustments[0] = 0.02
            hbar.fill.solid()
            hbar.fill.fore_color.rgb = header_bg
            hbar.line.fill.background()
            tb = slide.shapes.add_textbox(Inches(left + 0.18), Inches(top + 0.07), Inches(width - 0.36), Inches(0.35))
            p = tb.text_frame.paragraphs[0]
            p.text = title
            p.font.name = "Calibri"
            p.font.size = Pt(14)
            p.font.bold = True
            p.font.color.rgb = WHITE
            content_top = top + 0.54
        else:
            tb = slide.shapes.add_textbox(Inches(left + 0.18), Inches(top + 0.12), Inches(width - 0.36), Inches(0.38))
            p = tb.text_frame.paragraphs[0]
            p.text = title
            p.font.name = "Calibri"
            p.font.size = Pt(15)
            p.font.bold = True
            p.font.color.rgb = EMERALD_DARK
            content_top = top + 0.50
    return content_top

def add_bullets(slide, left, top, width, height, items, font_size=12.2, space_after=7, text_color=SLATE_DARK):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    for idx, item in enumerate(items):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.space_after = Pt(space_after)
        p.line_spacing = 1.18
        if isinstance(item, tuple):
            bold_prefix, rest = item
            r1 = p.add_run()
            r1.text = f"✦ {bold_prefix}: "
            r1.font.name = "Calibri"
            r1.font.size = Pt(font_size)
            r1.font.bold = True
            r1.font.color.rgb = EMERALD_DARK

            r2 = p.add_run()
            r2.text = rest
            r2.font.name = "Calibri"
            r2.font.size = Pt(font_size)
            r2.font.bold = False
            r2.font.color.rgb = text_color
        else:
            r = p.add_run()
            r.text = f"✦ {item}"
            r.font.name = "Calibri"
            r.font.size = Pt(font_size)
            r.font.color.rgb = text_color

def add_image_fitted(slide, img_name, box_left, box_top, box_w, box_h):
    path = os.path.join(MEDIA_DIR, img_name)
    if not os.path.exists(path):
        return
    with Image.open(path) as im:
        iw, ih = im.size
    aspect = iw / ih
    box_aspect = box_w / box_h
    if aspect > box_aspect:
        w = box_w
        h = box_w / aspect
    else:
        h = box_h
        w = box_h * aspect
    x = box_left + (box_w - w) / 2
    y = box_top + (box_h - h) / 2
    slide.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))

# ==============================================================================
# SLIDE 1: TITLE SLIDE (Cyber Emerald Theme)
# ==============================================================================
s1 = prs.slides.add_slide(blank_layout)
set_slide_bg(s1, OBSIDIAN_DEEP)

# Left decorative accent stripes
bar1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.28), Inches(7.5))
bar1.fill.solid()
bar1.fill.fore_color.rgb = EMERALD_MED
bar1.line.fill.background()

bar2 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.28), Inches(0), Inches(0.08), Inches(7.5))
bar2.fill.solid()
bar2.fill.fore_color.rgb = MINT_NEON
bar2.line.fill.background()

# Subject Badge Pill
badge = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), Inches(0.85), Inches(5.4), Inches(0.45))
badge.adjustments[0] = 0.5
badge.fill.solid()
badge.fill.fore_color.rgb = EMERALD_DARK
badge.line.color.rgb = MINT_NEON
p_b = badge.text_frame.paragraphs[0]
p_b.alignment = PP_ALIGN.CENTER
p_b.text = "AI-AUGMENTED BACKEND DEVELOPMENT - A&S"
p_b.font.name = "Calibri"
p_b.font.size = Pt(11)
p_b.font.bold = True
p_b.font.color.rgb = MINT_NEON

# Main Project Title
tb_main = s1.shapes.add_textbox(Inches(0.95), Inches(1.5), Inches(11.4), Inches(1.9))
tf_main = tb_main.text_frame
tf_main.word_wrap = True
p_m1 = tf_main.paragraphs[0]
p_m1.text = "Study Buddy AI"
p_m1.font.name = "Calibri"
p_m1.font.size = Pt(50)
p_m1.font.bold = True
p_m1.font.color.rgb = WHITE

p_m2 = tf_main.add_paragraph()
p_m2.space_before = Pt(4)
p_m2.text = "LearnMate — AI Assistant for Smarter Learning and Study Support"
p_m2.font.name = "Calibri"
p_m2.font.size = Pt(20)
p_m2.font.color.rgb = ICE_BLUE

p_m3 = tf_main.add_paragraph()
p_m3.space_before = Pt(8)
p_m3.text = "Node.js  •  Express.js  •  MongoDB & Mongoose  •  JWT & RBAC  •  Google Gemini 2.5 Flash API"
p_m3.font.name = "Calibri"
p_m3.font.size = Pt(13)
p_m3.font.bold = True
p_m3.font.color.rgb = MINT_NEON

# Team Information Card
team_card = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), Inches(3.8), Inches(11.4), Inches(3.05))
team_card.adjustments[0] = 0.03
team_card.fill.solid()
team_card.fill.fore_color.rgb = RGBColor(16, 29, 39)
team_card.line.color.rgb = EMERALD_MED
team_card.line.width = Pt(1.5)

# Team Leader Box (Left side of card)
tb_tl = s1.shapes.add_textbox(Inches(1.35), Inches(4.1), Inches(4.8), Inches(2.4))
tf_tl = tb_tl.text_frame
tf_tl.word_wrap = True
p_tl_h = tf_tl.paragraphs[0]
p_tl_h.text = "TEAM LEADER"
p_tl_h.font.name = "Calibri"
p_tl_h.font.size = Pt(12)
p_tl_h.font.bold = True
p_tl_h.font.color.rgb = MINT_NEON

p_tl_n = tf_tl.add_paragraph()
p_tl_n.space_before = Pt(10)
p_tl_n.text = "ILAZ AHAMED S K"
p_tl_n.font.name = "Calibri"
p_tl_n.font.size = Pt(22)
p_tl_n.font.bold = True
p_tl_n.font.color.rgb = WHITE

p_tl_sub = tf_tl.add_paragraph()
p_tl_sub.space_before = Pt(16)
p_tl_sub.text = "Role: Full Backend Architecture,\nGemini AI Integration & Security Systems"
p_tl_sub.font.name = "Calibri"
p_tl_sub.font.size = Pt(12)
p_tl_sub.font.color.rgb = ICE_BLUE

# Vertical divider
vdiv = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.4), Inches(4.1), Inches(0.02), Inches(2.45))
vdiv.fill.solid()
vdiv.fill.fore_color.rgb = EMERALD_MED
vdiv.line.fill.background()

# Teammates Box (Right side of card)
tb_tm = s1.shapes.add_textbox(Inches(6.75), Inches(4.1), Inches(5.2), Inches(2.4))
tf_tm = tb_tm.text_frame
tf_tm.word_wrap = True
p_tm_h = tf_tm.paragraphs[0]
p_tm_h.text = "TEAMMATES"
p_tm_h.font.name = "Calibri"
p_tm_h.font.size = Pt(12)
p_tm_h.font.bold = True
p_tm_h.font.color.rgb = MINT_NEON

new_teammates = [
    "BILL CLINTON J",
    "MANIKANDAN T",
    "DAVISH SHYAM S",
    "SAN MATHEW M"
]
for member in new_teammates:
    pm = tf_tm.add_paragraph()
    pm.space_before = Pt(6)
    pm.text = f"•  {member}"
    pm.font.name = "Calibri"
    pm.font.size = Pt(15)
    pm.font.bold = True
    pm.font.color.rgb = WHITE

# ==============================================================================
# SLIDE 2: PROBLEM STATEMENT & PROPOSED SOLUTION
# ==============================================================================
s2 = prs.slides.add_slide(blank_layout)
add_header(s2, "Scenario Case Study: Problem Statement & Proposed Solution", "CASE STUDY & SOLUTION", 2)

ct_prob = add_card(s2, 0.55, 1.3, 5.9, 5.65, "The Problem: Rahul's Academic Dilemma", header_bg=EMERALD_DARK)
add_bullets(s2, 0.75, ct_prob + 0.1, 5.5, 4.9, [
    ("Student Persona (Rahul)", "2nd-year engineering student preparing for university semester exams while studying new technical skills for internships."),
    ("Fragmented Materials", "Lecture notes, textbooks, PDFs, and handwritten notes are scattered, making revision disorganized."),
    ("Severe Time Drain", "Summarizing hundreds of pages manually and formulating revision questions is labor-intensive and repetitive."),
    ("No Adaptive Scheduling", "Lacks structured study plans tailored to available daily study hours and upcoming exam dates."),
    ("Self-Assessment Deficit", "Difficulty gauging conceptual mastery without automated, objective practice quizzes."),
    ("Productivity Loss", "Switching across disparate apps for note-taking, calendars, and test creation slows exam readiness.")
], font_size=12.2, space_after=8)

ct_sol = add_card(s2, 6.85, 1.3, 5.9, 5.65, "The Solution: Study Buddy AI Platform", header_bg=EMERALD_MED)
add_bullets(s2, 7.05, ct_sol + 0.1, 5.5, 4.9, [
    ("Centralized RESTful Backend", "Automates academic workflows through secure Express.js APIs connected to Google Gemini 2.5 Flash."),
    ("Material-Grounded AI", "AI generations are strictly anchored to the student's actual uploaded documents—no generic hallucinated templates."),
    ("High-Yield Summaries", "Generates concise, concept-dense revision summaries for accelerated learning."),
    ("Active Recall Flashcards", "Automatically extracts high-value Q&A pairs for spaced repetition learning."),
    ("Objective MCQ Quizzes", "Produces 4-option multiple-choice quizzes with verified answers to assert readiness."),
    ("Personalized Day-by-Day Plans", "Builds tailored preparation roadmaps based on user goals, daily hours, and exam timelines.")
], font_size=12.2, space_after=8)

# ==============================================================================
# SLIDE 3: SYSTEM ARCHITECTURE & TECHNOLOGY STACK
# ==============================================================================
s3 = prs.slides.add_slide(blank_layout)
add_header(s3, "System Architecture & Technology Stack", "SYSTEM DESIGN & TECH STACK", 3)

# Architecture Diagram on Left
c3_l = add_card(s3, 0.55, 1.3, 6.7, 5.65, "AI Study Buddy — System Architecture", header_bg=EMERALD_DARK)
add_image_fitted(s3, "image1.png", 0.7, 1.9, 6.4, 4.9)

# Technology & MVC on Right
c3_r = add_card(s3, 7.5, 1.3, 5.28, 5.65, "Technology Stack & MVC Pattern", header_bg=EMERALD_MED)
add_bullets(s3, 7.7, c3_r + 0.1, 4.88, 4.9, [
    ("Node.js & Express.js", "Runtime and web framework powering modular RESTful API routing, controllers, and error handling."),
    ("MongoDB & Mongoose ODM", "Document database storing users, uploaded study files, and generated AI learning resources."),
    ("Google Gemini 2.5 Flash API", "Integrated via @google/genai SDK for fast natural language generation and structured JSON outputs."),
    ("JWT Authentication & RBAC", "Stateless 7-day token security with strict separation between Student and Administrator roles."),
    ("Model-View-Controller (MVC)", "Clear separation between Mongoose Schemas (Model), JSON REST Responses (View), and Business Logic (Controller)."),
    ("Production Middleware Stack", "Helmet security headers, CORS, Morgan logger, Multer file upload handling, and 2-tier rate limiting.")
], font_size=12, space_after=8)

# ==============================================================================
# SLIDE 4: DATABASE DESIGN (ER DIAGRAM) & KEY FEATURES
# ==============================================================================
s4 = prs.slides.add_slide(blank_layout)
add_header(s4, "Entity-Relationship (ER) Data Model & Core Features", "DATABASE DESIGN & FEATURES", 4)

# ER Diagram on Left
c4_l = add_card(s4, 0.55, 1.3, 6.7, 5.65, "AI Study Buddy — ER Diagram (6 Entities)", header_bg=EMERALD_DARK)
add_image_fitted(s4, "image2.png", 0.7, 1.9, 6.4, 4.9)

# Key Features on Right
c4_r = add_card(s4, 7.5, 1.3, 5.28, 5.65, "Core Capabilities & Security", header_bg=EMERALD_MED)
add_bullets(s4, 7.7, c4_r + 0.1, 4.88, 4.9, [
    ("User & Auth Management", "User registration, password encryption with bcryptjs (10 rounds), and JWT login."),
    ("Study Material Management", "Upload notes (.txt, .md, .pdf), pagination, full-text search, subject filtering, and downloads."),
    ("AI Educational Synthesis", "On-demand Gemini 2.5 Flash generation for summaries, flashcards, MCQs, and custom study schedules."),
    ("Query-Level Data Isolation", "Every database query enforces { userId: req.user._id } to guarantee zero unauthorized cross-student data access."),
    ("Express 5 NoSQL Guard", "Custom sanitize.js middleware rejects operators ($ and .) to prevent injection vulnerabilities."),
    ("Admin Analytics & Oversight", "Privileged endpoints to monitor platform users, uploaded materials, and AI generation metrics.")
], font_size=12, space_after=8)

# ==============================================================================
# SLIDE 5: EXECUTION, API TESTING & CONCLUSION
# ==============================================================================
s5 = prs.slides.add_slide(blank_layout)
add_header(s5, "Project Execution, API Verification & Conclusion", "TESTING & CONCLUSION", 5)

# 3 Test Verification Cards
t1 = add_card(s5, 0.55, 1.25, 3.9, 4.5, "1. Auth & Material Upload", header_bg=EMERALD_DARK)
add_image_fitted(s5, "image14.png", 0.65, 1.82, 3.7, 1.9)
add_bullets(s5, 0.68, 3.85, 3.64, 1.8, [
    ("POST /api/auth/register", "Registers account, issues JWT token."),
    ("POST /api/material/upload", "Uploads study notes to MongoDB.")
], font_size=11, space_after=3)

t2 = add_card(s5, 4.7, 1.25, 3.9, 4.5, "2. AI Flashcards & Quiz", header_bg=EMERALD_MED)
add_image_fitted(s5, "image17.png", 4.8, 1.82, 3.7, 1.9)
add_bullets(s5, 4.83, 3.85, 3.64, 1.8, [
    ("POST /api/ai/flashcards", "Generates structured Q&A pairs."),
    ("POST /api/ai/quiz", "Generates 4-option MCQs & key.")
], font_size=11, space_after=3)

t3 = add_card(s5, 8.85, 1.25, 3.9, 4.5, "3. Personalized Study Plan", header_bg=EMERALD_DARK)
add_image_fitted(s5, "image19.png", 8.95, 1.82, 3.7, 1.9)
add_bullets(s5, 8.98, 3.85, 3.64, 1.8, [
    ("POST /api/ai/study-plan", "Creates tailored 7-day timetable."),
    ("Validated Endpoints", "Asserted via Postman test suite.")
], font_size=11, space_after=3)

# Bottom Conclusion & Links Banner
b_bot = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(5.9), Inches(12.2), Inches(1.1))
b_bot.adjustments[0] = 0.05
b_bot.fill.solid()
b_bot.fill.fore_color.rgb = OBSIDIAN_DEEP
b_bot.line.color.rgb = EMERALD_MED

tb_bot = s5.shapes.add_textbox(Inches(0.75), Inches(5.95), Inches(11.8), Inches(0.98))
tf_bot = tb_bot.text_frame
tf_bot.word_wrap = True
p_b1 = tf_bot.paragraphs[0]
p_b1.text = "PROJECT CONCLUSION & REPOSITORY LINKS"
p_b1.font.name = "Calibri"
p_b1.font.size = Pt(11)
p_b1.font.bold = True
p_b1.font.color.rgb = MINT_NEON

p_b2 = tf_bot.add_paragraph()
p_b2.space_before = Pt(2)
p_b2.text = "Study Buddy AI delivers a scalable, high-security backend transforming study notes into actionable learning tools via Google Gemini. Ready for seamless frontend UI integration."
p_b2.font.name = "Calibri"
p_b2.font.size = Pt(11)
p_b2.font.color.rgb = WHITE

p_b3 = tf_bot.add_paragraph()
p_b3.space_before = Pt(3)
p_b3.text = "GitHub: https://github.com/ILAZAHAMEDSK/AI-STUDYBUDDY   |   Demo Video: https://youtu.be/N5lNC06rn7w"
p_b3.font.name = "Calibri"
p_b3.font.size = Pt(11.5)
p_b3.font.bold = True
p_b3.font.color.rgb = ICE_BLUE

prs.save(OUT_PPTX_D)
prs.save(OUT_PPTX_DL)
print("Saved Cyber Emerald themed 5-slide PPTX to:", OUT_PPTX_D, "and", OUT_PPTX_DL)
