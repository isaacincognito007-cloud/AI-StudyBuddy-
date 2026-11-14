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

# NEW THEME: "Electric Violet & Cosmic Indigo AI"
SPACE_DARK    = RGBColor(15, 12, 33)    # #0F0C21 - Deepest Cosmic Midnight
VIOLET_DEEP   = RGBColor(76, 29, 149)   # #4C1D95 - Rich Royal Violet
VIOLET_MED    = RGBColor(124, 58, 237)  # #7C3AED - Electric Violet
INDIGO_VIBRANT= RGBColor(99, 102, 241)  # #6366F1 - Electric Indigo
CYAN_GLOW     = RGBColor(56, 189, 248)  # #38BDF8 - Glowing Cyan
LAVENDER_ICE  = RGBColor(237, 233, 254) # #EDE9FE - Soft Lavender Ice
SLATE_DARK    = RGBColor(15, 23, 42)    # #0F172A - Deep Text
SLATE_MUTED   = RGBColor(71, 85, 105)   # #475569 - Muted Text
BG_PEARL      = RGBColor(248, 247, 254) # #F8F7FE - Ultra-clean soft ambient lavender background
CARD_WHITE    = RGBColor(255, 255, 255) # #FFFFFF - Crisp white cards
CARD_BORDER   = RGBColor(221, 214, 254) # #DDD6FE - Violet-tinted border
WHITE         = RGBColor(255, 255, 255)

def set_slide_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_header(slide, title_text, category_text="STUDY BUDDY AI  |  BACKEND ARCHITECTURE", slide_num=None):
    set_slide_bg(slide, BG_PEARL)
    
    # Top banner bar - Cosmic Midnight
    banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.05))
    banner.fill.solid()
    banner.fill.fore_color.rgb = SPACE_DARK
    banner.line.fill.background()

    # Glowing Cyan & Electric Violet Accent stripe
    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.01), Inches(13.333), Inches(0.05))
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = VIOLET_MED
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
    p_cat.font.color.rgb = CYAN_GLOW

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
            p.font.color.rgb = VIOLET_DEEP
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
            r1.text = f"◆ {bold_prefix}: "
            r1.font.name = "Calibri"
            r1.font.size = Pt(font_size)
            r1.font.bold = True
            r1.font.color.rgb = VIOLET_DEEP

            r2 = p.add_run()
            r2.text = rest
            r2.font.name = "Calibri"
            r2.font.size = Pt(font_size)
            r2.font.bold = False
            r2.font.color.rgb = text_color
        else:
            r = p.add_run()
            r.text = f"◆ {item}"
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
# SLIDE 1: TITLE SLIDE (Electric Violet & Cosmic Indigo)
# ==============================================================================
s1 = prs.slides.add_slide(blank_layout)
set_slide_bg(s1, SPACE_DARK)

# Left decorative accent stripes
bar1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.28), Inches(7.5))
bar1.fill.solid()
bar1.fill.fore_color.rgb = VIOLET_MED
bar1.line.fill.background()

bar2 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.28), Inches(0), Inches(0.08), Inches(7.5))
bar2.fill.solid()
bar2.fill.fore_color.rgb = CYAN_GLOW
bar2.line.fill.background()

# Subject Badge Pill
badge = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), Inches(0.85), Inches(5.4), Inches(0.45))
badge.adjustments[0] = 0.5
badge.fill.solid()
badge.fill.fore_color.rgb = VIOLET_DEEP
badge.line.color.rgb = CYAN_GLOW
p_b = badge.text_frame.paragraphs[0]
p_b.alignment = PP_ALIGN.CENTER
p_b.text = "AI-AUGMENTED BACKEND DEVELOPMENT - A&S"
p_b.font.name = "Calibri"
p_b.font.size = Pt(11)
p_b.font.bold = True
p_b.font.color.rgb = CYAN_GLOW

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
p_m2.font.color.rgb = LAVENDER_ICE

p_m3 = tf_main.add_paragraph()
p_m3.space_before = Pt(8)
p_m3.text = "Node.js  •  Express.js  •  MongoDB & Mongoose  •  JWT & RBAC  •  Google Gemini 2.5 Flash API"
p_m3.font.name = "Calibri"
p_m3.font.size = Pt(13)
p_m3.font.bold = True
p_m3.font.color.rgb = CYAN_GLOW

# Team Information Card
team_card = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), Inches(3.8), Inches(11.4), Inches(3.05))
team_card.adjustments[0] = 0.03
team_card.fill.solid()
team_card.fill.fore_color.rgb = RGBColor(25, 20, 52)
team_card.line.color.rgb = VIOLET_MED
team_card.line.width = Pt(1.5)

# Team Leader Box (Left side of card)
tb_tl = s1.shapes.add_textbox(Inches(1.35), Inches(4.05), Inches(4.8), Inches(2.5))
tf_tl = tb_tl.text_frame
tf_tl.word_wrap = True
p_tl_h = tf_tl.paragraphs[0]
p_tl_h.text = "TEAM LEADER"
p_tl_h.font.name = "Calibri"
p_tl_h.font.size = Pt(12)
p_tl_h.font.bold = True
p_tl_h.font.color.rgb = CYAN_GLOW

p_tl_n = tf_tl.add_paragraph()
p_tl_n.space_before = Pt(8)
r_tln = p_tl_n.add_run()
r_tln.text = "TRAVIS FELIX PETERS"
r_tln.font.name = "Calibri"
r_tln.font.size = Pt(21)
r_tln.font.bold = True
r_tln.font.color.rgb = WHITE

p_tl_reg = tf_tl.add_paragraph()
p_tl_reg.space_before = Pt(4)
r_reg_lbl = p_tl_reg.add_run()
r_reg_lbl.text = "Reg. No: "
r_reg_lbl.font.name = "Calibri"
r_reg_lbl.font.size = Pt(14)
r_reg_lbl.font.bold = True
r_reg_lbl.font.color.rgb = LAVENDER_ICE

r_reg_val = p_tl_reg.add_run()
r_reg_val.text = "212404437"
r_reg_val.font.name = "Calibri"
r_reg_val.font.size = Pt(15)
r_reg_val.font.bold = True
r_reg_val.font.color.rgb = CYAN_GLOW

p_tl_sub = tf_tl.add_paragraph()
p_tl_sub.space_before = Pt(14)
p_tl_sub.text = "Role: Full Backend Architecture,\nGemini AI Integration & Security Systems"
p_tl_sub.font.name = "Calibri"
p_tl_sub.font.size = Pt(11.5)
p_tl_sub.font.color.rgb = LAVENDER_ICE

# Vertical divider
vdiv = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.4), Inches(4.05), Inches(0.02), Inches(2.55))
vdiv.fill.solid()
vdiv.fill.fore_color.rgb = VIOLET_MED
vdiv.line.fill.background()

# Teammates Box (Right side of card)
tb_tm = s1.shapes.add_textbox(Inches(6.75), Inches(4.05), Inches(5.2), Inches(2.55))
tf_tm = tb_tm.text_frame
tf_tm.word_wrap = True
p_tm_h = tf_tm.paragraphs[0]
p_tm_h.text = "TEAMMATES"
p_tm_h.font.name = "Calibri"
p_tm_h.font.size = Pt(12)
p_tm_h.font.bold = True
p_tm_h.font.color.rgb = CYAN_GLOW

new_teammates = [
    ("BALAJI K", "212404404"),
    ("SAINESH V", "212404429"),
    ("MOHITH R", "212404421"),
    ("SANKARA KRISHNAN R", "212404433")
]
for name, reg in new_teammates:
    pm = tf_tm.add_paragraph()
    pm.space_before = Pt(5)
    
    r_bullet = pm.add_run()
    r_bullet.text = "•  "
    r_bullet.font.name = "Calibri"
    r_bullet.font.size = Pt(14)
    r_bullet.font.bold = True
    r_bullet.font.color.rgb = CYAN_GLOW
    
    r_name = pm.add_run()
    r_name.text = f"{name} "
    r_name.font.name = "Calibri"
    r_name.font.size = Pt(14)
    r_name.font.bold = True
    r_name.font.color.rgb = WHITE
    
    r_reg = pm.add_run()
    r_reg.text = f"({reg})"
    r_reg.font.name = "Calibri"
    r_reg.font.size = Pt(13)
    r_reg.font.bold = True
    r_reg.font.color.rgb = CYAN_GLOW

# ==============================================================================
# SLIDE 2: PROBLEM STATEMENT & PROPOSED SOLUTION
# ==============================================================================
s2 = prs.slides.add_slide(blank_layout)
add_header(s2, "Scenario Case Study: Problem Statement & Proposed Solution", "CASE STUDY & SOLUTION", 2)

ct_prob = add_card(s2, 0.55, 1.3, 5.9, 5.65, "The Problem: Rahul's Academic Dilemma", header_bg=VIOLET_DEEP)
add_bullets(s2, 0.75, ct_prob + 0.1, 5.5, 4.9, [
    ("Student Persona (Rahul)", "2nd-year engineering student preparing for university semester exams while studying new technical skills for internships."),
    ("Fragmented Materials", "Lecture notes, textbooks, PDFs, and handwritten notes are scattered, making revision disorganized."),
    ("Severe Time Drain", "Summarizing hundreds of pages manually and formulating revision questions is labor-intensive and repetitive."),
    ("No Adaptive Scheduling", "Lacks structured study plans tailored to available daily study hours and upcoming exam dates."),
    ("Self-Assessment Deficit", "Difficulty gauging conceptual mastery without automated, objective practice quizzes."),
    ("Productivity Loss", "Switching across disparate apps for note-taking, calendars, and test creation slows exam readiness.")
], font_size=12.2, space_after=8)

ct_sol = add_card(s2, 6.85, 1.3, 5.9, 5.65, "The Solution: Study Buddy AI Platform", header_bg=INDIGO_VIBRANT)
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
c3_l = add_card(s3, 0.55, 1.3, 6.7, 5.65, "AI Study Buddy — System Architecture", header_bg=VIOLET_DEEP)
add_image_fitted(s3, "image1.png", 0.7, 1.9, 6.4, 4.9)

# Technology & MVC on Right
c3_r = add_card(s3, 7.5, 1.3, 5.28, 5.65, "Technology Stack & MVC Pattern", header_bg=INDIGO_VIBRANT)
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
c4_l = add_card(s4, 0.55, 1.3, 6.7, 5.65, "AI Study Buddy — ER Diagram (6 Entities)", header_bg=VIOLET_DEEP)
add_image_fitted(s4, "image2.png", 0.7, 1.9, 6.4, 4.9)

# Key Features on Right
c4_r = add_card(s4, 7.5, 1.3, 5.28, 5.65, "Core Capabilities & Security", header_bg=INDIGO_VIBRANT)
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
t1 = add_card(s5, 0.55, 1.25, 3.9, 4.5, "1. Auth & Material Upload", header_bg=VIOLET_DEEP)
add_image_fitted(s5, "image14.png", 0.65, 1.82, 3.7, 1.9)
add_bullets(s5, 0.68, 3.85, 3.64, 1.8, [
    ("POST /api/auth/register", "Registers account, issues JWT token."),
    ("POST /api/material/upload", "Uploads study notes to MongoDB.")
], font_size=11, space_after=3)

t2 = add_card(s5, 4.7, 1.25, 3.9, 4.5, "2. AI Flashcards & Quiz", header_bg=INDIGO_VIBRANT)
add_image_fitted(s5, "image17.png", 4.8, 1.82, 3.7, 1.9)
add_bullets(s5, 4.83, 3.85, 3.64, 1.8, [
    ("POST /api/ai/flashcards", "Generates structured Q&A pairs."),
    ("POST /api/ai/quiz", "Generates 4-option MCQs & key.")
], font_size=11, space_after=3)

t3 = add_card(s5, 8.85, 1.25, 3.9, 4.5, "3. Personalized Study Plan", header_bg=VIOLET_DEEP)
add_image_fitted(s5, "image19.png", 8.95, 1.82, 3.7, 1.9)
add_bullets(s5, 8.98, 3.85, 3.64, 1.8, [
    ("POST /api/ai/study-plan", "Creates tailored 7-day timetable."),
    ("Validated Endpoints", "Asserted via Postman test suite.")
], font_size=11, space_after=3)

# Bottom Conclusion & Links Banner
b_bot = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(5.9), Inches(12.2), Inches(1.1))
b_bot.adjustments[0] = 0.05
b_bot.fill.solid()
b_bot.fill.fore_color.rgb = SPACE_DARK
b_bot.line.color.rgb = VIOLET_MED

tb_bot = s5.shapes.add_textbox(Inches(0.75), Inches(5.95), Inches(11.8), Inches(0.98))
tf_bot = tb_bot.text_frame
tf_bot.word_wrap = True
p_b1 = tf_bot.paragraphs[0]
p_b1.text = "PROJECT CONCLUSION & REPOSITORY LINKS"
p_b1.font.name = "Calibri"
p_b1.font.size = Pt(11)
p_b1.font.bold = True
p_b1.font.color.rgb = CYAN_GLOW

p_b2 = tf_bot.add_paragraph()
p_b2.space_before = Pt(2)
p_b2.text = "Study Buddy AI delivers a scalable, high-security backend transforming study notes into actionable learning tools via Google Gemini. Ready for seamless frontend UI integration."
p_b2.font.name = "Calibri"
p_b2.font.size = Pt(11)
p_b2.font.color.rgb = WHITE

p_b3 = tf_bot.add_paragraph()
p_b3.space_before = Pt(3)
p_b3.text = "GitHub: https://github.com/travpete/AI-StudyBuddy   |   Demo Video: https://youtu.be/N5lNC06rn7w"
p_b3.font.name = "Calibri"
p_b3.font.size = Pt(11.5)
p_b3.font.bold = True
p_b3.font.color.rgb = LAVENDER_ICE

prs.save(OUT_PPTX_D)
print("Saved Electric Violet themed 5-slide PPTX to:", OUT_PPTX_D)

import shutil
try:
    shutil.copyfile(OUT_PPTX_D, OUT_PPTX_DL)
    print("Copied successfully to Downloads:", OUT_PPTX_DL)
except Exception as e:
    print("Could not copy directly to Downloads (might be locked):", e)
