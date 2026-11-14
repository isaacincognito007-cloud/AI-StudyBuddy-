import os
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

# Color palette
NAVY_DARK = RGBColor(15, 23, 42)      # #0F172A
NAVY_MED = RGBColor(30, 58, 138)      # #1E3A8A
BLUE_ACCENT = RGBColor(37, 99, 235)   # #2563EB
BLUE_LIGHT = RGBColor(219, 234, 254)  # #DBEAFE
CYAN_ACCENT = RGBColor(6, 182, 212)   # #06B6D4
EMERALD = RGBColor(16, 185, 129)      # #10B981
BG_LIGHT = RGBColor(248, 250, 252)    # #F8FAFC
CARD_BG = RGBColor(255, 255, 255)     # #FFFFFF
CARD_ALT = RGBColor(241, 245, 249)    # #F1F5F9
BORDER_GRAY = RGBColor(203, 213, 225) # #CBD5E1
TEXT_DARK = RGBColor(15, 23, 42)      # #0F172A
TEXT_MUTED = RGBColor(71, 85, 105)    # #475569
WHITE = RGBColor(255, 255, 255)

def set_slide_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_header(slide, title_text, category_text="STUDY BUDDY AI  |  BACKEND ARCHITECTURE", slide_num=None):
    set_slide_bg(slide, BG_LIGHT)
    # Top banner bar
    banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.05))
    banner.fill.solid()
    banner.fill.fore_color.rgb = NAVY_DARK
    banner.line.fill.background()

    # Accent stripe at bottom of banner
    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.01), Inches(13.333), Inches(0.05))
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = BLUE_ACCENT
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
    p_cat.font.color.rgb = CYAN_ACCENT

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

    # Slide number pill + footer
    if slide_num is not None:
        footer = slide.shapes.add_textbox(Inches(0.55), Inches(7.12), Inches(12.2), Inches(0.3))
        tf_f = footer.text_frame
        p_f = tf_f.paragraphs[0]
        p_f.text = f"Study Buddy AI (LearnMate)  •  Node.js + Express.js + MongoDB + Google Gemini AI                                                                                                  Slide {slide_num} of 16"
        p_f.font.name = "Calibri"
        p_f.font.size = Pt(9.5)
        p_f.font.color.rgb = TEXT_MUTED

def add_card(slide, left, top, width, height, title=None, bg_color=CARD_BG, border_color=BORDER_GRAY, header_bg=None):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.adjustments[0] = 0.02
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    if border_color:
        card.line.color.rgb = border_color
        card.line.width = Pt(1)
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
            p.font.color.rgb = NAVY_MED
            content_top = top + 0.50
    return content_top

def add_bullets(slide, left, top, width, height, items, font_size=12.5, space_after=6, text_color=TEXT_DARK):
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
            r1.text = f"▸ {bold_prefix}: "
            r1.font.name = "Calibri"
            r1.font.size = Pt(font_size)
            r1.font.bold = True
            r1.font.color.rgb = NAVY_MED

            r2 = p.add_run()
            r2.text = rest
            r2.font.name = "Calibri"
            r2.font.size = Pt(font_size)
            r2.font.bold = False
            r2.font.color.rgb = text_color
        else:
            r = p.add_run()
            r.text = f"▸ {item}"
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
# SLIDE 1: TITLE SLIDE (Project Name & Team Info)
# ==============================================================================
s1 = prs.slides.add_slide(blank_layout)
set_slide_bg(s1, NAVY_DARK)

# Left decorative accent bar
bar1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.28), Inches(7.5))
bar1.fill.solid()
bar1.fill.fore_color.rgb = BLUE_ACCENT
bar1.line.fill.background()

bar2 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.28), Inches(0), Inches(0.08), Inches(7.5))
bar2.fill.solid()
bar2.fill.fore_color.rgb = CYAN_ACCENT
bar2.line.fill.background()

# Project Badge
badge = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), Inches(0.95), Inches(4.8), Inches(0.45))
badge.adjustments[0] = 0.5
badge.fill.solid()
badge.fill.fore_color.rgb = NAVY_MED
badge.line.color.rgb = CYAN_ACCENT
p_b = badge.text_frame.paragraphs[0]
p_b.alignment = PP_ALIGN.CENTER
p_b.text = "AI-AUGMENTED BACKEND DEVELOPMENT"
p_b.font.name = "Calibri"
p_b.font.size = Pt(11.5)
p_b.font.bold = True
p_b.font.color.rgb = CYAN_ACCENT

# Main Project Title
tb_main = s1.shapes.add_textbox(Inches(0.95), Inches(1.65), Inches(11.4), Inches(1.8))
tf_main = tb_main.text_frame
tf_main.word_wrap = True
p_m1 = tf_main.paragraphs[0]
p_m1.text = "Study Buddy AI"
p_m1.font.name = "Calibri"
p_m1.font.size = Pt(50)
p_m1.font.bold = True
p_m1.font.color.rgb = WHITE

p_m2 = tf_main.add_paragraph()
p_m2.space_before = Pt(6)
p_m2.text = "LearnMate — AI Assistant for Smarter Learning and Study Support"
p_m2.font.name = "Calibri"
p_m2.font.size = Pt(20)
p_m2.font.color.rgb = BLUE_LIGHT

p_m3 = tf_main.add_paragraph()
p_m3.space_before = Pt(8)
p_m3.text = "Node.js  •  Express.js  •  MongoDB & Mongoose  •  JWT & RBAC  •  Google Gemini 2.5 Flash API"
p_m3.font.name = "Calibri"
p_m3.font.size = Pt(13)
p_m3.font.bold = True
p_m3.font.color.rgb = CYAN_ACCENT

# Team Information Card on Title Slide
team_card = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), Inches(3.85), Inches(11.4), Inches(2.95))
team_card.adjustments[0] = 0.03
team_card.fill.solid()
team_card.fill.fore_color.rgb = RGBColor(30, 41, 59)
team_card.line.color.rgb = BLUE_ACCENT
team_card.line.width = Pt(1.5)

# Team Leader Box (Left side of team card)
tb_tl = s1.shapes.add_textbox(Inches(1.35), Inches(4.15), Inches(4.8), Inches(2.3))
tf_tl = tb_tl.text_frame
tf_tl.word_wrap = True
p_tl_h = tf_tl.paragraphs[0]
p_tl_h.text = "TEAM LEADER"
p_tl_h.font.name = "Calibri"
p_tl_h.font.size = Pt(12)
p_tl_h.font.bold = True
p_tl_h.font.color.rgb = CYAN_ACCENT

p_tl_n = tf_tl.add_paragraph()
p_tl_n.space_before = Pt(10)
p_tl_n.text = "ISAAC SHILOAH HERBERT V"
p_tl_n.font.name = "Calibri"
p_tl_n.font.size = Pt(20)
p_tl_n.font.bold = True
p_tl_n.font.color.rgb = WHITE

p_tl_sub = tf_tl.add_paragraph()
p_tl_sub.space_before = Pt(18)
p_tl_sub.text = "Project: Study Buddy AI\nBackend Architecture & Gemini AI Integration"
p_tl_sub.font.name = "Calibri"
p_tl_sub.font.size = Pt(12)
p_tl_sub.font.color.rgb = BLUE_LIGHT

# Vertical divider inside team card
vdiv = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.4), Inches(4.15), Inches(0.02), Inches(2.35))
vdiv.fill.solid()
vdiv.fill.fore_color.rgb = BLUE_ACCENT
vdiv.line.fill.background()

# Teammates Box (Right side of team card)
tb_tm = s1.shapes.add_textbox(Inches(6.75), Inches(4.15), Inches(5.2), Inches(2.4))
tf_tm = tb_tm.text_frame
tf_tm.word_wrap = True
p_tm_h = tf_tm.paragraphs[0]
p_tm_h.text = "TEAMMATES"
p_tm_h.font.name = "Calibri"
p_tm_h.font.size = Pt(12)
p_tm_h.font.bold = True
p_tm_h.font.color.rgb = CYAN_ACCENT

for member in ["MOHAMMED IRFAN A", "BHUVANESWARI", "BENSAM GILBERT E", "JOEL EDWIN P"]:
    pm = tf_tm.add_paragraph()
    pm.space_before = Pt(7)
    pm.text = f"•  {member}"
    pm.font.name = "Calibri"
    pm.font.size = Pt(15.5)
    pm.font.bold = True
    pm.font.color.rgb = WHITE

# ==============================================================================
# SLIDE 2: EXECUTIVE OVERVIEW & DESCRIPTION
# ==============================================================================
s2 = prs.slides.add_slide(blank_layout)
add_header(s2, "Executive Overview & Platform Description", "1. INTRODUCTION", 2)

ct = add_card(s2, 0.55, 1.3, 6.0, 5.65, "What is Study Buddy AI (LearnMate)?", header_bg=NAVY_MED)
add_bullets(s2, 0.75, ct + 0.08, 5.6, 5.0, [
    ("AI-Powered Learning Platform", "Developed to simplify academic study workflows by combining a modern RESTful backend with Generative AI."),
    ("Material-Grounded Generation", "Unlike static template tools, Study Buddy AI dynamically processes the student's actual uploaded study notes via Google Gemini."),
    ("End-to-End Study Companion", "Enables students to upload study materials, generate concise summaries, build active-recall flashcards, take self-assessment quizzes, and create day-by-day study plans."),
    ("Persistent Account Storage", "All uploaded materials and AI-generated resources are stored in MongoDB under the student's account for anytime revision."),
    ("Role-Scoped Security", "Protected by JWT authentication, bcryptjs password encryption, and strict Role-Based Access Control (RBAC).")
], font_size=12.5, space_after=8)

ct2 = add_card(s2, 6.78, 1.3, 6.0, 5.65, "Core Technology Pillars", header_bg=BLUE_ACCENT)
add_bullets(s2, 6.98, ct2 + 0.08, 5.6, 5.0, [
    ("Node.js & Express.js", "Modular RESTful API architecture separating routes, controllers, middleware, models, and utility services."),
    ("MongoDB & Mongoose ODM", "Primary NoSQL database with 6 structured schemas, indexing, data validation, and relational references."),
    ("Google Gemini 2.5 Flash API", "Uses @google/genai SDK for high-speed NLP generation and guaranteed JSON schema output for flashcards & quizzes."),
    ("JWT & RBAC Security", "7-day stateless Bearer tokens with server-enforced Student/Admin authorization and bcryptjs salted hashing."),
    ("Production Middleware Chain", "Includes Helmet headers, CORS, Morgan logging, Multer file validation, Express 5 NoSQL injection guard, and 2-tier rate limiting.")
], font_size=12.5, space_after=8)

# ==============================================================================
# SLIDE 3: SCENARIO-BASED CASE STUDY (BACKGROUND & PROBLEM)
# ==============================================================================
s3 = prs.slides.add_slide(blank_layout)
add_header(s3, "Scenario-Based Case Study: Background & Problem Statement", "2. CASE STUDY & MOTIVATION", 3)

ct_bg = add_card(s3, 0.55, 1.3, 5.7, 5.65, "Background: The Student Dilemma", header_bg=NAVY_MED)
add_bullets(s3, 0.75, ct_bg + 0.1, 5.3, 4.9, [
    ("Persona — Rahul", "Second-year engineering student preparing for multiple semester exams while building technical skills for internships and placements."),
    ("Fragmented Materials", "Study content is scattered across lecture notes, PDFs, textbooks, online articles, and handwritten notes."),
    ("Severe Time Constraints", "High volume of syllabus content paired with limited preparation time before examinations."),
    ("Inefficient Manual Workflows", "Traditional revision requires hours of manual summarizing, question drafting, and schedule planning."),
    ("Inconsistent Outcomes", "Repetitive manual effort leads to burnout, incomplete revision, and difficulty gauging exam readiness.")
], font_size=13, space_after=10)

ct_pr = add_card(s3, 6.5, 1.3, 6.28, 5.65, "Key Problems Identified", header_bg=BLUE_ACCENT)
add_bullets(s3, 6.7, ct_pr + 0.1, 5.9, 4.9, [
    ("Information Overload", "Difficulty comprehending lengthy study materials within tight deadlines."),
    ("Time-Consuming Note-Taking", "Manual preparation of revision notes consumes time needed for concept mastery."),
    ("Labor-Intensive Flashcards", "Creating active-recall flashcards and practice questions manually is tedious."),
    ("No Personalized Study Plans", "Lack of adaptive schedules tailored to exam dates and daily available hours."),
    ("Lack of Centralized Storage", "No unified, secure platform to store and search AI-generated learning aids."),
    ("Self-Assessment Gap", "Difficulty evaluating comprehension without automated multiple-choice quizzes."),
    ("Context-Switching Friction", "Jumping between separate apps for notes, quizzes, and calendars lowers productivity.")
], font_size=12.5, space_after=7)

# ==============================================================================
# SLIDE 4: PROPOSED SOLUTION, USAGE & OUTCOMES
# ==============================================================================
s4 = prs.slides.add_slide(blank_layout)
add_header(s4, "Proposed Solution, System Usage & Learning Outcomes", "3. SOLUTION & IMPACT", 4)

c1 = add_card(s4, 0.55, 1.3, 4.0, 5.65, "The Study Buddy AI Solution", header_bg=NAVY_MED)
add_bullets(s4, 0.72, c1 + 0.08, 3.65, 5.0, [
    ("Centralized AI Backend", "Automates academic preparation through secure RESTful APIs."),
    ("JWT Security", "Authenticates users and isolates student resources."),
    ("MongoDB Persistence", "Stores profiles, materials, and generated AI assets."),
    ("Gemini AI Service", "Generates summaries, flashcards, MCQs, study plans, and Q&A."),
    ("Modular Separation", "Decouples auth, business logic, DB, and AI layers.")
], font_size=12, space_after=8)

c2 = add_card(s4, 4.68, 1.3, 4.0, 5.65, "Platform Usage Workflows", header_bg=BLUE_ACCENT)
add_bullets(s4, 4.85, c2 + 0.08, 3.65, 5.0, [
    ("Student Registration & Login", "Sign up securely and receive a JWT Bearer token."),
    ("Upload Study Material", "Submit text notes or upload .txt / .md / .pdf files."),
    ("On-Demand AI Generation", "Trigger summaries, flashcards, quizzes, or study plans."),
    ("Review & Search", "Paginate, filter by subject, download, or revisit saved outputs."),
    ("Administrator Oversight", "Manage users, update roles, and monitor AI usage analytics.")
], font_size=12, space_after=8)

c3 = add_card(s4, 8.81, 1.3, 3.97, 5.65, "Measurable Outcomes", header_bg=NAVY_MED)
add_bullets(s4, 8.98, c3 + 0.08, 3.62, 5.0, [
    ("Faster Revision Cycles", "Drastically cuts manual summarizing and note-making time."),
    ("Active Recall Mastery", "Automated flashcards & quizzes boost retention before exams."),
    ("Structured Exam Prep", "Personalized day-by-day plans keep students on track."),
    ("Secure & Consistent", "JWT + MongoDB + Gemini ensure reliable, private access."),
    ("Extensible Foundation", "Ready for React frontend and future AI modules.")
], font_size=12, space_after=8)

# ==============================================================================
# SLIDE 5: SOFTWARE & HARDWARE REQUIREMENTS
# ==============================================================================
s5 = prs.slides.add_slide(blank_layout)
add_header(s5, "System Specifications: Software & Hardware Requirements", "4. SYSTEM REQUIREMENTS", 5)

cs_sw = add_card(s5, 0.55, 1.3, 6.0, 5.65, "Software Requirements", header_bg=NAVY_MED)
add_bullets(s5, 0.75, cs_sw + 0.1, 5.6, 4.9, [
    ("Operating System", "Windows 10/11, macOS, or Linux (cross-platform runtime support)."),
    ("Node.js (v16 / v18+)", "JavaScript runtime environment executing server logic and async I/O."),
    ("npm (v8 or above)", "Node package manager managing dependencies and run scripts."),
    ("Express.js", "Lightweight web framework handling RESTful routing and middleware."),
    ("MongoDB & Mongoose", "NoSQL document database (local or MongoDB Atlas M0 cluster)."),
    ("Google Gemini API", "@google/genai SDK communicating with the gemini-2.5-flash model."),
    ("Postman / Thunder Client", "API testing suite to verify endpoints, schemas, and RBAC shields."),
    ("Visual Studio Code", "Primary IDE for backend development, debugging, and terminal execution.")
], font_size=12.5, space_after=7)

cs_hw = add_card(s5, 6.78, 1.3, 6.0, 2.55, "Hardware Requirements", header_bg=BLUE_ACCENT)
add_bullets(s5, 6.98, cs_hw + 0.08, 5.6, 1.8, [
    ("Processor", "Intel Core i5 (8th Gen or above) / AMD Ryzen 5 or equivalent."),
    ("Memory (RAM)", "8 GB minimum (16 GB recommended for concurrent DB & Node processes)."),
    ("Disk Storage", "~1 GB available workspace for dependencies and uploaded study files.")
], font_size=12.5, space_after=7)

cs_pkg = add_card(s5, 6.78, 4.05, 6.0, 2.9, "Key NPM Dependencies Installed", header_bg=NAVY_DARK)
add_bullets(s5, 6.98, cs_pkg + 0.08, 5.6, 2.15, [
    ("Core & DB", "express, mongoose, dotenv, cors, morgan"),
    ("Auth & Security", "jsonwebtoken, bcryptjs, helmet, express-rate-limit, express-validator"),
    ("AI & File Uploads", "@google/genai (Gemini 2.5 Flash), multer (multipart file uploads)"),
    ("Dev Tooling", "nodemon (automatic server reload during development)")
], font_size=12.5, space_after=7)

# ==============================================================================
# SLIDE 6: EPIC 1 — PROJECT ARCHITECTURE
# ==============================================================================
s6 = prs.slides.add_slide(blank_layout)
add_header(s6, "Epic 1: Modular RESTful Project Architecture", "5. EPIC 1 — ARCHITECTURE", 6)

ca_left = add_card(s6, 0.55, 1.3, 5.1, 5.65, "Architectural Components", header_bg=NAVY_MED)
add_bullets(s6, 0.73, ca_left + 0.06, 4.75, 5.0, [
    ("Client Application", "React.js frontend / Postman client sending HTTPS + JWT Bearer requests."),
    ("Express Server", "Central HTTP server handling routing, middleware chain, and error responses."),
    ("Auth Middleware", "Verifies JWT tokens (protect) and enforces Student/Admin roles (authorize)."),
    ("Route Handlers", "Modular endpoints for /api/auth, /api/materials, /api/ai, and /api/admin."),
    ("AI Service Layer", "Dedicated utility communicating with Google Gemini 2.5 Flash API."),
    ("Database Layer", "MongoDB + Mongoose ODM persisting users, materials, and AI outputs.")
], font_size=12, space_after=7)

ca_right = add_card(s6, 5.85, 1.3, 6.93, 5.65, "AI StudyBuddy — System Architecture Diagram", header_bg=BLUE_ACCENT)
add_image_fitted(s6, "image1.png", 6.0, 1.9, 6.63, 4.9)

# ==============================================================================
# SLIDE 7: MVC ARCHITECTURE PATTERN
# ==============================================================================
s7 = prs.slides.add_slide(blank_layout)
add_header(s7, "Model-View-Controller (MVC) Architecture Pattern", "5. EPIC 1 — MVC PATTERN", 7)

cm_left = add_card(s7, 0.55, 1.3, 5.1, 5.65, "Layered Separation of Concerns", header_bg=NAVY_MED)
add_bullets(s7, 0.73, cm_left + 0.08, 4.75, 5.0, [
    ("Model Layer (Data Layer)", "Defines 6 MongoDB collections using Mongoose schemas (User, StudyMaterial, Summary, Flashcard, Quiz, StudyPlan) with validation & indexing."),
    ("View Layer (API Response)", "As a headless RESTful API, structured JSON responses serve as the View layer consumed by React or Postman clients."),
    ("Controller Layer (Business Logic)", "Validates input, enforces ownership (userId scoping), calls Gemini AI utilities, and formats HTTP responses."),
    ("Key MVC Benefits", "Improves code readability, simplifies unit testing and debugging, and allows new AI features to be added cleanly.")
], font_size=12.2, space_after=9)

cm_right = add_card(s7, 5.85, 1.3, 6.93, 5.65, "MVC Architecture Diagram", header_bg=BLUE_ACCENT)
add_image_fitted(s7, "image4.png", 6.0, 1.9, 6.63, 4.9)

# ==============================================================================
# SLIDE 8: ENTITY-RELATIONSHIP (ER) DIAGRAM & DATA MODEL
# ==============================================================================
s8 = prs.slides.add_slide(blank_layout)
add_header(s8, "Database Design & Entity-Relationship (ER) Diagram", "5. EPIC 1 — DATA MODELING", 8)

ce_left = add_card(s8, 0.55, 1.3, 5.2, 5.65, "Entities & Cardinality Relationships", header_bg=NAVY_MED)
add_bullets(s8, 0.73, ce_left + 0.06, 4.85, 5.0, [
    ("1. User Entity", "_id (PK), name, email (unique), password (hashed), role (student/admin), createdAt."),
    ("2. StudyMaterial Entity", "_id (PK), userId (FK → User), title, subject, content, fileName, createdAt."),
    ("3. Summary Entity", "_id (PK), userId (FK), materialId (FK → StudyMaterial), summary, createdAt."),
    ("4. Flashcard Entity", "_id (PK), userId (FK), materialId (FK), question, answer, createdAt."),
    ("5. Quiz Entity", "_id (PK), userId (FK), materialId (FK), questions[] (options & answer), createdAt."),
    ("6. StudyPlan Entity", "_id (PK), userId (FK → User), studyPlan, examDate, createdAt."),
    ("Key Relationships", "User → StudyMaterial (1:N), Material → Summary (1:1), Material → Flashcards (1:N), Material → Quiz (1:1), User → StudyPlan (1:N).")
], font_size=11.5, space_after=6)

ce_right = add_card(s8, 5.95, 1.3, 6.83, 5.65, "AI StudyBuddy — ER Diagram", header_bg=BLUE_ACCENT)
add_image_fitted(s8, "image2.png", 6.1, 1.9, 6.53, 4.9)

# ==============================================================================
# SLIDE 9: ROLES, RESPONSIBILITIES & USER FLOW
# ==============================================================================
s9 = prs.slides.add_slide(blank_layout)
add_header(s9, "User Roles, Access Control & End-to-End User Flow", "5. EPIC 1 — ROLES & USER FLOW", 9)

cr_left = add_card(s9, 0.55, 1.3, 5.1, 5.65, "Roles & User Journey", header_bg=NAVY_MED)
add_bullets(s9, 0.73, cr_left + 0.08, 4.75, 5.0, [
    ("Student Role (Default)", "Registers/logs in, uploads study materials, generates AI summaries, flashcards, quizzes, and study plans, and manages personal history."),
    ("Administrator Role", "Monitors registered users, updates roles, oversees system-wide materials, and pulls platform-wide AI usage analytics."),
    ("Privilege Escalation Guard", "Self-registration strictly forces role='student' server-side; only an existing admin or DB owner can promote an account."),
    ("Step-by-Step User Flow", "Register/Login → JWT Issued → Upload Material → Select Gemini AI Feature → Auto-Save Output in MongoDB → Review/Download Anytime.")
], font_size=12.2, space_after=9)

cr_right = add_card(s9, 5.85, 1.3, 6.93, 5.65, "User Flow Diagram — AI Study Buddy", header_bg=BLUE_ACCENT)
add_image_fitted(s9, "image3.png", 6.0, 1.9, 6.63, 4.9)

# ==============================================================================
# SLIDE 10: KEY FEATURES & DEFENSE-IN-DEPTH SECURITY
# ==============================================================================
s10 = prs.slides.add_slide(blank_layout)
add_header(s10, "Key Platform Features & Security Engineering", "6. FEATURES & SECURITY", 10)

cf_left = add_card(s10, 0.55, 1.3, 6.0, 5.65, "8 Core Platform Capabilities", header_bg=NAVY_MED)
add_bullets(s10, 0.75, cf_left + 0.08, 5.6, 5.0, [
    ("Secure User Authentication", "JWT-based stateless auth (7-day expiry) with bcrypt password hashing."),
    ("Study Material Management", "Upload text or files (.txt, .md, .pdf), paginate, search by text, filter by subject, update, delete, or download."),
    ("AI Summary Generation", "Produces structured, easy-to-revise summaries from uploaded notes."),
    ("AI Flashcard Generation", "Creates Q&A flashcards for active recall using structured JSON schemas."),
    ("AI Quiz Generation", "Generates 4-option multiple-choice quizzes with verified answer indices."),
    ("AI Study Plan & Direct Q&A", "Builds custom day-by-day schedules and answers targeted student questions."),
    ("Admin Analytics Dashboard", "Tracks total users, materials, and AI generations broken down by type.")
], font_size=12, space_after=7)

cf_right = add_card(s10, 6.78, 1.3, 6.0, 5.65, "Defense-in-Depth Security Architecture", header_bg=BLUE_ACCENT)
add_bullets(s10, 6.98, cf_right + 0.08, 5.6, 5.0, [
    ("Express 5 NoSQL-Injection Guard", "Custom sanitize.js middleware recursively rejects keys starting with '$' or containing '.' (avoiding express-mongo-sanitize Express 5 crashes)."),
    ("Two-Tier Rate Limiting", "Global rate limiter (100 req / 15 min) plus a stricter AI limiter (20 req / 15 min) protecting Gemini API quota."),
    ("Query-Level Data Isolation", "Every student DB query filters by { userId: req.user._id }, preventing horizontal privilege escalation (IDOR)."),
    ("Strict Input Validation", "express-validator rules enforce email format, password length, and required fields on all write routes."),
    ("Centralized Error Handling", "Unified errorHandler.js normalizes Mongoose CastErrors, duplicate keys (11000), and JWT errors without leaking stack traces.")
], font_size=12, space_after=7)

# ==============================================================================
# SLIDE 11: EPIC 2 & 3 — SETUP, ENVIRONMENT & BACKEND STRUCTURE
# ==============================================================================
s11 = prs.slides.add_slide(blank_layout)
add_header(s11, "Epics 2 & 3: Project Setup, Environment & Backend Structure", "7. EPICS 2 & 3 — SETUP & STRUCTURE", 11)

cs11_l = add_card(s11, 0.55, 1.3, 4.9, 5.65, "Setup & Environment (.env)", header_bg=NAVY_MED)
add_bullets(s11, 0.73, cs11_l + 0.06, 4.55, 5.0, [
    ("1. Initialize Project", "npm init -y"),
    ("2. Install Dependencies", "npm install express mongoose bcryptjs jsonwebtoken cors dotenv multer @google/genai"),
    ("3. Dev Server Tooling", "npm install --save-dev nodemon"),
    ("4. Environment (.env)", "PORT=5000\n   MONGO_URI=your_mongodb_uri\n   JWT_SECRET=your_jwt_secret\n   GEMINI_API_KEY=your_gemini_key"),
    ("5. Startup Guarantee", "index.js connects to MongoDB via connectDB() before starting Express listener.")
], font_size=12, space_after=8)

cs11_m = add_card(s11, 5.62, 1.3, 2.75, 5.65, "Folder Structure", header_bg=BLUE_ACCENT)
add_image_fitted(s11, "image6.png", 5.72, 1.9, 2.55, 4.9)

cs11_r = add_card(s11, 8.55, 1.3, 4.23, 5.65, "Entry Point (index.js) & DB (db.js)", header_bg=NAVY_DARK)
add_image_fitted(s11, "image7.png", 8.68, 1.9, 3.97, 3.1)
add_image_fitted(s11, "image8.png", 8.68, 5.1, 3.97, 1.7)

# ==============================================================================
# SLIDE 12: EPIC 3 — CORE BACKEND IMPLEMENTATIONS
# ==============================================================================
s12 = prs.slides.add_slide(blank_layout)
add_header(s12, "Epic 3: Authentication, Gemini AI Service & File Upload Middleware", "8. EPIC 3 — CORE SCRIPTS", 12)

c12_1 = add_card(s12, 0.55, 1.3, 4.0, 5.65, "1. Auth Middleware (auth.js)", header_bg=NAVY_MED)
add_image_fitted(s12, "image9.png", 0.68, 1.9, 3.74, 3.1)
add_bullets(s12, 0.68, 5.12, 3.74, 1.7, [
    ("JWT Verification", "Extracts Bearer token, verifies signature, and attaches req.user."),
    ("Role Shield", "adminOnly / authorize restricts admin endpoints (403 Forbidden).")
], font_size=11.5, space_after=5)

c12_2 = add_card(s12, 4.68, 1.3, 4.0, 5.65, "2. AI Service (gemini.js)", header_bg=BLUE_ACCENT)
add_image_fitted(s12, "image10.png", 4.81, 1.9, 3.74, 2.2)
add_bullets(s12, 4.81, 4.25, 3.74, 2.5, [
    ("Gemini 2.5 Flash", "Integrates GoogleGenerativeAI / @google/genai SDK."),
    ("Dynamic Prompting", "Generates summaries, flashcards, quizzes, and study plans directly from uploaded content."),
    ("Structured JSON", "Enforces strict JSON schema on flashcards and quiz arrays.")
], font_size=11.5, space_after=5)

c12_3 = add_card(s12, 8.81, 1.3, 3.97, 5.65, "3. File Upload (upload.js)", header_bg=NAVY_MED)
add_image_fitted(s12, "image11.png", 8.94, 1.9, 3.71, 3.1)
add_bullets(s12, 8.94, 5.12, 3.71, 1.7, [
    ("Multer Disk Storage", "Generates unique timestamped filenames in uploads/."),
    ("Validation & Limits", "Restricts extensions (.txt, .md, .pdf) and caps size at 5 MB.")
], font_size=11.5, space_after=5)

# ==============================================================================
# SLIDE 13: EPIC 4 — DATABASE CONFIGURATION & MONGOOSE SCHEMAS
# ==============================================================================
s13 = prs.slides.add_slide(blank_layout)
add_header(s13, "Epic 4: Database Configuration & Mongoose Schema Definitions", "9. EPIC 4 — DATABASE & SCHEMAS", 13)

c13_l = add_card(s13, 0.55, 1.3, 6.0, 5.65, "1. User Schema (models/User.js)", header_bg=NAVY_MED)
add_image_fitted(s13, "image12.png", 0.75, 1.9, 5.6, 3.55)
add_bullets(s13, 0.75, 5.55, 5.6, 1.3, [
    ("Pre-Save Hook", "Automatically hashes modified passwords with bcryptjs (10 salt rounds)."),
    ("Password Verification", "matchPassword() instance method compares login credentials securely.")
], font_size=11.5, space_after=4)

c13_r = add_card(s13, 6.78, 1.3, 6.0, 5.65, "2. Study Material Schema (models/Material.js)", header_bg=BLUE_ACCENT)
add_image_fitted(s13, "image13.png", 6.98, 1.9, 5.6, 3.55)
add_bullets(s13, 6.98, 5.55, 5.6, 1.3, [
    ("Relational Reference", "Links each uploaded material to its owner via ObjectId ref: 'User'."),
    ("Dedicated Collections", "Also supported by modular Summary, Flashcard, Quiz & StudyPlan models.")
], font_size=11.5, space_after=4)

# ==============================================================================
# SLIDE 14: RESTFUL API ENDPOINTS REFERENCE TABLE
# ==============================================================================
s14 = prs.slides.add_slide(blank_layout)
add_header(s14, "Complete RESTful API Route Interface Endpoints", "10. API SPECIFICATION", 14)

rows_data = [
    ["Route Group", "HTTP Method & Endpoint", "Access", "Functional Description"],
    ["1. Auth APIs", "POST /api/auth/register", "Public", "Registers a new student account with bcrypt-hashed password & returns JWT."],
    ["1. Auth APIs", "POST /api/auth/login", "Public", "Authenticates email/password credentials and issues a 7-day JWT Bearer token."],
    ["1. Auth APIs", "GET / PUT /api/auth/profile", "Student/Admin", "Retrieves or updates the authenticated user's profile details."],
    ["2. Material APIs", "POST /api/materials/upload", "Student/Admin", "Uploads study material (text or .txt/.md/.pdf file) and stores it in MongoDB."],
    ["2. Material APIs", "GET /api/materials", "Student/Admin", "Lists user's study materials with pagination, subject filter, and full-text search."],
    ["2. Material APIs", "GET / PUT / DELETE /api/materials/:id", "Student/Admin", "Retrieves, updates, deletes, or downloads (/download) a specific study material."],
    ["3. AI Feature APIs", "POST /api/ai/summary", "Student/Admin", "Sends study material to Gemini 2.5 Flash to generate & store a concise summary."],
    ["3. AI Feature APIs", "POST /api/ai/flashcards", "Student/Admin", "Generates structured question-and-answer flashcards for active recall."],
    ["3. AI Feature APIs", "POST /api/ai/quiz", "Student/Admin", "Generates a multiple-choice quiz with 4 options and verified answer key."],
    ["3. AI Feature APIs", "POST /api/ai/study-plan  |  /ask", "Student/Admin", "Creates a personalized day-by-day study schedule or answers direct Q&A."],
    ["4. Admin APIs", "GET / DELETE /api/admin/users & /analytics", "Admin Only", "Manages registered users, updates roles, and retrieves system usage metrics."]
]

table_shape = s14.shapes.add_table(len(rows_data), 4, Inches(0.55), Inches(1.28), Inches(12.23), Inches(5.65))
tbl = table_shape.table
tbl.columns[0].width = Inches(1.85)
tbl.columns[1].width = Inches(3.35)
tbl.columns[2].width = Inches(1.45)
tbl.columns[3].width = Inches(5.58)

for r_idx, row in enumerate(rows_data):
    for c_idx, val in enumerate(row):
        cell = tbl.cell(r_idx, c_idx)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid()
        if r_idx == 0:
            cell.fill.fore_color.rgb = NAVY_DARK
        elif r_idx % 2 == 1:
            cell.fill.fore_color.rgb = WHITE
        else:
            cell.fill.fore_color.rgb = CARD_ALT
        p = cell.text_frame.paragraphs[0]
        p.text = val
        p.font.name = "Calibri"
        p.font.size = Pt(11 if r_idx > 0 else 12)
        p.font.bold = (r_idx == 0 or c_idx in (0, 1))
        if r_idx == 0:
            p.font.color.rgb = WHITE
        elif c_idx == 1:
            p.font.color.rgb = BLUE_ACCENT
        else:
            p.font.color.rgb = TEXT_DARK

# ==============================================================================
# SLIDE 15: EPIC 5 — API TESTING IN POSTMAN (AUTH & UPLOAD)
# ==============================================================================
s15 = prs.slides.add_slide(blank_layout)
add_header(s15, "Epic 5: Project Execution & Postman API Verification (Part 1)", "11. EPIC 5 — API TESTING", 15)

p15_1 = add_card(s15, 0.55, 1.3, 4.0, 5.65, "1. User Registration (201 Created)", header_bg=NAVY_MED)
add_image_fitted(s15, "image14.png", 0.65, 1.9, 3.8, 3.35)
add_bullets(s15, 0.68, 5.35, 3.74, 1.4, [
    ("POST /api/auth/register", "Creates account with hashed password and returns JWT accessToken.")
], font_size=11.5)

p15_2 = add_card(s15, 4.68, 1.3, 4.0, 5.65, "2. User Login (200 OK)", header_bg=BLUE_ACCENT)
add_image_fitted(s15, "image15.png", 4.78, 1.9, 3.8, 3.35)
add_bullets(s15, 4.81, 5.35, 3.74, 1.4, [
    ("POST /api/auth/login", "Verifies credentials and issues fresh JWT Bearer token for protected routes.")
], font_size=11.5)

p15_3 = add_card(s15, 8.81, 1.3, 3.97, 5.65, "3. Upload Material & Summarize", header_bg=NAVY_MED)
add_image_fitted(s15, "image16.png", 8.91, 1.9, 3.77, 3.35)
add_bullets(s15, 8.94, 5.35, 3.71, 1.4, [
    ("POST /api/materials/upload", "Uploads Node.js study notes via multipart form-data for Gemini AI processing.")
], font_size=11.5)

# ==============================================================================
# SLIDE 16: EPIC 5 — AI GENERATION TESTING & PROJECT LINKS
# ==============================================================================
s16 = prs.slides.add_slide(blank_layout)
add_header(s16, "Epic 5: Gemini AI Generation Verification & Project Links", "12. AI OUTPUTS & CONCLUSION", 16)

p16_1 = add_card(s16, 0.55, 1.25, 4.0, 4.65, "4. Generate AI Flashcards", header_bg=NAVY_MED)
add_image_fitted(s16, "image17.png", 0.65, 1.82, 3.8, 2.95)
add_bullets(s16, 0.68, 4.85, 3.74, 0.95, [
    ("POST /api/ai/flashcards", "Returns structured Q&A flashcards from Node.js notes.")
], font_size=11)

p16_2 = add_card(s16, 4.68, 1.25, 4.0, 4.65, "5. Generate AI Quiz (MCQs)", header_bg=BLUE_ACCENT)
add_image_fitted(s16, "image18.png", 4.78, 1.82, 3.8, 2.95)
add_bullets(s16, 4.81, 4.85, 3.74, 0.95, [
    ("POST /api/ai/quiz", "Produces 4-option multiple-choice questions & answer key.")
], font_size=11)

p16_3 = add_card(s16, 8.81, 1.25, 3.97, 4.65, "6. Personalized Study Plan", header_bg=NAVY_MED)
add_image_fitted(s16, "image19.png", 8.91, 1.82, 3.77, 2.95)
add_bullets(s16, 8.94, 4.85, 3.71, 0.95, [
    ("POST /api/ai/study-plan", "Creates tailored 7-day, 2 hrs/day exam revision schedule.")
], font_size=11)

# Bottom Links Banner
links_card = s16.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(6.02), Inches(12.23), Inches(0.98))
links_card.adjustments[0] = 0.05
links_card.fill.solid()
links_card.fill.fore_color.rgb = NAVY_DARK
links_card.line.color.rgb = BLUE_ACCENT
tb_l = s16.shapes.add_textbox(Inches(0.75), Inches(6.08), Inches(11.8), Inches(0.85))
tf_l = tb_l.text_frame
tf_l.word_wrap = True
pl1 = tf_l.paragraphs[0]
pl1.text = "PROJECT LINKS & REPOSITORY RESOURCES"
pl1.font.name = "Calibri"
pl1.font.size = Pt(11)
pl1.font.bold = True
pl1.font.color.rgb = CYAN_ACCENT

pl2 = tf_l.add_paragraph()
pl2.space_before = Pt(3)
pl2.text = "GitHub Repo: https://github.com/isaacincognito007-cloud/AI-StudyBuddy-   |   Demo Video: https://youtu.be/N5lNC06rn7w"
pl2.font.name = "Calibri"
pl2.font.size = Pt(12)
pl2.font.bold = True
pl2.font.color.rgb = WHITE

prs.save(OUT_PPTX_D)
prs.save(OUT_PPTX_DL)
print("Saved PPTX to:", OUT_PPTX_D, "and", OUT_PPTX_DL)
