import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

MEDIA_DIR = r"C:\Users\acer\.gemini\antigravity\brain\3df468b6-62f3-4271-a7b4-01a98becf01a\scratch\media"
CROPPED_DIR = os.path.join(MEDIA_DIR, "cropped")
OUT_PPTX_D = r"d:\AI-StudyBuddy\Study_Buddy_AI_Presentation.pptx"
OUT_PPTX_DL = r"c:\Users\acer\Downloads\Study_Buddy_AI_Presentation.pptx"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

# ==============================================================================
# MINIMALIST DESIGN THEME (White / Light Grey + Charcoal + Teal + Muted Grey)
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
WHITE         = RGBColor(255, 255, 255)

FONT_HEADING = "Calibri"
FONT_BODY    = "Calibri"

def set_slide_bg(slide, color=BG_PAGE):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_header(slide, title_text, category_text="AI-AUGMENTED BACKEND DEVELOPMENT", subtitle_text=None):
    set_slide_bg(slide, BG_PAGE)

    # Category Pill / Label at top
    tb_cat = slide.shapes.add_textbox(Inches(0.8), Inches(0.40), Inches(10), Inches(0.28))
    tf_cat = tb_cat.text_frame
    tf_cat.word_wrap = True
    p_cat = tf_cat.paragraphs[0]
    p_cat.text = category_text.upper()
    p_cat.font.name = FONT_HEADING
    p_cat.font.size = Pt(10)
    p_cat.font.bold = True
    p_cat.font.color.rgb = TEAL_PRIMARY

    # Main Slide Title
    tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.65), Inches(11.5), Inches(0.6))
    tf_title = tb_title.text_frame
    tf_title.word_wrap = True
    p_title = tf_title.paragraphs[0]
    p_title.text = title_text
    p_title.font.name = FONT_HEADING
    p_title.font.size = Pt(26)
    p_title.font.bold = True
    p_title.font.color.rgb = CHARCOAL_DARK

    if subtitle_text:
        tb_sub = slide.shapes.add_textbox(Inches(0.8), Inches(1.24), Inches(11.5), Inches(0.35))
        tf_sub = tb_sub.text_frame
        p_sub = tf_sub.paragraphs[0]
        p_sub.text = subtitle_text
        p_sub.font.name = FONT_BODY
        p_sub.font.size = Pt(13)
        p_sub.font.color.rgb = GREY_MUTED

    # Subtle bottom header divider rule
    divider = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.64), Inches(11.733), Inches(0.015))
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
# SLIDE 1: TITLE SLIDE (Study Buddy AI)
# ==============================================================================
s1 = prs.slides.add_slide(blank_layout)
set_slide_bg(s1, BG_WHITE)

# Top Left Category Tag
tb_c1 = s1.shapes.add_textbox(Inches(0.9), Inches(0.8), Inches(6.0), Inches(0.35))
tf_c1 = tb_c1.text_frame
p_c1 = tf_c1.paragraphs[0]
p_c1.text = "AI-AUGMENTED BACKEND DEVELOPMENT"
p_c1.font.name = FONT_HEADING
p_c1.font.size = Pt(11)
p_c1.font.bold = True
p_c1.font.color.rgb = TEAL_PRIMARY

# Main Title: Study Buddy AI
tb_t1 = s1.shapes.add_textbox(Inches(0.9), Inches(1.2), Inches(8.5), Inches(1.2))
tf_t1 = tb_t1.text_frame
tf_t1.word_wrap = True
p_t1 = tf_t1.paragraphs[0]
p_t1.text = "Study Buddy AI"
p_t1.font.name = FONT_HEADING
p_t1.font.size = Pt(46)
p_t1.font.bold = True
p_t1.font.color.rgb = CHARCOAL_DARK

# Subtitle: LearnMate — AI Assistant for Smarter Learning and Study Support
p_sub1 = tf_t1.add_paragraph()
p_sub1.space_before = Pt(4)
p_sub1.text = "LearnMate — AI Assistant for Smarter Learning and Study Support"
p_sub1.font.name = FONT_BODY
p_sub1.font.size = Pt(18)
p_sub1.font.color.rgb = TEAL_PRIMARY

# One-sentence project description
tb_d1 = s1.shapes.add_textbox(Inches(0.9), Inches(2.65), Inches(7.5), Inches(0.8))
tf_d1 = tb_d1.text_frame
tf_d1.word_wrap = True
p_d1 = tf_d1.paragraphs[0]
p_d1.text = "An intelligent backend system that dynamically transforms a student's uploaded study materials into concise summaries, active-recall flashcards, self-assessment quizzes, and personalized revision schedules using Google Gemini."
p_d1.font.name = FONT_BODY
p_d1.font.size = Pt(13.5)
p_d1.font.color.rgb = CHARCOAL_MID
p_d1.line_spacing = 1.25

# Abstract AI Connected-Node Graphic on Right
nodes = [
    (10.6, 2.2, 0.5, TEAL_PRIMARY, True),
    (9.4, 1.5, 0.35, TEAL_PRIMARY, False),
    (11.8, 1.6, 0.32, GREY_MUTED, False),
    (9.2, 3.0, 0.38, GREY_MUTED, False),
    (11.9, 2.9, 0.4, TEAL_PRIMARY, False),
    (10.7, 3.8, 0.34, TEAL_PRIMARY, False),
    (12.3, 3.9, 0.28, GREY_MUTED, False),
]
connections = [
    (0, 1), (0, 2), (0, 3), (0, 4), (0, 5),
    (1, 3), (2, 4), (4, 6), (5, 6)
]
for i1, i2 in connections:
    x1, y1, s1_w, _, _ = nodes[i1]
    x2, y2, s2_w, _, _ = nodes[i2]
    cx1, cy1 = x1 + s1_w/2, y1 + s1_w/2
    cx2, cy2 = x2 + s2_w/2, y2 + s2_w/2
    conn = s1.shapes.add_connector(1, Inches(cx1), Inches(cy1), Inches(cx2), Inches(cy2))
    conn.line.color.rgb = BORDER_LIGHT
    conn.line.width = Pt(1.2)

for x, y, size, col, is_center in nodes:
    node_circle = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(size), Inches(size))
    node_circle.fill.solid()
    node_circle.fill.fore_color.rgb = col if is_center else BG_WHITE
    node_circle.line.color.rgb = col
    node_circle.line.width = Pt(2 if is_center else 1.5)

# Horizontal Divider Line
hdiv1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.9), Inches(3.9), Inches(11.5), Inches(0.015))
hdiv1.fill.solid()
hdiv1.fill.fore_color.rgb = BORDER_LIGHT
hdiv1.line.fill.background()

# Bottom Two Balanced Columns: Tech Stack (Left) & Team (Right)
add_clean_card(s1, 0.9, 4.25, 5.6, 2.5, bg_color=BG_PAGE, border_color=BORDER_LIGHT)
tb_ts = s1.shapes.add_textbox(Inches(1.15), Inches(4.45), Inches(5.1), Inches(2.1))
tf_ts = tb_ts.text_frame
tf_ts.word_wrap = True
p_tsh = tf_ts.paragraphs[0]
p_tsh.text = "CORE TECHNOLOGY STACK"
p_tsh.font.name = FONT_HEADING
p_tsh.font.size = Pt(11)
p_tsh.font.bold = True
p_tsh.font.color.rgb = TEAL_PRIMARY

p_tsb = tf_ts.add_paragraph()
p_tsb.space_before = Pt(8)
p_tsb.text = "Node.js  •  Express.js  •  MongoDB  •  Mongoose ODM\nJWT & RBAC Security  •  Google Gemini 2.5 Flash API"
p_tsb.font.name = FONT_BODY
p_tsb.font.size = Pt(14)
p_tsb.font.bold = True
p_tsb.font.color.rgb = CHARCOAL_DARK
p_tsb.line_spacing = 1.35

p_tss = tf_ts.add_paragraph()
p_tss.space_before = Pt(8)
p_tss.text = "RESTful backend architecture enforcing strict query-level isolation."
p_tss.font.name = FONT_BODY
p_tss.font.size = Pt(11.5)
p_tss.font.color.rgb = GREY_MUTED

# Right: Team Information
add_clean_card(s1, 6.8, 4.25, 5.6, 2.5, bg_color=BG_PAGE, border_color=BORDER_LIGHT)
tb_tm1 = s1.shapes.add_textbox(Inches(7.05), Inches(4.45), Inches(5.1), Inches(2.1))
tf_tm1 = tb_tm1.text_frame
tf_tm1.word_wrap = True
p_tmh = tf_tm1.paragraphs[0]
p_tmh.text = "PROJECT TEAM"
p_tmh.font.name = FONT_HEADING
p_tmh.font.size = Pt(11)
p_tmh.font.bold = True
p_tmh.font.color.rgb = TEAL_PRIMARY

p_tl = tf_tm1.add_paragraph()
p_tl.space_before = Pt(6)
r_tll = p_tl.add_run()
r_tll.text = "Team Leader: "
r_tll.font.name = FONT_BODY
r_tll.font.size = Pt(13)
r_tll.font.bold = True
r_tll.font.color.rgb = CHARCOAL_DARK
r_tln = p_tl.add_run()
r_tln.text = "Isaac Shiloah Herbert V"
r_tln.font.name = FONT_BODY
r_tln.font.size = Pt(13)
r_tln.font.bold = True
r_tln.font.color.rgb = TEAL_PRIMARY

p_tmb = tf_tm1.add_paragraph()
p_tmb.space_before = Pt(6)
r_tml = p_tmb.add_run()
r_tml.text = "Team Members: "
r_tml.font.name = FONT_BODY
r_tml.font.size = Pt(12)
r_tml.font.bold = True
r_tml.font.color.rgb = CHARCOAL_DARK
r_tmn = p_tmb.add_run()
r_tmn.text = "\nMohammed Irfan A, Bhuvaneswari, Bensam Gilbert E, Joel Edwin P"
r_tmn.font.name = FONT_BODY
r_tmn.font.size = Pt(12)
r_tmn.font.color.rgb = CHARCOAL_MID
p_tmb.line_spacing = 1.25

# ==============================================================================
# SLIDE 2: THE PROBLEM AND OUR SOLUTION
# ==============================================================================
s2 = prs.slides.add_slide(blank_layout)
add_header(s2, "Smarter Study Preparation, One Platform", "THE PROBLEM AND OUR SOLUTION", 
           "Transforming fragmented, labor-intensive preparation into a streamlined AI-powered revision workflow.")

# Left Card: The Problem
add_clean_card(s2, 0.8, 1.95, 5.65, 5.05, bg_color=BG_WHITE, border_color=BORDER_LIGHT)
p_bar = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.95), Inches(5.65), Inches(0.06))
p_bar.fill.solid()
p_bar.fill.fore_color.rgb = RGBColor(239, 68, 68)
p_bar.line.fill.background()

tb_p = s2.shapes.add_textbox(Inches(1.15), Inches(2.2), Inches(5.0), Inches(4.6))
tf_p = tb_p.text_frame
tf_p.word_wrap = True
p_ph = tf_p.paragraphs[0]
p_ph.text = "THE PROBLEM"
p_ph.font.name = FONT_HEADING
p_ph.font.size = Pt(12)
p_ph.font.bold = True
p_ph.font.color.rgb = RGBColor(220, 38, 38)

p_ps = tf_p.add_paragraph()
p_ps.space_before = Pt(4)
p_ps.text = "Academic preparation friction & fragmented resources"
p_ps.font.name = FONT_HEADING
p_ps.font.size = Pt(16)
p_ps.font.bold = True
p_ps.font.color.rgb = CHARCOAL_DARK

prob_items = [
    ("Scattered Study Materials", "Notes, PDFs, textbooks, and articles are spread across different disconnected sources, making organized revision difficult."),
    ("Manual Preparation Bottlenecks", "Synthesizing lengthy materials into notes, summaries, and self-test questions consumes excessive study hours."),
    ("Lack of Structured Workflows", "Students lack adaptive timelines matched to exam dates and convenient self-assessment tools to evaluate readiness.")
]
for title, desc in prob_items:
    p_i = tf_p.add_paragraph()
    p_i.space_before = Pt(14)
    p_i.line_spacing = 1.22
    r_dot = p_i.add_run()
    r_dot.text = "• "
    r_dot.font.name = FONT_BODY
    r_dot.font.size = Pt(14)
    r_dot.font.bold = True
    r_dot.font.color.rgb = RGBColor(220, 38, 38)
    
    r_t = p_i.add_run()
    r_t.text = f"{title}: "
    r_t.font.name = FONT_BODY
    r_t.font.size = Pt(13)
    r_t.font.bold = True
    r_t.font.color.rgb = CHARCOAL_DARK
    
    r_d = p_i.add_run()
    r_d.text = desc
    r_d.font.name = FONT_BODY
    r_d.font.size = Pt(12.5)
    r_d.font.color.rgb = CHARCOAL_MID

# Center Arrow / Transition Pill
arrow_box = s2.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(6.55), Inches(4.3), Inches(0.4), Inches(0.35))
arrow_box.fill.solid()
arrow_box.fill.fore_color.rgb = TEAL_PRIMARY
arrow_box.line.fill.background()

# Right Card: The Solution
add_clean_card(s2, 7.05, 1.95, 5.5, 5.05, bg_color=TEAL_LIGHT, border_color=TEAL_BORDER)
s_bar = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.05), Inches(1.95), Inches(5.5), Inches(0.06))
s_bar.fill.solid()
s_bar.fill.fore_color.rgb = TEAL_PRIMARY
s_bar.line.fill.background()

tb_s = s2.shapes.add_textbox(Inches(7.35), Inches(2.2), Inches(4.9), Inches(4.6))
tf_s = tb_s.text_frame
tf_s.word_wrap = True
p_sh = tf_s.paragraphs[0]
p_sh.text = "OUR SOLUTION"
p_sh.font.name = FONT_HEADING
p_sh.font.size = Pt(12)
p_sh.font.bold = True
p_sh.font.color.rgb = TEAL_PRIMARY

p_ss = tf_s.add_paragraph()
p_ss.space_before = Pt(4)
p_ss.text = "Centralized, material-grounded AI learning assistant"
p_ss.font.name = FONT_HEADING
p_ss.font.size = Pt(16)
p_ss.font.bold = True
p_ss.font.color.rgb = CHARCOAL_DARK

sol_items = [
    ("Upload & Manage Materials", "Upload raw study notes and documents (.txt, .md, .pdf) directly into a secure, searchable repository."),
    ("AI Summaries, Flashcards & Quizzes", "Dynamically generate concise summaries, active-recall flashcards, and 4-option MCQs directly from your content."),
    ("Personalized Plans & Direct Q&A", "Generate tailored day-by-day revision schedules and ask contextual questions answered directly from your material."),
    ("Save Resources for Later Revision", "Every generated summary, quiz, and study plan is automatically saved under your account for anytime retrieval.")
]
for title, desc in sol_items:
    p_i = tf_s.add_paragraph()
    p_i.space_before = Pt(11)
    p_i.line_spacing = 1.2
    r_dot = p_i.add_run()
    r_dot.text = "✓ "
    r_dot.font.name = FONT_BODY
    r_dot.font.size = Pt(13)
    r_dot.font.bold = True
    r_dot.font.color.rgb = TEAL_PRIMARY
    
    r_t = p_i.add_run()
    r_t.text = f"{title}: "
    r_t.font.name = FONT_BODY
    r_t.font.size = Pt(12.5)
    r_t.font.bold = True
    r_t.font.color.rgb = CHARCOAL_DARK
    
    r_d = p_i.add_run()
    r_d.text = desc
    r_d.font.name = FONT_BODY
    r_d.font.size = Pt(12)
    r_d.font.color.rgb = CHARCOAL_MID

# ==============================================================================
# SLIDE 3: SYSTEM ARCHITECTURE ("How Study Buddy AI Works")
# ==============================================================================
s3 = prs.slides.add_slide(blank_layout)
add_header(s3, "How Study Buddy AI Works", "SYSTEM ARCHITECTURE", 
           "End-to-end request lifecycle: client requests, JWT verification, controllers, and dual backend services.")

def draw_arch_box(slide, x, y, w, h, title, subtitle=None, details=None, is_teal=False, is_accent=False):
    bg = TEAL_LIGHT if is_teal else (RGBColor(241, 245, 249) if is_accent else BG_WHITE)
    border = TEAL_PRIMARY if is_teal else (TEAL_BORDER if is_accent else BORDER_LIGHT)
    card = add_clean_card(slide, x, y, w, h, bg_color=bg, border_color=border, border_width=1.5 if is_teal else 1)
    
    tb = slide.shapes.add_textbox(Inches(x + 0.12), Inches(y + 0.08), Inches(w - 0.24), Inches(h - 0.16))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.06)
    tf.margin_right = Inches(0.06)
    tf.margin_top = Inches(0.04)
    tf.margin_bottom = Inches(0.04)
    
    p0 = tf.paragraphs[0]
    p0.text = title
    p0.font.name = FONT_HEADING
    p0.font.size = Pt(12.5)
    p0.font.bold = True
    p0.font.color.rgb = TEAL_PRIMARY if is_teal else CHARCOAL_DARK
    
    if subtitle:
        p1 = tf.add_paragraph()
        p1.space_before = Pt(2)
        p1.text = subtitle
        p1.font.name = FONT_BODY
        p1.font.size = Pt(10.5)
        p1.font.bold = True
        p1.font.color.rgb = GREY_MUTED
        
    if details:
        for d in details:
            pd = tf.add_paragraph()
            pd.space_before = Pt(2.5)
            pd.line_spacing = 1.12
            pd.text = f"• {d}"
            pd.font.name = FONT_BODY
            pd.font.size = Pt(9.5)
            pd.font.color.rgb = CHARCOAL_MID

# Main Horizontal Flow (Left side)
# Box 1: Student / Admin
draw_arch_box(s3, 0.8, 1.95, 2.3, 3.65, "Student / Admin", "Client Layer", 
              ["Postman / Web client", "HTTP requests + Bearer token", "Upload notes & trigger AI", "Access saved resources"])

# Arrow 1: Client -> REST API
a1 = s3.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(3.20), Inches(3.65), Inches(0.35), Inches(0.25))
a1.fill.solid()
a1.fill.fore_color.rgb = TEAL_PRIMARY
a1.line.fill.background()

# Box 2: REST API (Node.js + Express.js)
draw_arch_box(s3, 3.65, 1.95, 2.5, 3.65, "REST API Gateway", "Node.js + Express.js",
              ["Central HTTP entry point", "Helmet security headers", "CORS & Morgan logging", "Two-tier rate limiting", "Express 5 NoSQL guard"])

# Arrow 2: API -> Auth
a2 = s3.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(6.25), Inches(3.65), Inches(0.35), Inches(0.25))
a2.fill.solid()
a2.fill.fore_color.rgb = TEAL_PRIMARY
a2.line.fill.background()

# Box 3: Authentication & Controllers
draw_arch_box(s3, 6.7, 1.95, 2.6, 3.65, "Auth & Controllers", "JWT + RBAC Middleware",
              ["Verifies 7-day Bearer tokens", "Role check (Student / Admin)", "Scoped queries { userId }", "Input validation (validator)", "Dispatches to services"])

# Branch Arrows to Right (Gemini & MongoDB)
a3a = s3.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(9.40), Inches(2.70), Inches(0.4), Inches(0.22))
a3a.fill.solid()
a3a.fill.fore_color.rgb = TEAL_PRIMARY
a3a.line.fill.background()

# Box 4A: Google Gemini 2.5 Flash API (Top Right)
draw_arch_box(s3, 9.9, 1.95, 2.65, 1.72, "Google Gemini 2.5 Flash", "@google/genai Service",
              ["Material-grounded prompts", "Summaries & Study Plans (prose)", "Flashcards & Quizzes (JSON)"], is_teal=True)

# Arrow 3B: Controllers -> MongoDB
a3b = s3.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(9.40), Inches(4.65), Inches(0.4), Inches(0.22))
a3b.fill.solid()
a3b.fill.fore_color.rgb = TEAL_PRIMARY
a3b.line.fill.background()

# Box 4B: MongoDB + Mongoose (Bottom Right)
draw_arch_box(s3, 9.9, 3.88, 2.65, 1.72, "MongoDB & Mongoose", "Data Persistence Layer",
              ["User accounts & password hashes", "Uploaded notes & metadata", "Saves summaries, cards & quizzes", "Enables retrieval for later revision"], is_accent=True)

# Bottom Horizontal Summary Strip: Resource Retrieval Flow
bot_bar = add_clean_card(s3, 0.8, 5.82, 11.75, 1.10, bg_color=BG_WHITE, border_color=BORDER_LIGHT)
tb_bb = s3.shapes.add_textbox(Inches(1.0), Inches(5.86), Inches(11.35), Inches(0.98))
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
r_bbb.text = "Authorized students can re-access, search, filter, and review all their generated resources at any time. When an AI feature executes, its structured response is persisted directly to MongoDB associated with the user's account—guaranteeing that study assets are preserved for exam revision rather than lost after generation."
r_bbb.font.name = FONT_BODY
r_bbb.font.size = Pt(11)
r_bbb.font.color.rgb = CHARCOAL_MID
p_bb.line_spacing = 1.25

# ==============================================================================
# SLIDE 4: CORE FEATURES AND SECURITY ("Features Built for Focused Learning")
# ==============================================================================
s4 = prs.slides.add_slide(blank_layout)
add_header(s4, "Features Built for Focused Learning", "CORE FEATURES AND SECURITY", 
           "High-yield academic study utilities paired with multi-layer defense-in-depth security.")

# Left Section: Core Features (6 Clean Feature Cards in a 2x3 Grid)
feat_box = add_clean_card(s4, 0.8, 1.95, 7.5, 5.05, bg_color=BG_WHITE, border_color=BORDER_LIGHT)

tb_fh = s4.shapes.add_textbox(Inches(1.05), Inches(2.1), Inches(7.0), Inches(0.4))
tf_fh = tb_fh.text_frame
p_fh = tf_fh.paragraphs[0]
p_fh.text = "CORE PLATFORM FEATURES"
p_fh.font.name = FONT_HEADING
p_fh.font.size = Pt(12)
p_fh.font.bold = True
p_fh.font.color.rgb = TEAL_PRIMARY

features = [
    ("Secure Authentication", "User registration, bcrypt password hashing, and 7-day stateless JWT login with profile management."),
    ("Material Management", "Upload study notes as text or files (.txt, .md, .pdf), with full-text search, subject filtering, and downloads."),
    ("AI-Generated Summaries", "High-yield conceptual summaries synthesized from uploaded notes for accelerated exam revision."),
    ("Active-Recall Flashcards", "Structured question-and-answer pairs created via guaranteed JSON output for spaced repetition."),
    ("Multiple-Choice Quizzes", "Objective 4-option quizzes generated with verified answer keys to evaluate subject comprehension."),
    ("Personalized Plans & Q&A", "Custom day-by-day study schedules tailored to exam timelines, plus direct contextual study Q&A.")
]

for idx, (ftitle, fdesc) in enumerate(features):
    col = idx % 2
    row = idx // 2
    cx = 1.05 + col * 3.65
    cy = 2.55 + row * 1.35
    
    dot = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx), Inches(cy + 0.05), Inches(0.16), Inches(0.16))
    dot.fill.solid()
    dot.fill.fore_color.rgb = TEAL_PRIMARY
    dot.line.fill.background()
    
    tb_item = s4.shapes.add_textbox(Inches(cx + 0.28), Inches(cy), Inches(3.25), Inches(1.2))
    tf_item = tb_item.text_frame
    tf_item.word_wrap = True
    p_it = tf_item.paragraphs[0]
    p_it.text = ftitle
    p_it.font.name = FONT_HEADING
    p_it.font.size = Pt(13)
    p_it.font.bold = True
    p_it.font.color.rgb = CHARCOAL_DARK
    
    p_id = tf_item.add_paragraph()
    p_id.space_before = Pt(3)
    p_id.line_spacing = 1.15
    p_id.text = fdesc
    p_id.font.name = FONT_BODY
    p_id.font.size = Pt(11)
    p_id.font.color.rgb = CHARCOAL_MID

# Extra note at bottom of features box
tb_fn = s4.shapes.add_textbox(Inches(1.05), Inches(6.5), Inches(7.0), Inches(0.4))
tf_fn = tb_fn.text_frame
p_fn = tf_fn.paragraphs[0]
p_fn.text = "✦ Also includes administrator endpoints for user management and system-wide AI usage analytics."
p_fn.font.name = FONT_BODY
p_fn.font.size = Pt(10.5)
p_fn.font.bold = True
p_fn.font.color.rgb = GREY_MUTED

# Right Section: Security Architecture (Compact Security Card)
sec_box = add_clean_card(s4, 8.55, 1.95, 4.0, 5.05, bg_color=TEAL_LIGHT, border_color=TEAL_BORDER)

tb_sh = s4.shapes.add_textbox(Inches(8.8), Inches(2.1), Inches(3.5), Inches(0.4))
tf_sh = tb_sh.text_frame
p_sh = tf_sh.paragraphs[0]
p_sh.text = "SECURITY & DATA PROTECTION"
p_sh.font.name = FONT_HEADING
p_sh.font.size = Pt(12)
p_sh.font.bold = True
p_sh.font.color.rgb = TEAL_PRIMARY

sec_items = [
    ("JWT Authentication", "Stateless Bearer tokens with 7-day expiration protect private endpoints."),
    ("Bcrypt Password Hashing", "Passwords salted with 10 rounds; plaintext never reaches the database."),
    ("Role-Based Access Control", "Server-enforced Student vs. Admin privileges; role escalation blocked."),
    ("User-Scoped Data Access", "Queries enforce { userId: req.user._id }; prevents horizontal data access."),
    ("Input Guard & Rate Limiting", "Custom Express 5 NoSQL-injection guard plus two-tier rate limiting on AI quotas.")
]

tb_sb = s4.shapes.add_textbox(Inches(8.8), Inches(2.55), Inches(3.5), Inches(4.2))
tf_sb = tb_sb.text_frame
tf_sb.word_wrap = True
for idx, (stitle, sdesc) in enumerate(sec_items):
    p_s = tf_sb.paragraphs[0] if idx == 0 else tf_sb.add_paragraph()
    if idx > 0:
        p_s.space_before = Pt(11)
    p_s.line_spacing = 1.18
    
    r_check = p_s.add_run()
    r_check.text = "• "
    r_check.font.name = FONT_BODY
    r_check.font.size = Pt(12)
    r_check.font.bold = True
    r_check.font.color.rgb = TEAL_PRIMARY
    
    r_st = p_s.add_run()
    r_st.text = f"{stitle}: "
    r_st.font.name = FONT_BODY
    r_st.font.size = Pt(12)
    r_st.font.bold = True
    r_st.font.color.rgb = CHARCOAL_DARK
    
    r_sd = p_s.add_run()
    r_sd.text = sdesc
    r_sd.font.name = FONT_BODY
    r_sd.font.size = Pt(11)
    r_sd.font.color.rgb = CHARCOAL_MID

# ==============================================================================
# SLIDE 5: IMPLEMENTATION, VALIDATION AND PROJECT LINKS
# ==============================================================================
s5 = prs.slides.add_slide(blank_layout)
add_header(s5, "Implementation and Project Demonstration", "IMPLEMENTATION, VALIDATION & LINKS", 
           "Verified backend API endpoints, functional Postman validation, and project resources.")

# 3 Balanced Screenshots
shots = [
    ("User Authentication (201 Created)", "postman_auth.png", "Registration with bcrypt password encryption & JWT token issuance."),
    ("Material Upload (201 Created)", "postman_upload.png", "Multipart form-data upload (.txt) securely parsed and saved to MongoDB."),
    ("AI Flashcards (200 OK)", "postman_flashcards.png", "Google Gemini 2.5 Flash generating structured active-recall Q&A arrays.")
]

card_w = 3.65
card_h = 3.30
y_top = 1.95

for idx, (caption, filename, note) in enumerate(shots):
    x_pos = 0.8 + idx * 4.05
    add_clean_card(s5, x_pos, y_top, card_w, card_h, bg_color=BG_WHITE, border_color=BORDER_LIGHT)
    
    # Title
    tb_cap = s5.shapes.add_textbox(Inches(x_pos + 0.1), Inches(y_top + 0.1), Inches(card_w - 0.2), Inches(0.35))
    p_c = tb_cap.text_frame.paragraphs[0]
    p_c.text = caption
    p_c.font.name = FONT_HEADING
    p_c.font.size = Pt(11)
    p_c.font.bold = True
    p_c.font.color.rgb = TEAL_PRIMARY
    
    # Cropped Image
    img_path = os.path.join(CROPPED_DIR, filename)
    if os.path.exists(img_path):
        s5.shapes.add_picture(img_path, Inches(x_pos + 0.12), Inches(y_top + 0.45), Inches(card_w - 0.24), Inches(2.10))
        
    # Note / Annotation below image
    tb_n = s5.shapes.add_textbox(Inches(x_pos + 0.1), Inches(y_top + 2.60), Inches(card_w - 0.2), Inches(0.6))
    tf_n = tb_n.text_frame
    tf_n.word_wrap = True
    p_n = tf_n.paragraphs[0]
    p_n.text = note
    p_n.font.name = FONT_BODY
    p_n.font.size = Pt(10)
    p_n.font.color.rgb = GREY_MUTED
    p_n.line_spacing = 1.15

# Testing Statement Box
stmt_box = add_clean_card(s5, 0.8, 5.40, 11.75, 0.58, bg_color=BG_PAGE, border_color=BORDER_LIGHT)
tb_stmt = s5.shapes.add_textbox(Inches(0.95), Inches(5.42), Inches(11.45), Inches(0.52))
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
r_st2.text = "The backend API was comprehensively tested using Postman, asserting route protection, role authorization, and Gemini AI schema validation."
r_st2.font.name = FONT_BODY
r_st2.font.size = Pt(10.5)
r_st2.font.color.rgb = CHARCOAL_DARK

# Project Links & Closing Statement Banner
bot_box = add_clean_card(s5, 0.8, 6.12, 11.75, 1.05, bg_color=TEAL_LIGHT, border_color=TEAL_BORDER)

tb_links = s5.shapes.add_textbox(Inches(1.0), Inches(6.16), Inches(11.35), Inches(0.95))
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
print("Successfully generated redesigned 5-slide PPT to:", OUT_PPTX_D)

import shutil
try:
    shutil.copyfile(OUT_PPTX_D, OUT_PPTX_DL)
    print("Successfully copied to Downloads:", OUT_PPTX_DL)
except Exception as e:
    print("Notice on Downloads copy:", e)
