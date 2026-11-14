import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
import win32com.client
import fitz # PyMuPDF
import shutil

MEDIA_DIR = r"C:\Users\acer\.gemini\antigravity\brain\3df468b6-62f3-4271-a7b4-01a98becf01a\scratch\media"
CROPPED_DIR = os.path.join(MEDIA_DIR, "cropped")
SCRATCH_DIR = r"C:\Users\acer\.gemini\antigravity\brain\3df468b6-62f3-4271-a7b4-01a98becf01a\scratch"

OUT_DIR_D = r"d:\AI-StudyBuddy"
os.makedirs(OUT_DIR_D, exist_ok=True)

# ------------------------------------------------------------------------------
# COLOR PALETTE (Minimalist Clean Tech Theme)
# ------------------------------------------------------------------------------
BG_PAGE       = RGBColor(248, 250, 252) # #F8FAFC
BG_WHITE      = RGBColor(255, 255, 255) # #FFFFFF
CHARCOAL_DARK = RGBColor(15, 23, 42)    # #0F172A
CHARCOAL_MID  = RGBColor(51, 65, 85)    # #334155
TEAL_DARK     = RGBColor(15, 118, 110)  # #0F766E
TEAL_LIGHT    = RGBColor(240, 253, 250) # #F0FDFA
TEAL_BORDER   = RGBColor(153, 246, 228) # #99F6E4
GREY_MUTED    = RGBColor(100, 116, 139) # #64748B
BORDER_LIGHT  = RGBColor(226, 232, 240) # #E2E8F0
RED_ALERT     = RGBColor(225, 29, 72)   # #E11D48
RED_LIGHT     = RGBColor(255, 241, 242) # #FFF1F2
RED_BORDER    = RGBColor(254, 205, 211) # #FECDD3

FONT_HEADING = "Calibri"
FONT_BODY    = "Calibri"

def set_slide_bg(slide, color=BG_PAGE):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_header(slide, title_text, category_text="PATRICIAN COLLEGE OF ARTS AND SCIENCE  |  TN SKILLS", subtitle_text=None, slide_num=None):
    set_slide_bg(slide, BG_PAGE)

    # Category Label at top left
    tb_cat = slide.shapes.add_textbox(Inches(0.8), Inches(0.38), Inches(9.5), Inches(0.28))
    tf_cat = tb_cat.text_frame
    tf_cat.word_wrap = True
    p_cat = tf_cat.paragraphs[0]
    p_cat.text = category_text.upper()
    p_cat.font.name = FONT_HEADING
    p_cat.font.size = Pt(10)
    p_cat.font.bold = True
    p_cat.font.color.rgb = TEAL_DARK

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

    # Header Divider Line
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

def draw_arch_card(slide, x, y, w, h, title, subtitle, bullets, is_teal=False, is_accent=False):
    bg = TEAL_LIGHT if is_teal else (RGBColor(241, 245, 249) if is_accent else BG_WHITE)
    border = TEAL_DARK if is_teal else (TEAL_BORDER if is_accent else BORDER_LIGHT)
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
    p0.font.color.rgb = TEAL_DARK if is_teal else CHARCOAL_DARK
    
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

# ------------------------------------------------------------------------------
# FUNCTION TO GENERATE COMPLETE 8-SLIDE DECK FOR A GIVEN TEAM SPEC
# ------------------------------------------------------------------------------
def generate_8_slide_deck(team_info, output_path):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # ==========================================================================
    # SLIDE 1: TITLE SLIDE (College Name, Team, Roll No 1st, Points format, bold)
    # ==========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s1, BG_WHITE)

    # Top College & Category Tag
    tb_c1 = s1.shapes.add_textbox(Inches(0.85), Inches(0.65), Inches(10.0), Inches(0.35))
    tf_c1 = tb_c1.text_frame
    p_c1 = tf_c1.paragraphs[0]
    p_c1.text = "PATRICIAN COLLEGE OF ARTS AND SCIENCE  |  DEPARTMENT OF COMPUTER APPLICATIONS"
    p_c1.font.name = FONT_HEADING
    p_c1.font.size = Pt(11)
    p_c1.font.bold = True
    p_c1.font.color.rgb = TEAL_DARK

    p_c2 = tf_c1.add_paragraph()
    p_c2.space_before = Pt(2)
    p_c2.text = "AI-AUGMENTED BACKEND DEVELOPMENT - A&S (NMU88)  •  TN SKILLS PROJECT RECORD"
    p_c2.font.name = FONT_BODY
    p_c2.font.size = Pt(9.5)
    p_c2.font.color.rgb = GREY_MUTED

    # Main Project Title
    tb_t1 = s1.shapes.add_textbox(Inches(0.85), Inches(1.20), Inches(8.5), Inches(0.8))
    tf_t1 = tb_t1.text_frame
    tf_t1.word_wrap = True
    p_t1 = tf_t1.paragraphs[0]
    p_t1.text = "Study Buddy AI"
    p_t1.font.name = FONT_HEADING
    p_t1.font.size = Pt(44)
    p_t1.font.bold = True
    p_t1.font.color.rgb = CHARCOAL_DARK

    p_sub1 = tf_t1.add_paragraph()
    p_sub1.space_before = Pt(2)
    p_sub1.text = "LearnMate — AI Assistant for Smarter Learning and Study Support"
    p_sub1.font.name = FONT_BODY
    p_sub1.font.size = Pt(17)
    p_sub1.font.bold = True
    p_sub1.font.color.rgb = TEAL_DARK

    # Executive Project Overview - Formatted as Bold Bullet Points (NOT paragraph, NO highlight)
    tb_points = s1.shapes.add_textbox(Inches(0.85), Inches(2.50), Inches(8.2), Inches(1.75))
    tf_pts = tb_points.text_frame
    tf_pts.word_wrap = True
    tf_pts.margin_left = Inches(0.0)
    tf_pts.margin_right = Inches(0.0)

    overview_points = [
        ("Project Mandate", "Autonomous backend ecosystem converting static courseware into interactive active-recall assets"),
        ("Core Intelligence", "Integrated Google Gemini 2.5 Flash LLM delivering deterministic JSON educational structures"),
        ("System Blueprint", "Decoupled RESTful MVC micro-architecture built on Node.js, Express.js, and MongoDB"),
        ("Security Guarantee", "Cryptographic token governance (JWT) paired with query-level tenant isolation")
    ]

    for idx, (label, val) in enumerate(overview_points):
        p_pt = tf_pts.paragraphs[0] if idx == 0 else tf_pts.add_paragraph()
        if idx > 0:
            p_pt.space_before = Pt(4)
        p_pt.line_spacing = 1.15
        
        r_dot = p_pt.add_run()
        r_dot.text = "• "
        r_dot.font.name = FONT_BODY
        r_dot.font.size = Pt(11.5)
        r_dot.font.bold = True
        r_dot.font.color.rgb = TEAL_DARK
        
        r_lbl = p_pt.add_run()
        r_lbl.text = f"{label}: "
        r_lbl.font.name = FONT_BODY
        r_lbl.font.size = Pt(11.5)
        r_lbl.font.bold = True
        r_lbl.font.color.rgb = CHARCOAL_DARK
        
        r_val = p_pt.add_run()
        r_val.text = val
        r_val.font.name = FONT_BODY
        r_val.font.size = Pt(11.5)
        r_val.font.color.rgb = CHARCOAL_MID

    # Abstract Connected AI Node Graphic on the Right
    nodes = [
        (10.8, 2.1, 0.55, TEAL_DARK, True),
        (9.5, 1.4, 0.38, TEAL_DARK, False),
        (12.0, 1.5, 0.34, GREY_MUTED, False),
        (9.3, 3.0, 0.40, GREY_MUTED, False),
        (12.1, 2.9, 0.42, TEAL_DARK, False),
        (10.9, 3.8, 0.36, TEAL_DARK, False),
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
    p_tsh.font.color.rgb = TEAL_DARK

    p_tsb = tf_ts.add_paragraph()
    p_tsb.space_before = Pt(5)
    p_tsb.text = "Node.js  •  Express.js  •  MongoDB  •  Mongoose ODM\nJWT & RBAC Security  •  Google Gemini 2.5 Flash API"
    p_tsb.font.name = FONT_BODY
    p_tsb.font.size = Pt(12.5)
    p_tsb.font.bold = True
    p_tsb.font.color.rgb = CHARCOAL_DARK
    p_tsb.line_spacing = 1.3

    p_tss = tf_ts.add_paragraph()
    p_tss.space_before = Pt(5)
    p_tss.text = "Production RESTful architecture enforcing query-level data isolation."
    p_tss.font.name = FONT_BODY
    p_tss.font.size = Pt(10.5)
    p_tss.font.color.rgb = GREY_MUTED

    # Right: Team Information (Roll No 1st, College Name)
    add_clean_card(s1, 6.75, 4.60, 5.7, 2.35, bg_color=BG_PAGE, border_color=BORDER_LIGHT)
    tb_tm1 = s1.shapes.add_textbox(Inches(7.00), Inches(4.72), Inches(5.2), Inches(2.10))
    tf_tm1 = tb_tm1.text_frame
    tf_tm1.word_wrap = True
    p_tmh = tf_tm1.paragraphs[0]
    p_tmh.text = f"{team_info['college_short']} — {team_info['team_header']}".upper()
    p_tmh.font.name = FONT_HEADING
    p_tmh.font.size = Pt(10)
    p_tmh.font.bold = True
    p_tmh.font.color.rgb = TEAL_DARK

    p_tl = tf_tm1.add_paragraph()
    p_tl.space_before = Pt(4)
    r_tll = p_tl.add_run()
    r_tll.text = "Team Leader: "
    r_tll.font.name = FONT_BODY
    r_tll.font.size = Pt(11.5)
    r_tll.font.bold = True
    r_tll.font.color.rgb = CHARCOAL_DARK

    r_tln = p_tl.add_run()
    r_tln.text = team_info['leader_text'] # e.g. "212404437 — TRAVIS FELIX PETERS"
    r_tln.font.name = FONT_BODY
    r_tln.font.size = Pt(11.5)
    r_tln.font.bold = True
    r_tln.font.color.rgb = TEAL_DARK

    p_tl_role = tf_tm1.add_paragraph()
    p_tl_role.space_before = Pt(1)
    r_tlr = p_tl_role.add_run()
    r_tlr.text = team_info['leader_role']
    r_tlr.font.name = FONT_BODY
    r_tlr.font.size = Pt(9.5)
    r_tlr.font.color.rgb = GREY_MUTED

    p_tmb = tf_tm1.add_paragraph()
    p_tmb.space_before = Pt(4)
    r_tml = p_tmb.add_run()
    r_tml.text = "Teammates (Roll No 1st):\n"
    r_tml.font.name = FONT_BODY
    r_tml.font.size = Pt(11)
    r_tml.font.bold = True
    r_tml.font.color.rgb = CHARCOAL_DARK

    r_tmn = p_tmb.add_run()
    r_tmn.text = team_info['teammates_text'] # formatted 2 names per line with roll no first
    r_tmn.font.name = FONT_BODY
    r_tmn.font.size = Pt(10)
    r_tmn.font.color.rgb = CHARCOAL_MID
    p_tmb.line_spacing = 1.2

    # ==========================================================================
    # SLIDE 2: STUDENT PROBLEM SPACE & REAL-WORLD MOTIVATION
    # ==========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "Academic Pain Points & Cognitive Overload", "PROBLEM SPACE & CASE STUDY MOTIVATION", 
               "Deconstructing the critical inefficiencies of traditional, unassisted examination revision.", 2)

    # Case Study Context Banner
    add_clean_card(s2, 0.8, 1.85, 11.733, 1.05, bg_color=BG_WHITE, border_color=BORDER_LIGHT)
    tb_cs = s2.shapes.add_textbox(Inches(1.0), Inches(1.92), Inches(11.3), Inches(0.9))
    tf_cs = tb_cs.text_frame
    tf_cs.word_wrap = True
    p_csh = tf_cs.paragraphs[0]
    r_ch1 = p_csh.add_run()
    r_ch1.text = "The Undergraduate Dilemma (Case Scenario): "
    r_ch1.font.name = FONT_HEADING
    r_ch1.font.size = Pt(12)
    r_ch1.font.bold = True
    r_ch1.font.color.rgb = RED_ALERT

    r_ch2 = p_csh.add_run()
    r_ch2.text = (
        "Learners tackling multi-chapter academic syllabi concurrently balance coursework, technical skill-building, and exams. "
        "Traditional revision forces students into clerical transcription—manually summarizing dense texts, extracting keywords, and "
        "hand-crafting practice questions. This clerical overhead displaces deep comprehension and prevents quantitative self-assessment."
    )
    r_ch2.font.name = FONT_BODY
    r_ch2.font.size = Pt(11)
    r_ch2.font.color.rgb = CHARCOAL_DARK
    p_csh.line_spacing = 1.2

    # 3 Rephrased Problem Cards
    prob_cards = [
        ("Resource Fragmentation & Chaos", 
         "DISORGANIZED LEARNING INPUTS",
         "• Academic files are fractured across disparate channels: university portals, loose PDFs, physical notes, and chat links.\n"
         "• Complete lack of unified indexing prevents fast keyword retrieval during tight pre-exam study windows.\n"
         "• Incompatible document formats inhibit rapid cross-topic comparisons and centralized review."),
        ("High-Friction Manual Synthesis", 
         "COGNITIVE DRAIN & TIME LOSS",
         "• Condensing 50+ page technical modules demands hours of passive typing with minimal long-term memory consolidation.\n"
         "• Formulating active-recall flashcard pairs requires substantial manual effort that exhausts study energy prematurely.\n"
         "• Repetitive manual drafting consumes 60%+ of available preparation time before active testing even begins."),
        ("Diagnostic Gap & Static Timelines", 
         "ABSENCE OF EVALUATION FRAMEWORK",
         "• Passive reading induces an 'illusion of competence', masking knowledge gaps until actual exam sittings.\n"
         "• Absence of objective, automated 4-choice quizzes leaves students unable to gauge subject readiness.\n"
         "• Fixed study habits fail to adapt to remaining exam dates, leaving complex topics unreviewed under pressure.")
    ]

    for idx, (title, tag, body) in enumerate(prob_cards):
        cx = 0.8 + idx * 4.0
        add_clean_card(s2, cx, 3.10, 3.733, 3.85, bg_color=RED_LIGHT, border_color=RED_BORDER)
        
        tb_p = s2.shapes.add_textbox(Inches(cx + 0.2), Inches(3.25), Inches(3.333), Inches(3.5))
        tf_p = tb_p.text_frame
        tf_p.word_wrap = True
        
        p_tag = tf_p.paragraphs[0]
        p_tag.text = tag
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

    # ==========================================================================
    # SLIDE 3: STRATEGIC SOLUTION & VALUE PROPOSITION
    # ==========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "The LearnMate Paradigm: Automated Academic Synthesis", "STRATEGIC SOLUTION & VALUE PROPOSITION", 
               "Transforming raw educational artifacts into structured, persistent revision assets via Generative AI.", 3)

    sol_cards = [
        ("1. Omni-Format Material Ingestion", 
         "CENTRALIZED ASSET REPOSITORY",
         "• Ingests diverse student documents (.txt, .md, .pdf) through streaming multipart form-data pipelines.\n"
         "• Enforces server-side validation: MIME-type assertions, file-size quotas, and character buffer extraction.\n"
         "• Indexes courseware under unique user IDs in MongoDB, establishing an accessible digital knowledge vault.",
         TEAL_LIGHT, TEAL_BORDER),
        ("2. Algorithmic Active-Recall Engine", 
         "COGNITIVE ENHANCEMENT SUITE",
         "• Dynamic Synthesizer: Distills multi-page chapters into high-density conceptual revision summaries.\n"
         "• Schema-Locked Flashcards: Generates structured Q&A arrays via strict JSON constraints for spaced repetition.\n"
         "• Objective Diagnostic MCQs: Automatically formats 4-option quizzes complete with verified answer indices.",
         BG_WHITE, BORDER_LIGHT),
        ("3. Milestone-Driven Plans & Grounded Q&A", 
         "ADAPTIVE PREPARATION ROADMAPS",
         "• Calibrated Study Calendars: Generates realistic day-by-day revision milestones mapped to remaining exam dates.\n"
         "• Material-Grounded Q&A: Ingests specific student inquiries and answers strictly using uploaded context.\n"
         "• Hallucination Elimination: System prompts restrict AI responses to factual assertions in student notes.",
         BG_WHITE, BORDER_LIGHT),
        ("4. Permanent Revision Persistence", 
         "CONTINUOUS RETENTION & REUSE",
         "• Automatically catalogs every summary, quiz, and study schedule into user-partitioned MongoDB collections.\n"
         "• Enables multi-session review and recursive testing without incurring redundant LLM API token costs.\n"
         "• Guarantees students maintain complete ownership of their academic assets throughout their degree.",
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
        p_tag.font.color.rgb = TEAL_DARK
        
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

    # ==========================================================================
    # SLIDE 4: DISTRIBUTED SERVER ARCHITECTURE & EXECUTION PIPELINE
    # ==========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "Distributed Server Architecture & Execution Pipeline", "SYSTEM ARCHITECTURE & EXECUTION FLOW", 
               "Decoupled MVC pipeline: client requests, perimeter defenses, business logic, and dual service backends.", 4)

    # Box 1: Client Interfaces
    draw_arch_card(s4, 0.8, 1.95, 2.3, 3.65, "Client Interface Layer", "Web & Postman Consumers",
                   ["Student & Admin consumers", "HTTPS calls + Bearer JWT tokens", "Transmits raw notes & PDFs", "Requests on-demand AI tasks", "Retrieves persisted assets"])

    # Arrow 1: Client -> Gateway
    a1 = s4.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(3.20), Inches(3.65), Inches(0.35), Inches(0.25))
    a1.fill.solid()
    a1.fill.fore_color.rgb = TEAL_DARK
    a1.line.fill.background()

    # Box 2: API Gateway
    draw_arch_card(s4, 3.65, 1.95, 2.5, 3.65, "REST API Gateway", "Express.js Routing Tier",
                   ["Centralized endpoint dispatcher", "Helmet HTTP security headers", "CORS & Morgan access logger", "Two-tier rate-limit throttling", "Multer streaming file parser", "Express 5 NoSQL-injection guard"])

    # Arrow 2: Gateway -> Controllers
    a2 = s4.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(6.25), Inches(3.65), Inches(0.35), Inches(0.25))
    a2.fill.solid()
    a2.fill.fore_color.rgb = TEAL_DARK
    a2.line.fill.background()

    # Box 3: Controllers & Security
    draw_arch_card(s4, 6.7, 1.95, 2.6, 3.65, "Controllers & Middleware", "Business Logic & Auth Layer",
                   ["Validates 7-day JWT Bearer tokens", "Role gate (Student vs. Admin)", "Enforces tenant isolation { userId }", "express-validator schema guards", "Centralized exception handler", "Orchestrates AI & DB services"])

    # Arrow 3A: Controllers -> Gemini
    a3a = s4.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(9.40), Inches(2.70), Inches(0.4), Inches(0.22))
    a3a.fill.solid()
    a3a.fill.fore_color.rgb = TEAL_DARK
    a3a.line.fill.background()

    # Box 4A: Gemini Engine
    draw_arch_card(s4, 9.9, 1.95, 2.65, 1.72, "Google Gemini 2.5 Flash", "@google/genai Service Layer",
                   ["Material-grounded prompt tuning", "Conceptual summaries (prose)", "Structured JSON mode for MCQs/Cards"], is_teal=True)

    # Arrow 3B: Controllers -> MongoDB
    a3b = s4.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(9.40), Inches(4.65), Inches(0.4), Inches(0.22))
    a3b.fill.solid()
    a3b.fill.fore_color.rgb = TEAL_DARK
    a3b.line.fill.background()

    # Box 4B: MongoDB Persistence
    draw_arch_card(s4, 9.9, 3.88, 2.65, 1.72, "MongoDB Atlas & Mongoose", "Data Persistence Tier",
                   ["Bcrypt credentials & user roles", "Uploaded raw notes & metadata", "Saves summaries, cards & quizzes", "Enables anytime exam retrieval"], is_accent=True)

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
    r_bbh.text = "Persistent Asset Access Pipeline: "
    r_bbh.font.name = FONT_HEADING
    r_bbh.font.size = Pt(11.5)
    r_bbh.font.bold = True
    r_bbh.font.color.rgb = TEAL_DARK

    r_bbb = p_bb.add_run()
    r_bbb.text = (
        "Authenticated learners can securely search, filter, and review stored revision aids across examination cycles. "
        "When an AI generation completes, its structured payload is permanently persisted to MongoDB bound to the active user's account "
        "({ userId: req.user._id })—ensuring learning assets remain accessible on demand without redundant API invocation."
    )
    r_bbb.font.name = FONT_BODY
    r_bbb.font.size = Pt(10.5)
    r_bbb.font.color.rgb = CHARCOAL_MID
    p_bb.line_spacing = 1.25

    # ==========================================================================
    # SLIDE 5: ENGINEERING STACK & TECHNICAL ECOSYSTEM
    # ==========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "Engineering Stack & Technical Ecosystem", "TECHNOLOGY STACK & PROGRAMS USED", 
               "Exact packages, runtime environments, AI SDKs, and developer tooling powering the platform.", 5)

    tech_quadrants = [
        ("Server Runtime & Web Routing", "CORE COMPUTING INFRASTRUCTURE", [
            ("Node.js (v18+ LTS)", "High-throughput asynchronous JavaScript runtime engine with non-blocking event-driven I/O."),
            ("Express.js (v5.0)", "Minimalist web application framework orchestrating RESTful endpoints, middleware chaining, and HTTP controllers."),
            ("Mongoose ODM (v8.0)", "Object Data Modeling library providing strict schema validation, type casting, and MongoDB query construction.")
        ]),
        ("Generative AI & Cognitive Tier", "LLM INFERENCE ENGINE", [
            ("@google/genai SDK", "Official Google client library providing direct, low-latency streaming integration with Gemini 2.5 Flash."),
            ("Gemini 2.5 Flash", "Advanced multimodal LLM delivering rapid token inference, strong context comprehension, and cost efficiency."),
            ("Strict JSON Mode", "Deterministic schema enforcement (`responseSchema`) guaranteeing valid JSON arrays for flashcards and quizzes.")
        ]),
        ("Perimeter Defense & Utilities", "IDENTITY & SECURITY SUITE", [
            ("bcryptjs (10 rounds)", "Adaptive cryptographic password hashing with 10 salt rounds to eliminate plaintext exposure."),
            ("jsonwebtoken (JWT)", "Stateless authentication issuing 7-day HMAC-SHA256 Bearer tokens carrying decoded user ID and role."),
            ("helmet & express-rate-limit", "Protective HTTP headers mitigation (XSS, clickjacking) plus two-tier rate limiting against API abuse.")
        ]),
        ("File Streaming & Verification", "INGESTION & QUALITY ASSURANCE", [
            ("Multer Middleware", "Streaming multipart/form-data handler buffering incoming course documents (.txt, .md, .pdf) securely."),
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
        p_cat.font.color.rgb = TEAL_DARK
        
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

    # ==========================================================================
    # SLIDE 6: DATA MODELING & MONGOOSE ENTITY ARCHITECTURE
    # ==========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "Data Modeling & Mongoose Entity Architecture", "DATABASE CONFIGURATION & SCHEMAS", 
               "Relational data modeling in MongoDB NoSQL via Mongoose schemas, compound indexes, and entity hooks.", 6)

    schemas_2x2 = [
        ("User Identity Model (User.js)", "AUTHENTICATION & ACCESS ROLES", [
            ("Core Attributes", "name (String), email (String, unique, indexed), password (bcrypt hash), role (enum: ['student', 'admin']), createdAt."),
            ("Pre-Save Hook", "Mongoose middleware triggers bcrypt.hash(password, 10) automatically before committing modified credentials."),
            ("Cryptographic Method", "matchPassword(inputPassword) executes timing-attack resistant comparisons during session login.")
        ], TEAL_LIGHT, TEAL_BORDER),
        
        ("Course Material Model (Material.js)", "LEARNING ASSET REPOSITORY", [
            ("Core Attributes", "userId (ObjectId ref: 'User'), title (String), content (Text buffer), fileType (String), fileSize (Number), uploadedAt."),
            ("Relational Role", "Primary root anchor for all AI features; maintains 1-to-Many cascade relationships with generated study outputs."),
            ("Compound Indexing", "Indexed on { userId: 1, uploadedAt: -1 } ensuring microsecond-level query resolution for student documents.")
        ], BG_WHITE, BORDER_LIGHT),
        
        ("Active-Recall Learning Models", "FLASHCARDS & QUIZ ENTITIES", [
            ("Flashcard Schema", "userId, materialId, cards: [{ question: String, answer: String }], cardCount (Number)."),
            ("Quiz Schema", "userId, materialId, questions: [{ question: String, options: [String], correctAnswer: Number, explanation: String }]."),
            ("Output Guarantee", "Strict JSON schema enforcement (`responseSchema`) ensures zero parsing failures during downstream ingestion.")
        ], BG_WHITE, BORDER_LIGHT),
        
        ("Synthesis & Planning Models", "SUMMARIES & REVISION TIMELINES", [
            ("Summary Schema", "userId (ref User), materialId (ref Material), summaryText (String), keyTakeaways ([String]), generatedAt."),
            ("Study Plan Schema", "userId, targetExam: String, examDate: Date, dailyPlan: [{ day: Number, topics: [String], hours: Number }]."),
            ("Multi-Tenant Isolation", "Queries strictly scoped to authenticated user ID, preventing horizontal cross-student data exposure.")
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
        p_c.font.color.rgb = TEAL_DARK
        
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

    # ==========================================================================
    # SLIDE 7: ENTERPRISE SECURITY & ACCESS GOVERNANCE
    # ==========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "Enterprise Security & Access Governance", "DEFENSE-IN-DEPTH SECURITY MODEL", 
               "Multi-layered protective controls shielding user privacy, application state, and API integrity.", 7)

    sec_pillars = [
        ("Bcrypt Password Salting", "CRYPTOGRAPHIC STORAGE",
         "• Passwords salted with 10 rounds via bcryptjs prior to database commit.\n"
         "• Plaintext credentials never stored in database or exposed in server logs.\n"
         "• Resistant to precomputed rainbow-table and dictionary-based attacks.",
         BG_WHITE),
        ("Stateless Bearer JWT Tokens", "IDENTITY GOVERNANCE",
         "• 7-day HMAC-SHA256 signed JSON Web Tokens issued upon successful login.\n"
         "• Bearer tokens cryptographically extracted and asserted by authMiddleware.\n"
         "• Decoded user identity attached to req.user for downstream handler access.",
         TEAL_LIGHT),
        ("Role-Based Access Control (RBAC)", "PRIVILEGE SEGREGATION",
         "• Strict architectural partition between Student and Administrator privileges.\n"
         "• Server-side middleware blocks role manipulation and privilege escalation.\n"
         "• Dedicated administrative APIs for platform user governance and usage telemetry.",
         BG_WHITE),
        ("Query-Level Tenant Isolation", "HORIZONTAL PRIVACY",
         "• Every database query strictly enforces { userId: req.user._id }.\n"
         "• Insecure Direct Object Reference (IDOR) vulnerabilities completely eradicated.\n"
         "• Guarantees students cannot view, modify, or delete peer materials or notes.",
         BG_WHITE),
        ("NoSQL Guard & Rate Limiting", "ABUSE PREVENTION & HYGIENE",
         "• Custom Express 5 input filter strips $ and . operators from payloads.\n"
         "• express-validator enforces strict field types, lengths, and schemas.\n"
         "• Two-tier rate limiting: 100 req/15min general + 15 req/15min on AI inference.",
         TEAL_LIGHT),
        ("Telemetry & Error Masking", "MONITORING & RESILIENCE",
         "• Morgan logger records HTTP request methods, response codes, and latency.\n"
         "• Centralized error-handling middleware normalizes response payloads.\n"
         "• Internal database error stacks masked in production to prevent data leaks.",
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
        p_cat.font.color.rgb = TEAL_DARK
        
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

    # ==========================================================================
    # SLIDE 8: PROJECT EXECUTION, API VERIFICATION, DELIVERABLES & THANK YOU
    # ==========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "Project Execution, API Verification & Concluding Queries", "IMPLEMENTATION, VALIDATION & FINAL DISCUSSION", 
               "Verified endpoint testing in Postman, live demonstration evidence, and concluding Q&A discussion.", 8)

    shots = [
        ("User Authentication (201 Created)", "postman_auth.png", "Bcrypt password hashing & 7-day JWT Bearer token issuance."),
        ("Material Upload (201 Created)", "postman_upload.png", "Multipart form-data upload (.txt) securely parsed and saved to MongoDB."),
        ("AI Flashcards (200 OK)", "postman_flashcards.png", "Google Gemini 2.5 Flash generating structured active-recall Q&A arrays.")
    ]

    card_w = 3.65
    card_h = 3.10
    y_top = 1.82

    for idx, (caption, filename, note) in enumerate(shots):
        x_pos = 0.8 + idx * 4.05
        add_clean_card(s8, x_pos, y_top, card_w, card_h, bg_color=BG_WHITE, border_color=BORDER_LIGHT)
        
        # Title
        tb_cap = s8.shapes.add_textbox(Inches(x_pos + 0.1), Inches(y_top + 0.08), Inches(card_w - 0.2), Inches(0.32))
        p_c = tb_cap.text_frame.paragraphs[0]
        p_c.text = caption
        p_c.font.name = FONT_HEADING
        p_c.font.size = Pt(10.5)
        p_c.font.bold = True
        p_c.font.color.rgb = TEAL_DARK
        
        # Cropped Image
        img_path = os.path.join(CROPPED_DIR, filename)
        if os.path.exists(img_path):
            s8.shapes.add_picture(img_path, Inches(x_pos + 0.12), Inches(y_top + 0.40), Inches(card_w - 0.24), Inches(2.05))
            
        # Note / Annotation below image
        tb_n = s8.shapes.add_textbox(Inches(x_pos + 0.1), Inches(y_top + 2.50), Inches(card_w - 0.2), Inches(0.55))
        tf_n = tb_n.text_frame
        tf_n.word_wrap = True
        p_n = tf_n.paragraphs[0]
        p_n.text = note
        p_n.font.name = FONT_BODY
        p_n.font.size = Pt(9.5)
        p_n.font.color.rgb = GREY_MUTED
        p_n.line_spacing = 1.15

    # Testing Statement Box
    stmt_box = add_clean_card(s8, 0.8, 5.08, 11.75, 0.48, bg_color=BG_PAGE, border_color=BORDER_LIGHT)
    tb_stmt = s8.shapes.add_textbox(Inches(0.95), Inches(5.10), Inches(11.45), Inches(0.44))
    tf_stmt = tb_stmt.text_frame
    tf_stmt.word_wrap = True
    p_stmt = tf_stmt.paragraphs[0]
    r_st1 = p_stmt.add_run()
    r_st1.text = "API Verification: "
    r_st1.font.name = FONT_HEADING
    r_st1.font.size = Pt(10.5)
    r_st1.font.bold = True
    r_st1.font.color.rgb = TEAL_DARK

    r_st2 = p_stmt.add_run()
    r_st2.text = (
        "The backend RESTful API was comprehensively tested using Postman, asserting route protection, role authorization, "
        "multipart file uploads, and Gemini AI JSON schema validation across all endpoints."
    )
    r_st2.font.name = FONT_BODY
    r_st2.font.size = Pt(10)
    r_st2.font.color.rgb = CHARCOAL_DARK

    # THANK YOU & QUERIES BANNER (Prominent Concluding Box at Bottom)
    bot_box = add_clean_card(s8, 0.8, 5.72, 11.75, 1.45, bg_color=TEAL_LIGHT, border_color=TEAL_BORDER)

    tb_links = s8.shapes.add_textbox(Inches(1.0), Inches(5.78), Inches(11.35), Inches(1.35))
    tf_l = tb_links.text_frame
    tf_l.word_wrap = True
    tf_l.margin_left = Inches(0.0)
    tf_l.margin_right = Inches(0.0)
    tf_l.margin_top = Inches(0.02)
    tf_l.margin_bottom = Inches(0.02)

    # Line 1: THANK YOU & QUERIES
    p_ty = tf_l.paragraphs[0]
    r_ty = p_ty.add_run()
    r_ty.text = "THANK YOU!   •   ANY QUESTIONS & QUERIES?"
    r_ty.font.name = FONT_HEADING
    r_ty.font.size = Pt(14)
    r_ty.font.bold = True
    r_ty.font.color.rgb = TEAL_DARK

    # Line 2: Invitation to discuss
    p_inv = tf_l.add_paragraph()
    p_inv.space_before = Pt(2)
    r_inv = p_inv.add_run()
    r_inv.text = "We welcome your questions, technical inquiries, and feedback regarding Study Buddy AI."
    r_inv.font.name = FONT_BODY
    r_inv.font.size = Pt(10.5)
    r_inv.font.color.rgb = CHARCOAL_MID

    # Line 3: Links row
    p_l1 = tf_l.add_paragraph()
    p_l1.space_before = Pt(3)
    r_gh_lbl = p_l1.add_run()
    r_gh_lbl.text = "GitHub Repository: "
    r_gh_lbl.font.name = FONT_HEADING
    r_gh_lbl.font.size = Pt(10.5)
    r_gh_lbl.font.bold = True
    r_gh_lbl.font.color.rgb = TEAL_DARK

    r_gh_val = p_l1.add_run()
    r_gh_val.text = f"{team_info['github_url']}    "
    r_gh_val.font.name = FONT_BODY
    r_gh_val.font.size = Pt(10.5)
    r_gh_val.font.bold = True
    r_gh_val.font.color.rgb = CHARCOAL_DARK

    r_yt_lbl = p_l1.add_run()
    r_yt_lbl.text = "|   Demo Video: "
    r_yt_lbl.font.name = FONT_HEADING
    r_yt_lbl.font.size = Pt(10.5)
    r_yt_lbl.font.bold = True
    r_yt_lbl.font.color.rgb = TEAL_DARK

    r_yt_val = p_l1.add_run()
    r_yt_val.text = "https://youtu.be/N5lNC06rn7w"
    r_yt_val.font.name = FONT_BODY
    r_yt_val.font.size = Pt(10.5)
    r_yt_val.font.bold = True
    r_yt_val.font.color.rgb = CHARCOAL_DARK

    # Line 4: Closing statement
    p_l2 = tf_l.add_paragraph()
    p_l2.space_before = Pt(3)
    r_close = p_l2.add_run()
    r_close.text = "“Study Buddy AI — Turning Study Materials into Structured Learning Resources.”"
    r_close.font.name = FONT_HEADING
    r_close.font.size = Pt(11)
    r_close.font.bold = True
    r_close.font.color.rgb = TEAL_DARK

    prs.save(output_path)
    print(f"Saved PPTX to: {output_path}")

# ==============================================================================
# TEAM CONFIGURATIONS
# ==============================================================================
team_travis = {
    "name": "Travis",
    "college_short": "Patrician College",
    "team_header": "Project Team",
    "leader_text": "212404437 — TRAVIS FELIX PETERS",
    "leader_role": "Role: Full Backend Architecture, Gemini AI Integration & Security Systems",
    "teammates_text": "212404404 — BALAJI K   •   212404429 — SAINESH V\n212404421 — MOHITH R   •   212404433 — SANKARA KRISHNAN R",
    "github_url": "https://github.com/travpete/AI-StudyBuddy"
}

team_ilaz = {
    "name": "Ilaz",
    "college_short": "Patrician College",
    "team_header": "Project Team",
    "leader_text": "ILAZ AHAMED S K",
    "leader_role": "Role: Full Backend Architecture, Gemini AI Integration & Security Systems",
    "teammates_text": "BILL CLINTON J   •   MANIKANDAN T\nDAVISH SHYAM S   •   SAN MATHEW M",
    "github_url": "https://github.com/ILAZAHAMEDSK/AI-STUDYBUDDY"
}

team_isaac = {
    "name": "Isaac",
    "college_short": "Patrician College",
    "team_header": "Project Team",
    "leader_text": "212404414 — ISAAC SHILOAH HERBERT V",
    "leader_role": "Role: Full Backend Architecture, Gemini AI Integration & Security Systems",
    "teammates_text": "212404420 — MOHAMMED IRFAN A   •   212404441 — BHUVANESWARI A\n212404405 — BENSAM GILBERT E   •   212404416 — JOEL EDWIN P",
    "github_url": "https://github.com/isaacincognito007-cloud/AI-StudyBuddy-"
}

# Generate all 3 team decks
pptx_travis = os.path.join(OUT_DIR_D, "Study_Buddy_AI_Presentation_Travis.pptx")
pptx_ilaz = os.path.join(OUT_DIR_D, "Study_Buddy_AI_Presentation_Ilaz.pptx")
pptx_isaac = os.path.join(OUT_DIR_D, "Study_Buddy_AI_Presentation_Isaac.pptx")

generate_8_slide_deck(team_travis, pptx_travis)
generate_8_slide_deck(team_ilaz, pptx_ilaz)
generate_8_slide_deck(team_isaac, pptx_isaac)

# Also copy Travis deck to the default locations (as Travis was the latest requested in prompt)
shutil.copyfile(pptx_travis, r"d:\Study_Buddy_AI_Presentation.pptx")
shutil.copyfile(pptx_travis, r"c:\Users\acer\Downloads\Study_Buddy_AI_Presentation.pptx")
shutil.copyfile(pptx_travis, os.path.join(OUT_DIR_D, "Study_Buddy_AI_Presentation.pptx"))

# Copy all 3 team PPTXs to Downloads for convenience
shutil.copyfile(pptx_travis, r"c:\Users\acer\Downloads\Study_Buddy_AI_Presentation_Travis.pptx")
shutil.copyfile(pptx_ilaz, r"c:\Users\acer\Downloads\Study_Buddy_AI_Presentation_Ilaz.pptx")
shutil.copyfile(pptx_isaac, r"c:\Users\acer\Downloads\Study_Buddy_AI_Presentation_Isaac.pptx")

# Also copy to D:\ root
shutil.copyfile(pptx_travis, r"d:\Study_Buddy_AI_Presentation_Travis.pptx")
shutil.copyfile(pptx_ilaz, r"d:\Study_Buddy_AI_Presentation_Ilaz.pptx")
shutil.copyfile(pptx_isaac, r"d:\Study_Buddy_AI_Presentation_Isaac.pptx")

print("All PPTX presentations generated and copied successfully!")

# Convert Travis deck to PDF
pdf_travis_scratch = os.path.join(SCRATCH_DIR, "clean8_preview.pdf")
powerpoint = win32com.client.Dispatch("PowerPoint.Application")
deck = powerpoint.Presentations.Open(pptx_travis, WithWindow=False)
deck.SaveAs(pdf_travis_scratch, 32)
deck.Close()

# Also convert Ilaz and Isaac to PDF
deck_ilaz = powerpoint.Presentations.Open(pptx_ilaz, WithWindow=False)
pdf_ilaz_scratch = os.path.join(SCRATCH_DIR, "clean8_ilaz_preview.pdf")
deck_ilaz.SaveAs(pdf_ilaz_scratch, 32)
deck_ilaz.Close()

deck_isaac = powerpoint.Presentations.Open(pptx_isaac, WithWindow=False)
pdf_isaac_scratch = os.path.join(SCRATCH_DIR, "clean8_isaac_preview.pdf")
deck_isaac.SaveAs(pdf_isaac_scratch, 32)
deck_isaac.Close()

powerpoint.Quit()

# Copy PDFs to destination
shutil.copyfile(pdf_travis_scratch, r"c:\Users\acer\Downloads\Study_Buddy_AI_Presentation.pdf")
shutil.copyfile(pdf_travis_scratch, r"d:\Study_Buddy_AI_Presentation.pdf")
shutil.copyfile(pdf_travis_scratch, r"c:\Users\acer\Downloads\Study_Buddy_AI_Presentation_Travis.pdf")
shutil.copyfile(pdf_travis_scratch, r"d:\Study_Buddy_AI_Presentation_Travis.pdf")

shutil.copyfile(pdf_ilaz_scratch, r"c:\Users\acer\Downloads\Study_Buddy_AI_Presentation_Ilaz.pdf")
shutil.copyfile(pdf_ilaz_scratch, r"d:\Study_Buddy_AI_Presentation_Ilaz.pdf")

shutil.copyfile(pdf_isaac_scratch, r"c:\Users\acer\Downloads\Study_Buddy_AI_Presentation_Isaac.pdf")
shutil.copyfile(pdf_isaac_scratch, r"d:\Study_Buddy_AI_Presentation_Isaac.pdf")

print("All PDFs converted and copied successfully!")

# Render Travis slide 1 and slide 8 to PNG for visual inspection
doc = fitz.open(pdf_travis_scratch)
print(f"Total slides in Travis PDF: {len(doc)}")
doc[0].get_pixmap(dpi=150).save(os.path.join(SCRATCH_DIR, "clean8_slide_1.png"))
doc[7].get_pixmap(dpi=150).save(os.path.join(SCRATCH_DIR, "clean8_slide_8.png"))
doc.close()
print("Saved slide 1 and slide 8 previews!")
