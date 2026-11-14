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
DOWNLOADS_DIR = r"c:\Users\acer\Downloads"
os.makedirs(OUT_DIR_D, exist_ok=True)

# ------------------------------------------------------------------------------
# THEME DEFINITIONS: 3 COMPLETELY DISTINCT THEMES
# ------------------------------------------------------------------------------

# THEME 1: ISAAC (Original Mint & Teal Clean Tech Theme) - NMU88
THEME_ISAAC = {
    "key": "isaac",
    "um_code": "NMU88",
    "font_heading": "Calibri",
    "font_body": "Calibri",
    "bg_page": RGBColor(248, 250, 252),        # #F8FAFC
    "bg_card": RGBColor(255, 255, 255),        # #FFFFFF
    "charcoal_dark": RGBColor(15, 23, 42),     # #0F172A
    "charcoal_mid": RGBColor(51, 65, 85),      # #334155
    "accent_primary": RGBColor(15, 118, 110),  # #0F766E (Emerald Teal)
    "accent_light": RGBColor(240, 253, 250),   # #F0FDFA
    "accent_border": RGBColor(153, 246, 228),  # #99F6E4
    "grey_muted": RGBColor(100, 116, 139),     # #64748B
    "border_light": RGBColor(226, 232, 240),   # #E2E8F0
    "prob_alert": RGBColor(225, 29, 72),       # #E11D48
    "prob_light": RGBColor(255, 241, 242),
    "prob_border": RGBColor(254, 205, 211),
    "card_radius": 0.02,
    "graphic_style": "constellation"
}

# THEME 2: TRAVIS (Modern Royal Indigo & Electric Cobalt Tech Theme) - NMU93
THEME_TRAVIS = {
    "key": "travis",
    "um_code": "NMU93",
    "font_heading": "Segoe UI",
    "font_body": "Segoe UI",
    "bg_page": RGBColor(245, 247, 252),        # #F5F7FC (Crisp Cool Slate)
    "bg_card": RGBColor(255, 255, 255),        # #FFFFFF
    "charcoal_dark": RGBColor(15, 23, 42),     # #0F172A
    "charcoal_mid": RGBColor(71, 85, 105),     # #475569
    "accent_primary": RGBColor(67, 56, 202),   # #4338CA (Royal Indigo 700)
    "accent_light": RGBColor(238, 242, 255),   # #EEF2FF (Indigo 50)
    "accent_border": RGBColor(199, 210, 254),  # #C7D2FE (Indigo 200)
    "grey_muted": RGBColor(100, 116, 139),     # #64748B
    "border_light": RGBColor(226, 232, 240),   # #E2E8F0
    "prob_alert": RGBColor(219, 39, 119),      # #DB2777 (Vibrant Pink/Magenta)
    "prob_light": RGBColor(253, 242, 248),     # #FDF2F8
    "prob_border": RGBColor(251, 207, 232),    # #FBCFE8
    "card_radius": 0.015,
    "graphic_style": "neural_flow"
}

# THEME 3: ILAZ (Bold Executive Crimson Ruby & Slate Theme) - NMU93
THEME_ILAZ = {
    "key": "ilaz",
    "um_code": "NMU93",
    "font_heading": "Arial",
    "font_body": "Arial",
    "bg_page": RGBColor(250, 250, 250),        # #FAFAFA (Clean Neutral Zinc)
    "bg_card": RGBColor(255, 255, 255),        # #FFFFFF
    "charcoal_dark": RGBColor(24, 24, 27),     # #18181B (Zinc 900)
    "charcoal_mid": RGBColor(63, 63, 70),      # #3F3F46 (Zinc 700)
    "accent_primary": RGBColor(159, 18, 57),   # #9F1239 (Deep Rose / Ruby 800)
    "accent_light": RGBColor(255, 241, 242),   # #FFF1F2 (Rose 50)
    "accent_border": RGBColor(254, 205, 211),  # #FECDD3 (Rose 200)
    "grey_muted": RGBColor(113, 113, 122),     # #71717A (Zinc 500)
    "border_light": RGBColor(228, 228, 231),   # #E4E4E7 (Zinc 200)
    "prob_alert": RGBColor(194, 65, 12),       # #C2410C (Amber-Orange)
    "prob_light": RGBColor(255, 247, 237),     # #FFF7ED
    "prob_border": RGBColor(254, 215, 170),    # #FED7AA
    "card_radius": 0.025,
    "graphic_style": "radial_core"
}

# ------------------------------------------------------------------------------
# HELPER FUNCTIONS WITH DYNAMIC THEME SUPPORT
# ------------------------------------------------------------------------------
def set_slide_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_clean_card(slide, left, top, width, height, bg_color, border_color, border_width=1, radius=0.02):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.adjustments[0] = radius
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    if border_color:
        card.line.color.rgb = border_color
        card.line.width = Pt(border_width)
    else:
        card.line.fill.background()
    return card

def add_header(slide, title_text, category_text, subtitle_text, slide_num, theme):
    set_slide_bg(slide, theme["bg_page"])

    # Category Label at top left
    tb_cat = slide.shapes.add_textbox(Inches(0.8), Inches(0.38), Inches(9.5), Inches(0.28))
    tf_cat = tb_cat.text_frame
    tf_cat.word_wrap = True
    p_cat = tf_cat.paragraphs[0]
    p_cat.text = category_text.upper()
    p_cat.font.name = theme["font_heading"]
    p_cat.font.size = Pt(10)
    p_cat.font.bold = True
    p_cat.font.color.rgb = theme["accent_primary"]

    # Slide Number at top right
    if slide_num:
        tb_num = slide.shapes.add_textbox(Inches(10.5), Inches(0.38), Inches(2.0), Inches(0.28))
        tf_num = tb_num.text_frame
        p_num = tf_num.paragraphs[0]
        p_num.alignment = PP_ALIGN.RIGHT
        p_num.text = f"SLIDE {slide_num} OF 8"
        p_num.font.name = theme["font_heading"]
        p_num.font.size = Pt(10)
        p_num.font.bold = True
        p_num.font.color.rgb = theme["grey_muted"]

    # Main Slide Title
    tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.62), Inches(11.7), Inches(0.58))
    tf_title = tb_title.text_frame
    tf_title.word_wrap = True
    p_title = tf_title.paragraphs[0]
    p_title.text = title_text
    p_title.font.name = theme["font_heading"]
    p_title.font.size = Pt(25)
    p_title.font.bold = True
    p_title.font.color.rgb = theme["charcoal_dark"]

    if subtitle_text:
        tb_sub = slide.shapes.add_textbox(Inches(0.8), Inches(1.22), Inches(11.7), Inches(0.35))
        tf_sub = tb_sub.text_frame
        p_sub = tf_sub.paragraphs[0]
        p_sub.text = subtitle_text
        p_sub.font.name = theme["font_body"]
        p_sub.font.size = Pt(12)
        p_sub.font.color.rgb = theme["grey_muted"]

    # Header Divider Line
    divider = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.62), Inches(11.733), Inches(0.015))
    divider.fill.solid()
    divider.fill.fore_color.rgb = theme["border_light"]
    divider.line.fill.background()

def draw_arch_card(slide, x, y, w, h, title, subtitle, bullets, theme, is_accent=False, is_highlight=False):
    bg = theme["accent_light"] if is_highlight else (theme["bg_page"] if is_accent else theme["bg_card"])
    border = theme["accent_primary"] if is_highlight else (theme["accent_border"] if is_accent else theme["border_light"])
    add_clean_card(slide, x, y, w, h, bg_color=bg, border_color=border, border_width=1.5 if is_highlight else 1, radius=theme["card_radius"])
    
    tb = slide.shapes.add_textbox(Inches(x + 0.12), Inches(y + 0.08), Inches(w - 0.24), Inches(h - 0.16))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.04)
    tf.margin_bottom = Inches(0.04)
    
    p0 = tf.paragraphs[0]
    p0.text = title
    p0.font.name = theme["font_heading"]
    p0.font.size = Pt(12)
    p0.font.bold = True
    p0.font.color.rgb = theme["accent_primary"] if is_highlight else theme["charcoal_dark"]
    
    p1 = tf.add_paragraph()
    p1.space_before = Pt(2)
    p1.text = subtitle
    p1.font.name = theme["font_body"]
    p1.font.size = Pt(10)
    p1.font.bold = True
    p1.font.color.rgb = theme["grey_muted"]
    
    for b in bullets:
        pb = tf.add_paragraph()
        pb.space_before = Pt(2.5)
        pb.line_spacing = 1.12
        pb.text = f"• {b}"
        pb.font.name = theme["font_body"]
        pb.font.size = Pt(9.5)
        pb.font.color.rgb = theme["charcoal_mid"]

def draw_title_graphic(slide, theme):
    style = theme.get("graphic_style", "constellation")
    
    if style == "constellation":
        # Isaac: Organic constellation layout
        nodes = [
            (10.8, 2.1, 0.55, theme["accent_primary"], True),
            (9.5, 1.4, 0.38, theme["accent_primary"], False),
            (12.0, 1.5, 0.34, theme["grey_muted"], False),
            (9.3, 3.0, 0.40, theme["grey_muted"], False),
            (12.1, 2.9, 0.42, theme["accent_primary"], False),
            (10.9, 3.8, 0.36, theme["accent_primary"], False),
            (12.4, 3.9, 0.30, theme["grey_muted"], False),
        ]
        connections = [
            (0, 1), (0, 2), (0, 3), (0, 4), (0, 5),
            (1, 3), (2, 4), (4, 6), (5, 6)
        ]
        for i1, i2 in connections:
            x1, y1, s1_w, _, _ = nodes[i1]
            x2, y2, s2_w, _, _ = nodes[i2]
            conn = slide.shapes.add_connector(1, Inches(x1 + s1_w/2), Inches(y1 + s1_w/2), Inches(x2 + s2_w/2), Inches(y2 + s2_w/2))
            conn.line.color.rgb = theme["border_light"]
            conn.line.width = Pt(1.2)
        for x, y, size, col, is_center in nodes:
            nc = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(size), Inches(size))
            nc.fill.solid()
            nc.fill.fore_color.rgb = col if is_center else theme["bg_card"]
            nc.line.color.rgb = col
            nc.line.width = Pt(2 if is_center else 1.5)

    elif style == "neural_flow":
        # Travis: Sleek 3-Layer Neural/Pipeline Architecture Graphic
        # Layer 1: Inputs (x=9.4), Layer 2: Core LLM Engine (x=10.9), Layer 3: Outputs (x=12.3)
        l1 = [(9.4, 1.5), (9.4, 2.4), (9.4, 3.3)]
        l2 = [(10.9, 1.95), (10.9, 2.85)]
        l3 = [(12.3, 1.5), (12.3, 2.4), (12.3, 3.3)]
        
        # Connect L1 to L2
        for p1 in l1:
            for p2 in l2:
                conn = slide.shapes.add_connector(1, Inches(p1[0]+0.2), Inches(p1[1]+0.2), Inches(p2[0]+0.25), Inches(p2[1]+0.25))
                conn.line.color.rgb = theme["accent_border"]
                conn.line.width = Pt(1.0)
        # Connect L2 to L3
        for p2 in l2:
            for p3 in l3:
                conn = slide.shapes.add_connector(1, Inches(p2[0]+0.25), Inches(p2[1]+0.25), Inches(p3[0]+0.2), Inches(p3[1]+0.2))
                conn.line.color.rgb = theme["accent_border"]
                conn.line.width = Pt(1.0)
                
        # Draw L1 nodes (Document Inputs)
        for x, y in l1:
            nc = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(0.42), Inches(0.42))
            nc.adjustments[0] = 0.2
            nc.fill.solid()
            nc.fill.fore_color.rgb = theme["bg_card"]
            nc.line.color.rgb = theme["accent_primary"]
            nc.line.width = Pt(1.5)
            
        # Draw L2 nodes (Gemini Processing Hubs)
        for x, y in l2:
            nc = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(0.55), Inches(0.55))
            nc.fill.solid()
            nc.fill.fore_color.rgb = theme["accent_primary"]
            nc.line.color.rgb = theme["accent_primary"]
            nc.line.width = Pt(2.0)
            
        # Draw L3 nodes (Structured Assets: Cards, Quizzes, Summaries)
        for x, y in l3:
            nc = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(0.42), Inches(0.42))
            nc.adjustments[0] = 0.2
            nc.fill.solid()
            nc.fill.fore_color.rgb = theme["accent_light"]
            nc.line.color.rgb = theme["accent_primary"]
            nc.line.width = Pt(1.5)

    elif style == "radial_core":
        # Ilaz: Concentric Radial Intelligence Hub
        center_x, center_y = 10.9, 2.5
        # Concentric outer orbit rings
        r1 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(center_x - 1.5), Inches(center_y - 1.5), Inches(3.0), Inches(3.0))
        r1.fill.background()
        r1.line.color.rgb = theme["border_light"]
        r1.line.width = Pt(1.2)
        
        r2 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(center_x - 0.9), Inches(center_y - 0.9), Inches(1.8), Inches(1.8))
        r2.fill.background()
        r2.line.color.rgb = theme["accent_border"]
        r2.line.width = Pt(1.2)
        
        # Center Core Hub
        c_hub = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(center_x - 0.35), Inches(center_y - 0.35), Inches(0.70), Inches(0.70))
        c_hub.fill.solid()
        c_hub.fill.fore_color.rgb = theme["accent_primary"]
        c_hub.line.color.rgb = theme["accent_primary"]
        c_hub.line.width = Pt(2.0)
        
        # Satellites on outer orbits
        satellites = [
            (center_x - 1.5, center_y, 0.36),
            (center_x + 1.2, center_y - 0.6, 0.38),
            (center_x + 0.9, center_y + 0.9, 0.34),
            (center_x - 0.6, center_y - 1.1, 0.32),
            (center_x - 0.2, center_y + 1.3, 0.34)
        ]
        for sx, sy, s_diam in satellites:
            # line to center
            conn = slide.shapes.add_connector(1, Inches(center_x), Inches(center_y), Inches(sx + s_diam/2), Inches(sy + s_diam/2))
            conn.line.color.rgb = theme["accent_border"]
            conn.line.width = Pt(1.0)
            # satellite node
            sat = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(sx), Inches(sy), Inches(s_diam), Inches(s_diam))
            sat.fill.solid()
            sat.fill.fore_color.rgb = theme["bg_card"]
            sat.line.color.rgb = theme["accent_primary"]
            sat.line.width = Pt(1.8)

# ------------------------------------------------------------------------------
# COMPLETE 8-SLIDE DECK GENERATOR
# ------------------------------------------------------------------------------
def generate_8_slide_deck(team_info, theme, output_path):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # ==========================================================================
    # SLIDE 1: TITLE SLIDE (College Name, Team, Roll No 1st, Points format, bold)
    # ==========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s1, theme["bg_card"])

    # Top College & Category Tag
    tb_c1 = s1.shapes.add_textbox(Inches(0.85), Inches(0.65), Inches(10.0), Inches(0.35))
    tf_c1 = tb_c1.text_frame
    p_c1 = tf_c1.paragraphs[0]
    p_c1.text = "PATRICIAN COLLEGE OF ARTS AND SCIENCE  |  DEPARTMENT OF COMPUTER APPLICATIONS"
    p_c1.font.name = theme["font_heading"]
    p_c1.font.size = Pt(11)
    p_c1.font.bold = True
    p_c1.font.color.rgb = theme["accent_primary"]

    p_c2 = tf_c1.add_paragraph()
    p_c2.space_before = Pt(2)
    p_c2.text = f"AI-AUGMENTED BACKEND DEVELOPMENT - A&S ({theme['um_code']})  •  TN SKILLS PROJECT RECORD"
    p_c2.font.name = theme["font_body"]
    p_c2.font.size = Pt(9.5)
    p_c2.font.color.rgb = theme["grey_muted"]

    # Main Project Title
    tb_t1 = s1.shapes.add_textbox(Inches(0.85), Inches(1.20), Inches(8.5), Inches(0.8))
    tf_t1 = tb_t1.text_frame
    tf_t1.word_wrap = True
    p_t1 = tf_t1.paragraphs[0]
    p_t1.text = "Study Buddy AI"
    p_t1.font.name = theme["font_heading"]
    p_t1.font.size = Pt(44)
    p_t1.font.bold = True
    p_t1.font.color.rgb = theme["charcoal_dark"]

    p_sub1 = tf_t1.add_paragraph()
    p_sub1.space_before = Pt(2)
    p_sub1.text = "LearnMate — AI Assistant for Smarter Learning and Study Support"
    p_sub1.font.name = theme["font_body"]
    p_sub1.font.size = Pt(17)
    p_sub1.font.bold = True
    p_sub1.font.color.rgb = theme["accent_primary"]

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
        r_dot.font.name = theme["font_body"]
        r_dot.font.size = Pt(11.5)
        r_dot.font.bold = True
        r_dot.font.color.rgb = theme["accent_primary"]
        
        r_lbl = p_pt.add_run()
        r_lbl.text = f"{label}: "
        r_lbl.font.name = theme["font_body"]
        r_lbl.font.size = Pt(11.5)
        r_lbl.font.bold = True
        r_lbl.font.color.rgb = theme["charcoal_dark"]
        
        r_val = p_pt.add_run()
        r_val.text = val
        r_val.font.name = theme["font_body"]
        r_val.font.size = Pt(11.5)
        r_val.font.color.rgb = theme["charcoal_mid"]

    # Dynamic Right Graphic per theme
    draw_title_graphic(s1, theme)

    # Horizontal Divider Line
    hdiv1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.85), Inches(4.35), Inches(11.6), Inches(0.015))
    hdiv1.fill.solid()
    hdiv1.fill.fore_color.rgb = theme["border_light"]
    hdiv1.line.fill.background()

    # Bottom Two Balanced Cards: Tech Stack (Left) & Team Information (Right)
    add_clean_card(s1, 0.85, 4.60, 5.7, 2.35, bg_color=theme["bg_page"], border_color=theme["border_light"], radius=theme["card_radius"])
    tb_ts = s1.shapes.add_textbox(Inches(1.10), Inches(4.75), Inches(5.2), Inches(2.05))
    tf_ts = tb_ts.text_frame
    tf_ts.word_wrap = True
    p_tsh = tf_ts.paragraphs[0]
    p_tsh.text = "CORE TECHNOLOGY STACK"
    p_tsh.font.name = theme["font_heading"]
    p_tsh.font.size = Pt(11)
    p_tsh.font.bold = True
    p_tsh.font.color.rgb = theme["accent_primary"]

    p_tsb = tf_ts.add_paragraph()
    p_tsb.space_before = Pt(5)
    p_tsb.text = "Node.js  •  Express.js  •  MongoDB  •  Mongoose ODM\nJWT & RBAC Security  •  Google Gemini 2.5 Flash API"
    p_tsb.font.name = theme["font_body"]
    p_tsb.font.size = Pt(12.5)
    p_tsb.font.bold = True
    p_tsb.font.color.rgb = theme["charcoal_dark"]
    p_tsb.line_spacing = 1.3

    p_tss = tf_ts.add_paragraph()
    p_tss.space_before = Pt(5)
    p_tss.text = "Production RESTful architecture enforcing query-level data isolation."
    p_tss.font.name = theme["font_body"]
    p_tss.font.size = Pt(10.5)
    p_tss.font.color.rgb = theme["grey_muted"]

    # Right: Team Information (Roll No 1st, College Name)
    add_clean_card(s1, 6.75, 4.60, 5.7, 2.35, bg_color=theme["bg_page"], border_color=theme["border_light"], radius=theme["card_radius"])
    tb_tm1 = s1.shapes.add_textbox(Inches(7.00), Inches(4.72), Inches(5.2), Inches(2.10))
    tf_tm1 = tb_tm1.text_frame
    tf_tm1.word_wrap = True
    p_tmh = tf_tm1.paragraphs[0]
    p_tmh.text = f"{team_info['college_short']} — {team_info['team_header']}".upper()
    p_tmh.font.name = theme["font_heading"]
    p_tmh.font.size = Pt(10)
    p_tmh.font.bold = True
    p_tmh.font.color.rgb = theme["accent_primary"]

    p_tl = tf_tm1.add_paragraph()
    p_tl.space_before = Pt(4)
    r_tll = p_tl.add_run()
    r_tll.text = "Team Leader: "
    r_tll.font.name = theme["font_body"]
    r_tll.font.size = Pt(11.5)
    r_tll.font.bold = True
    r_tll.font.color.rgb = theme["charcoal_dark"]

    r_tln = p_tl.add_run()
    r_tln.text = team_info['leader_text']
    r_tln.font.name = theme["font_body"]
    r_tln.font.size = Pt(11.5)
    r_tln.font.bold = True
    r_tln.font.color.rgb = theme["accent_primary"]

    p_tl_role = tf_tm1.add_paragraph()
    p_tl_role.space_before = Pt(1)
    r_tlr = p_tl_role.add_run()
    r_tlr.text = team_info['leader_role']
    r_tlr.font.name = theme["font_body"]
    r_tlr.font.size = Pt(9.5)
    r_tlr.font.color.rgb = theme["grey_muted"]

    p_tmb = tf_tm1.add_paragraph()
    p_tmb.space_before = Pt(4)
    r_tml = p_tmb.add_run()
    r_tml.text = "Teammates (Roll No 1st):\n"
    r_tml.font.name = theme["font_body"]
    r_tml.font.size = Pt(11)
    r_tml.font.bold = True
    r_tml.font.color.rgb = theme["charcoal_dark"]

    r_tmn = p_tmb.add_run()
    r_tmn.text = team_info['teammates_text']
    r_tmn.font.name = theme["font_body"]
    r_tmn.font.size = Pt(10)
    r_tmn.font.color.rgb = theme["charcoal_mid"]
    p_tmb.line_spacing = 1.2

    # ==========================================================================
    # SLIDE 2: STUDENT PROBLEM SPACE & REAL-WORLD MOTIVATION
    # ==========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "Academic Pain Points & Cognitive Overload", "PROBLEM SPACE & CASE STUDY MOTIVATION", 
               "Deconstructing the critical inefficiencies of traditional, unassisted examination revision.", 2, theme)

    # Case Study Context Banner
    add_clean_card(s2, 0.8, 1.85, 11.733, 1.05, bg_color=theme["bg_card"], border_color=theme["border_light"], radius=theme["card_radius"])
    tb_cs = s2.shapes.add_textbox(Inches(1.0), Inches(1.92), Inches(11.3), Inches(0.9))
    tf_cs = tb_cs.text_frame
    tf_cs.word_wrap = True
    p_csh = tf_cs.paragraphs[0]
    r_ch1 = p_csh.add_run()
    r_ch1.text = "The Undergraduate Dilemma (Case Scenario): "
    r_ch1.font.name = theme["font_heading"]
    r_ch1.font.size = Pt(12)
    r_ch1.font.bold = True
    r_ch1.font.color.rgb = theme["prob_alert"]

    r_ch2 = p_csh.add_run()
    r_ch2.text = (
        "Learners tackling multi-chapter academic syllabi concurrently balance coursework, technical skill-building, and exams. "
        "Traditional revision forces students into clerical transcription—manually summarizing dense texts, extracting keywords, and "
        "hand-crafting practice questions. This clerical overhead displaces deep comprehension and prevents quantitative self-assessment."
    )
    r_ch2.font.name = theme["font_body"]
    r_ch2.font.size = Pt(11)
    r_ch2.font.color.rgb = theme["charcoal_dark"]
    p_csh.line_spacing = 1.2

    # 3 Problem Cards
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
        add_clean_card(s2, cx, 3.10, 3.733, 3.85, bg_color=theme["prob_light"], border_color=theme["prob_border"], radius=theme["card_radius"])
        
        tb_p = s2.shapes.add_textbox(Inches(cx + 0.2), Inches(3.25), Inches(3.333), Inches(3.5))
        tf_p = tb_p.text_frame
        tf_p.word_wrap = True
        
        p_tag = tf_p.paragraphs[0]
        p_tag.text = tag
        p_tag.font.name = theme["font_heading"]
        p_tag.font.size = Pt(10)
        p_tag.font.bold = True
        p_tag.font.color.rgb = theme["prob_alert"]
        
        p_t = tf_p.add_paragraph()
        p_t.space_before = Pt(3)
        p_t.text = title
        p_t.font.name = theme["font_heading"]
        p_t.font.size = Pt(13)
        p_t.font.bold = True
        p_t.font.color.rgb = theme["charcoal_dark"]
        
        p_b = tf_p.add_paragraph()
        p_b.space_before = Pt(8)
        p_b.line_spacing = 1.25
        p_b.text = body
        p_b.font.name = theme["font_body"]
        p_b.font.size = Pt(10.5)
        p_b.font.color.rgb = theme["charcoal_mid"]

    # ==========================================================================
    # SLIDE 3: STRATEGIC SOLUTION & VALUE PROPOSITION
    # ==========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "The LearnMate Paradigm: Automated Academic Synthesis", "STRATEGIC SOLUTION & VALUE PROPOSITION", 
               "Transforming raw educational artifacts into structured, persistent revision assets via Generative AI.", 3, theme)

    sol_cards = [
        ("1. Omni-Format Material Ingestion", 
         "CENTRALIZED ASSET REPOSITORY",
         "• Ingests diverse student documents (.txt, .md, .pdf) through streaming multipart form-data pipelines.\n"
         "• Enforces server-side validation: MIME-type assertions, file-size quotas, and character buffer extraction.\n"
         "• Indexes courseware under unique user IDs in MongoDB, establishing an accessible digital knowledge vault.",
         theme["accent_light"], theme["accent_border"]),
        ("2. Algorithmic Active-Recall Engine", 
         "COGNITIVE ENHANCEMENT SUITE",
         "• Dynamic Synthesizer: Distills multi-page chapters into high-density conceptual revision summaries.\n"
         "• Schema-Locked Flashcards: Generates structured Q&A arrays via strict JSON constraints for spaced repetition.\n"
         "• Objective Diagnostic MCQs: Automatically formats 4-option quizzes complete with verified answer indices.",
         theme["bg_card"], theme["border_light"]),
        ("3. Milestone-Driven Plans & Grounded Q&A", 
         "ADAPTIVE PREPARATION ROADMAPS",
         "• Calibrated Study Calendars: Generates realistic day-by-day revision milestones mapped to remaining exam dates.\n"
         "• Material-Grounded Q&A: Ingests specific student inquiries and answers strictly using uploaded context.\n"
         "• Hallucination Elimination: System prompts restrict AI responses to factual assertions in student notes.",
         theme["bg_card"], theme["border_light"]),
        ("4. Permanent Revision Persistence", 
         "CONTINUOUS RETENTION & REUSE",
         "• Automatically catalogs every summary, quiz, and study schedule into user-partitioned MongoDB collections.\n"
         "• Enables multi-session review and recursive testing without incurring redundant LLM API token costs.\n"
         "• Guarantees students maintain complete ownership of their academic assets throughout their degree.",
         theme["accent_light"], theme["accent_border"])
    ]

    for idx, (title, tag, body, bg_col, border_col) in enumerate(sol_cards):
        row = idx // 2
        col = idx % 2
        cx = 0.8 + col * 5.95
        cy = 1.85 + row * 2.60
        
        add_clean_card(s3, cx, cy, 5.75, 2.40, bg_color=bg_col, border_color=border_col, radius=theme["card_radius"])
        
        tb_s = s3.shapes.add_textbox(Inches(cx + 0.2), Inches(cy + 0.15), Inches(5.35), Inches(2.10))
        tf_s = tb_s.text_frame
        tf_s.word_wrap = True
        
        p_tag = tf_s.paragraphs[0]
        p_tag.text = tag
        p_tag.font.name = theme["font_heading"]
        p_tag.font.size = Pt(10)
        p_tag.font.bold = True
        p_tag.font.color.rgb = theme["accent_primary"]
        
        p_t = tf_s.add_paragraph()
        p_t.space_before = Pt(2)
        p_t.text = title
        p_t.font.name = theme["font_heading"]
        p_t.font.size = Pt(13)
        p_t.font.bold = True
        p_t.font.color.rgb = theme["charcoal_dark"]
        
        p_b = tf_s.add_paragraph()
        p_b.space_before = Pt(5)
        p_b.line_spacing = 1.2
        p_b.text = body
        p_b.font.name = theme["font_body"]
        p_b.font.size = Pt(10)
        p_b.font.color.rgb = theme["charcoal_mid"]

    # ==========================================================================
    # SLIDE 4: DISTRIBUTED SERVER ARCHITECTURE & EXECUTION PIPELINE
    # ==========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "Distributed Server Architecture & Execution Pipeline", "SYSTEM ARCHITECTURE & EXECUTION FLOW", 
               "Decoupled MVC pipeline: client requests, perimeter defenses, business logic, and dual service backends.", 4, theme)

    # Box 1: Client Interfaces
    draw_arch_card(s4, 0.8, 1.95, 2.3, 3.65, "Client Interface Layer", "Web & Postman Consumers",
                   ["Student & Admin consumers", "HTTPS calls + Bearer JWT tokens", "Transmits raw notes & PDFs", "Requests on-demand AI tasks", "Retrieves persisted assets"], theme)

    # Arrow 1: Client -> Gateway
    a1 = s4.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(3.20), Inches(3.65), Inches(0.35), Inches(0.25))
    a1.fill.solid()
    a1.fill.fore_color.rgb = theme["accent_primary"]
    a1.line.fill.background()

    # Box 2: API Gateway
    draw_arch_card(s4, 3.65, 1.95, 2.5, 3.65, "REST API Gateway", "Express.js Routing Tier",
                   ["Centralized endpoint dispatcher", "Helmet HTTP security headers", "CORS & Morgan access logger", "Two-tier rate-limit throttling", "Multer streaming file parser", "Express 5 NoSQL-injection guard"], theme)

    # Arrow 2: Gateway -> Controllers
    a2 = s4.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(6.25), Inches(3.65), Inches(0.35), Inches(0.25))
    a2.fill.solid()
    a2.fill.fore_color.rgb = theme["accent_primary"]
    a2.line.fill.background()

    # Box 3: Controllers & Security
    draw_arch_card(s4, 6.7, 1.95, 2.6, 3.65, "Controllers & Middleware", "Business Logic & Auth Layer",
                   ["Validates 7-day JWT Bearer tokens", "Role gate (Student vs. Admin)", "Enforces tenant isolation { userId }", "express-validator schema guards", "Centralized exception handler", "Orchestrates AI & DB services"], theme)

    # Arrow 3A: Controllers -> Gemini
    a3a = s4.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(9.40), Inches(2.70), Inches(0.4), Inches(0.22))
    a3a.fill.solid()
    a3a.fill.fore_color.rgb = theme["accent_primary"]
    a3a.line.fill.background()

    # Box 4A: Gemini Engine
    draw_arch_card(s4, 9.9, 1.95, 2.65, 1.72, "Google Gemini 2.5 Flash", "@google/genai Service Layer",
                   ["Material-grounded prompt tuning", "Conceptual summaries (prose)", "Structured JSON mode for MCQs/Cards"], theme, is_highlight=True)

    # Arrow 3B: Controllers -> MongoDB
    a3b = s4.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(9.40), Inches(4.65), Inches(0.4), Inches(0.22))
    a3b.fill.solid()
    a3b.fill.fore_color.rgb = theme["accent_primary"]
    a3b.line.fill.background()

    # Box 4B: MongoDB Persistence
    draw_arch_card(s4, 9.9, 3.88, 2.65, 1.72, "MongoDB Atlas & Mongoose", "Data Persistence Tier",
                   ["Bcrypt credentials & user roles", "Uploaded raw notes & metadata", "Saves summaries, cards & quizzes", "Enables anytime exam retrieval"], theme, is_accent=True)

    # Bottom Horizontal Persistent Retrieval Banner
    bot_bar = add_clean_card(s4, 0.8, 5.82, 11.75, 1.10, bg_color=theme["bg_card"], border_color=theme["border_light"], radius=theme["card_radius"])
    tb_bb = s4.shapes.add_textbox(Inches(1.0), Inches(5.86), Inches(11.35), Inches(0.98))
    tf_bb = tb_bb.text_frame
    tf_bb.word_wrap = True
    p_bb = tf_bb.paragraphs[0]
    r_bbh = p_bb.add_run()
    r_bbh.text = "Persistent Asset Access Pipeline: "
    r_bbh.font.name = theme["font_heading"]
    r_bbh.font.size = Pt(11.5)
    r_bbh.font.bold = True
    r_bbh.font.color.rgb = theme["accent_primary"]

    r_bbb = p_bb.add_run()
    r_bbb.text = (
        "Authenticated learners can securely search, filter, and review stored revision aids across examination cycles. "
        "When an AI generation completes, its structured payload is permanently persisted to MongoDB bound to the active user's account "
        "({ userId: req.user._id })—ensuring learning assets remain accessible on demand without redundant API invocation."
    )
    r_bbb.font.name = theme["font_body"]
    r_bbb.font.size = Pt(10.5)
    r_bbb.font.color.rgb = theme["charcoal_mid"]
    p_bb.line_spacing = 1.25

    # ==========================================================================
    # SLIDE 5: ENGINEERING STACK & TECHNICAL ECOSYSTEM
    # ==========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "Engineering Stack & Technical Ecosystem", "TECHNOLOGY STACK & PROGRAMS USED", 
               "Exact packages, runtime environments, AI SDKs, and developer tooling powering the platform.", 5, theme)

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
        
        add_clean_card(s5, cx, cy, 5.75, 2.40, bg_color=theme["bg_card"], border_color=theme["border_light"], radius=theme["card_radius"])
        
        tb_t = s5.shapes.add_textbox(Inches(cx + 0.2), Inches(cy + 0.15), Inches(5.35), Inches(2.10))
        tf_t = tb_t.text_frame
        tf_t.word_wrap = True
        
        p_cat = tf_t.paragraphs[0]
        p_cat.text = category
        p_cat.font.name = theme["font_heading"]
        p_cat.font.size = Pt(10)
        p_cat.font.bold = True
        p_cat.font.color.rgb = theme["accent_primary"]
        
        p_t = tf_t.add_paragraph()
        p_t.space_before = Pt(2)
        p_t.text = title
        p_t.font.name = theme["font_heading"]
        p_t.font.size = Pt(13)
        p_t.font.bold = True
        p_t.font.color.rgb = theme["charcoal_dark"]
        
        for item_name, item_desc in items:
            p_i = tf_t.add_paragraph()
            p_i.space_before = Pt(4)
            p_i.line_spacing = 1.15
            
            r_in = p_i.add_run()
            r_in.text = f"• {item_name}: "
            r_in.font.name = theme["font_body"]
            r_in.font.size = Pt(10)
            r_in.font.bold = True
            r_in.font.color.rgb = theme["charcoal_dark"]
            
            r_id = p_i.add_run()
            r_id.text = item_desc
            r_id.font.name = theme["font_body"]
            r_id.font.size = Pt(9.5)
            r_id.font.color.rgb = theme["charcoal_mid"]

    # ==========================================================================
    # SLIDE 6: DATA MODELING & MONGOOSE ENTITY ARCHITECTURE
    # ==========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "Data Modeling & Mongoose Entity Architecture", "DATABASE CONFIGURATION & SCHEMAS", 
               "Relational data modeling in MongoDB NoSQL via Mongoose schemas, compound indexes, and entity hooks.", 6, theme)

    schemas_2x2 = [
        ("User Identity Model (User.js)", "AUTHENTICATION & ACCESS ROLES", [
            ("Core Attributes", "name (String), email (String, unique, indexed), password (bcrypt hash), role (enum: ['student', 'admin']), createdAt."),
            ("Pre-Save Hook", "Mongoose middleware triggers bcrypt.hash(password, 10) automatically before committing modified credentials."),
            ("Cryptographic Method", "matchPassword(inputPassword) executes timing-attack resistant comparisons during session login.")
        ], theme["accent_light"], theme["accent_border"]),
        
        ("Course Material Model (Material.js)", "LEARNING ASSET REPOSITORY", [
            ("Core Attributes", "userId (ObjectId ref: 'User'), title (String), content (Text buffer), fileType (String), fileSize (Number), uploadedAt."),
            ("Relational Role", "Primary root anchor for all AI features; maintains 1-to-Many cascade relationships with generated study outputs."),
            ("Compound Indexing", "Indexed on { userId: 1, uploadedAt: -1 } ensuring microsecond-level query resolution for student documents.")
        ], theme["bg_card"], theme["border_light"]),
        
        ("Active-Recall Learning Models", "FLASHCARDS & QUIZ ENTITIES", [
            ("Flashcard Schema", "userId, materialId, cards: [{ question: String, answer: String }], cardCount (Number)."),
            ("Quiz Schema", "userId, materialId, questions: [{ question: String, options: [String], correctAnswer: Number, explanation: String }]."),
            ("Output Guarantee", "Strict JSON schema enforcement (`responseSchema`) ensures zero parsing failures during downstream ingestion.")
        ], theme["bg_card"], theme["border_light"]),
        
        ("Synthesis & Planning Models", "SUMMARIES & REVISION TIMELINES", [
            ("Summary Schema", "userId (ref User), materialId (ref Material), summaryText (String), keyTakeaways ([String]), generatedAt."),
            ("Study Plan Schema", "userId, targetExam: String, examDate: Date, dailyPlan: [{ day: Number, topics: [String], hours: Number }]."),
            ("Multi-Tenant Isolation", "Queries strictly scoped to authenticated user ID, preventing horizontal cross-student data exposure.")
        ], theme["accent_light"], theme["accent_border"])
    ]

    for idx, (title, category, items, bg_col, border_col) in enumerate(schemas_2x2):
        col = idx % 2
        row = idx // 2
        cx = 0.8 + col * 5.95
        cy = 1.85 + row * 2.60
        
        add_clean_card(s6, cx, cy, 5.75, 2.40, bg_color=bg_col, border_color=border_col, radius=theme["card_radius"])
        
        tb_sc = s6.shapes.add_textbox(Inches(cx + 0.2), Inches(cy + 0.15), Inches(5.35), Inches(2.10))
        tf_sc = tb_sc.text_frame
        tf_sc.word_wrap = True
        
        p_c = tf_sc.paragraphs[0]
        p_c.text = category
        p_c.font.name = theme["font_heading"]
        p_c.font.size = Pt(10)
        p_c.font.bold = True
        p_c.font.color.rgb = theme["accent_primary"]
        
        p_tit = tf_sc.add_paragraph()
        p_tit.space_before = Pt(2)
        p_tit.text = title
        p_tit.font.name = theme["font_heading"]
        p_tit.font.size = Pt(13)
        p_tit.font.bold = True
        p_tit.font.color.rgb = theme["charcoal_dark"]
        
        for iname, idesc in items:
            p_item = tf_sc.add_paragraph()
            p_item.space_before = Pt(4)
            p_item.line_spacing = 1.15
            
            r_in = p_item.add_run()
            r_in.text = f"• {iname}: "
            r_in.font.name = theme["font_body"]
            r_in.font.size = Pt(10)
            r_in.font.bold = True
            r_in.font.color.rgb = theme["charcoal_dark"]
            
            r_id = p_item.add_run()
            r_id.text = idesc
            r_id.font.name = theme["font_body"]
            r_id.font.size = Pt(9.5)
            r_id.font.color.rgb = theme["charcoal_mid"]

    # ==========================================================================
    # SLIDE 7: ENTERPRISE SECURITY & ACCESS GOVERNANCE
    # ==========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "Enterprise Security & Access Governance", "DEFENSE-IN-DEPTH SECURITY MODEL", 
               "Multi-layered protective controls shielding user privacy, application state, and API integrity.", 7, theme)

    sec_pillars = [
        ("Bcrypt Password Salting", "CRYPTOGRAPHIC STORAGE",
         "• Passwords salted with 10 rounds via bcryptjs prior to database commit.\n"
         "• Plaintext credentials never stored in database or exposed in server logs.\n"
         "• Resistant to precomputed rainbow-table and dictionary-based attacks.",
         theme["bg_card"]),
        ("Stateless Bearer JWT Tokens", "IDENTITY GOVERNANCE",
         "• 7-day HMAC-SHA256 signed JSON Web Tokens issued upon successful login.\n"
         "• Bearer tokens cryptographically extracted and asserted by authMiddleware.\n"
         "• Decoded user identity attached to req.user for downstream handler access.",
         theme["accent_light"]),
        ("Role-Based Access Control", "AUTHORIZATION GATEWAY",
         "• Authorize middleware segregates 'student' from 'admin' operational scopes.\n"
         "• Prevents unauthorized privilege escalation to platform telemetry endpoints.\n"
         "• Automated 403 Forbidden rejection for unauthorized permission requests.",
         theme["bg_card"]),
        ("Query-Level Data Isolation", "TENANT PRIVACY BOUNDARY",
         "• Every database operation explicitly scoped to `{ userId: req.user._id }`.\n"
         "• Guarantees students cannot view or mutate another learner's study assets.\n"
         "• Defense-in-depth isolation preventing horizontal data exposure.",
         theme["accent_light"])
    ]

    for idx, (title, tag, body, bg_col) in enumerate(sec_pillars):
        cx = 0.8 + idx * 3.0
        border_col = theme["accent_border"] if bg_col == theme["accent_light"] else theme["border_light"]
        add_clean_card(s7, cx, 1.85, 2.75, 4.95, bg_color=bg_col, border_color=border_col, radius=theme["card_radius"])
        
        tb_sec = s7.shapes.add_textbox(Inches(cx + 0.15), Inches(2.0), Inches(2.45), Inches(4.65))
        tf_sec = tb_sec.text_frame
        tf_sec.word_wrap = True
        
        p_t = tf_sec.paragraphs[0]
        p_t.text = tag
        p_t.font.name = theme["font_heading"]
        p_t.font.size = Pt(9.5)
        p_t.font.bold = True
        p_t.font.color.rgb = theme["accent_primary"]
        
        p_h = tf_sec.add_paragraph()
        p_h.space_before = Pt(4)
        p_h.text = title
        p_h.font.name = theme["font_heading"]
        p_h.font.size = Pt(12)
        p_h.font.bold = True
        p_h.font.color.rgb = theme["charcoal_dark"]
        
        p_b = tf_sec.add_paragraph()
        p_b.space_before = Pt(8)
        p_b.line_spacing = 1.2
        p_b.text = body
        p_b.font.name = theme["font_body"]
        p_b.font.size = Pt(9.5)
        p_b.font.color.rgb = theme["charcoal_mid"]

    # ==========================================================================
    # SLIDE 8: PROJECT EXECUTION, API VERIFICATION & THANK YOU / QUERIES
    # ==========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "Project Execution, API Verification & Concluding Queries", "IMPLEMENTATION, VALIDATION & FINAL DISCUSSION", 
               "Verified endpoint testing in Postman, live demonstration evidence, and concluding Q&A discussion.", 8, theme)

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
        add_clean_card(s8, x_pos, y_top, card_w, card_h, bg_color=theme["bg_card"], border_color=theme["border_light"], radius=theme["card_radius"])
        
        # Title
        tb_cap = s8.shapes.add_textbox(Inches(x_pos + 0.1), Inches(y_top + 0.08), Inches(card_w - 0.2), Inches(0.32))
        p_c = tb_cap.text_frame.paragraphs[0]
        p_c.text = caption
        p_c.font.name = theme["font_heading"]
        p_c.font.size = Pt(10.5)
        p_c.font.bold = True
        p_c.font.color.rgb = theme["accent_primary"]
        
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
        p_n.font.name = theme["font_body"]
        p_n.font.size = Pt(9.5)
        p_n.font.color.rgb = theme["grey_muted"]
        p_n.line_spacing = 1.15

    # Testing Statement Box
    add_clean_card(s8, 0.8, 5.08, 11.75, 0.48, bg_color=theme["bg_page"], border_color=theme["border_light"], radius=theme["card_radius"])
    tb_stmt = s8.shapes.add_textbox(Inches(0.95), Inches(5.10), Inches(11.45), Inches(0.44))
    tf_stmt = tb_stmt.text_frame
    tf_stmt.word_wrap = True
    p_stmt = tf_stmt.paragraphs[0]
    r_st1 = p_stmt.add_run()
    r_st1.text = "API Verification: "
    r_st1.font.name = theme["font_heading"]
    r_st1.font.size = Pt(10.5)
    r_st1.font.bold = True
    r_st1.font.color.rgb = theme["accent_primary"]

    r_st2 = p_stmt.add_run()
    r_st2.text = (
        "The backend RESTful API was comprehensively tested using Postman, asserting route protection, role authorization, "
        "multipart file uploads, and Gemini AI JSON schema validation across all endpoints."
    )
    r_st2.font.name = theme["font_body"]
    r_st2.font.size = Pt(10)
    r_st2.font.color.rgb = theme["charcoal_dark"]

    # THANK YOU & QUERIES BANNER (Theme-Specific Colors)
    add_clean_card(s8, 0.8, 5.72, 11.75, 1.45, bg_color=theme["accent_light"], border_color=theme["accent_border"], radius=theme["card_radius"])

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
    r_ty.font.name = theme["font_heading"]
    r_ty.font.size = Pt(14)
    r_ty.font.bold = True
    r_ty.font.color.rgb = theme["accent_primary"]

    # Line 2: Invitation to discuss
    p_inv = tf_l.add_paragraph()
    p_inv.space_before = Pt(2)
    r_inv = p_inv.add_run()
    r_inv.text = "We welcome your questions, technical inquiries, and feedback regarding Study Buddy AI."
    r_inv.font.name = theme["font_body"]
    r_inv.font.size = Pt(10.5)
    r_inv.font.color.rgb = theme["charcoal_mid"]

    # Line 3: Links row
    p_l1 = tf_l.add_paragraph()
    p_l1.space_before = Pt(3)
    r_gh_lbl = p_l1.add_run()
    r_gh_lbl.text = "GitHub Repository: "
    r_gh_lbl.font.name = theme["font_heading"]
    r_gh_lbl.font.size = Pt(10.5)
    r_gh_lbl.font.bold = True
    r_gh_lbl.font.color.rgb = theme["accent_primary"]

    r_gh_val = p_l1.add_run()
    r_gh_val.text = f"{team_info['github_url']}    "
    r_gh_val.font.name = theme["font_body"]
    r_gh_val.font.size = Pt(10.5)
    r_gh_val.font.bold = True
    r_gh_val.font.color.rgb = theme["charcoal_dark"]

    r_yt_lbl = p_l1.add_run()
    r_yt_lbl.text = "|   Demo Video: "
    r_yt_lbl.font.name = theme["font_heading"]
    r_yt_lbl.font.size = Pt(10.5)
    r_yt_lbl.font.bold = True
    r_yt_lbl.font.color.rgb = theme["accent_primary"]

    r_yt_val = p_l1.add_run()
    r_yt_val.text = "https://youtu.be/N5lNC06rn7w"
    r_yt_val.font.name = theme["font_body"]
    r_yt_val.font.size = Pt(10.5)
    r_yt_val.font.bold = True
    r_yt_val.font.color.rgb = theme["charcoal_dark"]

    # Line 4: Closing statement
    p_l2 = tf_l.add_paragraph()
    p_l2.space_before = Pt(3)
    r_close = p_l2.add_run()
    r_close.text = "“Study Buddy AI — Turning Study Materials into Structured Learning Resources.”"
    r_close.font.name = theme["font_heading"]
    r_close.font.size = Pt(11)
    r_close.font.bold = True
    r_close.font.color.rgb = theme["accent_primary"]

    prs.save(output_path)
    print(f"Saved PPTX to: {output_path}")

# ==============================================================================
# TEAM CONFIGURATIONS & THEME ASSIGNMENT
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

# Output file paths
pptx_travis = os.path.join(OUT_DIR_D, "Study_Buddy_AI_Presentation_Travis.pptx")
pptx_ilaz = os.path.join(OUT_DIR_D, "Study_Buddy_AI_Presentation_Ilaz.pptx")
pptx_isaac = os.path.join(OUT_DIR_D, "Study_Buddy_AI_Presentation_Isaac.pptx")

# Generate all 3 decks with their distinct themes and UM codes (Travis: NMU93, Ilaz: NMU93, Isaac: NMU88)
generate_8_slide_deck(team_travis, THEME_TRAVIS, pptx_travis)
generate_8_slide_deck(team_ilaz, THEME_ILAZ, pptx_ilaz)
generate_8_slide_deck(team_isaac, THEME_ISAAC, pptx_isaac)

# Replicate to Downloads & D:\ root
shutil.copyfile(pptx_travis, os.path.join(DOWNLOADS_DIR, "Study_Buddy_AI_Presentation_Travis.pptx"))
shutil.copyfile(pptx_ilaz, os.path.join(DOWNLOADS_DIR, "Study_Buddy_AI_Presentation_Ilaz.pptx"))
shutil.copyfile(pptx_isaac, os.path.join(DOWNLOADS_DIR, "Study_Buddy_AI_Presentation_Isaac.pptx"))

shutil.copyfile(pptx_travis, r"d:\Study_Buddy_AI_Presentation_Travis.pptx")
shutil.copyfile(pptx_ilaz, r"d:\Study_Buddy_AI_Presentation_Ilaz.pptx")
shutil.copyfile(pptx_isaac, r"d:\Study_Buddy_AI_Presentation_Isaac.pptx")

# Also copy Travis deck to default files
shutil.copyfile(pptx_travis, r"d:\Study_Buddy_AI_Presentation.pptx")
shutil.copyfile(pptx_travis, os.path.join(DOWNLOADS_DIR, "Study_Buddy_AI_Presentation.pptx"))
shutil.copyfile(pptx_travis, os.path.join(OUT_DIR_D, "Study_Buddy_AI_Presentation.pptx"))

print("All PPTX files generated with unique themes!")

# ------------------------------------------------------------------------------
# CONVERT ALL 3 TO PDF VIA POWERPOINT COM
# ------------------------------------------------------------------------------
powerpoint = win32com.client.Dispatch("PowerPoint.Application")

# Travis PDF
deck_t = powerpoint.Presentations.Open(pptx_travis, WithWindow=False)
pdf_t = os.path.join(OUT_DIR_D, "Study_Buddy_AI_Presentation_Travis.pdf")
deck_t.SaveAs(pdf_t, 32)
deck_t.Close()

# Ilaz PDF
deck_i = powerpoint.Presentations.Open(pptx_ilaz, WithWindow=False)
pdf_i = os.path.join(OUT_DIR_D, "Study_Buddy_AI_Presentation_Ilaz.pdf")
deck_i.SaveAs(pdf_i, 32)
deck_i.Close()

# Isaac PDF
deck_is = powerpoint.Presentations.Open(pptx_isaac, WithWindow=False)
pdf_is = os.path.join(OUT_DIR_D, "Study_Buddy_AI_Presentation_Isaac.pdf")
deck_is.SaveAs(pdf_is, 32)
deck_is.Close()

powerpoint.Quit()

# Replicate PDFs
shutil.copyfile(pdf_t, os.path.join(DOWNLOADS_DIR, "Study_Buddy_AI_Presentation_Travis.pdf"))
shutil.copyfile(pdf_i, os.path.join(DOWNLOADS_DIR, "Study_Buddy_AI_Presentation_Ilaz.pdf"))
shutil.copyfile(pdf_is, os.path.join(DOWNLOADS_DIR, "Study_Buddy_AI_Presentation_Isaac.pdf"))

shutil.copyfile(pdf_t, r"d:\Study_Buddy_AI_Presentation_Travis.pdf")
shutil.copyfile(pdf_i, r"d:\Study_Buddy_AI_Presentation_Ilaz.pdf")
shutil.copyfile(pdf_is, r"d:\Study_Buddy_AI_Presentation_Isaac.pdf")

shutil.copyfile(pdf_t, r"d:\Study_Buddy_AI_Presentation.pdf")
shutil.copyfile(pdf_t, os.path.join(DOWNLOADS_DIR, "Study_Buddy_AI_Presentation.pdf"))

print("All PDFs converted and copied successfully!")

# Render preview PNGs for inspection
for key, p_pdf in [("travis", pdf_t), ("ilaz", pdf_i), ("isaac", pdf_is)]:
    doc = fitz.open(p_pdf)
    doc[0].get_pixmap(dpi=150).save(os.path.join(SCRATCH_DIR, f"{key}_slide_1.png"))
    doc[7].get_pixmap(dpi=150).save(os.path.join(SCRATCH_DIR, f"{key}_slide_8.png"))
    doc.close()

print("Saved preview images for all 3 themes!")
