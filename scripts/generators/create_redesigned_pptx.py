import os
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def build_presentation(output_path):
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Typography & Palette
    FONT_HEADING = 'Aptos'
    FONT_BODY = 'Aptos'
    
    C_BG = RGBColor(248, 250, 252)          # #F8FAFC
    C_WHITE = RGBColor(255, 255, 255)       # #FFFFFF
    C_PRIMARY = RGBColor(23, 32, 51)        # #172033
    C_TEAL = RGBColor(15, 118, 110)         # #0F766E
    C_TEAL_LIGHT = RGBColor(240, 253, 250)  # #F0FDFA
    C_TEAL_BADGE = RGBColor(204, 251, 241)  # #CCFBF1
    C_TEAL_BORDER = RGBColor(153, 246, 228) # #99F6E4
    C_MUTED = RGBColor(100, 116, 139)       # #64748B
    C_SLATE = RGBColor(71, 85, 105)         # #475569
    C_BORDER = RGBColor(226, 232, 240)      # #E2E8F0
    C_ROSE_BG = RGBColor(255, 241, 242)     # #FFF1F2
    C_ROSE_BORDER = RGBColor(254, 205, 211) # #FECDD3
    C_ROSE_TEXT = RGBColor(190, 18, 60)     # #BE123C

    def set_slide_background(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = C_BG
        bg.line.fill.background()
        return bg

    def add_card(slide, x, y, w, h, bg_color=C_WHITE, border_color=C_BORDER, border_width=1.0, radius=0.035):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        if border_color:
            card.line.color.rgb = border_color
            card.line.width = Pt(border_width)
        else:
            card.line.fill.background()
        if radius is not None:
            card.adjustments[0] = radius
        return card

    def add_header(slide, kicker_text, title_text, subtitle_text):
        # Kicker
        kicker_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.45), Inches(11.733), Inches(0.32))
        tf_k = kicker_box.text_frame
        tf_k.word_wrap = True
        tf_k.margin_left = tf_k.margin_top = tf_k.margin_right = tf_k.margin_bottom = 0
        p_k = tf_k.paragraphs[0]
        p_k.text = kicker_text.upper()
        p_k.font.name = FONT_HEADING
        p_k.font.size = Pt(11)
        p_k.font.bold = True
        p_k.font.color.rgb = C_TEAL

        # Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.75), Inches(11.733), Inches(0.55))
        tf_t = title_box.text_frame
        tf_t.word_wrap = True
        tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0
        p_t = tf_t.paragraphs[0]
        p_t.text = title_text
        p_t.font.name = FONT_HEADING
        p_t.font.size = Pt(26)
        p_t.font.bold = True
        p_t.font.color.rgb = C_PRIMARY

        # Subtitle
        sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.33), Inches(11.733), Inches(0.35))
        tf_s = sub_box.text_frame
        tf_s.word_wrap = True
        tf_s.margin_left = tf_s.margin_top = tf_s.margin_right = tf_s.margin_bottom = 0
        p_s = tf_s.paragraphs[0]
        p_s.text = subtitle_text
        p_s.font.name = FONT_BODY
        p_s.font.size = Pt(14)
        p_s.font.color.rgb = C_MUTED

        # Thin divider
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.72), Inches(11.733), Inches(0.015))
        line.fill.solid()
        line.fill.fore_color.rgb = C_BORDER
        line.line.fill.background()

    # =========================================================================
    # SLIDE 1: TITLE SLIDE (Study Buddy AI)
    # =========================================================================
    slide1 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide1)

    # Category Pill Tag
    tag_bg = add_card(slide1, Inches(0.8), Inches(0.85), Inches(3.6), Inches(0.38), bg_color=C_TEAL_LIGHT, border_color=C_TEAL_BORDER, radius=0.25)
    tf_tag = tag_bg.text_frame
    tf_tag.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_tag = tf_tag.paragraphs[0]
    p_tag.alignment = PP_ALIGN.CENTER
    p_tag.text = "AI-AUGMENTED BACKEND DEVELOPMENT"
    p_tag.font.name = FONT_HEADING
    p_tag.font.size = Pt(10.5)
    p_tag.font.bold = True
    p_tag.font.color.rgb = C_TEAL

    # Main Title
    s1_title_box = slide1.shapes.add_textbox(Inches(0.8), Inches(1.35), Inches(7.2), Inches(0.85))
    tf_s1_t = s1_title_box.text_frame
    tf_s1_t.word_wrap = True
    tf_s1_t.margin_left = tf_s1_t.margin_top = tf_s1_t.margin_right = tf_s1_t.margin_bottom = 0
    p_s1_t = tf_s1_t.paragraphs[0]
    p_s1_t.text = "Study Buddy AI"
    p_s1_t.font.name = FONT_HEADING
    p_s1_t.font.size = Pt(44)
    p_s1_t.font.bold = True
    p_s1_t.font.color.rgb = C_PRIMARY

    # Subtitle
    s1_sub_box = slide1.shapes.add_textbox(Inches(0.8), Inches(2.32), Inches(7.5), Inches(0.45))
    tf_s1_sub = s1_sub_box.text_frame
    tf_s1_sub.word_wrap = True
    tf_s1_sub.margin_left = tf_s1_sub.margin_top = tf_s1_sub.margin_right = tf_s1_sub.margin_bottom = 0
    p_s1_sub = tf_s1_sub.paragraphs[0]
    p_s1_sub.text = "LearnMate — AI Assistant for Smarter Learning and Study Support"
    p_s1_sub.font.name = FONT_HEADING
    p_s1_sub.font.size = Pt(17.5)
    p_s1_sub.font.bold = True
    p_s1_sub.font.color.rgb = C_TEAL

    # One-sentence Project Description
    s1_desc_box = slide1.shapes.add_textbox(Inches(0.8), Inches(2.95), Inches(6.8), Inches(1.05))
    tf_s1_d = s1_desc_box.text_frame
    tf_s1_d.word_wrap = True
    tf_s1_d.margin_left = tf_s1_d.margin_top = tf_s1_d.margin_right = tf_s1_d.margin_bottom = 0
    p_s1_d = tf_s1_d.paragraphs[0]
    p_s1_d.text = (
        "An intelligent backend platform that dynamically transforms student study materials into "
        "concise summaries, active-recall flashcards, self-assessment quizzes, and personalized "
        "revision schedules using Google Gemini."
    )
    p_s1_d.font.name = FONT_BODY
    p_s1_d.font.size = Pt(14.5)
    p_s1_d.font.color.rgb = C_SLATE

    # Subtle divider
    s1_div = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(4.2), Inches(6.8), Inches(0.015))
    s1_div.fill.solid()
    s1_div.fill.fore_color.rgb = C_BORDER
    s1_div.line.fill.background()

    # Bottom Two Info Cards
    # Card 1: Core Technology Stack
    c1 = add_card(slide1, Inches(0.8), Inches(4.45), Inches(5.7), Inches(2.35), bg_color=C_WHITE, border_color=C_BORDER, radius=0.035)
    tf_c1 = c1.text_frame
    tf_c1.word_wrap = True
    tf_c1.margin_left = Inches(0.28)
    tf_c1.margin_top = Inches(0.22)
    tf_c1.margin_right = Inches(0.28)
    tf_c1.margin_bottom = Inches(0.2)
    
    p1 = tf_c1.paragraphs[0]
    p1.text = "CORE TECHNOLOGY STACK"
    p1.font.name = FONT_HEADING
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = C_TEAL
    p1.space_after = Pt(6)

    p2 = tf_c1.add_paragraph()
    p2.text = "Node.js  •  Express.js  •  MongoDB  •  Mongoose ODM"
    p2.font.name = FONT_HEADING
    p2.font.size = Pt(13.5)
    p2.font.bold = True
    p2.font.color.rgb = C_PRIMARY
    p2.space_after = Pt(2)

    p3 = tf_c1.add_paragraph()
    p3.text = "JWT & RBAC Security  •  Google Gemini 2.5 Flash API"
    p3.font.name = FONT_HEADING
    p3.font.size = Pt(13.5)
    p3.font.bold = True
    p3.font.color.rgb = C_PRIMARY
    p3.space_after = Pt(8)

    p4 = tf_c1.add_paragraph()
    p4.text = "Production-ready RESTful architecture with query-level student data isolation."
    p4.font.name = FONT_BODY
    p4.font.size = Pt(12)
    p4.font.color.rgb = C_MUTED

    # Card 2: Project Team
    c2 = add_card(slide1, Inches(6.8), Inches(4.45), Inches(5.733), Inches(2.35), bg_color=C_WHITE, border_color=C_BORDER, radius=0.035)
    tf_c2 = c2.text_frame
    tf_c2.word_wrap = True
    tf_c2.margin_left = Inches(0.28)
    tf_c2.margin_top = Inches(0.22)
    tf_c2.margin_right = Inches(0.28)
    tf_c2.margin_bottom = Inches(0.2)

    t1 = tf_c2.paragraphs[0]
    t1.text = "PROJECT TEAM"
    t1.font.name = FONT_HEADING
    t1.font.size = Pt(11)
    t1.font.bold = True
    t1.font.color.rgb = C_TEAL
    t1.space_after = Pt(6)

    t2 = tf_c2.add_paragraph()
    t2.text = "Team Leader: Isaac Shiloah Herbert V"
    t2.font.name = FONT_HEADING
    t2.font.size = Pt(14)
    t2.font.bold = True
    t2.font.color.rgb = C_PRIMARY
    t2.space_after = Pt(6)

    t3 = tf_c2.add_paragraph()
    t3.text = "Team Members:"
    t3.font.name = FONT_BODY
    t3.font.size = Pt(12.5)
    t3.font.bold = True
    t3.font.color.rgb = C_SLATE
    t3.space_after = Pt(3)

    t4 = tf_c2.add_paragraph()
    t4.text = "Mohammed Irfan A   •   Bhuvaneswari"
    t4.font.name = FONT_BODY
    t4.font.size = Pt(12)
    t4.font.color.rgb = C_MUTED
    t4.space_after = Pt(2)

    t5 = tf_c2.add_paragraph()
    t5.text = "Bensam Gilbert E   •   Joel Edwin P"
    t5.font.name = FONT_BODY
    t5.font.size = Pt(12)
    t5.font.color.rgb = C_MUTED

    # Right side: Connected AI network graphic
    hub_x, hub_y = Inches(9.65), Inches(2.25)
    satellites = [
        ("Study Notes", Inches(8.2), Inches(1.25), C_SLATE),
        ("Summaries", Inches(11.1), Inches(1.25), C_TEAL),
        ("Flashcards", Inches(11.35), Inches(2.6), C_TEAL),
        ("MCQ Quizzes", Inches(10.7), Inches(3.65), C_TEAL),
        ("Study Plans", Inches(8.65), Inches(3.65), C_TEAL),
        ("MongoDB", Inches(8.0), Inches(2.55), C_SLATE)
    ]

    for label, sx, sy, col in satellites:
        conn = slide1.shapes.add_connector(
            pptx.enum.shapes.MSO_CONNECTOR.STRAIGHT,
            hub_x + Inches(0.58), hub_y + Inches(0.58),
            sx + Inches(0.62), sy + Inches(0.28)
        )
        conn.line.color.rgb = RGBColor(203, 213, 225)
        conn.line.width = Pt(1.5)

    hub = slide1.shapes.add_shape(MSO_SHAPE.OVAL, hub_x, hub_y, Inches(1.15), Inches(1.15))
    hub.fill.solid()
    hub.fill.fore_color.rgb = C_TEAL
    hub.line.color.rgb = C_WHITE
    hub.line.width = Pt(3)
    tf_hub = hub.text_frame
    tf_hub.word_wrap = False
    tf_hub.margin_left = tf_hub.margin_top = tf_hub.margin_right = tf_hub.margin_bottom = 0
    tf_hub.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_h = tf_hub.paragraphs[0]
    p_h.alignment = PP_ALIGN.CENTER
    p_h.text = "Gemini\n2.5 Flash"
    p_h.font.name = FONT_HEADING
    p_h.font.size = Pt(10)
    p_h.font.bold = True
    p_h.font.color.rgb = C_WHITE

    for label, sx, sy, col in satellites:
        node = add_card(slide1, sx, sy, Inches(1.25), Inches(0.56), bg_color=C_WHITE, border_color=col, border_width=1.5, radius=0.25)
        tf_n = node.text_frame
        tf_n.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_n = tf_n.paragraphs[0]
        p_n.alignment = PP_ALIGN.CENTER
        p_n.text = label
        p_n.font.name = FONT_HEADING
        p_n.font.size = Pt(10)
        p_n.font.bold = True
        p_n.font.color.rgb = col

    # =========================================================================
    # SLIDE 2: THE PROBLEM AND OUR SOLUTION
    # =========================================================================
    slide2 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide2)
    add_header(
        slide2,
        kicker_text="THE PROBLEM AND OUR SOLUTION",
        title_text="Smarter Study Preparation, One Platform",
        subtitle_text="Transforming fragmented, manual preparation into a streamlined AI-powered revision workflow."
    )

    card_y = Inches(1.98)
    card_h = Inches(5.0)

    # Left Card: THE PROBLEM
    p_card = add_card(slide2, Inches(0.8), card_y, Inches(5.4), card_h, bg_color=C_WHITE, border_color=C_ROSE_BORDER, border_width=1.5, radius=0.035)
    
    # Header Badge inside card
    p_badge = add_card(slide2, Inches(1.15), Inches(2.25), Inches(1.6), Inches(0.35), bg_color=C_ROSE_BG, border_color=C_ROSE_BORDER, radius=0.2)
    tf_pb = p_badge.text_frame
    tf_pb.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_b = tf_pb.paragraphs[0]
    p_b.alignment = PP_ALIGN.CENTER
    p_b.text = "THE PROBLEM"
    p_b.font.name = FONT_HEADING
    p_b.font.size = Pt(10.5)
    p_b.font.bold = True
    p_b.font.color.rgb = C_ROSE_TEXT

    # Problem Card Text
    p_tb = slide2.shapes.add_textbox(Inches(1.15), Inches(2.75), Inches(4.7), Inches(4.0))
    tf_p = p_tb.text_frame
    tf_p.word_wrap = True
    tf_p.margin_left = tf_p.margin_top = tf_p.margin_right = tf_p.margin_bottom = 0

    p_sub = tf_p.paragraphs[0]
    p_sub.text = "Fragmented Materials & Manual Bottlenecks"
    p_sub.font.name = FONT_HEADING
    p_sub.font.size = Pt(17)
    p_sub.font.bold = True
    p_sub.font.color.rgb = C_PRIMARY
    p_sub.space_after = Pt(20)

    prob_bullets = [
        ("Study materials are scattered across different sources", "Lecture slides, textbooks, PDFs, and notes are dispersed across multiple apps without unified organization."),
        ("Manually creating notes and questions takes time", "Synthesizing comprehensive summaries and drafting active-recall questions consumes critical hours needed for concept mastery."),
        ("Students lack structured workflows & self-assessment", "No adaptive preparation timelines matched to exam deadlines, and no convenient self-assessment tools to evaluate readiness.")
    ]

    for title, desc in prob_bullets:
        pb = tf_p.add_paragraph()
        r1 = pb.add_run()
        r1.text = "• " + title + ": "
        r1.font.name = FONT_HEADING
        r1.font.size = Pt(13)
        r1.font.bold = True
        r1.font.color.rgb = C_PRIMARY
        
        r2 = pb.add_run()
        r2.text = desc
        r2.font.name = FONT_BODY
        r2.font.size = Pt(12)
        r2.font.color.rgb = C_SLATE
        pb.space_after = Pt(14)

    # Center Flow Element: Sleek Arrow & Badge
    c_flow = add_card(slide2, Inches(6.32), Inches(4.1), Inches(0.68), Inches(0.42), bg_color=C_TEAL, border_color=None, radius=0.2)
    tf_cf = c_flow.text_frame
    tf_cf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_cf = tf_cf.paragraphs[0]
    p_cf.alignment = PP_ALIGN.CENTER
    p_cf.text = "➔"
    p_cf.font.name = FONT_HEADING
    p_cf.font.size = Pt(16)
    p_cf.font.bold = True
    p_cf.font.color.rgb = C_WHITE

    # Right Card: OUR SOLUTION
    s_card = add_card(slide2, Inches(7.133), card_y, Inches(5.4), card_h, bg_color=C_TEAL_LIGHT, border_color=C_TEAL, border_width=1.5, radius=0.035)

    # Solution Header Badge
    s_badge = add_card(slide2, Inches(7.48), Inches(2.25), Inches(1.6), Inches(0.35), bg_color=C_TEAL_BADGE, border_color=C_TEAL_BORDER, radius=0.2)
    tf_sb = s_badge.text_frame
    tf_sb.vertical_anchor = MSO_ANCHOR.MIDDLE
    s_b = tf_sb.paragraphs[0]
    s_b.alignment = PP_ALIGN.CENTER
    s_b.text = "OUR SOLUTION"
    s_b.font.name = FONT_HEADING
    s_b.font.size = Pt(10.5)
    s_b.font.bold = True
    s_b.font.color.rgb = C_TEAL

    # Solution Card Text
    s_tb = slide2.shapes.add_textbox(Inches(7.48), Inches(2.75), Inches(4.7), Inches(4.0))
    tf_s = s_tb.text_frame
    tf_s.word_wrap = True
    tf_s.margin_left = tf_s.margin_top = tf_s.margin_right = tf_s.margin_bottom = 0

    s_sub = tf_s.paragraphs[0]
    s_sub.text = "Centralized, Material-Grounded AI Platform"
    s_sub.font.name = FONT_HEADING
    s_sub.font.size = Pt(17)
    s_sub.font.bold = True
    s_sub.font.color.rgb = C_PRIMARY
    s_sub.space_after = Pt(16)

    sol_bullets = [
        ("Upload and manage study materials", "Centralize text notes and documents (.txt, .md, .pdf) in a secure, searchable repository."),
        ("Generate AI summaries, flashcards & quizzes", "Dynamically create concise conceptual summaries, active-recall flashcards, and 4-option MCQs directly from uploaded notes."),
        ("Personalized study plans and direct Q&A", "Construct custom day-by-day revision schedules matched to exam dates and ask targeted study questions."),
        ("Save learning resources for later revision", "Automatically persist all generated summaries, flashcards, and quizzes in MongoDB for convenient retrieval anytime.")
    ]

    for title, desc in sol_bullets:
        sb = tf_s.add_paragraph()
        r1 = sb.add_run()
        r1.text = "✓ " + title + ": "
        r1.font.name = FONT_HEADING
        r1.font.size = Pt(12.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEAL
        
        r2 = sb.add_run()
        r2.text = desc
        r2.font.name = FONT_BODY
        r2.font.size = Pt(11.5)
        r2.font.color.rgb = C_SLATE
        sb.space_after = Pt(9)

    # =========================================================================
    # SLIDE 3: SYSTEM ARCHITECTURE
    # =========================================================================
    slide3 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide3)
    add_header(
        slide3,
        kicker_text="SYSTEM ARCHITECTURE",
        title_text="How Study Buddy AI Works",
        subtitle_text="End-to-end request lifecycle: client requests, JWT verification, controllers, and dual backend services."
    )

    box_w = Inches(2.45)
    box_h = Inches(3.55)
    top_y = Inches(1.95)

    # Card 1: Client Layer
    c_client = add_card(slide3, Inches(0.8), top_y, box_w, box_h, bg_color=C_WHITE, border_color=C_BORDER, radius=0.035)
    tf = c_client.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.2)
    tf.margin_top = Inches(0.2)
    
    p = tf.paragraphs[0]
    p.text = "Student / Admin"
    p.font.name = FONT_HEADING
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = C_PRIMARY
    
    p = tf.add_paragraph()
    p.text = "Client Interface Layer"
    p.font.name = FONT_HEADING
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = C_TEAL
    p.space_after = Pt(10)

    for item in [
        "Web App (React) / Postman",
        "HTTPS + Bearer tokens",
        "Upload study materials",
        "Request summaries & quizzes",
        "Retrieve saved study aids"
    ]:
        p = tf.add_paragraph()
        p.text = "• " + item
        p.font.name = FONT_BODY
        p.font.size = Pt(11.5)
        p.font.color.rgb = C_SLATE
        p.space_after = Pt(4)

    # Arrow 1: Client -> Gateway (centered at y = 3.65)
    arr1 = slide3.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(3.32), Inches(3.62), Inches(0.42), Inches(0.26))
    arr1.fill.solid()
    arr1.fill.fore_color.rgb = C_TEAL
    arr1.line.fill.background()

    # Card 2: REST API Gateway
    c_gw = add_card(slide3, Inches(3.8), top_y, box_w, box_h, bg_color=C_WHITE, border_color=C_BORDER, radius=0.035)
    tf = c_gw.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.2)
    tf.margin_top = Inches(0.2)
    
    p = tf.paragraphs[0]
    p.text = "REST API Gateway"
    p.font.name = FONT_HEADING
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = C_PRIMARY
    
    p = tf.add_paragraph()
    p.text = "Node.js + Express.js"
    p.font.name = FONT_HEADING
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = C_TEAL
    p.space_after = Pt(10)

    for item in [
        "Central HTTP request router",
        "Helmet HTTP security headers",
        "CORS & Morgan access logs",
        "Two-tier API rate limiting",
        "Express 5 NoSQL guard",
        "Multer multipart file parser"
    ]:
        p = tf.add_paragraph()
        p.text = "• " + item
        p.font.name = FONT_BODY
        p.font.size = Pt(11.5)
        p.font.color.rgb = C_SLATE
        p.space_after = Pt(4)

    # Arrow 2: Gateway -> Auth & Controllers
    arr2 = slide3.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(6.32), Inches(3.62), Inches(0.42), Inches(0.26))
    arr2.fill.solid()
    arr2.fill.fore_color.rgb = C_TEAL
    arr2.line.fill.background()

    # Card 3: Auth & Controllers
    c_app = add_card(slide3, Inches(6.8), top_y, box_w, box_h, bg_color=C_WHITE, border_color=C_TEAL, border_width=1.5, radius=0.035)
    tf = c_app.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.2)
    tf.margin_top = Inches(0.2)
    
    p = tf.paragraphs[0]
    p.text = "Auth & Controllers"
    p.font.name = FONT_HEADING
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = C_PRIMARY
    
    p = tf.add_paragraph()
    p.text = "JWT + RBAC Middleware"
    p.font.name = FONT_HEADING
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = C_TEAL
    p.space_after = Pt(10)

    for item in [
        "7-day Bearer token validation",
        "Student / Admin role guard",
        "Query-level data isolation",
        "express-validator schema rules",
        "Normalized error handling",
        "Dispatches AI & DB services"
    ]:
        p = tf.add_paragraph()
        p.text = "• " + item
        p.font.name = FONT_BODY
        p.font.size = Pt(11.5)
        p.font.color.rgb = C_SLATE
        p.space_after = Pt(4)

    # Arrows from Auth/Controllers to Dual Backend Services
    arr_gem = slide3.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(9.32), Inches(2.65), Inches(0.42), Inches(0.24))
    arr_gem.fill.solid()
    arr_gem.fill.fore_color.rgb = C_TEAL
    arr_gem.line.fill.background()

    arr_db = slide3.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(9.32), Inches(4.55), Inches(0.42), Inches(0.24))
    arr_db.fill.solid()
    arr_db.fill.fore_color.rgb = C_TEAL
    arr_db.line.fill.background()

    # Upper Service Card: Google Gemini 2.5 Flash
    box_r_w = Inches(2.75)
    c_gem = add_card(slide3, Inches(9.8), Inches(1.95), box_r_w, Inches(1.72), bg_color=C_TEAL_LIGHT, border_color=C_TEAL, border_width=1.5, radius=0.035)
    tf = c_gem.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.18)
    tf.margin_top = Inches(0.14)

    p = tf.paragraphs[0]
    p.text = "Google Gemini 2.5 Flash"
    p.font.name = FONT_HEADING
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = C_TEAL
    
    p = tf.add_paragraph()
    p.text = "@google/genai Service Layer"
    p.font.name = FONT_HEADING
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = C_MUTED
    p.space_after = Pt(3)

    for item in [
        "Material-grounded dynamic prompts",
        "Summaries & study plans (prose)",
        "Flashcards & quizzes (strict JSON)"
    ]:
        p = tf.add_paragraph()
        p.text = "• " + item
        p.font.name = FONT_BODY
        p.font.size = Pt(10)
        p.font.color.rgb = C_SLATE
        p.space_after = Pt(1.5)

    # Lower Service Card: MongoDB & Mongoose ODM (Height 1.72", well padded)
    c_db = add_card(slide3, Inches(9.8), Inches(3.78), box_r_w, Inches(1.72), bg_color=C_WHITE, border_color=C_BORDER, radius=0.035)
    tf = c_db.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.18)
    tf.margin_top = Inches(0.14)

    p = tf.paragraphs[0]
    p.text = "MongoDB & Mongoose ODM"
    p.font.name = FONT_HEADING
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = C_PRIMARY
    
    p = tf.add_paragraph()
    p.text = "Data Persistence Layer"
    p.font.name = FONT_HEADING
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = C_TEAL
    p.space_after = Pt(3)

    for item in [
        "User credentials with bcrypt hashing",
        "Uploaded notes & files (.txt, .md, .pdf)",
        "Saves summaries, flashcards & quizzes"
    ]:
        p = tf.add_paragraph()
        p.text = "• " + item
        p.font.name = FONT_BODY
        p.font.size = Pt(10)
        p.font.color.rgb = C_SLATE
        p.space_after = Pt(1.5)

    # Bottom Full-Width Card: Persistent Resource Retrieval Flow
    # Y position 5.72 to 7.05 (Zero overlap, generous 0.22" margin)
    c_ret = add_card(slide3, Inches(0.8), Inches(5.72), Inches(11.75), Inches(1.32), bg_color=C_WHITE, border_color=C_TEAL_BORDER, border_width=1.5, radius=0.035)
    tf = c_ret.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.25)
    tf.margin_right = Inches(0.25)
    tf.margin_top = Inches(0.16)

    p = tf.paragraphs[0]
    r1 = p.add_run()
    r1.text = "PERSISTENT RESOURCE RETRIEVAL FLOW: "
    r1.font.name = FONT_HEADING
    r1.font.size = Pt(12)
    r1.font.bold = True
    r1.font.color.rgb = C_TEAL

    r2 = p.add_run()
    r2.text = "MongoDB ➔ Auth & Scoped Queries ({ userId }) ➔ REST API Gateway ➔ Authorized Student"
    r2.font.name = FONT_HEADING
    r2.font.size = Pt(12)
    r2.font.bold = True
    r2.font.color.rgb = C_PRIMARY
    p.space_after = Pt(4)

    p2 = tf.add_paragraph()
    p2.text = (
        "Authorized students can re-access, search, filter, and review all their generated resources at any time. "
        "When an AI feature executes, its structured response is persisted directly to MongoDB under the user's account "
        "— guaranteeing that learning assets are preserved for revision rather than lost after generation."
    )
    p2.font.name = FONT_BODY
    p2.font.size = Pt(11.5)
    p2.font.color.rgb = C_SLATE

    # =========================================================================
    # SLIDE 4: CORE FEATURES AND SECURITY
    # =========================================================================
    slide4 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide4)
    add_header(
        slide4,
        kicker_text="CORE FEATURES AND SECURITY",
        title_text="Features Built for Focused Learning",
        subtitle_text="High-yield academic study utilities paired with multi-layer defense-in-depth security."
    )

    feat_x = Inches(0.8)
    feat_w = Inches(3.55)
    feat_h = Inches(1.32)
    
    features = [
        # Col 1
        ("Secure Authentication", "User registration, bcrypt password hashing, and 7-day stateless JWT login with profile management.", 0, 0),
        ("Study Material Upload", "Upload study notes as text or files (.txt, .md, .pdf), with full-text search, subject filtering, and downloads.", 0, 1),
        ("AI-Generated Summaries", "High-yield conceptual summaries synthesized from uploaded notes for accelerated exam revision.", 0, 2),
        # Col 2
        ("Active-Recall Flashcards", "Structured question-and-answer pairs created via guaranteed JSON output for spaced repetition.", 1, 0),
        ("Multiple-Choice Quizzes", "Objective 4-option quizzes generated with verified answer keys to evaluate subject comprehension.", 1, 1),
        ("Personalized Plans & Q&A", "Custom day-by-day study schedules tailored to exam timelines, plus direct contextual study Q&A.", 1, 2),
    ]

    for title, desc, col, row in features:
        cx = feat_x + col * Inches(3.72)
        cy = Inches(2.0) + row * Inches(1.45)
        
        card = add_card(slide4, cx, cy, feat_w, feat_h, bg_color=C_WHITE, border_color=C_BORDER, radius=0.035)
        
        # Tiny teal vertical indicator on left
        acc = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx + Inches(0.12), cy + Inches(0.16), Inches(0.05), Inches(0.98))
        acc.fill.solid()
        acc.fill.fore_color.rgb = C_TEAL
        acc.line.fill.background()

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.28)
        tf.margin_top = Inches(0.16)
        tf.margin_right = Inches(0.18)
        tf.margin_bottom = Inches(0.12)

        p = tf.paragraphs[0]
        p.text = title
        p.font.name = FONT_HEADING
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = C_PRIMARY
        p.space_after = Pt(3)

        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.name = FONT_BODY
        p2.font.size = Pt(11)
        p2.font.color.rgb = C_SLATE

    # Admin note at bottom of features section
    f_note = slide4.shapes.add_textbox(feat_x, Inches(6.45), Inches(7.2), Inches(0.35))
    tf_fn = f_note.text_frame
    tf_fn.word_wrap = True
    tf_fn.margin_left = tf_fn.margin_top = tf_fn.margin_right = tf_fn.margin_bottom = 0
    p = tf_fn.paragraphs[0]
    p.text = "✦ Includes administrator user oversight, role updating, and platform-wide AI usage analytics."
    p.font.name = FONT_BODY
    p.font.size = Pt(11.5)
    p.font.color.rgb = C_TEAL

    # Right Section: Security & Data Protection Card
    sec_x = Inches(8.45)
    sec_w = Inches(4.08)
    sec_h = Inches(4.75)
    
    sec_card = add_card(slide4, sec_x, Inches(2.0), sec_w, sec_h, bg_color=C_TEAL_LIGHT, border_color=C_TEAL, border_width=1.5, radius=0.035)

    tf_sec = sec_card.text_frame
    tf_sec.word_wrap = True
    tf_sec.margin_left = Inches(0.28)
    tf_sec.margin_top = Inches(0.24)
    tf_sec.margin_right = Inches(0.28)
    tf_sec.margin_bottom = Inches(0.2)

    p = tf_sec.paragraphs[0]
    p.text = "SECURITY & DATA PROTECTION"
    p.font.name = FONT_HEADING
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = C_TEAL
    p.space_after = Pt(2)

    p_sub = tf_sec.add_paragraph()
    p_sub.text = "Defense-in-Depth Security"
    p_sub.font.name = FONT_HEADING
    p_sub.font.size = Pt(16)
    p_sub.font.bold = True
    p_sub.font.color.rgb = C_PRIMARY
    p_sub.space_after = Pt(12)

    sec_items = [
        ("JWT Authentication", "Stateless Bearer tokens with 7-day expiration protect private endpoints."),
        ("Bcrypt Password Hashing", "Passwords salted with 10 rounds; plaintext credentials never reach the database."),
        ("Role-Based Access Control", "Server-enforced Student vs Admin roles; self-promotion strictly blocked."),
        ("User-Scoped Data Access", "Queries enforce { userId: req.user._id }; prevents horizontal unauthorized access."),
        ("Input Guard & Rate Limiting", "Express 5 NoSQL-injection guard plus two-tier rate limiting on AI quotas.")
    ]

    for title, desc in sec_items:
        p = tf_sec.add_paragraph()
        r1 = p.add_run()
        r1.text = "• " + title + ": "
        r1.font.name = FONT_HEADING
        r1.font.size = Pt(12)
        r1.font.bold = True
        r1.font.color.rgb = C_PRIMARY

        r2 = p.add_run()
        r2.text = desc
        r2.font.name = FONT_BODY
        r2.font.size = Pt(11)
        r2.font.color.rgb = C_SLATE
        p.space_after = Pt(8)

    # =========================================================================
    # SLIDE 5: IMPLEMENTATION, VALIDATION AND PROJECT LINKS
    # =========================================================================
    slide5 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide5)
    add_header(
        slide5,
        kicker_text="IMPLEMENTATION & VALIDATION",
        title_text="Implementation and Project Demonstration",
        subtitle_text="Verified backend API endpoints, functional Postman validation, and project resources."
    )

    card_w = Inches(3.72)
    card_h = Inches(3.2)
    card_y = Inches(1.9)

    screens = [
        ("User Authentication (201 Created)", "D:/slide_assets/crop_registration.png", "Bcrypt password hashing & 7-day JWT Bearer token issuance.", 0),
        ("Material Upload (201 Created)", "D:/slide_assets/crop_upload.png", "Multipart form-data (.txt) upload securely parsed & stored in MongoDB.", 1),
        ("AI Flashcards (200 OK)", "D:/slide_assets/crop_flashcards.png", "Gemini 2.5 Flash generating structured active-recall Q&A arrays.", 2),
    ]

    for title, img_path, caption, idx in screens:
        cx = Inches(0.8) + idx * Inches(4.01)
        
        sc = add_card(slide5, cx, card_y, card_w, card_h, bg_color=C_WHITE, border_color=C_BORDER, radius=0.035)

        # Card Title
        tb_t = slide5.shapes.add_textbox(cx + Inches(0.15), card_y + Inches(0.1), card_w - Inches(0.3), Inches(0.3))
        tf_t = tb_t.text_frame
        tf_t.word_wrap = True
        tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0
        p = tf_t.paragraphs[0]
        p.text = title
        p.font.name = FONT_HEADING
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = C_PRIMARY

        # Cropped Screenshot Image
        if os.path.exists(img_path):
            slide5.shapes.add_picture(img_path, cx + Inches(0.15), card_y + Inches(0.42), card_w - Inches(0.3), Inches(2.15))

        # Caption
        tb_c = slide5.shapes.add_textbox(cx + Inches(0.15), card_y + Inches(2.65), card_w - Inches(0.3), Inches(0.5))
        tf_c = tb_c.text_frame
        tf_c.word_wrap = True
        tf_c.margin_left = tf_c.margin_top = tf_c.margin_right = tf_c.margin_bottom = 0
        p = tf_c.paragraphs[0]
        p.text = caption
        p.font.name = FONT_BODY
        p.font.size = Pt(10.5)
        p.font.color.rgb = C_SLATE

    # Postman API Verification Statement Box
    v_box = add_card(slide5, Inches(0.8), Inches(5.25), Inches(11.75), Inches(0.58), bg_color=C_WHITE, border_color=C_BORDER, radius=0.035)
    tf_v = v_box.text_frame
    tf_v.word_wrap = True
    tf_v.margin_left = Inches(0.2)
    tf_v.margin_right = Inches(0.2)
    tf_v.margin_top = Inches(0.12)
    
    p = tf_v.paragraphs[0]
    r1 = p.add_run()
    r1.text = "API Verification: "
    r1.font.name = FONT_HEADING
    r1.font.size = Pt(12)
    r1.font.bold = True
    r1.font.color.rgb = C_TEAL

    r2 = p.add_run()
    r2.text = (
        "The backend RESTful API was comprehensively tested using Postman, verifying route protection, "
        "role authorization, multipart uploads, and Gemini AI JSON schema validation across all endpoints."
    )
    r2.font.name = FONT_BODY
    r2.font.size = Pt(11.5)
    r2.font.color.rgb = C_SLATE

    # Links & Closing Statement Card
    link_card = add_card(slide5, Inches(0.8), Inches(5.95), Inches(11.75), Inches(1.15), bg_color=C_TEAL_LIGHT, border_color=C_TEAL_BORDER, border_width=1.5, radius=0.035)

    tf_l = link_card.text_frame
    tf_l.word_wrap = True
    tf_l.margin_left = Inches(0.25)
    tf_l.margin_right = Inches(0.25)
    tf_l.margin_top = Inches(0.14)

    # Clickable Links Paragraph
    p_link = tf_l.paragraphs[0]
    
    r_gh_lbl = p_link.add_run()
    r_gh_lbl.text = "GitHub Repository: "
    r_gh_lbl.font.name = FONT_HEADING
    r_gh_lbl.font.size = Pt(12)
    r_gh_lbl.font.bold = True
    r_gh_lbl.font.color.rgb = C_PRIMARY

    r_gh_url = p_link.add_run()
    r_gh_url.text = "https://github.com/isaacincognito007-cloud/AI-StudyBuddy-"
    r_gh_url.font.name = FONT_HEADING
    r_gh_url.font.size = Pt(12)
    r_gh_url.font.bold = True
    r_gh_url.font.color.rgb = C_TEAL
    r_gh_url.font.underline = True
    r_gh_url.hyperlink.address = "https://github.com/isaacincognito007-cloud/AI-StudyBuddy-"

    r_sep = p_link.add_run()
    r_sep.text = "    |    "
    r_sep.font.name = FONT_HEADING
    r_sep.font.size = Pt(12)
    r_sep.font.color.rgb = C_MUTED

    r_vid_lbl = p_link.add_run()
    r_vid_lbl.text = "Demo Video: "
    r_vid_lbl.font.name = FONT_HEADING
    r_vid_lbl.font.size = Pt(12)
    r_vid_lbl.font.bold = True
    r_vid_lbl.font.color.rgb = C_PRIMARY

    r_vid_url = p_link.add_run()
    r_vid_url.text = "https://youtu.be/N5lNC06rn7w"
    r_vid_url.font.name = FONT_HEADING
    r_vid_url.font.size = Pt(12)
    r_vid_url.font.bold = True
    r_vid_url.font.color.rgb = C_TEAL
    r_vid_url.font.underline = True
    r_vid_url.hyperlink.address = "https://youtu.be/N5lNC06rn7w"
    
    p_link.space_after = Pt(4)

    # Closing statement
    p_close = tf_l.add_paragraph()
    p_close.text = "“Study Buddy AI — Turning Study Materials into Structured Learning Resources.”"
    p_close.font.name = FONT_HEADING
    p_close.font.size = Pt(15.5)
    p_close.font.bold = True
    p_close.font.color.rgb = C_TEAL

    prs.save(output_path)
    print(f"Successfully created presentation at {output_path}")

if __name__ == '__main__':
    build_presentation(r'D:\AI-StudyBuddy\Study_Buddy_AI_Presentation.pptx')
