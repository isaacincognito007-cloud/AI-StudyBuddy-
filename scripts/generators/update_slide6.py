with open(r'C:\Users\acer\.gemini\antigravity\brain\3df468b6-62f3-4271-a7b4-01a98becf01a\scratch\build_creative_8_slides_ppt.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_s6 = """schemas_2x2 = [
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
"""

new_lines = lines[:642] + [new_s6 + '\n'] + lines[701:]
with open(r'C:\Users\acer\.gemini\antigravity\brain\3df468b6-62f3-4271-a7b4-01a98becf01a\scratch\build_creative_8_slides_ppt.py', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print('Updated Slide 6 to 2x2 grid successfully!')
