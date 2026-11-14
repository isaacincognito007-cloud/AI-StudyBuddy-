with open(r'C:\Users\acer\.gemini\antigravity\brain\3df468b6-62f3-4271-a7b4-01a98becf01a\scratch\create_one_page_abstract.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

replacement_code = """# -------------------------------------------------------------
# 2. PROJECT TITLE & METADATA (As clean bold points, no highlight box)
# -------------------------------------------------------------
meta_items = [
    ("Project Title", "STUDY BUDDY AI (LearnMate) — AI Assistant for Smarter Learning and Study Support"),
    ("Subject & Degree", "AI-Augmented Backend Development - A&S (NMU88)  |  BCA (III Year – V Semester)"),
    ("Team Leader", "ISAAC SHILOAH HERBERT V (212404414)"),
    ("Team Members", "MOHAMMED IRFAN A (212404420), BHUVANESWARI A (212404441), BENSAM GILBERT E (212404405), JOEL EDWIN P (212404416)")
]

for idx, (lbl, val) in enumerate(meta_items):
    p_item = doc.add_paragraph()
    p_item.paragraph_format.space_before = Pt(1)
    p_item.paragraph_format.space_after = Pt(2)
    p_item.paragraph_format.line_spacing = 1.12
    p_item.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    r_dot = p_item.add_run("• ")
    r_dot.font.name = "Calibri"
    r_dot.font.size = Pt(9.5)
    r_dot.font.bold = True
    r_dot.font.color.rgb = COLOR_TEAL
    
    r_lbl = p_item.add_run(lbl + ": ")
    r_lbl.font.name = "Calibri"
    r_lbl.font.size = Pt(9.5)
    r_lbl.font.bold = True
    r_lbl.font.color.rgb = COLOR_NAVY
    
    r_val = p_item.add_run(val)
    r_val.font.name = "Calibri"
    r_val.font.size = Pt(9)
    r_val.font.bold = (idx == 0 or idx == 2)
    r_val.font.color.rgb = COLOR_CHARCOAL
"""

# Replace lines 80 to 133
new_lines = lines[:80] + [replacement_code + '\n'] + lines[133:]
with open(r'C:\Users\acer\.gemini\antigravity\brain\3df468b6-62f3-4271-a7b4-01a98becf01a\scratch\create_one_page_abstract.py', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print('Successfully updated create_one_page_abstract.py!')
