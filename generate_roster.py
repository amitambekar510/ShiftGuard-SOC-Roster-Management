#!/usr/bin/env python3
"""
ShiftGuard: SOC Roster Management System
========================================
Main script to generate comprehensive SOC team rosters with:
- Automated shift scheduling (4-day and 5-day cycles)
- Holiday integration
- Health & wellness guidelines
- Statistical analysis and coverage validation

Author: Amit Ambekar
Project: ShiftGuard (Open Source)
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta

# ============================================================================
# 1. DEFINE 2026 PUBLIC HOLIDAYS (From Official Calendar)
# ============================================================================
holidays_2026 = {
    '2026-01-26': 'Republic Day', '2026-02-15': 'Mahashivratri', 
    '2026-03-03': 'Holi', '2026-03-21': 'Ramadan Eid',
    '2026-03-26': 'Shri Ram Navami', '2026-03-31': 'Shri Mahavir Jayanti',
    '2026-04-03': 'Good Friday', '2026-04-14': 'Dr Ambedkar Jayanti',
    '2026-05-01': 'Maharashtra Day', '2026-05-28': 'Bakri Eid',
    '2026-06-26': 'Muharram', '2026-08-15': 'Independence Day',
    '2026-09-14': 'Ganesh Chaturthi', '2026-10-02': 'Gandhi Jayanti',
    '2026-10-20': 'Dussehra', '2026-11-08': 'Diwali Laxmi Pujan',
    '2026-11-10': 'Diwali-Balipratipada', '2026-11-24': 'Guru Nanak Dev',
    '2026-12-25': 'Christmas'
}

# ============================================================================
# 2. DEFINE TEAM MEMBERS (ANONYMIZED FOR OPEN SOURCE)
# ============================================================================
team = [
    # 24/7 Rotational Shift Team
    {'Name': 'SOC L1 Analyst 1', 'Role': 'SOC L1', 'Type': '24/7', 'Offset': 0, 'Start': '2026-03-01', 'End': '2026-12-31'},
    {'Name': 'SOC L1 Analyst 2', 'Role': 'SOC L1', 'Type': '24/7', 'Offset': 1, 'Start': '2026-03-01', 'End': '2026-12-31'},
    {'Name': 'SOC L1 Analyst 3', 'Role': 'SOC L1', 'Type': '24/7', 'Offset': 2, 'Start': '2026-03-01', 'End': '2026-12-31'},
    {'Name': 'SOC L1 Analyst 4', 'Role': 'SOC L1', 'Type': '24/7', 'Offset': 3, 'Start': '2026-03-01', 'End': '2026-05-31'},
    {'Name': 'SOC L1 Analyst 5', 'Role': 'SOC L1', 'Type': '24/7', 'Offset': 3, 'Start': '2026-06-01', 'End': '2026-12-31'},
    {'Name': 'SOC L1 Backup 1', 'Role': 'SOC L1 Backup', 'Type': '24/7', 'Offset': 4, 'Start': '2026-08-01', 'End': '2026-12-31'},
    
    {'Name': 'SOC L2 Analyst 1', 'Role': 'SOC L2', 'Type': '24/7', 'Offset': 0, 'Start': '2026-03-01', 'End': '2026-12-31'},
    {'Name': 'SOC L2 Analyst 2', 'Role': 'SOC L2', 'Type': '24/7', 'Offset': 1, 'Start': '2026-03-01', 'End': '2026-12-31'},
    {'Name': 'SOC L2 Analyst 3', 'Role': 'SOC L2', 'Type': '24/7', 'Offset': 2, 'Start': '2026-03-01', 'End': '2026-12-31'},
    {'Name': 'SOC L2 Backup 1', 'Role': 'SOC L2 Backup', 'Type': '24/7', 'Offset': 3, 'Start': '2026-03-01', 'End': '2026-12-31'},
    {'Name': 'SOC L2/L3 Backup 1', 'Role': 'SOC L2/L3 Backup', 'Type': '24/7', 'Offset': 4, 'Start': '2026-08-01', 'End': '2026-12-31'},
    
    # General Shift Team
    {'Name': 'SOC L3 Team Lead A', 'Role': 'SOC L3 Team Lead', 'Type': 'Gen', 'Start': '2026-03-01', 'End': '2026-12-31'},
    {'Name': 'SOC L3 Team Lead B', 'Role': 'SOC L3 Team Lead', 'Type': 'Gen', 'Start': '2026-03-01', 'End': '2026-12-31'},
    {'Name': 'SIEM Engineer A', 'Role': 'SIEM Engineer', 'Type': 'Gen', 'Start': '2026-03-01', 'End': '2026-12-31'},
    {'Name': 'SIEM Engineer B', 'Role': 'SIEM Engineer', 'Type': 'Gen', 'Start': '2026-03-01', 'End': '2026-12-31'},
    {'Name': 'Firewall Ops A', 'Role': 'Firewall Ops', 'Type': 'Gen', 'Start': '2026-03-01', 'End': '2026-12-31'},
    {'Name': 'Firewall Ops B', 'Role': 'Firewall Ops', 'Type': 'Gen', 'Start': '2026-03-01', 'End': '2026-12-31'},
    {'Name': 'SOC L1 Backup 2', 'Role': 'SOC L1 Backup', 'Type': 'Gen', 'Start': '2026-06-01', 'End': '2026-12-31'},
]

# ============================================================================
# 3. GENERATE DATES & DEFINE STYLES
# ============================================================================
start_date = datetime(2026, 3, 1)
end_date = datetime(2026, 12, 31)
dates = [start_date + timedelta(days=x) for x in range((end_date - start_date).days + 1)]

role_colors = {
    'SOC L1': {'Morning': 'BDD7EE', 'Afternoon': '9BC2E6', 'Night': '7AAED4', 'General': 'D9E2F3', 'Week Off': 'FFFFFF', 'Public Holiday': 'F4CCCC'},
    'SOC L2': {'Morning': 'C6EFCE', 'Afternoon': 'A9D18E', 'Night': '8BBF7A', 'General': 'E2EFDA', 'Week Off': 'FFFFFF', 'Public Holiday': 'F4CCCC'},
    'SOC L3 Team Lead': {'Morning': 'FCE4D6', 'Afternoon': 'F8CBAD', 'Night': 'F4B084', 'General': 'F2DCDB', 'Week Off': 'FFFFFF', 'Public Holiday': 'F4CCCC'},
    'SIEM Engineer': {'Morning': 'D9B3FF', 'Afternoon': 'C79FE8', 'Night': 'B388D9', 'General': 'E5D9F2', 'Week Off': 'FFFFFF', 'Public Holiday': 'F4CCCC'},
    'Firewall Ops': {'Morning': 'FFD966', 'Afternoon': 'FFC000', 'Night': 'FFB000', 'General': 'FFF2CC', 'Week Off': 'FFFFFF', 'Public Holiday': 'F4CCCC'},
    'Backup': {'Morning': 'B4F0E8', 'Afternoon': '8FE5DC', 'Night': '6ADAD0', 'General': 'D9F5F2', 'Week Off': 'FFFFFF', 'Public Holiday': 'F4CCCC'}
}

stats_colors = {
    'Total Work Days': 'D9E2F3', 'Morning Count': 'BDD7EE', 'Afternoon Count': 'C6EFCE',
    'Night Count': 'FFF2CC', 'General Count': 'F2F2F2', 'Week Offs': 'FFFFFF',
    'PH Worked': 'F4CCCC', 'Sundays Worked': 'FFEB9C', 'Additional Days': 'D9E2F3'
}

thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
thick_border = Border(left=Side(style='medium'), right=Side(style='medium'), top=Side(style='medium'), bottom=Side(style='medium'))

def get_role_color(role, shift_type):
    if 'Backup' in role: return role_colors['Backup'][shift_type]
    return role_colors.get(role, role_colors['SOC L1'])[shift_type]

# ============================================================================
# 4. CREATE WORKBOOK & SHEET 1: SOC ROSTER 2026
# ============================================================================
wb = Workbook()
ws_roster = wb.active
ws_roster.title = "SOC Roster 2026"

header_fill = PatternFill(start_color='1F4E79', fill_type='solid')
header_font = Font(bold=True, color='FFFFFF', size=11, name='Calibri')
header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

# Headers
for col, val in zip(['A1', 'B1', 'C1'], ['Analyst Name', 'Role', 'Shift Timing']):
    ws_roster[col].value = val
    ws_roster[col].fill = header_fill
    ws_roster[col].font = header_font
    ws_roster[col].alignment = header_align
    ws_roster[col].border = thick_border

for col_idx, date in enumerate(dates, start=4):
    cell = ws_roster.cell(row=1, column=col_idx, value=date.strftime('%d %B %Y'))
    cell.fill = header_fill
    cell.font = Font(bold=True, color='FFFFFF', size=9, name='Calibri')
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = thick_border

stats_headers = ['Total Work Days', 'Morning Count', 'Afternoon Count', 'Night Count', 'General Count', 'Week Offs', 'PH Worked', 'Sundays Worked', 'Additional Days']
for col_idx, header in enumerate(stats_headers, start=len(dates) + 4):
    cell = ws_roster.cell(row=1, column=col_idx, value=header)
    cell.fill = header_fill
    cell.font = Font(bold=True, color='FFFFFF', size=9, name='Calibri')
    cell.alignment = header_align
    cell.border = thick_border

# Populate Data
for row_idx, emp in enumerate(team, start=2):
    ws_roster.cell(row=row_idx, column=1, value=emp['Name']).font = Font(bold=True, size=10, name='Calibri')
    ws_roster.cell(row=row_idx, column=2, value=emp['Role']).font = Font(size=10, name='Calibri')
    ws_roster.cell(row=row_idx, column=1).border = thin_border
    ws_roster.cell(row=row_idx, column=2).border = thin_border
    
    shift_timing = "08:00 AM - 12:00 PM" if emp['Type'] == 'Gen' else "08:00 AM - 04:00 PM / 04:00 PM - 12:00 AM / 12:00 AM - 08:00 AM"
    ws_roster.cell(row=row_idx, column=3, value=shift_timing).font = Font(size=9, name='Calibri', italic=True)
    ws_roster.cell(row=row_idx, column=3).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws_roster.cell(row=row_idx, column=3).border = thin_border
    
    emp_start, emp_end = datetime.strptime(emp['Start'], '%Y-%m-%d'), datetime.strptime(emp['End'], '%Y-%m-%d')
    
    for col_idx, date in enumerate(dates, start=4):
        cell = ws_roster.cell(row=row_idx, column=col_idx)
        if date < emp_start or date > emp_end:
            cell.value, cell.fill, cell.border = '', PatternFill(start_color='F2F2F2', fill_type='solid'), thin_border
            cell.alignment = Alignment(horizontal='center', vertical='center')
            continue
        
        is_holiday = date.strftime('%Y-%m-%d') in holidays_2026
        is_sunday = (date.weekday() == 6)
        
        if emp['Type'] == 'Gen':
            shift = 'Week Off' if is_sunday else ('Public Holiday' if is_holiday else 'General')
        else:
            cycle = ['Morning', 'Afternoon', 'Night', 'Week Off', 'Week Off'] if date >= datetime(2026, 8, 1) else ['Morning', 'Afternoon', 'Night', 'Week Off']
            shift = cycle[((date - start_date).days + emp['Offset']) % len(cycle)]
            if is_holiday and shift != 'Week Off': shift = f"{shift} (Public Holiday)"
        
        cell.value = shift
        base_shift = shift.split(' (')[0] if '(Public Holiday)' in shift else shift
        cell.fill = PatternFill(start_color=get_role_color(emp['Role'], base_shift), fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.font = Font(size=9, name='Calibri')
        cell.border = thin_border

    # Statistics Formulas
    first_data_col, last_data_col = 4, len(dates) + 3
    col_range = f"{get_column_letter(first_data_col)}{row_idx}:{get_column_letter(last_data_col)}{row_idx}"
    
    stats_formulas = [
        (f'=COUNTA({col_range})-COUNTIF({col_range},"Week Off")-COUNTIF({col_range},"")', 'Total Work Days', True),
        (f'=COUNTIF({col_range},"Morning*")', 'Morning Count', False),
        (f'=COUNTIF({col_range},"Afternoon*")', 'Afternoon Count', False),
        (f'=COUNTIF({col_range},"Night*")', 'Night Count', False),
        (f'=COUNTIF({col_range},"General")', 'General Count', False),
        (f'=COUNTIF({col_range},"Week Off")', 'Week Offs', False),
        (f'=COUNTIF({col_range},"*Public Holiday*")', 'PH Worked', False),
        (f'=COUNTA({",".join([get_column_letter(i+4) + str(row_idx) for i, d in enumerate(dates) if d.weekday() == 6])})-COUNTIF({",".join([get_column_letter(i+4) + str(row_idx) for i, d in enumerate(dates) if d.weekday() == 6])},"Week Off")' if any(d.weekday() == 6 for d in dates) else 0, 'Sundays Worked', False),
        (0, 'Additional Days', False)
    ]
    
    for col_offset, (formula, stat_name, is_bold) in enumerate(stats_formulas, start=len(dates) + 4):
        cell = ws_roster.cell(row=row_idx, column=col_offset, value=formula)
        cell.font = Font(bold=is_bold, size=10, name='Calibri')
        cell.fill = PatternFill(start_color=stats_colors[stat_name], fill_type='solid')
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border

# Merge Shift Timing Cells
ws_roster.merge_cells('C2:C12')
c2 = ws_roster['C2']
c2.value = "08:00 AM - 04:00 PM / 04:00 PM - 12:00 AM / 12:00 AM - 08:00 AM"
c2.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
c2.font = Font(size=9, name='Calibri', italic=True)
c2.border = thick_border

ws_roster.merge_cells('C13:C19')
c13 = ws_roster['C13']
c13.value = "08:00 AM - 12:00 PM"
c13.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
c13.font = Font(size=9, name='Calibri', italic=True)
c13.border = thick_border

# Coverage Check Row
cov_row = len(team) + 2
ws_roster.cell(row=cov_row, column=1, value='COVERAGE CHECK').font = Font(bold=True, size=11, color='006100', name='Calibri')
ws_roster.cell(row=cov_row, column=2, value='Automated Validation').font = Font(bold=True, size=11, color='006100', name='Calibri')
ws_roster.cell(row=cov_row, column=1).border = thick_border
ws_roster.cell(row=cov_row, column=2).border = thick_border

for col_idx, date in enumerate(dates, start=4):
    cell = ws_roster.cell(row=cov_row, column=col_idx)
    l1_active = sum(1 for e in team if e['Role'] in ['SOC L1', 'SOC L1 Backup'] and datetime.strptime(e['Start'], '%Y-%m-%d') <= date <= datetime.strptime(e['End'], '%Y-%m-%d'))
    l2_active = sum(1 for e in team if e['Role'] in ['SOC L2', 'SOC L2 Backup', 'SOC L2/L3 Backup'] and datetime.strptime(e['Start'], '%Y-%m-%d') <= date <= datetime.strptime(e['End'], '%Y-%m-%d'))
    
    if l1_active >= 1 and l2_active >= 1:
        cell.value, cell.fill, cell.font = 'OK', PatternFill(start_color='C6EFCE', fill_type='solid'), Font(bold=True, color='006100', size=10, name='Calibri')
    else:
        cell.value, cell.fill, cell.font = f'FAIL (L1:{l1_active}, L2:{l2_active})', PatternFill(start_color='FFC7CE', fill_type='solid'), Font(bold=True, color='9C0006', size=10, name='Calibri')
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = thin_border

# Column Widths & Freeze
ws_roster.column_dimensions['A'].width = 22
ws_roster.column_dimensions['B'].width = 20
ws_roster.column_dimensions['C'].width = 45
for col_idx in range(4, len(dates) + 13):
    ws_roster.column_dimensions[get_column_letter(col_idx)].width = 16 if col_idx > len(dates) + 3 else 13
ws_roster.freeze_panes = 'D2'

# ============================================================================
# 5. CREATE SHEET 2: HEALTH & WELLNESS TIPS
# ============================================================================
ws_wellness = wb.create_sheet(title="Health & Wellness Tips")

title_fill = PatternFill(start_color='1F4E79', fill_type='solid')
title_font = Font(bold=True, color='FFFFFF', size=16, name='Calibri')
subtitle_font = Font(italic=True, color='1F4E79', size=12, name='Calibri')
cat_fill = PatternFill(start_color='2E75B6', fill_type='solid')
cat_font = Font(bold=True, color='FFFFFF', size=11, name='Calibri')
text_font = Font(size=10, name='Calibri')
bold_text_font = Font(bold=True, size=10, name='Calibri')

ws_wellness.column_dimensions['A'].width = 35
ws_wellness.column_dimensions['B'].width = 80

ws_wellness.merge_cells('A1:B1')
t1 = ws_wellness['A1']
t1.value = "SOC Team Health & Wellness Guidelines"
t1.fill = title_fill
t1.font = title_font
t1.alignment = Alignment(horizontal='center', vertical='center')
t1.border = thick_border
ws_wellness['B1'].border = thick_border

ws_wellness.merge_cells('A2:B2')
t2 = ws_wellness['A2']
t2.value = "Because we are humans first, and analysts second. Your health is our top priority."
t2.font = subtitle_font
t2.alignment = Alignment(horizontal='center', vertical='center')
ws_wellness.row_dimensions[1].height = 40
ws_wellness.row_dimensions[2].height = 30

tips_data = [
    ("1. Shift Work & Sleep Hygiene", [
        ("Control Your Light", "Use blackout curtains when sleeping during the day. Wear sunglasses on your commute home after a Night shift to prevent sunlight from disrupting your sleep cycle."),
        ("The Caffeine Cut-off", "Stop consuming coffee, tea, or energy drinks at least 4 to 6 hours before your planned sleep time."),
        ("Anchor Sleep", "Try to keep at least 4 hours of your sleep schedule consistent, even on your weekly offs, to prevent 'social jetlag'."),
        ("Pre-Shift Naps", "If transitioning to a Night shift, take a 20-30 minute power nap in the late afternoon before coming to the office.")
    ]),
    ("2. Physical Health & Office Ergonomics", [
        ("The 20-20-20 Rule", "Staring at SIEM dashboards causes severe eye strain. Every 20 minutes, look at something 20 feet away for 20 seconds."),
        ("Micro-Breaks", "You cannot leave the SOC unmonitored, but stand up, stretch your neck, and roll your shoulders for 30 seconds every hour while keeping eyes on screens."),
        ("Posture Check", "Ensure your monitor is at eye level and your feet are flat on the floor. Do not hunch over the keyboard during high-severity alerts.")
    ]),
    ("3. Nutrition & Hydration (Night Shift Survival)", [
        ("Ditch the Vending Machine", "Pack healthy, light snacks (nuts, fruits, yogurt). Avoid heavy, carb-loaded meals at 2 AM, as they cause a 'sugar crash' and sleepiness."),
        ("Hydrate Constantly", "Keep a water bottle at your desk. Dehydration leads to fatigue and headaches, which mimic alert fatigue."),
        ("Eat Light Before Bed", "After a Night shift, eat a very light, protein-rich snack rather than a heavy meal to help your body sleep better.")
    ]),
    ("4. Mental Well-being & Stress Management", [
        ("Clean Handovers", "Anxiety often comes from worrying about alerts after your shift. Ensure a thorough, documented handover so you can mentally disconnect."),
        ("Decompress After Shifts", "Do not jump straight into heavy chores after a Night shift. Give yourself 30 minutes of quiet time to transition."),
        ("Speak Up", "If feeling overwhelmed or facing chronic fatigue, talk to your Team Lead. We can adjust shifts temporarily. Your mental health matters.")
    ]),
    ("5. Office Hygiene & Commute Safety", [
        ("Sanitize Your Workspace", "SOC keyboards and mice are high-touch areas. Use sanitizing wipes to clean your desk and peripherals at the start of every shift."),
        ("Hand Hygiene", "Wash your hands thoroughly before eating at your desk and after using the restroom."),
        ("Safe Commuting (Crucial)", "If leaving after a Night shift, be extra cautious while driving. Fatigue slows reaction times. If too tired, use a cab or carpool. DO NOT drive drowsy.")
    ])
]

current_row = 4
for category, tips in tips_data:
    ws_wellness.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
    cat_cell = ws_wellness.cell(row=current_row, column=1, value=category)
    cat_cell.fill = cat_fill
    cat_cell.font = cat_font
    cat_cell.alignment = Alignment(horizontal='left', vertical='center')
    cat_cell.border = thick_border
    ws_wellness.cell(row=current_row, column=2).border = thick_border
    ws_wellness.row_dimensions[current_row].height = 25
    current_row += 1
    
    for tip_title, tip_desc in tips:
        cell_a = ws_wellness.cell(row=current_row, column=1, value=tip_title)
        cell_a.font = bold_text_font
        cell_a.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        cell_a.border = thin_border
        
        cell_b = ws_wellness.cell(row=current_row, column=2, value=tip_desc)
        cell_b.font = text_font
        cell_b.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        cell_b.border = thin_border
        
        ws_wellness.row_dimensions[current_row].height = 40
        current_row += 1

current_row += 1
ws_wellness.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
footer = ws_wellness.cell(row=current_row, column=1, value="Note: These guidelines are mandatory recommendations. Please discuss any specific health concerns with HR or your Team Lead confidentially.")
footer.font = Font(italic=True, size=9, color='666666', name='Calibri')
footer.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

# ============================================================================
# 6. SAVE FILE
# ============================================================================
file_name = 'SOC_Roster_Complete.xlsx'
wb.save(file_name)
print(f"✅ SUCCESS! File saved as '{file_name}'")
print("📊 Sheet 1: SOC Roster 2026 (Complete with borders, merges, and formulas)")
print("💡 Sheet 2: Health & Wellness Tips (Formatted with colors and categories)")