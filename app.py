#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pattern Worksheet Generator - Final Layout Adjusted
- A4 Size specified
- Unscramble: Added space for writing (between text and line)
- Unscramble: Optimized item spacing to fit footer on single page
- Footer included
"""

from flask import Flask, render_template, request, send_file, jsonify
import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch, mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
import os
from datetime import datetime

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
DB_PATH = os.path.join(BASE_DIR, 'database.xlsx')

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# --- 폰트 설정 ---
def setup_korean_font():
    try:
        local_font = os.path.join(BASE_DIR, 'fonts', 'NanumGothic.ttf')
        if os.path.exists(local_font):
            pdfmetrics.registerFont(TTFont('KoreanFont', local_font))
            return 'KoreanFont'
        return 'Helvetica' 
    except:
        return 'Helvetica'

KOREAN_FONT = setup_korean_font()

# --- 엑셀 읽기 ---
def load_patterns_from_excel(excel_path):
    wb = openpyxl.load_workbook(excel_path)
    
    # Overview
    ws_overview = wb["Pattern Overview"]
    pattern_info = {}
    for row in ws_overview.iter_rows(min_row=2, values_only=True):
        if row[0] is not None:
            pattern_info[int(row[0])] = {
                'number': int(row[0]),
                'name': str(row[1]),
                'unit': str(row[3]) if len(row) > 3 and row[3] else 'Level A'
            }
            
    # Details
    ws_detail = wb["Pattern Details"]
    patterns = {}
    for row in ws_detail.iter_rows(min_row=2, values_only=True):
        try:
            p_num = int(row[0])
            section = row[2]
            content = row[4]
            
            if p_num not in patterns:
                patterns[p_num] = {
                    'pattern_num': p_num,
                    'pattern_name': pattern_info.get(p_num, {}).get('name', ''),
                    'unit': pattern_info.get(p_num, {}).get('unit', 'Level A'),
                    'speaking1': [], 'speaking2': [], 'unscramble': []
                }
            
            if section == 'Speaking I':
                patterns[p_num]['speaking1'].append(content)
            elif section == 'Speaking II':
                patterns[p_num]['speaking2'].append(content)
            elif section == 'Unscramble':
                scrambled = row[6].strip('()') if row[6] else ""
                patterns[p_num]['unscramble'].append((content, scrambled))
        except:
            continue
            
    return patterns

def distribute_questions(selected_patterns, target_count=5):
    result = {'speaking1': [], 'speaking2': [], 'unscramble': []}
    if not selected_patterns: return result
    
    pattern_count = len(selected_patterns)
    items_per = target_count // pattern_count
    remainder = target_count % pattern_count
    
    for key in result.keys():
        for i, p in enumerate(selected_patterns):
            count = items_per + (1 if i < remainder else 0)
            result[key].extend(p[key][:count])
            
    return result

# --- PDF 생성 (간격 미세 조정) ---
def create_worksheet(pattern_data, selected_patterns, output_path):
    # A4 용지 사용 (Letter보다 세로가 조금 더 김)
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        topMargin=15*mm,    # 위쪽 여백 약간 축소
        bottomMargin=15*mm, # 아래쪽 여백 약간 축소
        leftMargin=15*mm,
        rightMargin=15*mm
    )
    
    story = []
    
    # Info
    p_nums = ", ".join([str(p['pattern_num']) for p in selected_patterns])
    unit_name = selected_patterns[0]['unit'] if selected_patterns else "Level A"
    
    # Styles
    title_style = ParagraphStyle('Title', fontSize=12, fontName='Helvetica-Bold', alignment=TA_CENTER, spaceBefore=0, spaceAfter=5)
    section_style = ParagraphStyle('Section', fontSize=10, fontName='Helvetica-Bold', spaceBefore=0, spaceAfter=0)
    item_style = ParagraphStyle('Item', fontSize=9, fontName='Helvetica', leftIndent=0, spaceBefore=2, spaceAfter=2)
    item_kr_style = ParagraphStyle('ItemKr', fontSize=9, fontName=KOREAN_FONT, leftIndent=0, spaceBefore=2, spaceAfter=2)
    line_style = ParagraphStyle('Line', fontSize=9, fontName='Helvetica', spaceAfter=0)
    
    # 1. Header
    story.append(Paragraph("<b>Weekly Test</b>", title_style))
    story.append(Paragraph(f"<b>Pattern {unit_name} - Patterns: {p_nums}</b>", title_style))
    
    # Name/Date
    name_date_data = [[
        Paragraph("NAME: _______________________________", ParagraphStyle('Name', fontSize=12, fontName='Helvetica')),
        Paragraph("DATE: _____ / _____", ParagraphStyle('Date', fontSize=12, fontName='Helvetica', alignment=TA_RIGHT))
    ]]
    name_date_table = Table(name_date_data, colWidths=[120*mm, 50*mm])
    name_date_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
    ]))
    story.append(name_date_table)
    story.append(Spacer(1, 5*mm))
    
    # 2. Speaking I
    story.append(Paragraph("<b>◈ Speaking I - Answer the questions</b>", section_style))
    story.append(Spacer(1, 2*mm))
    
    for idx, question in enumerate(pattern_data['speaking1'][:5], 1):
        story.append(Paragraph(f"{idx}. {question}", item_style))
    
    story.append(Spacer(1, 6*mm))
    
    # 3. Speaking II
    story.append(Paragraph("<b>◈ Speaking II - Say in English</b>", section_style))
    story.append(Spacer(1, 2*mm))
    
    for idx, korean in enumerate(pattern_data['speaking2'][:5], 1):
        story.append(Paragraph(f"{idx}. {korean}", item_kr_style))
    
    story.append(Spacer(1, 6*mm))
    
    # 4. Speaking III
    story.append(Paragraph("<b>◈ Speaking III - With your teacher</b>", section_style))
    story.append(Spacer(1, 2*mm))
    
    for idx in range(1, 6):
        story.append(Paragraph(f"{idx}. Pattern {idx}", item_style))
    
    story.append(Spacer(1, 6*mm))
    
    # 5. Unscramble (핵심 수정 부분)
    story.append(Paragraph("<b>◈ Unscramble</b>", section_style))
    story.append(Spacer(1, 3*mm))
    
    for idx, (korean, words) in enumerate(pattern_data['unscramble'][:5], 1):
        # 문제 텍스트
        story.append(Paragraph(f"{idx}. {korean} ({words})", item_kr_style))
        
        # [수정 1] 글씨 쓸 공간 확보 (문제와 밑줄 사이의 간격 추가)
        story.append(Spacer(1, 8*mm)) 
        
        # 밑줄
        story.append(Paragraph("_" * 85, line_style))
        
        # [수정 2] 다음 문제와의 간격 (너무 넓으면 페이지 넘어가므로 적절히 조절)
        story.append(Spacer(1, 5*mm))
    
    # 6. Footer (GRADE / REMARK)
    # 남은 공간을 계산하기 어려우므로 Spacer를 사용하여 아래쪽으로 밀어줍니다.
    # A4 페이지 내에 들어오도록 적당한 간격 추가
    story.append(Spacer(1, 5*mm))

    footer_data = [[
        Paragraph("<b>GRADE:</b>", ParagraphStyle('Footer', fontSize=12, fontName='Helvetica-Bold')),
        "",
        Paragraph("<b>REMARK:</b>", ParagraphStyle('Footer', fontSize=12, fontName='Helvetica-Bold'))
    ]]
    
    footer_table = Table(footer_data, colWidths=[40*mm, 40*mm, 90*mm])
    footer_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (2, 0), (2, 0), 'LEFT'),
    ]))
    story.append(footer_table)
    
    doc.build(story)

# --- 라우트 설정 ---
@app.route('/')
def index():
    if not os.path.exists(DB_PATH):
        return "<h3>Error: database.xlsx not found.</h3>"
    try:
        patterns = load_patterns_from_excel(DB_PATH)
        pattern_list = []
        for p_num in sorted(patterns.keys()):
            pattern_list.append({
                'number': p_num,
                'name': patterns[p_num]['pattern_name'],
                'unit': patterns[p_num]['unit']
            })
        return render_template('index.html', patterns=pattern_list)
    except Exception as e:
        return f"<h3>Error loading DB: {str(e)}</h3>"

@app.route('/generate', methods=['POST'])
def generate():
    try:
        selected_nums = request.json.get('patterns', [])
        if not selected_nums: return jsonify({'error': 'No patterns selected'}), 400
            
        all_patterns = load_patterns_from_excel(DB_PATH)
        selected_data = []
        for num in selected_nums:
            if int(num) in all_patterns:
                selected_data.append(all_patterns[int(num)])
                
        final_questions = distribute_questions(selected_data)
        
        filename = f"Worksheet_{datetime.now().strftime('%m%d_%H%M%S')}.pdf"
        output_path = os.path.join(OUTPUT_FOLDER, filename)
        
        create_worksheet(final_questions, selected_data, output_path)
        return send_file(output_path, as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=3000, debug=True)