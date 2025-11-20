#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pattern Worksheet Generator - Final Layout Version
- Embedded Database (No upload required)
- Original Layout restored
- Footer (GRADE/REMARK) restored
- 'PATTERN' label removed in Speaking I
- Unscramble spacing increased to fill A4
"""

from flask import Flask, render_template, request, send_file, jsonify
import openpyxl
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.units import inch
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

# --- PDF 생성 (레이아웃 수정됨) ---
def create_worksheet(pattern_data, selected_patterns, output_path):
    # 여백 설정 (원본과 동일하게 조정)
    doc = SimpleDocTemplate(
        output_path,
        pagesize=letter,
        topMargin=0.4*inch,
        bottomMargin=0.4*inch,
        leftMargin=0.5*inch,
        rightMargin=0.5*inch
    )
    
    story = []
    
    # 1. Header Info
    p_nums = ", ".join([str(p['pattern_num']) for p in selected_patterns])
    unit_name = selected_patterns[0]['unit'] if selected_patterns else "Level A"
    
    # Styles
    title_style = ParagraphStyle('Title', fontSize=12, fontName='Helvetica-Bold', alignment=TA_CENTER, spaceBefore=0, spaceAfter=5)
    section_style = ParagraphStyle('Section', fontSize=10, fontName='Helvetica-Bold', spaceBefore=0, spaceAfter=0)
    item_style = ParagraphStyle('Item', fontSize=9, fontName='Helvetica', leftIndent=0, spaceBefore=3, spaceAfter=3)
    item_kr_style = ParagraphStyle('ItemKr', fontSize=9, fontName=KOREAN_FONT, leftIndent=0, spaceBefore=3, spaceAfter=3)
    
    # Title
    story.append(Paragraph("<b>Weekly Test</b>", title_style))
    story.append(Paragraph(f"<b>Pattern {unit_name} - Patterns: {p_nums}</b>", title_style))
    
    # Name/Date Table
    name_date_data = [[
        Paragraph("NAME: _______________________________", ParagraphStyle('Name', fontSize=12, fontName='Helvetica')),
        Paragraph("DATE: _____ / _____", ParagraphStyle('Date', fontSize=12, fontName='Helvetica', alignment=TA_RIGHT))
    ]]
    name_date_table = Table(name_date_data, colWidths=[5*inch, 2*inch])
    name_date_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
    ]))
    story.append(name_date_table)
    story.append(Spacer(1, 0.2*inch))
    
    # 2. Speaking I
    story.append(Paragraph("<b>◈ Speaking I - Answer the questions</b>", section_style))
    story.append(Spacer(1, 0.05*inch))
    
    # [삭제됨] PATTERN 라벨 제거
    # story.append(Paragraph("<b>PATTERN</b>", ParagraphStyle('sub', fontSize=9, fontName='Helvetica-Bold')))
    
    for idx, question in enumerate(pattern_data['speaking1'][:5], 1):
        story.append(Paragraph(f"{idx}. {question}", item_style))
    
    story.append(Spacer(1, 0.2*inch))
    
    # 3. Speaking II
    story.append(Paragraph("<b>◈ Speaking II - Say in English</b>", section_style))
    story.append(Spacer(1, 0.05*inch))
    
    for idx, korean in enumerate(pattern_data['speaking2'][:5], 1):
        story.append(Paragraph(f"{idx}. {korean}", item_kr_style))
    
    story.append(Spacer(1, 0.2*inch))
    
    # 4. Speaking III
    story.append(Paragraph("<b>◈ Speaking III - With your teacher</b>", section_style))
    story.append(Spacer(1, 0.05*inch))
    
    for idx in range(1, 6):
        story.append(Paragraph(f"{idx}. Pattern {idx}", item_style))
    
    story.append(Spacer(1, 0.2*inch))
    
    # 5. Unscramble
    story.append(Paragraph("<b>◈ Unscramble</b>", section_style))
    story.append(Spacer(1, 0.1*inch))
    
    for idx, (korean, words) in enumerate(pattern_data['unscramble'][:5], 1):
        # 문제
        story.append(Paragraph(f"{idx}. {korean} ({words})", item_kr_style))
        # 밑줄 (답 적는 곳)
        story.append(Paragraph("_" * 85, ParagraphStyle('Line', fontSize=9, fontName='Helvetica', spaceAfter=0)))
        
        # [수정됨] 간격을 0.55 inch로 늘려서 페이지를 채움 (기존 0.35)
        story.append(Spacer(1, 0.55*inch)) 
    
    # 6. Footer (GRADE / REMARK) - 복구됨
    # Unscramble 루프가 끝난 후 약간의 간격이 있을 수 있으므로 Spacer 추가
    story.append(Spacer(1, 0.1*inch))

    footer_data = [[
        Paragraph("<b>GRADE:</b>", ParagraphStyle('Footer', fontSize=12, fontName='Helvetica-Bold')),
        "",
        Paragraph("<b>REMARK:</b>", ParagraphStyle('Footer', fontSize=12, fontName='Helvetica-Bold'))
    ]]
    
    footer_table = Table(footer_data, colWidths=[1.5*inch, 1.5*inch, 4*inch])
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