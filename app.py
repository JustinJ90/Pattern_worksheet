#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pattern Worksheet Generator - Multi-Book Version (Final Layout)
- Support for multiple Excel files in 'databases' folder
- Select Book -> Select Patterns
- Layout: A4 Optimized, Footer Fixed, Grade/Remark included
"""

from flask import Flask, render_template, request, send_file, jsonify
import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
import os
import glob
from datetime import datetime

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
DB_FOLDER = os.path.join(BASE_DIR, 'databases')  # DB 폴더 경로

os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(DB_FOLDER, exist_ok=True)

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

# --- 엑셀 읽기 (파일명 받아서 처리) ---
def load_patterns_from_excel(filename):
    file_path = os.path.join(DB_FOLDER, filename)
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"DB 파일을 찾을 수 없습니다: {filename}")

    wb = openpyxl.load_workbook(file_path)
    
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

# --- PDF 생성 (레이아웃 유지) ---
def create_worksheet(pattern_data, selected_patterns, output_path, book_title):
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        topMargin=10*mm,
        bottomMargin=10*mm,
        leftMargin=15*mm,
        rightMargin=15*mm
    )
    
    story = []
    
    p_nums = ", ".join([str(p['pattern_num']) for p in selected_patterns])
    # 책 제목(파일명)에서 확장자 제거 (.xlsx)
    clean_book_title = book_title.replace('.xlsx', '')
    
    # Styles
    title_style = ParagraphStyle('Title', fontSize=12, fontName='Helvetica-Bold', alignment=TA_CENTER, spaceAfter=5)
    section_style = ParagraphStyle('Section', fontSize=11, fontName='Helvetica-Bold', spaceBefore=0, spaceAfter=0)
    item_style = ParagraphStyle('Item', fontSize=10, fontName='Helvetica', leftIndent=0, spaceBefore=2, spaceAfter=2)
    item_kr_style = ParagraphStyle('ItemKr', fontSize=10, fontName=KOREAN_FONT, leftIndent=0, spaceBefore=2, spaceAfter=2)
    line_style = ParagraphStyle('Line', fontSize=10, fontName='Helvetica', spaceAfter=0)
    
    # 1. Header
    story.append(Paragraph("<b>Weekly Test</b>", title_style))
    # 제목에 책 이름(Level A 등)을 포함시킴
    story.append(Paragraph(f"<b>{clean_book_title} - Patterns: {p_nums}</b>", title_style))
    
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
    story.append(Spacer(1, 4*mm))
    
    # 2. Speaking I
    story.append(Paragraph("<b>◈ Speaking I - Answer the questions</b>", section_style))
    story.append(Spacer(1, 2*mm))
    for idx, question in enumerate(pattern_data['speaking1'][:5], 1):
        story.append(Paragraph(f"{idx}. {question}", item_style))
    story.append(Spacer(1, 4*mm))
    
    # 3. Speaking II
    story.append(Paragraph("<b>◈ Speaking II - Say in English</b>", section_style))
    story.append(Spacer(1, 2*mm))
    for idx, korean in enumerate(pattern_data['speaking2'][:5], 1):
        story.append(Paragraph(f"{idx}. {korean}", item_kr_style))
    story.append(Spacer(1, 4*mm))
    
    # 4. Speaking III
    story.append(Paragraph("<b>◈ Speaking III - With your teacher</b>", section_style))
    story.append(Spacer(1, 2*mm))
    for idx in range(1, 6):
        story.append(Paragraph(f"{idx}. Pattern {idx}", item_style))
    story.append(Spacer(1, 4*mm))
    
    # 5. Unscramble
    story.append(Paragraph("<b>◈ Unscramble</b>", section_style))
    story.append(Spacer(1, 2*mm))
    for idx, (korean, words) in enumerate(pattern_data['unscramble'][:5], 1):
        story.append(Paragraph(f"{idx}. {korean} ({words})", item_kr_style))
        story.append(Spacer(1, 7*mm)) 
        story.append(Paragraph("_" * 85, line_style))
        story.append(Spacer(1, 3*mm))
    
    # 6. Footer
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
    # databases 폴더 안의 .xlsx 파일 목록을 가져옵니다.
    files = glob.glob(os.path.join(DB_FOLDER, "*.xlsx"))
    books = sorted([os.path.basename(f) for f in files])
    return render_template('index.html', books=books)

@app.route('/get_patterns/<filename>')
def get_patterns(filename):
    # 선택된 책(파일명)의 패턴 목록을 반환합니다.
    try:
        patterns = load_patterns_from_excel(filename)
        pattern_list = []
        for p_num in sorted(patterns.keys()):
            pattern_list.append({
                'number': p_num,
                'name': patterns[p_num]['pattern_name'],
                'unit': patterns[p_num]['unit']
            })
        return jsonify({'success': True, 'patterns': pattern_list})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.json
        book_filename = data.get('book')
        selected_nums = data.get('patterns', [])
        
        if not book_filename or not selected_nums:
            return jsonify({'error': 'Book or Patterns missing'}), 400
            
        # 해당 책에서 데이터 로드
        all_patterns = load_patterns_from_excel(book_filename)
        selected_data = []
        for num in selected_nums:
            if int(num) in all_patterns:
                selected_data.append(all_patterns[int(num)])
                
        final_questions = distribute_questions(selected_data)
        
        filename = f"Worksheet_{datetime.now().strftime('%m%d_%H%M%S')}.pdf"
        output_path = os.path.join(OUTPUT_FOLDER, filename)
        
        create_worksheet(final_questions, selected_data, output_path, book_filename)
        return send_file(output_path, as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=3000, debug=True)