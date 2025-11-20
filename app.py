#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from flask import Flask, render_template, request, send_file, jsonify
import openpyxl
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
import os
import platform
from datetime import datetime

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
# DB 파일 경로를 고정합니다 (같은 폴더의 database.xlsx)
DB_PATH = os.path.join(BASE_DIR, 'database.xlsx')

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# --- 폰트 설정 (기존과 동일) ---
def setup_korean_font():
    try:
        local_font = os.path.join(BASE_DIR, 'fonts', 'NanumGothic.ttf')
        if os.path.exists(local_font):
            pdfmetrics.registerFont(TTFont('KoreanFont', local_font))
            return 'KoreanFont'
        # (서버용) 폰트가 없을 경우 대비
        return 'Helvetica' 
    except:
        return 'Helvetica'

KOREAN_FONT = setup_korean_font()

# --- 엑셀 읽기 함수 ---
def load_patterns_from_excel(excel_path):
    wb = openpyxl.load_workbook(excel_path)
    
    # 1. Overview 시트 읽기
    ws_overview = wb["Pattern Overview"]
    pattern_info = {}
    for row in ws_overview.iter_rows(min_row=2, values_only=True):
        if row[0] is not None:
            pattern_info[int(row[0])] = {
                'number': int(row[0]),
                'name': str(row[1]),
                'unit': str(row[3]) if len(row) > 3 and row[3] else 'Level A'
            }
            
    # 2. Details 시트 읽기
    ws_detail = wb["Pattern Details"]
    patterns = {}
    for row in ws_detail.iter_rows(min_row=2, values_only=True):
        try:
            p_num = int(row[0])
            section = row[2]
            content = row[4] # Korean or Question
            
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

# --- PDF 생성 함수 (기존과 동일, 생략된 부분 없음) ---
def create_worksheet(pattern_data, selected_patterns, output_path):
    doc = SimpleDocTemplate(output_path, pagesize=letter,
                            topMargin=0.4*inch, bottomMargin=0.4*inch,
                            leftMargin=0.5*inch, rightMargin=0.5*inch)
    story = []
    
    # 헤더 정보
    p_nums = ", ".join([str(p['pattern_num']) for p in selected_patterns])
    unit_name = selected_patterns[0]['unit'] if selected_patterns else "Level A"
    
    # 스타일
    styles = {
        'Title': ParagraphStyle('Title', fontSize=12, fontName='Helvetica-Bold', alignment=TA_CENTER, spaceAfter=5),
        'Section': ParagraphStyle('Section', fontSize=10, fontName='Helvetica-Bold', spaceBefore=10),
        'Item': ParagraphStyle('Item', fontSize=9, fontName='Helvetica', spaceBefore=3),
        'ItemKr': ParagraphStyle('ItemKr', fontSize=9, fontName=KOREAN_FONT, spaceBefore=3)
    }
    
    # 제목
    story.append(Paragraph("<b>Weekly Test</b>", styles['Title']))
    story.append(Paragraph(f"<b>Pattern {unit_name} - Patterns: {p_nums}</b>", styles['Title']))
    
    # 이름/날짜 표
    data = [[Paragraph("NAME: ___________________", styles['Item']), 
             Paragraph("DATE: ____ / ____", ParagraphStyle('D', parent=styles['Item'], alignment=TA_RIGHT))]]
    story.append(Table(data, colWidths=[5*inch, 2*inch]))
    story.append(Spacer(1, 0.15*inch))
    
    # Sections...
    sections = [
        ('Speaking I - Answer the questions', 'speaking1', styles['Item']),
        ('Speaking II - Say in English', 'speaking2', styles['ItemKr']),
    ]
    
    for title, key, style in sections:
        story.append(Paragraph(f"<b>◈ {title}</b>", styles['Section']))
        if key == 'speaking1': story.append(Paragraph("<b>PATTERN</b>", ParagraphStyle('sub', fontSize=8, fontName='Helvetica-Bold')))
        
        for i, q in enumerate(pattern_data[key][:5], 1):
            story.append(Paragraph(f"{i}. {q}", style))
        story.append(Spacer(1, 0.1*inch))
        
    # Speaking III (패턴만 나열)
    story.append(Paragraph("<b>◈ Speaking III - With your teacher</b>", styles['Section']))
    for i in range(1, 6):
        story.append(Paragraph(f"{i}. Pattern {i}", styles['Item']))
    story.append(Spacer(1, 0.1*inch))

    # Unscramble
    story.append(Paragraph("<b>◈ Unscramble</b>", styles['Section']))
    for i, (q, hint) in enumerate(pattern_data['unscramble'][:5], 1):
        story.append(Paragraph(f"{i}. {q} ({hint})", styles['ItemKr']))
        story.append(Paragraph("_"*85, ParagraphStyle('Line', fontSize=8, spaceAfter=5)))
        
    doc.build(story)

# --- 라우트 설정 ---

@app.route('/')
def index():
    # 접속하자마자 엑셀 파일을 읽어서 리스트를 만듭니다.
    if not os.path.exists(DB_PATH):
        return "<h3>Error: database.xlsx 파일을 찾을 수 없습니다. 서버에 파일을 업로드했는지 확인하세요.</h3>"
    
    try:
        patterns = load_patterns_from_excel(DB_PATH)
        # 패턴 리스트 정리 (번호, 이름, 유닛)
        pattern_list = []
        for p_num in sorted(patterns.keys()):
            pattern_list.append({
                'number': p_num,
                'name': patterns[p_num]['pattern_name'],
                'unit': patterns[p_num]['unit']
            })
        
        # index.html로 패턴 목록을 바로 보냅니다.
        return render_template('index.html', patterns=pattern_list)
        
    except Exception as e:
        return f"<h3>DB 로드 중 오류 발생: {str(e)}</h3>"

@app.route('/generate', methods=['POST'])
def generate():
    try:
        # 1. 선택된 패턴 번호 받기
        selected_nums = request.json.get('patterns', [])
        if not selected_nums:
            return jsonify({'error': '패턴을 선택해주세요.'}), 400
            
        # 2. DB 다시 읽기 (혹시 모를 오류 방지 및 최신 상태 유지)
        all_patterns = load_patterns_from_excel(DB_PATH)
        
        selected_data = []
        for num in selected_nums:
            if int(num) in all_patterns:
                selected_data.append(all_patterns[int(num)])
                
        # 3. 문제 섞기 및 PDF 생성
        final_questions = distribute_questions(selected_data)
        
        filename = f"Worksheet_{datetime.now().strftime('%m%d_%H%M%S')}.pdf"
        output_path = os.path.join(OUTPUT_FOLDER, filename)
        
        create_worksheet(final_questions, selected_data, output_path)
        
        return send_file(output_path, as_attachment=True)
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=3000, debug=True)