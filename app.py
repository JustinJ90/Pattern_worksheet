#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pattern Worksheet Generator - Multi-Pattern Version with Original Layout
10월 31일 원본 레이아웃 + 여러 패턴 선택 기능
"""

from flask import Flask, render_template, request, send_file, jsonify
import openpyxl
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER
import os
import platform
from datetime import datetime
from werkzeug.utils import secure_filename

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app.config['UPLOAD_FOLDER'] = os.path.join(BASE_DIR, 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(BASE_DIR, 'outputs')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# 현재 데이터베이스 경로
CURRENT_DB_PATH = None

def setup_korean_font():
    """Setup Korean font - works on Windows, Mac, Linux"""
    try:
        # 1. 프로젝트 내부 fonts 폴더
        local_font = os.path.join(BASE_DIR, 'fonts', 'NanumGothic.ttf')
        if os.path.exists(local_font):
            pdfmetrics.registerFont(TTFont('KoreanFont', local_font))
            print(f"✅ 폰트 로드 성공: {local_font}")
            return 'KoreanFont'
        
        # 2. Windows fonts
        if platform.system() == 'Windows':
            for font_path in [r'C:\Windows\Fonts\malgun.ttf', 
                            r'C:\Windows\Fonts\gulim.ttc',
                            r'C:\Windows\Fonts\batang.ttc']:
                if os.path.exists(font_path):
                    pdfmetrics.registerFont(TTFont('KoreanFont', font_path))
                    print(f"✅ 폰트 로드 성공: {font_path}")
                    return 'KoreanFont'
        
        # 3. Mac fonts
        elif platform.system() == 'Darwin':
            for font_path in ['/System/Library/Fonts/AppleSDGothicNeo.ttc',
                            '/Library/Fonts/AppleGothic.ttf']:
                if os.path.exists(font_path):
                    pdfmetrics.registerFont(TTFont('KoreanFont', font_path))
                    print(f"✅ 폰트 로드 성공: {font_path}")
                    return 'KoreanFont'
        
        # 4. Linux fonts
        else:
            for font_path in ['/usr/share/fonts/truetype/nanum/NanumGothic.ttf',
                            '/usr/share/fonts/truetype/nanum/NanumBarunGothic.ttf']:
                if os.path.exists(font_path):
                    pdfmetrics.registerFont(TTFont('KoreanFont', font_path))
                    print(f"✅ 폰트 로드 성공: {font_path}")
                    return 'KoreanFont'
    except Exception as e:
        print(f"⚠️ 폰트 로드 실패: {e}")
    
    print("⚠️ 한글 폰트를 찾지 못했습니다. Helvetica 사용")
    return 'Helvetica'

KOREAN_FONT = setup_korean_font()


def load_patterns_from_excel(excel_path):
    """Load pattern data from Excel file"""
    wb = openpyxl.load_workbook(excel_path)
    
    # Load pattern overview
    ws_overview = wb["Pattern Overview"]
    pattern_info = {}
    
    for row in ws_overview.iter_rows(min_row=2, values_only=True):
        # 유연한 컬럼 처리
        if len(row) >= 3:
            pattern_num, pattern_name, total_q = row[0], row[1], row[2]
            unit = row[3] if len(row) > 3 else ''
        else:
            continue
        
        if pattern_num is not None:
            pattern_info[int(pattern_num)] = {
                'number': int(pattern_num),
                'name': str(pattern_name),
                'unit': str(unit) if unit else 'Level A',
                'total_questions': int(total_q) if total_q else 0
            }
    
    # Load pattern details
    ws_detail = wb["Pattern Details"]
    patterns = {}
    
    for row in ws_detail.iter_rows(min_row=2, values_only=True):
        pattern_num, pattern_name, section, q_num, col_e, col_f, col_g = row
        pattern_num = int(pattern_num)
        
        if pattern_num not in patterns:
            patterns[pattern_num] = {
                'pattern_num': pattern_num,
                'pattern_name': pattern_name,
                'unit': pattern_info.get(pattern_num, {}).get('unit', 'Level A'),
                'speaking1': [],
                'speaking2': [],
                'unscramble': []
            }
        
        # Speaking I: Questions only
        if section == 'Speaking I':
            patterns[pattern_num]['speaking1'].append(col_e)
        # Speaking II: Korean
        elif section == 'Speaking II':
            patterns[pattern_num]['speaking2'].append(col_e)
        # Unscramble: Korean + scrambled words
        elif section == 'Unscramble':
            words_str = col_g.strip('()') if col_g else ""
            patterns[pattern_num]['unscramble'].append((col_e, words_str))
    
    return patterns


def distribute_questions(selected_patterns, target_count=5):
    """Distribute questions evenly across patterns"""
    result = {'speaking1': [], 'speaking2': [], 'unscramble': []}
    pattern_count = len(selected_patterns)
    
    items_per_pattern = target_count // pattern_count
    remainder = target_count % pattern_count
    
    for section in ['speaking1', 'speaking2', 'unscramble']:
        for i, pattern in enumerate(selected_patterns):
            take_count = items_per_pattern + (1 if i < remainder else 0)
            result[section].extend(pattern[section][:take_count])
        result[section] = result[section][:target_count]
    
    return result


def create_worksheet(pattern_data, selected_patterns, output_path):
    """Create worksheet PDF matching original layout EXACTLY"""
    doc = SimpleDocTemplate(
        output_path,
        pagesize=letter,
        topMargin=0.4*inch,
        bottomMargin=0.4*inch,
        leftMargin=0.5*inch,
        rightMargin=0.5*inch
    )
    
    story = []
    
    # === HEADER: Title centered, then NAME and DATE on same line ===
    pattern_nums = ", ".join([str(p['pattern_num']) for p in selected_patterns])
    unit_name = selected_patterns[0]['unit'] if selected_patterns else "Level A"
    
    # Title centered at top
    title_style = ParagraphStyle('Title', fontSize=12, fontName='Helvetica-Bold', 
                                alignment=TA_CENTER, spaceBefore=0, spaceAfter=5)
    story.append(Paragraph("<b>Weekly Test</b>", title_style))
    story.append(Paragraph(f"<b>Pattern {unit_name} - Patterns: {pattern_nums}</b>", title_style))
    
    # NAME and DATE on same line below title
    name_date_data = [[
        Paragraph("NAME: _______________________________", 
                 ParagraphStyle('Name', fontSize=12, fontName='Helvetica')),
        Paragraph("DATE: _____ / _____", 
                 ParagraphStyle('Date', fontSize=12, fontName='Helvetica', alignment=TA_RIGHT))
    ]]
    
    name_date_table = Table(name_date_data, colWidths=[5*inch, 2*inch])
    name_date_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
    ]))
    story.append(name_date_table)
    story.append(Spacer(1, 0.15*inch))
    
    # === SPEAKING I ===
    story.append(Paragraph("<b>◈ Speaking I - Answer the questions</b>", 
                          ParagraphStyle('Section', fontSize=10, fontName='Helvetica-Bold')))
    story.append(Spacer(1, 0.05*inch))
    
    # PATTERN 라벨
    story.append(Paragraph("<b>PATTERN</b>", 
                          ParagraphStyle('Pattern', fontSize=9, fontName='Helvetica-Bold')))
    story.append(Spacer(1, 0.05*inch))
    
    # Speaking I 질문들
    for idx, question in enumerate(pattern_data['speaking1'][:5], 1):
        story.append(Paragraph(f"{idx}. {question}", 
                              ParagraphStyle('Item', fontSize=9, fontName='Helvetica', 
                                           leftIndent=0, spaceBefore=3, spaceAfter=3)))
    
    story.append(Spacer(1, 0.15*inch))
    
    # === SPEAKING II ===
    story.append(Paragraph("<b>◈ Speaking II - Say in English</b>", 
                          ParagraphStyle('Section', fontSize=10, fontName='Helvetica-Bold')))
    story.append(Spacer(1, 0.05*inch))
    
    for idx, korean in enumerate(pattern_data['speaking2'][:5], 1):
        story.append(Paragraph(f"{idx}. {korean}", 
                              ParagraphStyle('Item', fontSize=9, fontName=KOREAN_FONT, 
                                           leftIndent=0, spaceBefore=3, spaceAfter=3)))
    story.append(Spacer(1, 0.15*inch))
    
    # === SPEAKING III ===
    story.append(Paragraph("<b>◈ Speaking III - With your teacher</b>", 
                          ParagraphStyle('Section', fontSize=10, fontName='Helvetica-Bold')))
    story.append(Spacer(1, 0.05*inch))
    
    # Show "Pattern 1", "Pattern 2", etc.
    for idx in range(1, 6):
        story.append(Paragraph(f"{idx}. Pattern {idx}", 
                              ParagraphStyle('Item', fontSize=9, fontName='Helvetica', 
                                           leftIndent=0, spaceBefore=3, spaceAfter=3)))
    story.append(Spacer(1, 0.15*inch))
    
    # === UNSCRAMBLE ===
    story.append(Paragraph("<b>◈ Unscramble</b>", 
                          ParagraphStyle('Section', fontSize=10, fontName='Helvetica-Bold')))
    story.append(Spacer(1, 0.08*inch))
    
    for idx, (korean, words) in enumerate(pattern_data['unscramble'][:5], 1):
        story.append(Paragraph(f"{idx}. {korean} ({words})", 
                              ParagraphStyle('Item', fontSize=9, fontName=KOREAN_FONT, 
                                           leftIndent=0, spaceBefore=4, spaceAfter=3)))
        story.append(Paragraph("_" * 80, 
                              ParagraphStyle('Line', fontSize=9, fontName='Helvetica', 
                                           spaceAfter=10)))
    
    story.append(Spacer(1, 0.35*inch))
    
    # === FOOTER: GRADE and REMARK on same line ===
    footer_data = [[
        Paragraph("<b>GRADE:</b>", ParagraphStyle('Footer', fontSize=12, fontName='Helvetica-Bold')),
        "",
        Paragraph("<b>REMARK:</b>", ParagraphStyle('Footer', fontSize=12, fontName='Helvetica-Bold'))
    ]]
    
    footer_table = Table(footer_data, colWidths=[1*inch, 2*inch, 4*inch])
    footer_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (2, 0), (2, 0), 'LEFT'),
    ]))
    story.append(footer_table)
    
    # Build PDF
    doc.build(story)
    return output_path


@app.route('/')
def index():
    """메인 페이지"""
    return render_template('index.html')


@app.route('/upload_database', methods=['POST'])
def upload_database():
    """데이터베이스 업로드 및 패턴 정보 반환"""
    global CURRENT_DB_PATH
    
    try:
        if 'database' not in request.files:
            return jsonify({'error': '파일이 없습니다.'}), 400
        
        file = request.files['database']
        if file.filename == '':
            return jsonify({'error': '파일을 선택해주세요.'}), 400
        
        if not file.filename.endswith('.xlsx'):
            return jsonify({'error': 'Excel 파일(.xlsx)만 업로드 가능합니다.'}), 400
        
        # 파일 저장
        filename = 'uploaded_database.xlsx'
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        CURRENT_DB_PATH = filepath
        
        # 패턴 정보 로드
        patterns = load_patterns_from_excel(filepath)
        
        pattern_list = []
        for pattern_num in sorted(patterns.keys()):
            pattern = patterns[pattern_num]
            pattern_list.append({
                'number': pattern['pattern_num'],
                'name': pattern['pattern_name']
            })
        
        return jsonify({
            'success': True,
            'patterns': pattern_list,
            'message': f'{len(pattern_list)}개의 패턴이 로드되었습니다.'
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'데이터베이스 로드 실패: {str(e)}'}), 500


@app.route('/generate', methods=['POST'])
def generate_worksheet():
    """활동지 생성 (여러 패턴 지원)"""
    global CURRENT_DB_PATH
    
    try:
        selected_pattern_nums = request.json.get('patterns', [])
        
        if not selected_pattern_nums:
            return jsonify({'error': '패턴을 선택해주세요.'}), 400
        
        if len(selected_pattern_nums) > 5:
            return jsonify({'error': '최대 5개 패턴까지 선택 가능합니다.'}), 400
        
        # 데이터베이스 확인
        if not CURRENT_DB_PATH or not os.path.exists(CURRENT_DB_PATH):
            return jsonify({'error': '데이터베이스를 먼저 업로드해주세요.'}), 400
        
        # 패턴 로드
        all_patterns = load_patterns_from_excel(CURRENT_DB_PATH)
        
        # 선택된 패턴 추출
        selected_patterns = []
        for num in selected_pattern_nums:
            pattern_num = int(num)
            if pattern_num not in all_patterns:
                return jsonify({'error': f'패턴 {pattern_num}을 찾을 수 없습니다.'}), 404
            selected_patterns.append(all_patterns[pattern_num])
        
        # 5문항으로 분배
        distributed_data = distribute_questions(selected_patterns, target_count=5)
        
        # PDF 생성
        pattern_nums_str = '_'.join([str(num) for num in selected_pattern_nums])
        output_filename = f"worksheet_patterns_{pattern_nums_str}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        
        create_worksheet(distributed_data, selected_patterns, output_path)
        
        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/pdf'
        )
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    print("=" * 60)
    print("🎓 Pattern Worksheet Generator - Original Layout")
    print("=" * 60)
    print(f"✅ 한글 폰트: {KOREAN_FONT}")
    print(f"✅ 작업 폴더: {BASE_DIR}")
    print("=" * 60)
    print("🌐 웹 브라우저에서 다음 주소로 접속하세요:")
    print("   http://127.0.0.1:3000")
    print("=" * 60)
    print("\n종료하려면 Ctrl + C 를 누르세요.\n")
    
    app.run(host='0.0.0.0', port=3000, debug=True)
