# Pattern Worksheet Generator - Original Layout Version

**10월 31일 원본 레이아웃** + **여러 패턴 동시 선택 기능**

## 📋 주요 기능

✅ **원본 레이아웃 완벽 재현** - 10월 31일 버전과 100% 동일  
✅ **여러 패턴 동시 선택** - 최대 5개 패턴 선택 가능  
✅ **드래그앤드롭 업로드** - Excel 데이터베이스 쉽게 업로드  
✅ **자동 문항 분배** - 선택한 패턴에서 5문항씩 균등 분배  
✅ **한글 지원** - NanumGothic 폰트로 한글 완벽 지원

## 📄 PDF 레이아웃

```
Weekly Test
Pattern Level A - Patterns: 1, 2, 3

NAME: _______________________________     DATE: _____ / _____

◈ Speaking I - Answer the questions
PATTERN
1. [질문 1]
2. [질문 2]
...

◈ Speaking II - Say in English
1. [한글 문장 1]
2. [한글 문장 2]
...

◈ Speaking III - With your teacher
1. Pattern 1
2. Pattern 2
...

◈ Unscramble
1. [한글] (scrambled words)
   ________________________________________________________________________________
2. ...

GRADE:              REMARK:
```

## 🚀 실행 방법

### 1. 필요한 패키지 설치
```bash
pip install flask openpyxl reportlab werkzeug
```

### 2. 프로그램 실행
```bash
python app.py
```

### 3. 웹 브라우저에서 접속
```
http://127.0.0.1:3000
```

## 📁 폴더 구조

```
final_multi_pattern_CORRECT/
├── app.py                                      # Flask 웹 애플리케이션
├── templates/
│   └── index.html                              # 웹 인터페이스
├── fonts/
│   └── NanumGothic.ttf                        # 한글 폰트
├── uploads/                                    # 업로드된 데이터베이스 저장
├── outputs/                                    # 생성된 PDF 저장
├── pattern_database_COMPLETE_10items_each.xlsx # 샘플 데이터베이스
├── requirements.txt
└── README.md                                   # 이 파일
```

## 📊 데이터베이스 형식

Excel 파일에는 다음 두 개의 시트가 필요합니다:

### 1. Pattern Overview
| Pattern Number | Pattern Name | Total Items |
|---------------|-------------|-------------|
| 1 | My name is .... | 30 |
| 2 | I am .... | 30 |

### 2. Pattern Details
| Pattern # | Pattern Name | Section | Question # | Korean/Question | English/Answer | Scrambled |
|-----------|-------------|---------|------------|-----------------|----------------|-----------|
| 1 | My name is .... | Speaking I | 1 | What's your name? | | |
| 1 | My name is .... | Speaking II | 1 | 내 이름은 Jade야. | My name is Jade. | |
| 1 | My name is .... | Unscramble | 1 | 내 이름은 Jade야. | My name is Jade. | My / is / name / Jade |

## 🎯 사용 방법

1. **데이터베이스 업로드**
   - 웹 페이지에서 Excel 파일을 드래그앤드롭하거나 클릭하여 업로드
   - 업로드가 완료되면 패턴 목록이 자동으로 표시됩니다

2. **패턴 선택**
   - 체크박스를 클릭하여 원하는 패턴 선택 (최대 5개)
   - 선택한 패턴은 파란색으로 강조 표시됩니다

3. **워크시트 생성**
   - "워크시트 생성" 버튼 클릭
   - PDF 파일이 자동으로 다운로드됩니다

## ⚙️ 기술 스택

- **Backend**: Flask (Python)
- **PDF Generation**: ReportLab
- **Excel Processing**: OpenPyXL
- **Frontend**: HTML + JavaScript + CSS
- **Font**: NanumGothic (한글 지원)

## 🔧 포트 변경

기본 포트는 3000입니다. 변경하려면 `app.py` 마지막 줄을 수정하세요:
```python
app.run(host='0.0.0.0', port=3000, debug=True)  # 포트 번호 변경
```

## 📝 주의사항

- 데이터베이스 파일은 반드시 `.xlsx` 형식이어야 합니다
- "Pattern Overview"와 "Pattern Details" 시트가 필수입니다
- 각 패턴마다 최소 5개 이상의 문항이 있어야 합니다
- 한글이 포함된 경우 NanumGothic.ttf 폰트 파일이 필요합니다

## 🆚 차이점

**원본 (10월 31일) vs 이전 버전:**
- ✅ Title: "Weekly Test" + "Pattern Level A - Patterns: X, Y"
- ✅ NAME과 DATE가 같은 줄
- ✅ "◈ Speaking I - Answer the questions" 형식
- ✅ PATTERN 라벨 표시
- ✅ Speaking III에 "Pattern 1, Pattern 2..." 표시
- ✅ Unscramble 아래 밑줄 표시
- ✅ GRADE: 와 REMARK: 같은 줄

---
**Version**: Original Layout Multi-Pattern 1.0  
**Based on**: 2024-10-31 worksheet_FINAL version  
**Last Updated**: 2025-11-18
