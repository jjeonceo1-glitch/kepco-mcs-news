import streamlit as st
from PIL import Image, ImageOps
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
from datetime import datetime
import json
import time
import base64

from streamlit_geolocation import streamlit_geolocation
from geopy.geocoders import Nominatim
import pytz
import google.generativeai as genai

# ==========================================
# 1. 페이지 설정 (단 한 번만 맨 처음에 실행)
# ==========================================
icon_img = Image.open("safety_check_app/app_icon.png")
st.set_page_config(
    page_title="한전MCS 위험성평가 AI",
    page_icon=icon_img,
    layout="centered"
)

# 스마트폰 홈 화면 추가 시 아이콘 적용을 위한 메타 태그 강제 주입
app_icon_url = "https://raw.githubusercontent.com/jjeonceo1-glitch/kepco-mcs-news/main/safety_check_app/app_icon.png"
st.markdown(f"""
    <link rel="apple-touch-icon" href="{app_icon_url}">
    <link rel="icon" sizes="192x192" href="{app_icon_url}">
""", unsafe_allow_html=True)

# ==========================================
# 2. 커스텀 CSS 디자인 적용 (여백 및 스타일)
# ==========================================
st.markdown("""
    <style>
    .block-container {
        padding-top: 1.5rem;
    }
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    
    div.stButton > button:first-child {
        border-radius: 8px;
        font-weight: bold;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        transition: all 0.3s ease;
    }
    div.stButton > button:first-child:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 10px rgba(0,0,0,0.1);
    }
    hr {
        margin-top: 1.5em;
        margin-bottom: 1.5em;
        border-color: #E0E6ED;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 3. 문서 서식 자동화 함수
# ==========================================
def apply_cell_shading(cell, color="D9D9D9"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        tcPr.append(shd)
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)

def format_table(table, is_header_top=True):
    for row in table.rows:
        for cell in row.cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if is_header_top and len(table.rows) > 0:
        for cell in table.rows[0].cells:
            apply_cell_shading(cell, "D9D9D9")

# ==========================================
# 4. 보안 로그인 함수 (4자리 비밀번호: 0325)
# ==========================================
def check_password():
    def password_entered():
        if st.session_state["password"] == "0325":
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.markdown("### 🔒 한전MCS 안전점검 시스템")
        st.text_input(
            "관리자 비밀번호 4자리를 입력하세요", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        return False
        
    elif not st.session_state["password_correct"]:
        st.markdown("### 🔒 한전MCS 안전점검 시스템")
        st.text_input(
            "관리자 비밀번호 4자리를 입력하세요", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        st.error("비밀번호가 일치하지 않습니다. 다시 시도해 주세요.")
        return False
        
    else:
        return True

# ==========================================
# 5. 메인 앱 실행 (비밀번호 통과 시 작동)
# ==========================================
if check_password():
    # API 키 및 모델 설정
    try:
        API_KEY = st.secrets["GEMINI_API_KEY"]
        genai.configure(api_key=API_KEY)
        # 안정적인 1.5-flash 모델로 설정 (오타 수정)
        model = genai.GenerativeModel("gemini-3.8-flash")
    except Exception as e:
        st.error(f"GEMINI_API_KEY 설정 오류: {e}")

    # --- 좌측 사이드바 구성 ---
    with st.sidebar:
        try:
            st.image("safety_check_app/logo.png", use_container_width=True)
        except Exception:
            st.title("한전MCS 안전재난부")
            
        st.divider()
        st.header("📌 사용 방법")
        st.markdown("""
        **1단계:** 본문의 📍[위치 수집] 버튼을 눌러 현재 현장 주소를 불러옵니다.  
        **2단계:** 점검할 현장의 사진을 1장 업로드합니다.  
        **3단계:** [평가표 작성 시작] 버튼을 누릅니다.  
        **4단계:** 분석이 끝나면 워드(.docx) 파일을 다운로드하여 서명 후 보관합니다.
        """)
        st.divider()
        st.info("💡 **Tip:** 사진에 작업자의 행동이나 주요 설비(전신주, 개구부 등)가 잘 보일수록 AI가 더 정확한 법적 근거를 찾아냅니다.")
        st.caption("© 2027 한전MCS. All rights reserved.")

    # --- 메인 화면 헤더 (모바일 최적화 Flexbox 레이아웃) ---
    try:
        with open("safety_check_app/logo.png", "rb") as image_file:
            logo_base64 = base64.b64encode(image_file.read()).decode()
        
        st.markdown(f"""
            <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 15px;">
                <img src="data:image/png;base64,{logo_base64}" width="60">
                <h1 style="margin: 0; font-size: 26px; letter-spacing: -0.5px;">위험성평가 AI</h1>
            </div>
        """, unsafe_allow_html=True)
    except Exception:
        st.title("위험성평가 AI")

    st.markdown("고용노동부고시 제2024-76호 지침에 따른 **위험성평가**를 지원합니다.")
    st.markdown("<p style='font-size: 14px; color: #555;'>본 AI기반 위험성평가는 참고자료이며 관리자가 필히 재확인 해야합니다.</p>", unsafe_allow_html=True)
    st.divider()

    # 이 아래부터 사진 업로드, 제미나이 분석, 워드 다운로드 등 추가 기능을 이어서 작성하시면 됩니다.
    st.write("✨ 로그인 및 레이아웃 초기화 완료. 핵심 기능을 구현할 차례입니다!")

# --- GPS 위치 수집 ---
st.subheader("📍 1단계: 현장 위치 수집")
    st.write("아래 아이콘을 눌러 현재 위치(GPS)를 수집해 주세요.")
    
    # 위치 수집 모듈 실행 (버튼 생성)
    location = streamlit_geolocation()
    
    # 위치 정보가 정상적으로 들어왔는지 확인
    if location is not None and location.get('latitude') is not None:
        lat = location['latitude']
        lon = location['longitude']
        st.success("✅ 위치 수집 완료!")
        st.info(f"위도: {lat} / 경도: {lon}")
    
    st.divider()

# --- 사진 업로드 ---
st.subheader("2. 현장 사진 업로드")
uploaded_file = st.file_uploader("작업 현장 사진을 선택하거나 촬영해 주세요.", type=["jpg", "jpeg", "png"])

if uploaded_file is not None:
    # 1. 원본 이미지 열기 및 방향(EXIF) 복구 (딱 한 번만 실행)
    image = Image.open(uploaded_file)
    image = ImageOps.exif_transpose(image)

    # 2. 화면을 두 칸(좌/우)으로 나누어 보기 좋게 배치
    col1, col2 = st.columns([1, 1])
    
    with col1:
        # 위에서 방향을 맞춘 image를 그대로 가져와서 크기만 줄이고 화면에 띄웁니다.
        image.thumbnail((800, 800))
        st.image(image, caption="업로드된 현장 사진", use_container_width=True)
        
    with col2:
        st.warning("⚠️ **주의:** 사진이 흐리거나 너무 멀리서 찍힌 경우 위험 요소를 판독하기 어려울 수 있습니다.")
        st.write("사진 확인이 완료되었다면 아래 버튼을 눌러 문서를 생성하세요.")
        
        # 버튼을 돋보이게 만들기 (type="primary")
        if st.button("🚀 위험성평가 작성 시작", type="primary", use_container_width=True):
            with st.spinner("AI가 산업안전보건법령을 검토하고 문서를 조립 중입니다... (약 10~15초 소요)"):
                prompt = """
                당신은 한전MCS의 산업안전보건 및 현장점검 전문가입니다.
                사진을 분석하여 '위험성평가표' 작성을 위한 데이터를 JSON 형식으로 반환해주세요.
                
                [JSON 구조]
                {
                  "criteria": {
                    "high": {"condition": "사진 현장 특성에 맞는 '상' 수준 판단 기준 (예: 활선 노출, 2m 이상 개구부 등)", "action": "상 수준 조치 사항"},
                    "medium": {"condition": "사진 현장 특성에 맞는 '중' 수준 판단 기준", "action": "중 수준 조치 사항"},
                    "low": {"condition": "사진 현장 특성에 맞는 '하' 수준 판단 기준", "action": "하 수준 조치 사항"},
                    "legal_basis": "수준 판단의 기준이 된 구체적 법적 근거 (예: 산업안전보건기준에 관한 규칙 제42조 추락의 방지 등)"
                  },
                  "hazards": [
                    {
                      "no": 1,
                      "factor": "유해·위험요인",
                      "type": "재해형태",
                      "current_measure": "현재 조치 (예: 보이지 않음)",
                      "level": "상, 중, 하 중 택 1",
                      "countermeasure": "감소대책(후보·제12조 순서)"
                    }
                  ],
                  "tbm": [
                    "작업 전 공유(TBM) 사항 1"
                  ]
                }
                사진에 보이는 작업 환경(예: 고소작업, 전기작업 등)을 파악하여, 그 현장에 가장 적합한 '위험성 수준 기준(criteria)'과 '위험 요인(hazards)'을 도출해주세요.
                """
                
                try:
                    max_retries = 3
                    for attempt in range(max_retries):
                        try:
                            response = model.generate_content(
                                [prompt, image],
                                generation_config={"response_mime_type": "application/json"}
                            )
                            break
                        except Exception as e:
                            if "429" in str(e) and attempt < max_retries - 1:
                                time.sleep(16)
                            else:
                                raise e
                    
                    data = json.loads(response.text)
                    hazards = data.get("hazards", [])
                    total_count = len(hazards)
                    count_high = sum(1 for h in hazards if h.get("level") == "상")
                    count_mid = sum(1 for h in hazards if h.get("level") == "중")
                    count_low = sum(1 for h in hazards if h.get("level") == "하")
                    
                    korea_tz = pytz.timezone('Asia/Seoul')
                    current_time_str = datetime.now(korea_tz).strftime('%Y. %m. %d. %H:%M')
                    
                    # --- Word 문서 조립 ---
                    doc = Document()
                
                    # 1. 페이지 여백 설정 (공공기관 보고서 기본 여백: 상하 15mm, 좌우 20mm)
                    for section in doc.sections:
                        section.top_margin = Mm(15)
                        section.bottom_margin = Mm(15)
                        section.left_margin = Mm(20)
                        section.right_margin = Mm(20)
                
                    # 2. 기본 폰트 완벽 적용 (한글 깨짐 방지 및 맑은 고딕 강제 세팅)
                    style = doc.styles['Normal']
                    style.font.name = '맑은 고딕'
                    style._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
                    style.font.size = Pt(11) # 가독성을 위해 기존 10pt에서 11pt로 확대
                
                    # 3. 메인 타이틀 디자인 고급화
                    title = doc.add_heading('현장 위험성 평가표', 0)
                    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    title_run = title.runs[0]
                    title_run.font.name = '맑은 고딕'
                    title_run._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
                    title_run.font.bold = True
                    title_run.font.color.rgb = None # 파란색 띠는 기본 제목 색을 검정으로 강제 변경
                
                    # 부제목 서식
                    subtitle = doc.add_paragraph('사업장 위험성평가에 관한 지침(고용노동부고시 제2024-76호)의 절차를 따름')
                    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    # 1. 헤더 테이블
                    table1 = doc.add_table(rows=4, cols=4)
                    table1.style = 'Table Grid'
                    table1.cell(0,0).text, table1.cell(0,1).text = '사업장', '[          지점]'
                    table1.cell(0,2).text, table1.cell(0,3).text = '평가 일시', current_time_str
                    table1.cell(1,0).text, table1.cell(1,1).text = '평가 대상', current_address
                    table1.cell(1,2).text, table1.cell(1,3).text = '평가 종류', '□ 최초 □ 수시\n□ 정기 □ 상시'
                    table1.cell(2,0).text, table1.cell(2,1).text = '평가 방법', '현장점검·사진 1장'
                    table1.cell(2,2).text, table1.cell(2,3).text = '평가 결과', f'{total_count}건, 상 {count_high}·중 {count_mid}·하 {count_low}'
                    table1.cell(3,0).text, table1.cell(3,1).text = '평가 책임자', '         (서명)'
                    table1.cell(3,2).text, table1.cell(3,3).text = '참여 근로자', '         (서명)'
                    
                    format_table(table1, is_header_top=False)
                    for r in range(4):
                        apply_cell_shading(table1.cell(r, 0), "D9D9D9")
                        apply_cell_shading(table1.cell(r, 2), "D9D9D9")
                    doc.add_paragraph()
                    
                    # 2. 절차 테이블
                    doc.add_heading('1. 절차', level=2)
                    table2 = doc.add_table(rows=7, cols=3)
                    table2.style = 'Table Grid'
                    hdr = table2.rows[0].cells
                    hdr[0].text, hdr[1].text, hdr[2].text = '절차', 'AI 판독', '관리자 확인'
                    table2.cell(1,0).text = '가. 사전준비(제9조)'
                    table2.cell(1,1).text = '위험성 수준 기준 기본안'
                    table2.cell(1,2).text = '위험성의 수준과 그 수준을 판단하는 기준, 허용 가능한 위험성의 수준을 확정(제2항)'
                    table2.cell(2,0).text = '나. 유해·위험요인 파악(제10조)'
                    table2.cell(2,1).text = f'현장점검(제10조 제1호)간 첨부된 사진에서 {total_count}건'
                    table2.cell(2,2).text = '사진에 포함되지 않은 현장 위험물(화학물질·소음·중량물 등)과 근로자가 제안한 위험요인도 추가하여 파악'
                    table2.cell(3,0).text = '다. 위험성 결정(제11조)'
                    table2.cell(3,1).text = '상·중·하 수준으로 분류'
                    table2.cell(3,2).text = '허용 가능한지 결정, 현장 근로자 참여(제6조)'
                    table2.cell(4,0).text = '라. 위험성 감소대책 수립·실행(제12조)'
                    table2.cell(4,1).text = '감소대책 수립(제거·대체 → 공학적 대책 → 관리적 대책 → 보호구 착용)'
                    table2.cell(4,2).text = '감소대책에 대한 담당자 및 완료일 작성, 실행 후 위험성 수준 재확인'
                    table2.cell(5,0).text = '마. 공유(제13조)'
                    table2.cell(5,1).text = '작업 전 위험요인과 감소대책을 공유'
                    table2.cell(5,2).text = '작업 전 안전점검회의(TBM)에서 공유하고 서명'
                    table2.cell(6,0).text = '바. 기록·보존(제14조)'
                    table2.cell(6,1).text = '본 문서'
                    table2.cell(6,2).text = '서명한 본 문서를 보관'
                    
                    format_table(table2)
                    doc.add_paragraph("※ 매월 1회 이상 위험개소 순회점검 통해 유해·위험요인을 발굴하여 제11조의 위험성 결정 및 제12조의 위험성 감소대책을 수립·실행 할 것.(사업장 위험성 평가에 관한 지침 제15조 제4항)")
                    
                    # 3. 법이 정한 의무 테이블
                    doc.add_heading('2. 법이 정한 의무', level=2)
                    table3 = doc.add_table(rows=4, cols=4)
                    table3.style = 'Table Grid'
                    hdr = table3.rows[0].cells
                    hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text = '의무', '조문', '과태료', '비고'
                    table3.cell(1,0).text, table3.cell(1,1).text, table3.cell(1,2).text, table3.cell(1,3).text = '위험성평가 실시', '제36조 제1항', '1천만원 이하', '위험성평가표 전체'
                    table3.cell(2,0).text, table3.cell(2,1).text, table3.cell(2,2).text, table3.cell(2,3).text = '근로자의 참여\n근로자대표 참여\n결과 알림', '제36조 제2항ㆍ제3항ㆍ제4항 전단', '500만원 이하', '참여 근로자 서명\n작업 전 공유'
                    table3.cell(3,0).text, table3.cell(3,1).text, table3.cell(3,2).text, table3.cell(3,3).text = '결과 기록·보존', '제36조 제5항', '300만원 이하', '서명된 본 평가표 보관'
                    
                    format_table(table3)
                    doc.add_paragraph("※ 과태료 시행 : 상시 50명 이상(건설업은 공사 금액 50억원 이상) 2027년 1월 1일 시행(산업안전보건법 부칙 제1조). 중대재해로 이어질 수 있는 위험은 작업 전 안전점검회의(TBM) 등을 통해 상시 알리도록 노력해야 합니다(제36조 제4항 후단).")
                    
                    # 4. 위험성 수준 기준 테이블
                    doc.add_heading('3. 위험성 수준 기준(AI 판독)', level=2)
                    table4 = doc.add_table(rows=6, cols=4)
                    table4.style = 'Table Grid'
                    hdr = table4.rows[0].cells
                    hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text = '수준', '기준 제시', '판단', '조치'
                    
                    criteria_data = data.get("criteria", {})
                    high_c = criteria_data.get("high", {})
                    mid_c = criteria_data.get("medium", {})
                    low_c = criteria_data.get("low", {})
                    legal_basis = criteria_data.get("legal_basis", "산업안전보건기준에 관한 규칙 참조")
                    
                    table4.cell(1,0).text, table4.cell(1,1).text, table4.cell(1,2).text, table4.cell(1,3).text = '상', high_c.get("condition", ""), '허용 불가', high_c.get("action", "")
                    table4.cell(2,0).text, table4.cell(2,1).text, table4.cell(2,2).text, table4.cell(2,3).text = '중', mid_c.get("condition", ""), '허용 불가', mid_c.get("action", "")
                    table4.cell(3,0).text, table4.cell(3,1).text, table4.cell(3,2).text, table4.cell(3,3).text = '하', low_c.get("condition", ""), '허용 가능', low_c.get("action", "")
                    table4.cell(4,0).text = '판단 근거'
                    table4.cell(4,1).text = legal_basis
                    table4.cell(4,1).merge(table4.cell(4,3))
                    table4.cell(5,0).merge(table4.cell(5,3)).text = '위 판독 기준으로 정함.', '     근로자              (서명)'
                    
                    format_table(table4)
                    apply_cell_shading(table4.cell(5, 0), "D9D9D9")
                    doc.add_page_break()
                    
                    # 5. 위험성평가표
                    doc.add_heading('4. 위험성 평가표', level=2)
                    table5 = doc.add_table(rows=1, cols=7)
                    table5.style = 'Table Grid'
                    hdr = table5.rows[0].cells
                    hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text, hdr[4].text, hdr[5].text, hdr[6].text = '번호', '유해·위험요인\n(사진에서 보이는 것)', '재해형태', '현재 조치', '위험 수준', '감소 대책\n(제12조)', '담당\n완료일'
                    
                    for item in hazards:
                        row = table5.add_row().cells
                        row[0].text = str(item.get("no", ""))
                        row[1].text = item.get("factor", "")
                        row[2].text = item.get("type", "")
                        row[3].text = item.get("current_measure", "")
                        row[4].text = item.get("level", "")
                        row[5].text = item.get("countermeasure", "")
                        row[6].text = '담당자\n\n완료일\n년  월  일\n\n(서명)'
                    
                    format_table(table5)
                    
                    # 6. TBM
                    doc.add_heading('5. 안전점검회의(TBM) 공유', level=2)
                    for i, tbm_text in enumerate(data.get("tbm", [])):
                        doc.add_paragraph(f"{i+1}. {tbm_text}")
                    
                    table6 = doc.add_table(rows=2, cols=3)
                    table6.style = 'Table Grid'
                    hdr = table6.rows[0].cells
                    hdr[0].text, hdr[1].text, hdr[2].text = '공유 일시', '공유한 사람', '참석자 서명'
                    format_table(table6)
                    
                    doc.add_paragraph("\n※ 이 문서는 초안입니다. 유해·위험요인은 본 AI 모델이 사진에서 분석한 결과이고, 위험성 수준과 감소대책 후보는 코드가 정해진 규칙으로 고른 것입니다. 허용 가능 여부의 결정, 대책 선택과 실행, 근로자 참여는 관리자가 합니다.(사업장 위험성 평가에 관한 지침 제6조·제11조·제12조)")
                    doc.add_paragraph("※ 사진에 보이지 않는 위험(화학물질 소음 분진·중량물 취급 작업 자세 등)은 이 표에 들어 있지 않습니다. 파악 단계에서 추가 작성해 주세요.")
                    
                    # 파일 저장
                    buffer = io.BytesIO()
                    doc.save(buffer)
                    buffer.seek(0)
                    
                    # 문서 생성 완료 표시
                    st.success("✅ 문서 작성이 완료되었습니다! 아래 버튼을 눌러 다운로드하세요.")
                    st.download_button(
                        label="📄 현장 위험성 평가표(.docx) 다운로드",
                        data=buffer,
                        file_name=f"현장_위험성_평가표_{datetime.now(korea_tz).strftime('%Y%m%d_%H%M%S')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary" # 다운로드 버튼 색상 강조
                    )
                        
                except Exception as e:
                    st.error(f"오류가 발생했습니다: {e}")