import os
import time
import html
import re
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==========================================
# [1] 사용자 설정 구역 (텔레그램 정보 입력)
# ==========================================
KEYWORD = "한전MCS"
CHECK_INTERVAL = 300  # 검색 주기 (300초 = 5분)

# 깃허브 시크릿(Environment Variables)에서 토큰을 안전하게 불러옵니다.
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")

# 현재 파이썬 파일(C:\전코딩) 위치로 저장 경로 고정 (권한 오류 방지)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, "sent_links.txt")
EXCEL_FILE = os.path.join(BASE_DIR, "한전MCS_뉴스모니터링_대장.xlsx")
# ==========================================

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7"
}

def clean_text(raw_text):
    if not raw_text:
        return ""
    text = re.sub(r"<[^>]+>", "", raw_text)
    text = html.unescape(text)
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return text.strip()

def clean_for_excel(text):
    return text.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")

# [중복 방지] 안전한 링크 파일 읽기 및 저장
def load_sent_links():
    if not os.path.exists(DB_FILE) or os.path.isdir(DB_FILE):
        return set()
    try:
        with open(DB_FILE, "r", encoding="utf-8") as f:
            return set(line.strip() for line in f if line.strip())
    except Exception as e:
        print(f"⚠️ 링크 파일 읽기 오류: {e}")
        return set()

def save_sent_link(link):
    try:
        with open(DB_FILE, "a", encoding="utf-8") as f:
            f.write(link + "\n")
    except PermissionError:
        print(f"⚠️ [권한 오류] '{DB_FILE}' 파일이 잠겨 있거나 읽기 전용입니다.")
    except Exception as e:
        print(f"⚠️ 링크 저장 오류: {e}")

# [수집 1] 네이버 뉴스 크롤링 (언론사명 + 작성시간 포함)
def fetch_naver_news(keyword):
    url = "https://search.naver.com/search.naver"
    params = {
        "where": "news",
        "query": keyword,
        "sort": "1"  # 최신순
    }
    
    news_list = []
    try:
        res = requests.get(url, params=params, headers=HEADERS, timeout=10)
        if res.status_code == 200:
            soup = BeautifulSoup(res.text, "html.parser")
            items = soup.select("div.news_wrap, li.bx, div.sds-comps-vertical-layout")
            
            for item in items:
                # 1. 제목 및 링크 추출
                title_el = item.select_one("a.news_tit, a.sds-comps-text-type-headline1, a[data-heatmap-target='.tit']")
                if not title_el:
                    continue
                title = clean_text(title_el.get_text())
                link = title_el.get("href", "")
                if not link or not link.startswith("http"):
                    continue

                # 2. 본문 요약 추출
                desc_el = item.select_one("div.news_dsc, a.api_txt_lines.dsc_txt_wrap, a.sds-comps-text-type-body1, a[data-heatmap-target='.body']")
                desc = clean_text(desc_el.get_text()) if desc_el else "본문 요약 없음"

                # 3. 언론사명 추출
                press_el = item.select_one("a.info.press, span.sds-comps-profile-info-title-text, a[data-heatmap-target='.prof']")
                press = clean_text(press_el.get_text()).replace("언론사 선정", "").strip() if press_el else "네이버 뉴스"

                # 4. 기사 작성 시간 추출 (예: '2시간 전', '2026.10.26.')
                pub_date = "최근"
                info_spans = item.select("span.info, span.sds-comps-profile-info-subtext")
                for sp in info_spans:
                    txt = clean_text(sp.get_text())
                    if any(k in txt for k in ["전", "분", "시간", "일", "202"]):
                        pub_date = txt
                        break

                news_list.append({
                    "portal": "네이버 뉴스",
                    "press": press,
                    "pub_date": pub_date,
                    "title": title,
                    "summary": desc,
                    "link": link
                })
                if len(news_list) >= 10:
                    break
        else:
            print(f"[네이버 크롤링 오류] 상태 코드: {res.status_code}")
    except Exception as e:
        print(f"[네이버 크롤링 중 예외 발생] {e}")
        
    return news_list

# [수집 2] 다음(Daum) 뉴스 크롤링 (언론사명 + 작성시간 포함)
def fetch_daum_news(keyword):
    url = "https://search.daum.net/search"
    params = {
        "w": "news",
        "q": keyword,
        "sort": "recency"
    }
    
    news_list = []
    try:
        res = requests.get(url, params=params, headers=HEADERS, timeout=10)
        if res.status_code == 200:
            soup = BeautifulSoup(res.text, "html.parser")
            items = soup.select("ul.c-list-basic > li")
            
            for item in items[:10]:
                # 1. 제목 및 링크
                title_el = item.select_one("div.item-title a")
                if not title_el:
                    continue
                title = clean_text(title_el.get_text())
                link = title_el.get("href", "")
                if not link:
                    continue

                # 2. 본문 요약
                desc_el = item.select_one("p.conts-desc")
                desc = clean_text(desc_el.get_text()) if desc_el else "요약 내용 없음"

                # 3. 언론사명 추출
                press_el = item.select_one("div.c-tit-doc span.tit_item, span.txt_info")
                press = clean_text(press_el.get_text()) if press_el else "다음 뉴스"

                # 4. 기사 작성 시간 추출
                date_el = item.select_one("span.gem-subinfo, span.txt_info:nth-of-type(2)")
                pub_date = clean_text(date_el.get_text()) if date_el else "최근"
                
                news_list.append({
                    "portal": "다음 뉴스",
                    "press": press,
                    "pub_date": pub_date,
                    "title": title,
                    "summary": desc,
                    "link": link
                })
    except Exception as e:
        print(f"[다음 크롤링 중 예외 발생] {e}")
        
    return news_list

# [분석] 키워드 기반 자동 논조 판별기
def analyze_sentiment(title, summary):
    positive_words = ["혁신", "성공", "표창", "수상", "협약", "봉사", "기부", "나눔", "강화", "개선", "최우수", "선도", "확대", "호평"]
    negative_words = ["논란", "사고", "비판", "갈등", "파업", "적자", "우려", "지적", "처벌", "위반", "고발", "중단"]
    
    combined = title + " " + summary
    pos_score = sum(1 for w in positive_words if w in combined)
    neg_score = sum(1 for w in negative_words if w in combined)
    
    if pos_score > neg_score:
        return "긍정 (우호적 보도)"
    elif neg_score > pos_score:
        return "부정 (이슈 점검)"
    else:
        return "중립 (스트레이트)"

# [엑셀 저장] 언론사명과 발행일시를 포함하여 8열로 기록
def save_to_excel(article, sentiment):
    now_str = time.strftime("%Y-%m-%d %H:%M:%S")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    try:
        if not os.path.exists(EXCEL_FILE):
            wb = Workbook()
            ws = wb.active
            ws.title = "뉴스모니터링_내역"
            
            headers = ["수집일시", "포털 구분", "언론사명", "기사 발행일시", "논조 분류", "기사 제목", "핵심 요약", "원문 링크"]
            ws.append(headers)
            
            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            header_font = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")
            
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            
            col_widths = {"A": 20, "B": 13, "C": 18, "D": 16, "E": 18, "F": 45, "G": 55, "H": 40}
            for col_letter, width in col_widths.items():
                ws.column_dimensions[col_letter].width = width
            ws.freeze_panes = "A2"
        else:
            wb = load_workbook(EXCEL_FILE)
            ws = wb.active

        row_data = [
            now_str,
            article["portal"],
            clean_for_excel(article["press"]),
            clean_for_excel(article["pub_date"]),
            sentiment,
            clean_for_excel(article["title"]),
            clean_for_excel(article["summary"]),
            article["link"]
        ]
        ws.append(row_data)
        curr_row = ws.max_row

        for col_idx in range(1, 9):
            cell = ws.cell(row=curr_row, column=col_idx)
            cell.font = Font(name="맑은 고딕", size=10)
            cell.border = thin_border
            if col_idx in [1, 2, 3, 4, 5]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 8:
                cell.hyperlink = article["link"]
                cell.font = Font(name="맑은 고딕", size=10, color="0563C1", underline="single")
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        wb.save(EXCEL_FILE)
    except PermissionError:
        print(f"⚠️ [주의] '{EXCEL_FILE}' 파일이 열려 있어 저장하지 못했습니다. 엑셀 창을 닫아주세요!")
    except Exception as e:
        print(f"[엑셀 저장 오류] {e}")

# [전송] 텔레그램 메시지 발송 (언론사명 + 발행시간 표시)
def send_telegram_alert(article, sentiment):
    portal = article["portal"]
    press = article["press"]
    pub_date = article["pub_date"]
    title = article["title"]
    summary = article["summary"]
    link = article["link"]
    
    icon = "🟢" if "네이버" in portal else "🔵"
    
    message = (
        f"🚨 <b>[{icon} {portal}] {KEYWORD} 신규 기사</b>\n"
        f"━━━━━━━━━━━━━━━━━━\n"
        f"🏢 <b>매체 및 시간:</b> {press} ({pub_date})\n"
        f"📰 <b>제목:</b> <a href='{link}'>{title}</a>\n"
        f"📊 <b>자동 분류:</b> {sentiment}\n\n"
        f"💡 <b>핵심 내용 요약:</b>\n{summary}\n\n"
        f"🔗 <a href='{link}'>👉 기사 원문 바로가기 (클릭)</a>"
    )
    
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": message,
        "parse_mode": "HTML",
        "disable_web_page_preview": False
    }
    
    res = requests.post(url, json=payload, timeout=10)
    return res.status_code == 200

# [메인 루프] 주기적 실행
def main():
    print(f"🚀 [{KEYWORD}] 네이버·다음 뉴스 모니터링 (매체명·발행시간 포함) 시작...")
    sent_links = load_sent_links()
    
    INITIAL_SKIP = False

    while True:
        naver_articles = fetch_naver_news(KEYWORD)
        daum_articles = fetch_daum_news(KEYWORD)
        print(f"🔍 검색된 기사 수 -> 네이버: {len(naver_articles)}건 / 다음: {len(daum_articles)}건")
        
        all_articles = naver_articles + daum_articles
        new_count = 0
        
        for article in reversed(all_articles):
            link = article["link"]
            if link not in sent_links:
                sent_links.add(link)
                save_sent_link(link)
                
                sentiment = analyze_sentiment(article["title"], article["summary"])
                save_to_excel(article, sentiment)
                
                if not INITIAL_SKIP:
                    success = send_telegram_alert(article, sentiment)
                    if success:
                        new_count += 1
                        print(f"✅ [{article['press']}] {article['title']}")
                    time.sleep(1)
        
        INITIAL_SKIP = False
        
        current_time = time.strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{current_time}] 점검 완료 (신규 발송: {new_count}건) - {CHECK_INTERVAL}초 후 재검색합니다.\n")
        time.sleep(CHECK_INTERVAL)

if __name__ == "__main__":
    main()
