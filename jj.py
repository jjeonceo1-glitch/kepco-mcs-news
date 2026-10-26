import requests

# 1단계와 3단계에서 복사한 값을 큰따옴표 안에 붙여넣으세요.
TELEGRAM_TOKEN = "8747050640:AAHVjxpVytdjMjvwXzuwXwSyk6jjYoagRDs"
TELEGRAM_CHAT_ID = "8166249552"

def send_test_alert():
    # 테스트용 가상 데이터
    portal = "네이버 뉴스"
    title = "한전MCS, 전력 서비스 현장 디지털 혁신 가속화"
    ai_summary = (
        "• 현장 맞춤형 스마트 시스템 도입으로 업무 효율성 대폭 향상\n"
        "• 도서·산간 지역 전력 서비스 품질 제고 및 안전 관리 강화\n"
        "• ESG 경영 실천을 통한 공공기관 사회적 책임 이행"
    )
    sentiment = "🟢 긍정 (우호적 보도)"
    link = "https://news.naver.com"

    # 텔레그램 메시지 디자인 (HTML 태그 적용)
    message = (
        f"🚨 <b>[{portal}] 한전MCS 신규 기사 알림</b>\n"
        f"━━━━━━━━━━━━━━━━━━\n"
        f"📰 <b>제목:</b> <a href='{link}'>{title}</a>\n"
        f"📊 <b>논조:</b> {sentiment}\n\n"
        f"💡 <b>핵심 3줄 요약:</b>\n{ai_summary}\n\n"
        f"🔗 <a href='{link}'>👉 기사 원문 바로가기 (클릭)</a>"
    )

    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": message,
        "parse_mode": "HTML",               # 굵은 글씨(<b>)와 하이퍼링크(<a>) 활성화
        "disable_web_page_preview": False   # 하단에 링크 썸네일 미리보기 카드 표시
    }

    response = requests.post(url, json=payload)
    
    if response.status_code == 200:
        print("✅ 성공! 텔레그램 채팅방을 확인해 보세요.")
    else:
        print(f"❌ 전송 실패 (오류 코드 {response.status_code}): {response.text}")

if __name__ == "__main__":
    send_test_alert()