import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_mcs_to_calculator(filename="한전MCS_지점별_적정TO_및_SLA총액산출_시뮬레이터.xlsx"):
    wb = openpyxl.Workbook()
    
    # ---------------------------------------------------------
    # 1. [시트 1] 산출 기준 및 파라미터 설정 시트 (Config)
    # ---------------------------------------------------------
    ws_cfg = wb.active
    ws_cfg.title = "1_산출기준_파라미터"
    
    # 스타일 정의
    title_font = Font(name="맑은 고딕", size=14, bold=True, color="FFFFFF")
    header_font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    sub_header_font = Font(name="맑은 고딕", size=10, bold=True, color="1F4E79")
    bold_font = Font(name="맑은 고딕", size=10, bold=True)
    regular_font = Font(name="맑은 고딕", size=10)
    
    navy_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    blue_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    input_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # 노란색 (입력셀)
    calc_fill = PatternFill(start_color="E9EDF4", end_color="E9EDF4", fill_type="solid")  # 연청색 (핵심산출셀)
    point_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") # 연녹색 (최종결과셀)
    warn_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")  # 연주황 (현행비교셀)
    
    thin_border = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF")
    )
    
    # 시트1 타이틀
    ws_cfg.merge_cells("A1:D1")
    ws_cfg["A1"] = "한전MCS 권역별 적정 T/O 및 SLA 총액제 산출 기준표 (Control Panel)"
    ws_cfg["A1"].font = title_font
    ws_cfg["A1"].fill = navy_fill
    ws_cfg["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_cfg.row_dimensions[1].height = 30
    
    cfg_headers = ["구분", "파라미터 항목", "설정 기준값", "단위 / 설명"]
    for col_idx, text in enumerate(cfg_headers, 1):
        cell = ws_cfg.cell(row=3, column=col_idx, value=text)
        cell.font = header_font
        cell.fill = blue_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    ws_cfg.row_dimensions[3].height = 24
    
    # 기준값 데이터 (노란색 셀을 수정하면 시트2 전체에 자동 반영됨)
    cfg_data = [
        ("근로 및 예비율", "1인당 월 실근로시간", 9600, "분 (월 160시간 기준)"),
        ("근로 및 예비율", "법정휴가·교육 백업 예비율", 1.15, "배 (15% 대체인력 계수)"),
        ("과업 표준품셈", "Ⅰ급(도심) 호당 난검침 시간", 5.0, "분/호 (이동+지침확인)"),
        ("과업 표준품셈", "Ⅱ급(도농) 호당 난검침 시간", 9.0, "분/호 (중거리 이동 포함)"),
        ("과업 표준품셈", "Ⅲ급(산간) 호당 난검침 시간", 15.0, "분/호 (장거리 점재 이동)"),
        ("과업 표준품셈", "AMI에러 검침", 12.0, "분/건 (검침)"),
        ("과업 표준품셈", "단전·재공급·송달·복지", 15.0, "분/건 (현장출동 평균 소요시간)"),
        ("인건비 및 경비", "전력매니저 월 인건비 원가", 285.0, "만원/월 (퇴직금·4대보험 포함)"),
        ("인건비 및 경비", "지점장/차장 월 인건비 원가", 350.0, "만원/월 (퇴직금·4대보험 포함)"),
        ("인건비 및 경비", "회사차량 월 고정 유지비", 32.0, "만원/대 (리스·보험·기본정비)"),
        ("인건비 및 경비", "산간용 SUV 월 고정 유지비", 42.0, "만원/대 (험로·장거리 할증)"),
        ("SLA 계약비율", "고정 기본료 인건비 보장률", 0.85, "85% (물량 무관 확정 지급 비율)"),
        ("SLA 계약비율", "SLA 성과금 및 적정이윤율", 0.055, "5.5% (총원가 대비 목표 마진율)")
    ]
    
    for idx, (cat, name, val, desc) in enumerate(cfg_data, start=4):
        ws_cfg.row_dimensions[idx].height = 20
        c1 = ws_cfg.cell(row=idx, column=1, value=cat)
        c2 = ws_cfg.cell(row=idx, column=2, value=name)
        c3 = ws_cfg.cell(row=idx, column=3, value=val)
        c4 = ws_cfg.cell(row=idx, column=4, value=desc)
        
        c1.font, c2.font, c3.font, c4.font = bold_font, regular_font, bold_font, regular_font
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c3.alignment = Alignment(horizontal="right", vertical="center")
        c3.fill = input_fill
        if isinstance(val, float) and val < 1.0:
            c3.number_format = "0.0%"
        elif isinstance(val, float):
            c3.number_format = "#,##0.0"
        else:
            c3.number_format = "#,##0"
            
        for c in (c1, c2, c3, c4):
            c.border = thin_border
            
    for col_idx, width in enumerate([18, 30, 16, 35], 1):
        ws_cfg.column_dimensions[get_column_letter(col_idx)].width = width

    # ---------------------------------------------------------
    # 2. [시트 2] 지소별 적정 T/O 및 SLA 총액 산출 시뮬레이터
    # ---------------------------------------------------------
    ws_sim = wb.create_sheet(title="2_지점별_TO_총액산출표")
    
    ws_sim.merge_cells("A1:W1")
    ws_sim["A1"] = "전국 지점별 권역 등급 자동 판정 · 적정 최소 정원(T/O) · SLA 기반 월 총액 산출 시뮬레이터 (단위: 호, 분, 명, 만원)"
    ws_sim["A1"].font = title_font
    ws_sim["A1"].fill = navy_fill
    ws_sim["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_sim.row_dimensions[1].height = 30
    
    headers = [
        "지점명", "관할 총호수\n(호)", "관할 면적\n(km²)", "도서포함\n(Y/N)", "AMI 보급률\n(%)",
        "월 모뎀점검\n(건)", "월 단전·복지\n(건)", "수용가 밀도\n(호/km²)", "자동 판정\n권역 등급", "잔여 난검침\n(호)",
        "호당 시간\n(분)", "월 총업무량\n(분)", "권역 최소하한\n(명)", "산출 매니저\n(명)", "관리·내근\n(명)",
        "최종 적정 T/O\n(명)", "배정 차량\n(대)", "지점 실투입원가\n(만원/월)", "① 월 고정기본료\n(만원/월)", "② 활동 변동급\n(만원/월)",
        "③ SLA 성과금\n(만원/월)", "제안 월 정산총액\n(만원/월)", "현행 단가제 대비\n수지개선(만원)"
    ]
    
    for col_idx, h_text in enumerate(headers, 1):
        cell = ws_sim.cell(row=2, column=col_idx, value=h_text)
        cell.font = header_font
        cell.fill = blue_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws_sim.row_dimensions[2].height = 32
    
    # 샘플 지소 10개 데이터 (도심형, 도농복합형, 산간도서형 대표 사례)
    sample_branches = [
        ("노원도봉지점", 135000, 65, "N", 0.90, 2200, 2600),
        ("경기지사직할지점", 115000, 78, "N", 0.88, 1900, 2300),
        ("서초지점", 125000, 92, "N", 0.91, 2000, 2400),
        ("고양지점", 62000, 340, "N", 0.89, 1300, 1100),
        ("강원지사직할지점", 75000, 290, "N", 0.90, 1400, 1300),
        ("동청주지점", 52000, 310, "N", 0.88, 1100, 950),
        ("세종지점", 48000, 480, "N", 0.87, 1000, 900),
        ("상주지점", 24000, 720, "N", 0.90, 800, 500),
        ("의성지점", 16000, 650, "N", 0.92, 650, 400),
        ("성주지점", 19000, 430, "Y", 0.89, 750, 480),
    ]
    
    cfg_sheet = "'1_산출기준_파라미터'"
    
    for idx, b_data in enumerate(sample_branches, start=3):
        ws_sim.row_dimensions[idx].height = 22
        r = idx
        
        # A~G열: 기초 입력 데이터 (노란색 배경)
        for c_i, val in enumerate(b_data, 1):
            cell = ws_sim.cell(row=r, column=c_i, value=val)
            cell.fill = input_fill
            cell.font = bold_font if c_i == 1 else regular_font
            cell.border = thin_border
            if c_i in [1, 4]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_i == 5:
                cell.number_format = "0.0%"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
                
        # H열: 수용가 밀도 (호/km²)
        c_h = ws_sim.cell(row=r, column=8, value=f"=ROUND(B{r}/C{r}, 1)")
        c_h.number_format = "#,##0.0"
        
        # I열: 권역 등급 자동 판정
        c_i = ws_sim.cell(
            row=r, column=9,
            value=f'=IF(OR(D{r}="Y",C{r}>=500,H{r}<200),"Ⅲ등급(산간도서)",IF(AND(C{r}<100,H{r}>=1200),"Ⅰ등급(도심밀집)","Ⅱ등급(도농복합)"))'
        )
        c_i.alignment = Alignment(horizontal="center", vertical="center")
        c_i.font = sub_header_font
        c_i.fill = calc_fill
        
        # J열: 잔여 난검침 가구수
        c_j = ws_sim.cell(row=r, column=10, value=f"=ROUND(B{r}*(1-E{r}), 0)")
        c_j.number_format = "#,##0"
        
        # K열: 등급별 호당 소요시간 (시트1 참조)
        c_k = ws_sim.cell(
            row=r, column=11,
            value=f'=IF(LEFT(I{r},2)="Ⅰ등",{cfg_sheet}!$C$6,IF(LEFT(I{r},2)="Ⅱ등",{cfg_sheet}!$C$7,{cfg_sheet}!$C$8))'
        )
        c_k.number_format = "0.0"
        
        # L열: 월 총 업무시간(분)
        c_l = ws_sim.cell(
            row=r, column=12,
            value=f"=(J{r}*K{r})+(F{r}*{cfg_sheet}!$C$9)+(G{r}*{cfg_sheet}!$C$10)"
        )
        c_l.number_format = "#,##0"
        
        # M열: 권역 최소 하한 매니저 수
        c_m = ws_sim.cell(row=r, column=13, value=f'=IF(LEFT(I{r},2)="Ⅰ등",8,6)')
        c_m.number_format = "#,##0"
        c_m.alignment = Alignment(horizontal="center", vertical="center")
        
        # N열: 산출 현장 매니저 정원
        c_n = ws_sim.cell(
            row=r, column=14,
            value=f"=MAX(ROUNDUP((L{r}/{cfg_sheet}!$C$4)*{cfg_sheet}!$C$5, 0), M{r})"
        )
        c_n.number_format = "#,##0"
        c_n.font = bold_font
        
        # O열: 관리자 및 내근 정원
        c_o = ws_sim.cell(row=r, column=15, value=f'=IF(LEFT(I{r},2)="Ⅰ등",2,1)')
        c_o.number_format = "#,##0"
        
        # P열: 최종 적정 유지 정원 (T/O)
        c_p = ws_sim.cell(row=r, column=16, value=f"=N{r}+O{r}")
        c_p.number_format = "#,##0"
        c_p.font = bold_font
        c_p.fill = point_fill
        
        # Q열: 배정 기동차량 대수
        c_q = ws_sim.cell(
            row=r, column=17,
            value=f'=IF(LEFT(I{r},2)="Ⅰ등",ROUNDUP(N{r}/3,0),IF(LEFT(I{r},2)="Ⅱ등",ROUNDUP(N{r}*0.65,0),N{r}))'
        )
        c_q.number_format = "#,##0"
        
        # R열: 지소 실제 투입 총원가 (만원)
        c_r = ws_sim.cell(
            row=r, column=18,
            value=f'=(N{r}*{cfg_sheet}!$C$11)+(O{r}*{cfg_sheet}!$C$12)+(Q{r}*IF(LEFT(I{r},2)="Ⅲ등",{cfg_sheet}!$C$14,{cfg_sheet}!$C$13))+IF(LEFT(I{r},2)="Ⅰ등",394,IF(LEFT(I{r},2)="Ⅱ등",218,133))'
        )
        c_r.number_format = "#,##0.0"
        c_r.font = bold_font
        
        # S열: ① 월 고정 기본료 (Base Capacity Fee)
        c_s = ws_sim.cell(
            row=r, column=19,
            value=f'=ROUND(((N{r}*{cfg_sheet}!$C$11+O{r}*{cfg_sheet}!$C$12)*{cfg_sheet}!$C$15)+(Q{r}*IF(LEFT(I{r},2)="Ⅲ등",{cfg_sheet}!$C$14,{cfg_sheet}!$C$13))+(IF(LEFT(I{r},2)="Ⅰ등",394,IF(LEFT(I{r},2)="Ⅱ등",218,133))*0.5), 1)'
        )
        c_s.number_format = "#,##0.0"
        c_s.font = bold_font
        c_s.fill = calc_fill
        
        # T열: ② 활동 변동급 (Activity Fee)
        c_t = ws_sim.cell(row=r, column=20, value=f"=ROUND((R{r}-S{r})*1.05, 1)")
        c_t.number_format = "#,##0.0"
        
        # U열: ③ SLA 성과금 및 이윤
        c_u = ws_sim.cell(row=r, column=21, value=f"=ROUND(R{r}*{cfg_sheet}!$C$16, 1)")
        c_u.number_format = "#,##0.0"
        
        # V열: 제안 월 정산 총액 (①+②+③)
        c_v = ws_sim.cell(row=r, column=22, value=f"=S{r}+T{r}+U{r}")
        c_v.number_format = "#,##0.0"
        c_v.font = bold_font
        c_v.fill = point_fill
        
        # W열: 현행 호당단가제 대비 월 수지 개선액
        c_w = ws_sim.cell(
            row=r, column=23,
            value=f"=V{r}-ROUND((J{r}*0.07)+(F{r}*0.15)+(G{r}*0.35), 1)"
        )
        c_w.number_format = "+#,##0.0;-#,##0.0;0"
        c_w.font = bold_font
        c_w.fill = warn_fill
        
        for col_i in range(8, 24):
            ws_sim.cell(row=r, column=col_i).border = thin_border

    # 하단 합계 / 평균 요약 행 (13행)
    sum_row = len(sample_branches) + 3
    ws_sim.row_dimensions[sum_row].height = 25
    ws_sim.cell(row=sum_row, column=1, value="합계 / 평균").alignment = Alignment(horizontal="center", vertical="center")
    
    sum_cols = [2, 3, 6, 7, 10, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23]
    for col_i in range(1, 24):
        cell = ws_sim.cell(row=sum_row, column=col_i)
        col_let = get_column_letter(col_i)
        cell.font = bold_font
        cell.fill = calc_fill
        cell.border = thin_border
        if col_i in sum_cols:
            cell.value = f"=SUM({col_let}3:{col_let}{sum_row-1})"
            cell.number_format = "#,##0" if col_i < 18 else "#,##0.0"
        elif col_i == 5:
            cell.value = f"=AVERAGE(E3:E{sum_row-1})"
            cell.number_format = "0.0%"

    # 열 너비 자동 최적화 및 틀 고정
    col_widths = [18, 13, 11, 10, 12, 12, 12, 12, 16, 13, 10, 14, 12, 12, 10, 14, 11, 16, 16, 14, 14, 16, 16]
    for col_idx, width in enumerate(col_widths, 1):
        ws_sim.column_dimensions[get_column_letter(col_idx)].width = width
    ws_sim.freeze_panes = "B3"
    
    wb.save(filename)
    print(f"[생성 완료] '{filename}' 파일이 성공적으로 생성되었습니다.")

if __name__ == "__main__":
    create_mcs_to_calculator()