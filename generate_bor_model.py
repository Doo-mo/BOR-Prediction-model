"""
LNG 탱크 BOR(Boil-Off Rate) 예측 엑셀 모델 자동 생성 스크립트
---------------------------------------------------------------
정상상태(steady-state) 열전달 기반의 설계 예측용 모델입니다.
BOR = Boil-Off Rate (LNG 증발률, 단위: %/day)
  - 외부에서 탱크 내부로 침투하는 열(Q)에 의해 LNG가 증발하는 가스(BOG)를 계산합니다.
  - 병상이용률(Bed Occupancy Rate)과 전혀 관계없습니다.

실행 방법:
    pip install -r requirements.txt
    python generate_bor_model.py

출력: LNG_BOR_Model.xlsx
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

# ── 색상 상수 ──────────────────────────────────────────────────
CLR_HEADER_BG  = "1F4E79"
CLR_HEADER_FG  = "FFFFFF"
CLR_INPUT_BG   = "EBF3FB"
CLR_RESULT_BG  = "FFF2CC"
CLR_RESULT_EM  = "C00000"
CLR_DB_HDR_BG  = "2E75B6"
CLR_SECTION_BG = "D9E1F2"
CLR_CALC_BG    = "E2EFDA"

INPUT_ROWS = {
    "sec_geo": 3,
    "diameter": 4,
    "length": 5,
    "fill_ratio": 6,
    "tank_volume": 7,
    "lng_volume": 8,
    "sec_oper": 9,
    "pressure": 10,
    "ambient_temp": 11,
    "sec_ins": 12,
    "insulation": 13,
    "insulation_thickness": 14,
    "vacuum_pa": 15,
    "sec_pipe": 16,
    "pipe_state": 17,
    "pipe_ratio": 18,
    "sec_measure": 19,
    "stabilization_time": 20,
    "measure_time": 21,
    "sec_comp": 22,
    "comp_header": 23,
    "comp1": 24,
    "comp2": 25,
    "comp3": 26,
}


def _thin():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)


def _hdr(cell, text, bg=CLR_HEADER_BG, fg=CLR_HEADER_FG, size=11):
    cell.value = text
    cell.font = Font(bold=True, size=size, color=fg)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = _thin()


def _lbl(cell, text, bold=False):
    cell.value = text
    cell.font = Font(bold=bold, size=10)
    cell.alignment = Alignment(vertical="center")
    cell.border = _thin()


def _inp(cell, value):
    cell.value = value
    cell.fill = PatternFill("solid", fgColor=CLR_INPUT_BG)
    cell.font = Font(size=10)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = _thin()


def _calc(cell, formula):
    cell.value = formula
    cell.fill = PatternFill("solid", fgColor=CLR_CALC_BG)
    cell.font = Font(size=10)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = _thin()


def _result(cell, formula, emph=False):
    cell.value = formula
    cell.fill = PatternFill("solid", fgColor=CLR_RESULT_BG)
    color = CLR_RESULT_EM if emph else "000000"
    cell.font = Font(bold=emph, size=11 if emph else 10, color=color)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = _thin()


def _sec(cell, text):
    cell.value = text
    cell.font = Font(bold=True, size=10, color="1F3864")
    cell.fill = PatternFill("solid", fgColor=CLR_SECTION_BG)
    cell.alignment = Alignment(vertical="center")
    cell.border = _thin()


def _cw(ws, col, w):
    ws.column_dimensions[col].width = w


# ════════════════════════════════════════════════════════════════
# 시트 1: Input (입력조건)
# ════════════════════════════════════════════════════════════════
def build_sheet1(wb):
    ws = wb.create_sheet("Input")
    _cw(ws, "A", 32); _cw(ws, "B", 22); _cw(ws, "C", 14); _cw(ws, "D", 18)
    _cw(ws, "F", 30); _cw(ws, "G", 18); _cw(ws, "H", 14)
    _cw(ws, "I", 14); _cw(ws, "J", 14); _cw(ws, "K", 14); _cw(ws, "L", 18)
    r = INPUT_ROWS

    # 타이틀
    ws.merge_cells("A1:D1")
    c = ws["A1"]
    c.value = "LNG 탱크 BOR 예측 모델 — 입력 조건"
    c.font = Font(bold=True, size=14, color=CLR_HEADER_FG)
    c.fill = PatternFill("solid", fgColor=CLR_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # 헤더 행
    for j, t in enumerate(["항목", "값", "단위", "비고"], 1):
        _hdr(ws.cell(2, j), t)
    ws.row_dimensions[2].height = 20

    # ── 행 번호 매핑 (수식 참조에 사용) ──
    # 섹션 행
    ws.merge_cells(f"A{r['sec_geo']}:D{r['sec_geo']}");  _sec(ws[f"A{r['sec_geo']}"], "▶ 탱크 기하 조건")
    # Row 4 : 탱크 내경
    _lbl(ws.cell(r["diameter"], 1), "탱크 내경 (직경) [m]")
    _inp(ws.cell(r["diameter"], 2), 4.0)
    _lbl(ws.cell(r["diameter"], 3), "m"); _lbl(ws.cell(r["diameter"], 4), "원통형 가정")
    # Row 5 : 탱크 길이
    _lbl(ws.cell(r["length"], 1), "탱크 길이 (원통부) [m]")
    _inp(ws.cell(r["length"], 2), 10.0)
    _lbl(ws.cell(r["length"], 3), "m"); _lbl(ws.cell(r["length"], 4), "")
    # Row 6 : 충전율
    _lbl(ws.cell(r["fill_ratio"], 1), "LNG 충전율 [%]")
    _inp(ws.cell(r["fill_ratio"], 2), 90)
    _lbl(ws.cell(r["fill_ratio"], 3), "%"); _lbl(ws.cell(r["fill_ratio"], 4), "0~100")

    _lbl(ws.cell(r["tank_volume"], 1), "탱크 내부 체적 [m³]")
    _calc(
        ws.cell(r["tank_volume"], 2),
        f"=PI()*(B{r['diameter']}/2)^2*B{r['length']}+(4/3)*PI()*(B{r['diameter']}/2)^3",
    )
    ws.cell(r["tank_volume"], 2).number_format = "0.00"
    _lbl(ws.cell(r["tank_volume"], 3), "m³"); _lbl(ws.cell(r["tank_volume"], 4), "내경·길이로 자동 계산")

    _lbl(ws.cell(r["lng_volume"], 1), "LNG 체적 [m³]")
    _calc(ws.cell(r["lng_volume"], 2), f"=B{r['tank_volume']}*B{r['fill_ratio']}/100")
    ws.cell(r["lng_volume"], 2).number_format = "0.00"
    _lbl(ws.cell(r["lng_volume"], 3), "m³"); _lbl(ws.cell(r["lng_volume"], 4), "체적 × 충전율")

    ws.merge_cells(f"A{r['sec_oper']}:D{r['sec_oper']}");  _sec(ws[f"A{r['sec_oper']}"], "▶ 운전 조건")
    # Row 8 : 압력
    _lbl(ws.cell(r["pressure"], 1), "탱크 내부 압력 [bar(a)]")
    _inp(ws.cell(r["pressure"], 2), 1.13)
    _lbl(ws.cell(r["pressure"], 3), "bar(a)"); _lbl(ws.cell(r["pressure"], 4), "절대압")
    # Row 9 : 외부 온도
    _lbl(ws.cell(r["ambient_temp"], 1), "외부(주위) 온도 [℃]")
    _inp(ws.cell(r["ambient_temp"], 2), 25)
    _lbl(ws.cell(r["ambient_temp"], 3), "℃"); _lbl(ws.cell(r["ambient_temp"], 4), "")

    ws.merge_cells(f"A{r['sec_ins']}:D{r['sec_ins']}"); _sec(ws[f"A{r['sec_ins']}"], "▶ 단열재 조건")
    # Row 11 : 단열재 종류
    _lbl(ws.cell(r["insulation"], 1), "단열재 종류")
    _inp(ws.cell(r["insulation"], 2), "Perlite(진공)")
    _lbl(ws.cell(r["insulation"], 3), "-"); _lbl(ws.cell(r["insulation"], 4), "드롭다운 선택")
    # Row 12 : 두께
    _lbl(ws.cell(r["insulation_thickness"], 1), "단열재 두께 [m]")
    _inp(ws.cell(r["insulation_thickness"], 2), 0.30)
    _lbl(ws.cell(r["insulation_thickness"], 3), "m"); _lbl(ws.cell(r["insulation_thickness"], 4), "")
    # Row 13 : 진공도
    _lbl(ws.cell(r["vacuum_pa"], 1), "진공도 [Pa]")
    _inp(ws.cell(r["vacuum_pa"], 2), 1.0)
    _lbl(ws.cell(r["vacuum_pa"], 3), "Pa"); _lbl(ws.cell(r["vacuum_pa"], 4), "진공단열 시 적용")

    ws.merge_cells(f"A{r['sec_pipe']}:D{r['sec_pipe']}"); _sec(ws[f"A{r['sec_pipe']}"], "▶ 배관 조건")
    # Row 15 : 배관 보냉 여부
    _lbl(ws.cell(r["pipe_state"], 1), "배관 보냉/진공 여부")
    _inp(ws.cell(r["pipe_state"], 2), "진공배관")
    _lbl(ws.cell(r["pipe_state"], 3), "-"); _lbl(ws.cell(r["pipe_state"], 4), "드롭다운 선택")
    # Row 16 : 배관 추가 열유입
    _lbl(ws.cell(r["pipe_ratio"], 1), "배관 추가 열유입 비율 [%]")
    _inp(ws.cell(r["pipe_ratio"], 2), 10)
    _lbl(ws.cell(r["pipe_ratio"], 3), "%"); _lbl(ws.cell(r["pipe_ratio"], 4), "탱크 열유입 대비")

    ws.merge_cells(f"A{r['sec_measure']}:D{r['sec_measure']}"); _sec(ws[f"A{r['sec_measure']}"], "▶ 계측 조건")
    # Row 18 : 안정화 시간
    _lbl(ws.cell(r["stabilization_time"], 1), "안정화 시간 [h]")
    _inp(ws.cell(r["stabilization_time"], 2), 12)
    _lbl(ws.cell(r["stabilization_time"], 3), "h"); _lbl(ws.cell(r["stabilization_time"], 4), "BOG 안정화 대기 시간")
    # Row 19 : 계측 시간
    _lbl(ws.cell(r["measure_time"], 1), "BOG 계측 시간 [h]")
    _inp(ws.cell(r["measure_time"], 2), 1.0)
    _lbl(ws.cell(r["measure_time"], 3), "h"); _lbl(ws.cell(r["measure_time"], 4), "실측 수집 시간")

    ws.merge_cells(f"A{r['sec_comp']}:D{r['sec_comp']}"); _sec(ws[f"A{r['sec_comp']}"], "▶ LNG 조성 (상위 3성분)")
    # Row 21 : 조성 헤더
    for j, t in enumerate(["성분 이름", "몰분율 [mol%]", "단위", "비고"], 1):
        _hdr(ws.cell(r["comp_header"], j), t, bg=CLR_DB_HDR_BG)
    # Row 22 : 성분 1
    _inp(ws.cell(r["comp1"], 1), "Methane");  _inp(ws.cell(r["comp1"], 2), 90)
    _lbl(ws.cell(r["comp1"], 3), "mol%");    _lbl(ws.cell(r["comp1"], 4), "주성분")
    # Row 23 : 성분 2
    _inp(ws.cell(r["comp2"], 1), "Ethane");   _inp(ws.cell(r["comp2"], 2), 7)
    _lbl(ws.cell(r["comp2"], 3), "mol%");    _lbl(ws.cell(r["comp2"], 4), "")
    # Row 24 : 성분 3
    _inp(ws.cell(r["comp3"], 1), "Propane");  _inp(ws.cell(r["comp3"], 2), 3)
    _lbl(ws.cell(r["comp3"], 3), "mol%");    _lbl(ws.cell(r["comp3"], 4), "합계 100%")

    # 행 높이
    for row_no in range(3, r["comp3"] + 1):
        ws.row_dimensions[row_no].height = 18

    # ── 드롭다운 유효성 검사 ──
    insul_list = (
        '"Polyurethane Foam (PUF),Perlite(대기압),Perlite(진공),'
        'MLI(다층진공단열),Glass Wool,Vacuum Panel"'
    )
    dv1 = DataValidation(type="list", formula1=insul_list, allow_blank=False)
    dv1.error = "목록에서 선택하세요"; dv1.errorTitle = "입력 오류"
    ws.add_data_validation(dv1); dv1.add(ws[f"B{r['insulation']}"])

    dv2 = DataValidation(type="list", formula1='"보냉없음,일반보냉,진공배관"',
                         allow_blank=False)
    dv2.error = "목록에서 선택하세요"; dv2.errorTitle = "입력 오류"
    ws.add_data_validation(dv2); dv2.add(ws[f"B{r['pipe_state']}"])

    dv3 = DataValidation(type="list",
                         formula1='"Methane,Ethane,Propane,Nitrogen,i-Butane,n-Butane"',
                         allow_blank=False)
    dv3.error = "목록에서 선택하세요"; dv3.errorTitle = "입력 오류"
    ws.add_data_validation(dv3); dv3.add(f"B{r['comp1']}:B{r['comp3']}")

    # ── 우측 탱크 도식 + 사양 패널 (고정 형상, 셀 연동 값) ──
    ws.merge_cells("F3:L3"); _sec(ws["F3"], "▶ 탱크 형상(고정) 및 사양 정보")

    ws.merge_cells("G5:K9")
    for row_no in range(5, 10):
        for col in ("G", "H", "I", "J", "K"):
            c = ws[f"{col}{row_no}"]
            c.fill = PatternFill("solid", fgColor="D9E1F2")
            c.border = _thin()
    for row_no in range(6, 9):
        for col in ("F", "L"):
            c = ws[f"{col}{row_no}"]
            c.fill = PatternFill("solid", fgColor="B4C6E7")
            c.border = _thin()
    ws["G5"].value = "원통형 LNG 탱크 (고정)"
    ws["G5"].font = Font(bold=True, size=10, color="1F3864")
    ws["G5"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("F11:H11"); _sec(ws["F11"], "■ 사양 패널 (셀 연동)")
    panel_rows = [
        ("탱크 내경", f"=B{r['diameter']}", "m"),
        ("탱크 길이", f"=B{r['length']}", "m"),
        ("충전율", f"=B{r['fill_ratio']}", "%"),
        ("탱크 내부 체적", f"=B{r['tank_volume']}", "m³"),
        ("LNG 체적", f"=B{r['lng_volume']}", "m³"),
        ("내부 압력", f"=B{r['pressure']}", "bar(a)"),
        ("외부 온도", f"=B{r['ambient_temp']}", "℃"),
        ("단열재 종류", f"=B{r['insulation']}", "-"),
        ("단열재 두께", f"=B{r['insulation_thickness']}", "m"),
        ("진공도", f"=B{r['vacuum_pa']}", "Pa"),
        ("배관 보냉/진공", f"=B{r['pipe_state']}", "-"),
        ("배관 열유입 비율", f"=B{r['pipe_ratio']}", "%"),
        ("성분1 이름", f"=A{r['comp1']}", "-"),
        ("성분1 몰분율", f"=B{r['comp1']}", "mol%"),
        ("성분2 이름", f"=A{r['comp2']}", "-"),
        ("성분2 몰분율", f"=B{r['comp2']}", "mol%"),
        ("성분3 이름", f"=A{r['comp3']}", "-"),
        ("성분3 몰분율", f"=B{r['comp3']}", "mol%"),
    ]
    start_row = 12
    for idx, (label, formula, unit) in enumerate(panel_rows):
        row_no = start_row + idx
        _lbl(ws.cell(row_no, 6), label)
        _calc(ws.cell(row_no, 7), formula)
        _lbl(ws.cell(row_no, 8), unit)

    ws.freeze_panes = "A3"
    return ws


# ════════════════════════════════════════════════════════════════
# 시트 2: PropertyDB (물성DB)
# ════════════════════════════════════════════════════════════════
def build_sheet2(wb):
    ws = wb.create_sheet("PropertyDB")
    _cw(ws, "A", 28); _cw(ws, "B", 16); _cw(ws, "C", 18)
    _cw(ws, "D", 16); _cw(ws, "E", 14); _cw(ws, "F", 14)

    ws.merge_cells("A1:F1")
    c = ws["A1"]
    c.value = "LNG BOR 예측 모델 — 물성 데이터베이스 (2_물성DB)"
    c.font = Font(bold=True, size=13, color=CLR_HEADER_FG)
    c.fill = PatternFill("solid", fgColor=CLR_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # ── 단열재 물성표 (A2:B9) ──
    ws.merge_cells("A2:B2"); _sec(ws["A2"], "■ 단열재 열전도율 (k값)")
    for j, t in enumerate(["단열재 종류", "열전도율 k [W/m·K]"], 1):
        _hdr(ws.cell(3, j), t, bg=CLR_DB_HDR_BG)

    insulation = [
        ("Polyurethane Foam (PUF)", 0.025),
        ("Perlite(대기압)",          0.040),
        ("Perlite(진공)",            0.0015),
        ("MLI(다층진공단열)",         0.0005),
        ("Glass Wool",               0.038),
        ("Vacuum Panel",             0.008),
    ]
    # Rows 4-9 : 단열재 데이터
    for i, (name, k) in enumerate(insulation, start=4):
        ws.cell(i, 1).value = name
        ws.cell(i, 1).border = _thin()
        ws.cell(i, 1).alignment = Alignment(vertical="center")
        ws.cell(i, 2).value = k
        ws.cell(i, 2).border = _thin()
        ws.cell(i, 2).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(i, 2).number_format = "0.0000"

    # ── LNG 성분 물성표 (A11:F18) ──
    # Row 10 : 공백
    # Row 11 : 섹션 헤더
    ws.merge_cells("A11:F11"); _sec(ws["A11"], "■ LNG 성분 물성 (비등점 기준 대표값)")
    lng_hdrs = ["성분", "분자량\n[g/mol]", "잠열 ΔHvap\n[kJ/kg]",
                "액밀도\n[kg/m³]", "비등점\n[℃]", "비고"]
    # Row 12 : 컬럼 헤더
    for j, t in enumerate(lng_hdrs, 1):
        c = ws.cell(12, j)
        _hdr(c, t, bg=CLR_DB_HDR_BG)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[12].height = 30

    lng_data = [
        ("Methane",  16.04, 511, 422.6, -161.5, "주성분"),
        ("Ethane",   30.07, 489, 544.1,  -88.6, ""),
        ("Propane",  44.10, 426, 581.0,  -42.1, ""),
        ("Nitrogen", 28.01, 199, 806.6, -195.8, ""),
        ("i-Butane", 58.12, 366, 593.4,  -11.7, ""),
        ("n-Butane", 58.12, 385, 601.8,   -0.5, ""),
    ]
    # Rows 13-18 : LNG 성분 데이터
    for i, row in enumerate(lng_data, start=13):
        for j, v in enumerate(row, 1):
            c = ws.cell(i, j)
            c.value = v
            c.border = _thin()
            c.alignment = Alignment(
                horizontal="center" if j > 1 else "left", vertical="center"
            )
        ws.row_dimensions[i].height = 16

    # ── 배관 보냉계수표 (A20:B23) ──
    ws.merge_cells("A20:B20"); _sec(ws["A20"], "■ 배관 보냉계수 (열유입 보정)")
    for j, t in enumerate(["배관 상태", "열유입 계수"], 1):
        _hdr(ws.cell(21, j), t, bg=CLR_DB_HDR_BG)
    pipe_data = [("보냉없음", 1.0), ("일반보냉", 0.4), ("진공배관", 0.1)]
    # Rows 22-24 : 배관 데이터
    for i, (st, cf) in enumerate(pipe_data, start=22):
        ws.cell(i, 1).value = st
        ws.cell(i, 1).border = _thin()
        ws.cell(i, 1).alignment = Alignment(vertical="center")
        ws.cell(i, 2).value = cf
        ws.cell(i, 2).border = _thin()
        ws.cell(i, 2).alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "A2"
    return ws


# ════════════════════════════════════════════════════════════════
# 시트 3: BOR_Calc (BOR계산)
# ════════════════════════════════════════════════════════════════
def build_sheet3(wb):
    ws = wb.create_sheet("BOR_Calc")
    _cw(ws, "A", 36); _cw(ws, "B", 22); _cw(ws, "C", 14); _cw(ws, "D", 30)

    ws.merge_cells("A1:D1")
    c = ws["A1"]
    c.value = "LNG 탱크 BOR 예측 계산 시트 (정상상태 열전달 모델)"
    c.font = Font(bold=True, size=13, color=CLR_HEADER_FG)
    c.fill = PatternFill("solid", fgColor=CLR_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    for j, t in enumerate(["계산 항목", "값", "단위", "수식/비고"], 1):
        _hdr(ws.cell(2, j), t)
    ws.row_dimensions[2].height = 20

    # 참조 시트 별칭
    I = "Input"
    D = "PropertyDB"
    R_I = INPUT_ROWS

    # ──────────────────────────────────────────────────────────────
    # 행 번호 상수 정의 (수식 참조 오류 방지)
    # ──────────────────────────────────────────────────────────────
    R_SEC_GEO  = 3
    R_RADIUS   = 4    # 탱크 반지름 r
    R_A_CYL    = 5    # 원통부 표면적
    R_A_HEMI   = 6    # 양단 반구 표면적
    R_A_TOT    = 7    # 총 표면적 A
    R_V_TANK   = 8    # 탱크 내부 체적
    R_V_LNG    = 9    # LNG 체적

    R_SEC_INS  = 10
    R_K_INS    = 11   # 단열재 k
    R_VAC_CF   = 12   # 진공보정계수
    R_K_EFF    = 13   # keff
    R_R_TH     = 14   # 열저항 R

    R_SEC_LNG  = 15
    R_HV1      = 16   # 성분1 잠열
    R_HV2      = 17   # 성분2 잠열
    R_HV3      = 18   # 성분3 잠열
    R_HV_AVG   = 19   # 평균 잠열
    R_RHO1     = 20   # 성분1 밀도
    R_RHO2     = 21   # 성분2 밀도
    R_RHO3     = 22   # 성분3 밀도
    R_RHO_AVG  = 23   # 평균 밀도
    R_TBP1     = 24   # 성분1 비등점
    R_TBP2     = 25   # 성분2 비등점
    R_TBP3     = 26   # 성분3 비등점
    R_TBP_AVG  = 27   # 평균 비등점

    R_SEC_PRES = 28
    R_P_CORR   = 29   # 압력보정계수

    R_SEC_Q    = 30
    R_DT       = 31   # 온도차 ΔT
    R_Q_TANK   = 32   # Q_tank
    R_PIPE_CF  = 33   # 배관 보냉계수
    R_Q_PIPE   = 34   # Q_pipe
    R_Q_TOT    = 35   # Q_total

    R_SEC_BOG  = 36
    R_BOG_KGS  = 37   # BOG [kg/s]
    R_BOG_DAY  = 38   # BOG 일일 [kg/day]
    R_LNG_M    = 39   # LNG 총 질량
    R_BOR      = 40   # ★ BOR [%/day]
    R_BOG_H    = 41   # ★ BOG [kg/h]
    R_BOG_DAY2 = 42   # ★ BOG [kg/day] (표시용)

    # ──────────────────────────────────────────────────────────────
    # 각 행 데이터: (행번호, 라벨, 수식, 단위, 비고, 섹션여부)
    # ──────────────────────────────────────────────────────────────
    items = [
        (R_SEC_GEO, "▶ 탱크 기하 계산", None, None, None, True),
        (R_RADIUS,  "탱크 반지름 r [m]",
         f"={I}!B{R_I['diameter']}/2", "m", "직경 / 2", False),
        (R_A_CYL,   "원통부 외면 표면적 [m²]",
         f"=2*PI()*B{R_RADIUS}*{I}!B{R_I['length']}", "m²", "2π r L", False),
        (R_A_HEMI,  "양단 반구 표면적 [m²]",
         f"=4*PI()*B{R_RADIUS}^2", "m²", "4π r²  (양쪽 반구 합)", False),
        (R_A_TOT,   "총 표면적 A [m²]",
         f"=B{R_A_CYL}+B{R_A_HEMI}", "m²", "원통 + 양단 반구", False),
        (R_V_TANK,  "탱크 내부 체적 V [m³]",
         f"=PI()*B{R_RADIUS}^2*{I}!B{R_I['length']}+(4/3)*PI()*B{R_RADIUS}^3",
         "m³", "π r² L + (4/3)π r³", False),
        (R_V_LNG,   "LNG 체적 V_LNG [m³]",
         f"=B{R_V_TANK}*{I}!B{R_I['fill_ratio']}/100", "m³", "V_tank × 충전율(%)", False),

        (R_SEC_INS, "▶ 단열 열저항 계산", None, None, None, True),
        (R_K_INS,   "단열재 k값 (VLOOKUP) [W/m·K]",
         f"=VLOOKUP({I}!B{R_I['insulation']},{D}!A4:B9,2,FALSE)",
         "W/m·K", "2_물성DB A4:B9 참조", False),
        (R_VAC_CF,  "진공보정계수 (고진공 시 keff 감소)",
         f"=IF({I}!B{R_I['vacuum_pa']}<10,0.7,IF({I}!B{R_I['vacuum_pa']}<100,0.85,1.0))",
         "-", "진공도<10Pa→0.7, <100Pa→0.85, 대기→1.0", False),
        (R_K_EFF,   "유효 열전도율 keff [W/m·K]",
         f"=B{R_K_INS}*B{R_VAC_CF}",
         "W/m·K", "k × 진공보정계수", False),
        (R_R_TH,    "열저항 R [K/W]",
         f"={I}!B{R_I['insulation_thickness']}/(B{R_K_EFF}*B{R_A_TOT})",
         "K/W", "두께 / (keff × A)", False),

        (R_SEC_LNG, "▶ LNG 조성 가중평균 물성 계산", None, None, None, True),
        (R_HV1,     "성분1 잠열 [kJ/kg]",
         f"=VLOOKUP({I}!A{R_I['comp1']},{D}!A13:E18,3,FALSE)",
         "kJ/kg", "2_물성DB A13:E18 참조", False),
        (R_HV2,     "성분2 잠열 [kJ/kg]",
         f"=VLOOKUP({I}!A{R_I['comp2']},{D}!A13:E18,3,FALSE)",
         "kJ/kg", "", False),
        (R_HV3,     "성분3 잠열 [kJ/kg]",
         f"=VLOOKUP({I}!A{R_I['comp3']},{D}!A13:E18,3,FALSE)",
         "kJ/kg", "", False),
        (R_HV_AVG,  "평균 잠열 ΔHvap [kJ/kg]",
         f"=(B{R_HV1}*{I}!B{R_I['comp1']}+B{R_HV2}*{I}!B{R_I['comp2']}+B{R_HV3}*{I}!B{R_I['comp3']})"
         f"/({I}!B{R_I['comp1']}+{I}!B{R_I['comp2']}+{I}!B{R_I['comp3']})",
         "kJ/kg", "몰분율 가중평균", False),
        (R_RHO1,    "성분1 액밀도 [kg/m³]",
         f"=VLOOKUP({I}!A{R_I['comp1']},{D}!A13:E18,4,FALSE)",
         "kg/m³", "", False),
        (R_RHO2,    "성분2 액밀도 [kg/m³]",
         f"=VLOOKUP({I}!A{R_I['comp2']},{D}!A13:E18,4,FALSE)",
         "kg/m³", "", False),
        (R_RHO3,    "성분3 액밀도 [kg/m³]",
         f"=VLOOKUP({I}!A{R_I['comp3']},{D}!A13:E18,4,FALSE)",
         "kg/m³", "", False),
        (R_RHO_AVG, "평균 밀도 ρ [kg/m³]",
         f"=(B{R_RHO1}*{I}!B{R_I['comp1']}+B{R_RHO2}*{I}!B{R_I['comp2']}+B{R_RHO3}*{I}!B{R_I['comp3']})"
         f"/({I}!B{R_I['comp1']}+{I}!B{R_I['comp2']}+{I}!B{R_I['comp3']})",
         "kg/m³", "몰분율 가중평균", False),
        (R_TBP1,    "성분1 비등점 [℃]",
         f"=VLOOKUP({I}!A{R_I['comp1']},{D}!A13:E18,5,FALSE)",
         "℃", "", False),
        (R_TBP2,    "성분2 비등점 [℃]",
         f"=VLOOKUP({I}!A{R_I['comp2']},{D}!A13:E18,5,FALSE)",
         "℃", "", False),
        (R_TBP3,    "성분3 비등점 [℃]",
         f"=VLOOKUP({I}!A{R_I['comp3']},{D}!A13:E18,5,FALSE)",
         "℃", "", False),
        (R_TBP_AVG, "평균 비등점 T_bp [℃]",
         f"=(B{R_TBP1}*{I}!B{R_I['comp1']}+B{R_TBP2}*{I}!B{R_I['comp2']}+B{R_TBP3}*{I}!B{R_I['comp3']})"
         f"/({I}!B{R_I['comp1']}+{I}!B{R_I['comp2']}+{I}!B{R_I['comp3']})",
         "℃", "몰분율 가중평균", False),

        (R_SEC_PRES,"▶ 압력 보정 계산", None, None, None, True),
        (R_P_CORR,  "압력보정계수",
         f"=MAX(0.8,1-(({I}!B{R_I['pressure']}-1.013)*0.05))",
         "-", "압력↑→비등점↑→BOG↓ (근사, 최소 0.8)", False),

        (R_SEC_Q,   "▶ 열침투량 Q 계산", None, None, None, True),
        (R_DT,      "온도차 ΔT [K]",
         f"={I}!B{R_I['ambient_temp']}-B{R_TBP_AVG}",
         "K", "외부온도 - LNG 평균비등점", False),
        (R_Q_TANK,  "탱크 열침투 Q_tank [W]",
         f"=B{R_DT}/B{R_R_TH}",
         "W", "ΔT / R", False),
        (R_PIPE_CF, "배관 보냉계수 (VLOOKUP)",
         f"=VLOOKUP({I}!B{R_I['pipe_state']},{D}!A22:B24,2,FALSE)",
         "-", "2_물성DB A22:B24 참조", False),
        (R_Q_PIPE,  "배관 열침투 Q_pipe [W]",
         f"=B{R_Q_TANK}*({I}!B{R_I['pipe_ratio']}/100)*B{R_PIPE_CF}",
         "W", "Q_tank × 배관비율 × 배관계수", False),
        (R_Q_TOT,   "총 열침투 Q_total [W]",
         f"=B{R_Q_TANK}+B{R_Q_PIPE}",
         "W", "Q_tank + Q_pipe", False),

        (R_SEC_BOG, "▶ BOG / BOR 결과", None, None, None, True),
        (R_BOG_KGS, "BOG 질량유량 [kg/s]",
         f"=B{R_Q_TOT}/(B{R_HV_AVG}*1000)",
         "kg/s", "Q_total / ΔHvap[J/kg]", False),
        (R_BOG_DAY, "BOG 일일 증발량 [kg/day]",
         f"=B{R_BOG_KGS}*86400*B{R_P_CORR}",
         "kg/day", "BOG유량 × 86400 × 압력보정", False),
        (R_LNG_M,   "LNG 총 질량 [kg]",
         f"=B{R_V_LNG}*B{R_RHO_AVG}",
         "kg", "V_LNG × 평균밀도", False),
        (R_BOR,     "★ BOR [%/day]",
         f"=B{R_BOG_DAY}/B{R_LNG_M}*100",
         "%/day", "최종 예측 BOR", False),
        (R_BOG_H,   "★ BOG [kg/h]",
         f"=B{R_BOG_DAY}/24",
         "kg/h", "시간당 증발량", False),
        (R_BOG_DAY2,"★ BOG [kg/day]",
         f"=B{R_BOG_DAY}",
         "kg/day", "일일 증발량", False),
    ]

    for row_num, label, formula, unit, note, is_sec in items:
        if is_sec:
            ws.merge_cells(f"A{row_num}:D{row_num}")
            _sec(ws[f"A{row_num}"], label)
            ws.row_dimensions[row_num].height = 18
        else:
            _lbl(ws.cell(row_num, 1), label)
            emph = label.startswith("★")
            if emph:
                _result(ws.cell(row_num, 2), formula, emph=True)
            else:
                _calc(ws.cell(row_num, 2), formula)
            _lbl(ws.cell(row_num, 3), unit)
            _lbl(ws.cell(row_num, 4), note)
            ws.row_dimensions[row_num].height = 16

    ws.freeze_panes = "A3"
    # 시트4에서 참조할 행 번호 반환
    return ws, R_BOR, R_LNG_M


# ════════════════════════════════════════════════════════════════
# 시트 4: Measure (계측보정)
# ════════════════════════════════════════════════════════════════
def build_sheet4(wb, row_bor, row_lng_mass):
    ws = wb.create_sheet("Measure")
    _cw(ws, "A", 34); _cw(ws, "B", 22); _cw(ws, "C", 14); _cw(ws, "D", 32)

    ws.merge_cells("A1:D1")
    c = ws["A1"]
    c.value = "실측 BOG → BOR 역산 및 예측 비교 (4_계측보정)"
    c.font = Font(bold=True, size=13, color=CLR_HEADER_FG)
    c.fill = PatternFill("solid", fgColor=CLR_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    for j, t in enumerate(["항목", "값", "단위", "비고"], 1):
        _hdr(ws.cell(2, j), t)
    ws.row_dimensions[2].height = 20

    I = "Input"
    C = "BOR_Calc"
    R_I = INPUT_ROWS

    # ── 실측 입력 섹션 ──
    ws.merge_cells("A3:D3"); _sec(ws["A3"], "▶ 실측 BOG 입력")
    # Row 4 : 실측 BOG량 (사용자 입력)
    _lbl(ws.cell(4, 1), "실측 BOG량 [kg]")
    _inp(ws.cell(4, 2), 0.0)
    _lbl(ws.cell(4, 3), "kg"); _lbl(ws.cell(4, 4), "← 사용자 직접 입력 (계측 동안 수집량)")
    # Row 5 : 계측 시간 (시트1 참조)
    _lbl(ws.cell(5, 1), "계측 시간 [h]")
    _calc(ws.cell(5, 2), f"={I}!B{R_I['measure_time']}")
    _lbl(ws.cell(5, 3), "h"); _lbl(ws.cell(5, 4), "1_입력조건 참조")

    # ── 보정 계산 섹션 ──
    ws.merge_cells("A6:D6"); _sec(ws["A6"], "▶ 안정화 보정 및 실측 BOR 계산")
    # Row 7 : 안정화 보정계수
    _lbl(ws.cell(7, 1), "안정화 보정계수")
    _calc(ws.cell(7, 2), f"=IF({I}!B{R_I['stabilization_time']}<6,1.15,IF({I}!B{R_I['stabilization_time']}<12,1.05,1.0))")
    _lbl(ws.cell(7, 3), "-")
    _lbl(ws.cell(7, 4), "안정화<6h→1.15 / <12h→1.05 / ≥12h→1.0")
    # Row 8 : 실측 BOG [kg/h]
    _lbl(ws.cell(8, 1), "실측 BOG [kg/h]")
    _calc(ws.cell(8, 2), "=B4/B5*B7")
    _lbl(ws.cell(8, 3), "kg/h"); _lbl(ws.cell(8, 4), "실측량 / 계측시간 × 보정계수")
    # Row 9 : 실측 BOG [kg/day]
    _lbl(ws.cell(9, 1), "실측 BOG [kg/day]")
    _calc(ws.cell(9, 2), "=B8*24")
    _lbl(ws.cell(9, 3), "kg/day"); _lbl(ws.cell(9, 4), "시간당 × 24")
    # Row 10 : LNG 총 질량 (시트3 참조)
    _lbl(ws.cell(10, 1), "LNG 총 질량 [kg]")
    _calc(ws.cell(10, 2), f"={C}!B{row_lng_mass}")
    _lbl(ws.cell(10, 3), "kg"); _lbl(ws.cell(10, 4), "3_BOR계산 참조")
    # Row 11 : ★ 실측 BOR
    _lbl(ws.cell(11, 1), "★ 실측 BOR [%/day]")
    _result(ws.cell(11, 2), "=B9/B10*100", emph=True)
    _lbl(ws.cell(11, 3), "%/day"); _lbl(ws.cell(11, 4), "실측BOG일일 / LNG총질량 × 100")

    # ── 예측 비교 섹션 ──
    ws.merge_cells("A12:D12"); _sec(ws["A12"], "▶ 예측 BOR vs 실측 BOR 비교")
    # Row 13 : ★ 예측 BOR (시트3 참조)
    _lbl(ws.cell(13, 1), "★ 예측 BOR [%/day]")
    _result(ws.cell(13, 2), f"={C}!B{row_bor}", emph=True)
    _lbl(ws.cell(13, 3), "%/day"); _lbl(ws.cell(13, 4), "3_BOR계산 최종 결과")
    # Row 14 : 오차율
    _lbl(ws.cell(14, 1), "오차율 [%]")
    _result(ws.cell(14, 2), "=(B11-B13)/B13*100")
    _lbl(ws.cell(14, 3), "%"); _lbl(ws.cell(14, 4), "(실측 - 예측) / 예측 × 100")
    # Row 15 : 판정
    _lbl(ws.cell(15, 1), "판정")
    _result(ws.cell(15, 2), '=IF(ABS(B14)<10,"✅ 양호","⚠ 재계측 권장")')
    _lbl(ws.cell(15, 3), "-"); _lbl(ws.cell(15, 4), "오차 10% 미만이면 양호")

    # 행 높이
    for r in range(3, 16):
        ws.row_dimensions[r].height = 18

    # 주의 박스
    ws.merge_cells("A17:D19")
    nc = ws["A17"]
    nc.value = (
        "📌 주의: 이 모델은 정상상태(steady-state) 열전달 기반의 설계 예측용 모델입니다.\n"
        "실측값과의 오차가 10% 이상이면 단열재 물성, 충전율, 탱크 형상 등 입력값을 재확인하거나\n"
        "일사량, 지면 열전도 등 현장 조건을 추가 반영하세요."
    )
    nc.font = Font(size=9, italic=True, color="595959")
    nc.alignment = Alignment(wrap_text=True, vertical="top")
    nc.fill = PatternFill("solid", fgColor="FFF9E6")
    nc.border = _thin()
    for r in range(17, 20):
        ws.row_dimensions[r].height = 14

    ws.freeze_panes = "A3"
    return ws


# ════════════════════════════════════════════════════════════════
# 메인 실행
# ════════════════════════════════════════════════════════════════
def main():
    wb = Workbook()
    del wb[wb.sheetnames[0]]   # 기본 빈 시트 제거

    print("📄 시트 생성 중...")
    build_sheet1(wb)
    print("  ✔ Input 완료")
    build_sheet2(wb)
    print("  ✔ PropertyDB 완료")
    _, row_bor, row_lng_mass = build_sheet3(wb)
    print("  ✔ BOR_Calc 완료")
    build_sheet4(wb, row_bor, row_lng_mass)
    print("  ✔ Measure 완료")

    output = "LNG_BOR_Model.xlsx"
    wb.save(output)
    print(f"\n✅ 저장 완료: {output}")
    print("   → 엑셀을 열고 Input 시트의 값을 변경하면 BOR이 자동 계산됩니다.")


if __name__ == "__main__":
    main()
