"""
LNG BOR 예측 엑셀 모델 검증 스크립트
--------------------------------------
generate_bor_model.py로 생성된 LNG_BOR_Model.xlsx의 구조, 수식, 수치,
드롭다운, XML 수식 보존을 검사합니다.

실행 방법:
    python validate_model.py

검증 통과 시 "✅ 모든 검증 통과" 출력 후 exit code 0.
실패 시 오류 메시지 출력 후 exit code 1.
"""

import math
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from openpyxl import load_workbook

from generate_bor_model import main as generate_workbook

MODEL_PATH = Path("LNG_BOR_Model.xlsx")

# 워크시트 탭 이름 (영문 ASCII)
REQUIRED_SHEETS = ["Input", "PropertyDB", "BOR_Calc", "Measure"]

EXPECTED_LOOKUP_FORMULAS = {
    "B16": "=VLOOKUP(Input!A22,PropertyDB!A13:E18,3,FALSE)",
    "B17": "=VLOOKUP(Input!A23,PropertyDB!A13:E18,3,FALSE)",
    "B18": "=VLOOKUP(Input!A24,PropertyDB!A13:E18,3,FALSE)",
    "B20": "=VLOOKUP(Input!A22,PropertyDB!A13:E18,4,FALSE)",
    "B21": "=VLOOKUP(Input!A23,PropertyDB!A13:E18,4,FALSE)",
    "B22": "=VLOOKUP(Input!A24,PropertyDB!A13:E18,4,FALSE)",
    "B24": "=VLOOKUP(Input!A22,PropertyDB!A13:E18,5,FALSE)",
    "B25": "=VLOOKUP(Input!A23,PropertyDB!A13:E18,5,FALSE)",
    "B26": "=VLOOKUP(Input!A24,PropertyDB!A13:E18,5,FALSE)",
}

EXPECTED_WEIGHTED_FORMULAS = {
    "B19": "=(B16*Input!B22+B17*Input!B23+B18*Input!B24)/(Input!B22+Input!B23+Input!B24)",
    "B23": "=(B20*Input!B22+B21*Input!B23+B22*Input!B24)/(Input!B22+Input!B23+Input!B24)",
    "B27": "=(B24*Input!B22+B25*Input!B23+B26*Input!B24)/(Input!B22+Input!B23+Input!B24)",
}
# 기본 입력값(직경 4 m, 길이 10 m, 충전율 90%, Perlite(진공), 외기 25 ℃,
# Methane 90 / Ethane 7 / Propane 3)으로 계산한 BOR는 약 0.03 %/day여야 한다.
MIN_EXPECTED_BOR = 0.01
MAX_EXPECTED_BOR = 0.10
HIGH_VACUUM_THRESHOLD_PA = 10
MEDIUM_VACUUM_THRESHOLD_PA = 100
HIGH_VACUUM_FACTOR = 0.7
MEDIUM_VACUUM_FACTOR = 0.85
ATMOSPHERIC_VACUUM_FACTOR = 1.0
REFERENCE_PRESSURE_BAR = 1.013
PRESSURE_CORRECTION_SLOPE = 0.05
MIN_PRESSURE_CORRECTION = 0.8

# OOXML 네임스페이스
NS_SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def validate_sheets(wb):
    """필수 시트 존재 여부 확인."""
    missing = [s for s in REQUIRED_SHEETS if s not in wb.sheetnames]
    if missing:
        print(f"❌ 누락된 시트: {missing}")
        return False
    print(f"  ✔ 시트 구성 확인: {REQUIRED_SHEETS}")
    return True


def validate_formula_references(wb):
    """BOR_Calc LNG 조성 계산 수식이 올바른 열을 참조하는지 확인."""
    ws = wb["BOR_Calc"]
    errors = []

    for cell_ref, expected in EXPECTED_LOOKUP_FORMULAS.items():
        actual = ws[cell_ref].value
        if actual != expected:
            errors.append(f"{cell_ref} lookup 수식 오류: {actual!r} != {expected!r}")

    for cell_ref, expected in EXPECTED_WEIGHTED_FORMULAS.items():
        actual = ws[cell_ref].value
        if actual != expected:
            errors.append(f"{cell_ref} 가중평균 수식 오류: {actual!r} != {expected!r}")

    if errors:
        for e in errors:
            print(f"❌ {e}")
        return False

    print("  ✔ BOR_Calc LNG 조성 수식 열 참조 확인")
    return True


def validate_formulas_xml(path):
    """xlsx ZIP 내부 sheet XML들에 <f> 태그가 존재하는지 확인."""
    total_formula_tags = 0
    worksheet_files = []

    with zipfile.ZipFile(path, "r") as zf:
        for name in zf.namelist():
            if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
                worksheet_files.append(name)
                sheet_root = ET.fromstring(zf.read(name))
                total_formula_tags += len(
                    list(sheet_root.iter(f"{{{NS_SPREADSHEET}}}f"))
                )

    if not worksheet_files:
        print("❌ xlsx 내부에서 xl/worksheets/sheet*.xml 파일을 찾을 수 없습니다.")
        return False
    if total_formula_tags == 0:
        print("❌ xlsx 내부 시트 XML들에 <f> 수식 태그가 없습니다.")
        return False

    print(f"  ✔ XML 수식 태그 확인: {len(worksheet_files)}개 시트 XML, <f> 태그 {total_formula_tags}개")
    return True


def validate_dropdown(wb):
    """Input 시트에 드롭다운(데이터 유효성) 검사가 존재하는지 확인."""
    ws = wb["Input"]
    data_validation_rules = list(ws.data_validations.dataValidation)
    if not data_validation_rules:
        print("❌ Input 시트에 드롭다운(데이터 유효성) 검사가 없습니다.")
        return False
    ranges = {str(rule.sqref) for rule in data_validation_rules}
    expected_ranges = {"B11", "B15", "B22:B24"}
    if not expected_ranges.issubset(ranges):
        print(f"❌ Input 드롭다운 범위가 예상과 다릅니다. 현재: {sorted(ranges)}")
        return False
    print(f"  ✔ Input 드롭다운 유효성 검사 확인: {len(data_validation_rules)}개 존재")
    return True


def validate_bor_range(wb):
    """기본 입력값으로 독립 계산한 BOR 값이 예상 범위인지 확인."""
    input_ws = wb["Input"]
    prop_ws = wb["PropertyDB"]

    insulation = {
        prop_ws.cell(row, 1).value: float(prop_ws.cell(row, 2).value)
        for row in range(4, 10)
    }
    components = {
        prop_ws.cell(row, 1).value: {
            "hvap": float(prop_ws.cell(row, 3).value),
            "rho": float(prop_ws.cell(row, 4).value),
            "tbp": float(prop_ws.cell(row, 5).value),
        }
        for row in range(13, 19)
    }
    pipe_factors = {
        prop_ws.cell(row, 1).value: float(prop_ws.cell(row, 2).value)
        for row in range(22, 25)
    }

    diameter = float(input_ws["B4"].value)
    length = float(input_ws["B5"].value)
    fill_ratio = float(input_ws["B6"].value)
    pressure = float(input_ws["B8"].value)
    ambient_temp = float(input_ws["B9"].value)
    insulation_name = input_ws["B11"].value
    insulation_thickness = float(input_ws["B12"].value)
    vacuum_pa = float(input_ws["B13"].value)
    pipe_state = input_ws["B15"].value
    pipe_heat_ratio = float(input_ws["B16"].value)

    composition = [
        (input_ws["A22"].value, float(input_ws["B22"].value)),
        (input_ws["A23"].value, float(input_ws["B23"].value)),
        (input_ws["A24"].value, float(input_ws["B24"].value)),
    ]
    total_fraction = sum(fraction for _, fraction in composition)

    avg_hvap = sum(components[name]["hvap"] * fraction for name, fraction in composition) / total_fraction
    avg_rho = sum(components[name]["rho"] * fraction for name, fraction in composition) / total_fraction
    avg_tbp = sum(components[name]["tbp"] * fraction for name, fraction in composition) / total_fraction

    radius = diameter / 2
    area_total = (2 * math.pi * radius * length) + (4 * math.pi * radius**2)
    volume_tank = (math.pi * radius**2 * length) + ((4 / 3) * math.pi * radius**3)
    volume_lng = volume_tank * fill_ratio / 100

    vacuum_factor = (
        HIGH_VACUUM_FACTOR
        if vacuum_pa < HIGH_VACUUM_THRESHOLD_PA
        else MEDIUM_VACUUM_FACTOR
        if vacuum_pa < MEDIUM_VACUUM_THRESHOLD_PA
        else ATMOSPHERIC_VACUUM_FACTOR
    )
    k_eff = insulation[insulation_name] * vacuum_factor
    thermal_resistance = insulation_thickness / (k_eff * area_total)

    pressure_correction = max(
        MIN_PRESSURE_CORRECTION,
        1 - ((pressure - REFERENCE_PRESSURE_BAR) * PRESSURE_CORRECTION_SLOPE),
    )
    delta_t = ambient_temp - avg_tbp
    q_tank = delta_t / thermal_resistance
    q_pipe = q_tank * (pipe_heat_ratio / 100) * pipe_factors[pipe_state]
    q_total = q_tank + q_pipe

    bog_kg_s = q_total / (avg_hvap * 1000)
    bog_kg_day = bog_kg_s * 86400 * pressure_correction
    lng_mass = volume_lng * avg_rho
    bor_percent_day = bog_kg_day / lng_mass * 100

    if not (MIN_EXPECTED_BOR <= bor_percent_day <= MAX_EXPECTED_BOR):
        print(f"❌ 독립 계산 BOR 값이 예상 범위를 벗어났습니다: {bor_percent_day:.5f} %/day")
        return False

    print(f"  ✔ 독립 계산 BOR 확인: {bor_percent_day:.5f} %/day")
    return True


def main():
    print(f"🔍 검증 시작: {MODEL_PATH}\n")

    before_stat = MODEL_PATH.stat() if MODEL_PATH.exists() else None

    print("[0] 엑셀 모델 재생성")
    generate_workbook()
    if not MODEL_PATH.exists():
        print(f"❌ generate_bor_model.main() 실행 후 파일이 생성되지 않았습니다: {MODEL_PATH}")
        sys.exit(1)
    if before_stat is not None:
        after_stat = MODEL_PATH.stat()
        if (
            after_stat.st_mtime_ns == before_stat.st_mtime_ns
            and after_stat.st_size == before_stat.st_size
        ):
            print(f"❌ generate_bor_model.main() 실행 후 파일이 갱신되지 않았습니다: {MODEL_PATH}")
            sys.exit(1)
    print(f"  ✔ 파일 생성 확인: {MODEL_PATH}")

    try:
        wb = load_workbook(MODEL_PATH)
    except FileNotFoundError:
        print(f"❌ 파일을 찾을 수 없습니다: {MODEL_PATH}")
        print("   먼저 python generate_bor_model.py 를 실행하세요.")
        sys.exit(1)

    results = []

    print("[1] 시트 구성 확인")
    results.append(validate_sheets(wb))

    print("[2] BOR_Calc 수식 열 참조 확인")
    results.append(validate_formula_references(wb))

    print("[3] 시트 XML 수식 태그 확인")
    results.append(validate_formulas_xml(MODEL_PATH))

    print("[4] 기본 입력값 독립 계산 BOR 확인")
    results.append(validate_bor_range(wb))

    print("[5] Input 시트 드롭다운 유효성 확인")
    results.append(validate_dropdown(wb))

    print()
    if all(results):
        print("✅ 모든 검증 통과")
        sys.exit(0)
    else:
        failed = sum(1 for r in results if not r)
        print(f"❌ 검증 실패: {failed}개 항목 오류")
        sys.exit(1)


if __name__ == "__main__":
    main()
