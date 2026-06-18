"""
LNG BOR 예측 엑셀 모델 검증 스크립트
--------------------------------------
생성된 LNG_BOR_Model.xlsx의 구조, 수식 보존, 드롭다운 유효성을 검사합니다.

실행 방법:
    python validate_model.py

검증 통과 시 "✅ 모든 검증 통과" 출력 후 exit code 0.
실패 시 오류 메시지 출력 후 exit code 1.
"""

import sys
import zipfile
import xml.etree.ElementTree as ET
from openpyxl import load_workbook

MODEL_PATH = "LNG_BOR_Model.xlsx"

# 워크시트 탭 이름 (영문 ASCII)
REQUIRED_SHEETS = ["Input", "PropertyDB", "BOR_Calc", "Measure"]

# BOR_Calc 시트의 결과 라벨 (수식 보존 확인용)
RESULT_LABELS = ["★ BOR [%/day]", "★ BOG [kg/h]"]


def find_label_row(ws, label):
    """워크시트에서 특정 라벨 텍스트가 있는 행 번호를 반환. 없으면 None."""
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == label:
                return cell.row
    return None


def validate_sheets(wb):
    """필수 시트 존재 여부 확인."""
    missing = [s for s in REQUIRED_SHEETS if s not in wb.sheetnames]
    if missing:
        print(f"❌ 누락된 시트: {missing}")
        return False
    print(f"  ✔ 시트 구성 확인: {REQUIRED_SHEETS}")
    return True


def validate_formulas_openpyxl(wb):
    """openpyxl로 BOR_Calc 시트의 결과 셀에 수식이 보존되는지 확인."""
    ws = wb["BOR_Calc"]
    errors = []
    for label in RESULT_LABELS:
        row = find_label_row(ws, label)
        if row is None:
            errors.append(f"라벨 '{label}'을 BOR_Calc 시트에서 찾을 수 없습니다.")
            continue
        cell = ws.cell(row, 2)
        val = cell.value
        if not (isinstance(val, str) and val.startswith("=")):
            errors.append(
                f"'{label}' (B{row}) 수식이 보존되지 않았습니다. 현재 값: {repr(val)}"
            )
    if errors:
        for e in errors:
            print(f"❌ {e}")
        return False
    print(f"  ✔ BOR_Calc 수식 보존 확인 (openpyxl): {RESULT_LABELS}")
    return True


def validate_formulas_xml(path):
    """xlsx ZIP 내부의 BOR_Calc 시트 XML에 <f> 태그(수식)가 존재하는지 확인."""
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    with zipfile.ZipFile(path, "r") as zf:
        # workbook.xml에서 시트 이름과 파일 경로 매핑
        wb_xml = zf.read("xl/workbook.xml")
        wb_root = ET.fromstring(wb_xml)
        sheet_map = {}
        for sheet in wb_root.iter(f"{{{ns}}}sheet"):
            name = sheet.get("name")
            rid = sheet.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            sheet_map[name] = rid

        if "BOR_Calc" not in sheet_map:
            print("❌ BOR_Calc 시트를 workbook.xml에서 찾을 수 없습니다.")
            return False

        # workbook.xml.rels에서 rid → 실제 파일 경로 매핑
        rels_xml = zf.read("xl/_rels/workbook.xml.rels")
        rels_root = ET.fromstring(rels_xml)
        rid_to_path = {}
        for rel in rels_root:
            rid_to_path[rel.get("Id")] = rel.get("Target")

        bor_calc_rid = sheet_map["BOR_Calc"]
        bor_calc_target = rid_to_path.get(bor_calc_rid)
        if bor_calc_target is None:
            print(f"❌ BOR_Calc 시트의 파일 경로를 찾을 수 없습니다. (rid={bor_calc_rid})")
            return False

        # 절대 경로(/로 시작)는 앞의 /를 제거, 상대 경로는 xl/ 접두사 추가
        if bor_calc_target.startswith("/"):
            bor_calc_target = bor_calc_target.lstrip("/")
        elif not bor_calc_target.startswith("xl/"):
            bor_calc_target = "xl/" + bor_calc_target

        sheet_xml = zf.read(bor_calc_target)
        sheet_root = ET.fromstring(sheet_xml)

        # <f> 태그 개수 세기
        formula_tags = list(sheet_root.iter(f"{{{ns}}}f"))
        if not formula_tags:
            print(f"❌ BOR_Calc 시트 XML ({bor_calc_target})에 <f> 수식 태그가 없습니다.")
            return False

        print(f"  ✔ BOR_Calc XML 수식 태그 확인: <f> 태그 {len(formula_tags)}개 존재")
        return True


def validate_dropdown(wb):
    """Input 시트에 드롭다운(데이터 유효성) 검사가 존재하는지 확인."""
    ws = wb["Input"]
    dvs = list(ws.data_validations.dataValidation)
    if not dvs:
        print("❌ Input 시트에 드롭다운(데이터 유효성) 검사가 없습니다.")
        return False
    print(f"  ✔ Input 드롭다운 유효성 검사 확인: {len(dvs)}개 존재")
    return True


def main():
    print(f"🔍 검증 시작: {MODEL_PATH}\n")

    try:
        wb = load_workbook(MODEL_PATH)
    except FileNotFoundError:
        print(f"❌ 파일을 찾을 수 없습니다: {MODEL_PATH}")
        print("   먼저 python generate_bor_model.py 를 실행하세요.")
        sys.exit(1)

    results = []

    print("[1] 시트 구성 확인")
    results.append(validate_sheets(wb))

    print("[2] BOR_Calc 수식 보존 확인 (openpyxl)")
    results.append(validate_formulas_openpyxl(wb))

    print("[3] BOR_Calc 수식 보존 확인 (XML <f> 태그)")
    results.append(validate_formulas_xml(MODEL_PATH))

    print("[4] Input 시트 드롭다운 유효성 확인")
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
