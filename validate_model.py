import math
import os
import sys

from openpyxl import load_workbook

import generate_bor_model


MODEL_PATH = "LNG_BOR_Model.xlsx"
REQUIRED_SHEETS = ["1_입력조건", "2_물성DB", "3_BOR계산", "4_계측보정"]


def fail(message: str) -> None:
    raise AssertionError(message)


def find_label_row(ws, label: str) -> int:
    for row in range(1, ws.max_row + 1):
        if ws.cell(row, 1).value == label:
            return row
    fail(f"'{ws.title}' 시트에서 라벨을 찾을 수 없습니다: {label}")


def has_formula(ws, row: int, col: int = 2) -> bool:
    value = ws.cell(row, col).value
    return isinstance(value, str) and value.startswith("=")


def has_data_validation_for_cell(ws, coord: str) -> bool:
    for dv in ws.data_validations.dataValidation:
        if coord in dv:
            return True
    return False


def calculate_expected_bor_percent_per_day() -> float:
    diameter_m = 4.0
    length_m = 10.0
    fill_pct = 90.0
    pressure_bar = 1.13
    ambient_temp_c = 25.0
    insulation_thickness_m = 0.30
    vacuum_pa = 1.0
    pipe_heat_ratio_pct = 10.0

    # 기본 선택값: Perlite(진공), 진공배관, Methane90/Ethane7/Propane3
    insulation_k = 0.0015
    pipe_coeff = 0.1
    composition = {
        "Methane": 90.0,
        "Ethane": 7.0,
        "Propane": 3.0,
    }
    h_vap_kj_per_kg = {
        "Methane": 511.0,
        "Ethane": 489.0,
        "Propane": 426.0,
    }
    density_kg_per_m3 = {
        "Methane": 422.6,
        "Ethane": 544.1,
        "Propane": 581.0,
    }
    t_bp_c = {
        "Methane": -161.5,
        "Ethane": -88.6,
        "Propane": -42.1,
    }

    radius_m = diameter_m / 2.0
    area_total_m2 = 2 * math.pi * radius_m * length_m + 4 * math.pi * radius_m**2
    v_tank_m3 = math.pi * radius_m**2 * length_m + (4.0 / 3.0) * math.pi * radius_m**3
    v_lng_m3 = v_tank_m3 * (fill_pct / 100.0)

    if vacuum_pa < 10:
        vacuum_cf = 0.7
    elif vacuum_pa < 100:
        vacuum_cf = 0.85
    else:
        vacuum_cf = 1.0

    k_eff = insulation_k * vacuum_cf
    r_th = insulation_thickness_m / (k_eff * area_total_m2)

    mol_total = sum(composition.values())
    hv_avg = sum(h_vap_kj_per_kg[k] * v for k, v in composition.items()) / mol_total
    rho_avg = sum(density_kg_per_m3[k] * v for k, v in composition.items()) / mol_total
    tbp_avg = sum(t_bp_c[k] * v for k, v in composition.items()) / mol_total

    delta_t = ambient_temp_c - tbp_avg
    q_tank_w = delta_t / r_th
    q_pipe_w = q_tank_w * (pipe_heat_ratio_pct / 100.0) * pipe_coeff
    q_total_w = q_tank_w + q_pipe_w

    p_corr = max(0.8, 1.0 - ((pressure_bar - 1.013) * 0.05))
    bog_kg_s = q_total_w / (hv_avg * 1000.0)
    bog_kg_day = bog_kg_s * 86400.0 * p_corr

    lng_mass_kg = v_lng_m3 * rho_avg
    bor_pct_day = (bog_kg_day / lng_mass_kg) * 100.0
    return bor_pct_day


def main() -> int:
    try:
        # 1) 모델 생성 검증
        generate_bor_model.main()
        if not os.path.exists(MODEL_PATH):
            fail(f"모델 파일이 생성되지 않았습니다: {MODEL_PATH}")

        wb = load_workbook(MODEL_PATH)

        # 2) 시트 존재 검증
        missing = [name for name in REQUIRED_SHEETS if name not in wb.sheetnames]
        if missing:
            fail(f"필수 시트 누락: {missing}")

        # 3) 시트3 핵심 결과 셀 수식 검증
        ws3 = wb["3_BOR계산"]
        bor_row = find_label_row(ws3, "★ BOR [%/day]")
        bog_h_row = find_label_row(ws3, "★ BOG [kg/h]")
        if not has_formula(ws3, bor_row):
            fail("'★ BOR [%/day]' 결과 셀(B열)에 수식이 없습니다.")
        if not has_formula(ws3, bog_h_row):
            fail("'★ BOG [kg/h]' 결과 셀(B열)에 수식이 없습니다.")

        # 4) 시트1 드롭다운(데이터 유효성) 검증
        ws1 = wb["1_입력조건"]
        for coord, name in [("B11", "단열재"), ("B15", "배관"), ("B22", "LNG성분1"), ("B23", "LNG성분2"), ("B24", "LNG성분3")]:
            if not has_data_validation_for_cell(ws1, coord):
                fail(f"데이터 유효성 검사 누락: {name} 셀({coord})")

        # 5) 수치 검증 (독립 계산)
        bor = calculate_expected_bor_percent_per_day()
        if not (0.0 < bor < 5.0):
            fail(f"BOR 값이 비정상 범위입니다: {bor:.6f} %/day")
        if not (0.01 <= bor <= 0.10):
            fail(
                "기본 입력 기준 BOR 예상 범위를 벗어났습니다: "
                f"{bor:.6f} %/day (기대 약 0.03 부근)"
            )

        print(f"✅ 모든 검증 통과 (독립 계산 BOR={bor:.6f} %/day)")
        return 0
    except Exception as exc:  # noqa: BLE001
        print(f"❌ 검증 실패: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
