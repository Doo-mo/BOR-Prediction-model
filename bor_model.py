"""
BOR (Bed Occupancy Rate) 예측 모델
병상이용률 = (재원일수 / (병상수 × 해당 월 일수)) × 100
"""

import calendar
import os

import numpy as np
import pandas as pd
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from sklearn.model_selection import train_test_split

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    GradientFill,
    PatternFill,
    Side,
)
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────
# 1. 데이터 로드 및 전처리
# ─────────────────────────────────────────────

def load_data(filepath: str) -> pd.DataFrame:
    """CSV 파일에서 병원 데이터를 로드하고 BOR을 계산합니다."""
    df = pd.read_csv(filepath)

    # 날짜 파싱
    df["연월"] = pd.to_datetime(df["연월"], format="%Y-%m")
    df["연도"] = df["연월"].dt.year
    df["월"] = df["연월"].dt.month

    # 해당 월의 일수 계산
    df["월일수"] = df.apply(
        lambda r: calendar.monthrange(r["연도"], r["월"])[1], axis=1
    )

    # BOR 계산 (%)
    df["BOR"] = (df["재원일수"] / (df["병상수"] * df["월일수"])) * 100
    df["BOR"] = df["BOR"].round(2)

    return df


def add_features(df: pd.DataFrame) -> pd.DataFrame:
    """예측에 사용할 피처를 추가합니다."""
    df = df.copy()

    # 계절 (1=봄, 2=여름, 3=가을, 4=겨울)
    df["계절"] = df["월"].apply(
        lambda m: 1 if m in [3, 4, 5] else
                  2 if m in [6, 7, 8] else
                  3 if m in [9, 10, 11] else 4
    )

    # 시간 인덱스 (트렌드 반영)
    df["시간인덱스"] = range(len(df))

    # 월별 사인/코사인 (계절성 반영)
    df["월_sin"] = np.sin(2 * np.pi * df["월"] / 12)
    df["월_cos"] = np.cos(2 * np.pi * df["월"] / 12)

    return df


# ─────────────────────────────────────────────
# 2. 모델 학습 및 예측
# ─────────────────────────────────────────────

FEATURES = ["시간인덱스", "병상수", "월_sin", "월_cos", "입원환자수"]


def train_model(df: pd.DataFrame):
    """선형 회귀 모델을 학습합니다."""
    X = df[FEATURES]
    y = df["BOR"]

    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.2, random_state=42, shuffle=False
    )

    model = LinearRegression()
    model.fit(X_train, y_train)

    y_pred = model.predict(X_test)

    metrics = {
        "R²": round(r2_score(y_test, y_pred), 4),
        "MAE": round(mean_absolute_error(y_test, y_pred), 4),
        "RMSE": round(np.sqrt(mean_squared_error(y_test, y_pred)), 4),
    }

    return model, X_train, X_test, y_train, y_test, y_pred, metrics


def predict_future(
    model: LinearRegression,
    df: pd.DataFrame,
    months: int = 12,
) -> pd.DataFrame:
    """향후 N개월치 BOR을 예측합니다.

    Note:
        미래 병상수와 입원환자수는 마지막 실적 월의 값을 그대로 사용합니다.
        실제 병상 증감이나 환자 수 변화가 예상되면 CSV 데이터를 업데이트하세요.
    """
    last_row = df.iloc[-1]
    last_time_idx = int(last_row["시간인덱스"])
    last_beds = int(last_row["병상수"])
    last_patients = int(last_row["입원환자수"])

    last_date = last_row["연월"]

    future_rows = []
    for i in range(1, months + 1):
        future_date = last_date + pd.DateOffset(months=i)
        month = future_date.month
        future_rows.append({
            "연월": future_date,
            "연도": future_date.year,
            "월": month,
            "시간인덱스": last_time_idx + i,
            "병상수": last_beds,
            "입원환자수": last_patients,
            "월_sin": np.sin(2 * np.pi * month / 12),
            "월_cos": np.cos(2 * np.pi * month / 12),
        })

    future_df = pd.DataFrame(future_rows)
    future_df["예측BOR"] = model.predict(future_df[FEATURES]).round(2)

    clipped = future_df["예측BOR"].gt(100).sum() + future_df["예측BOR"].lt(0).sum()
    if clipped:
        print(f"  ⚠️  예측값 {clipped}건이 0~100% 범위를 벗어나 보정되었습니다.")

    future_df["예측BOR"] = future_df["예측BOR"].clip(0, 100)

    return future_df


# ─────────────────────────────────────────────
# 3. Excel 출력
# ─────────────────────────────────────────────

# 색상 팔레트
COLOR_HEADER_BLUE = "1F497D"
COLOR_HEADER_LIGHT = "DCE6F1"
COLOR_ACCENT = "4472C4"
COLOR_PRED = "ED7D31"
COLOR_GOOD = "C6EFCE"
COLOR_WARN = "FFEB9C"
COLOR_BAD = "FFC7CE"

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _header_style(cell, bg: str = COLOR_HEADER_BLUE, font_color: str = "FFFFFF"):
    cell.font = Font(bold=True, color=font_color, size=11)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER


def _data_style(cell, bg: str = "FFFFFF"):
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = BORDER


def _bor_fill(bor_value: float) -> str:
    """BOR 수준에 따른 배경색 반환"""
    if bor_value >= 90:
        return COLOR_BAD
    if bor_value >= 75:
        return COLOR_GOOD
    return COLOR_WARN


def write_data_sheet(ws, df: pd.DataFrame):
    """실적 데이터 시트 작성"""
    ws.title = "실적데이터"

    headers = ["연월", "병상수", "재원일수", "월일수", "입원환자수", "외래환자수", "BOR(%)"]
    col_widths = [14, 10, 12, 10, 12, 12, 12]

    # 제목
    ws.merge_cells("A1:G1")
    title_cell = ws["A1"]
    title_cell.value = "병상이용률(BOR) 실적 데이터"
    title_cell.font = Font(bold=True, size=14, color=COLOR_HEADER_BLUE)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # 헤더
    for col_idx, (header, width) in enumerate(zip(headers, col_widths), start=1):
        cell = ws.cell(row=2, column=col_idx, value=header)
        _header_style(cell)
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[2].height = 22

    # 데이터
    for row_idx, (_, row) in enumerate(df.iterrows(), start=3):
        values = [
            row["연월"].strftime("%Y-%m"),
            int(row["병상수"]),
            int(row["재원일수"]),
            int(row["월일수"]),
            int(row["입원환자수"]),
            int(row.get("외래환자수", 0)),
            row["BOR"],
        ]
        bg = _bor_fill(row["BOR"])
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            _data_style(cell, bg if col_idx == 7 else "FFFFFF")

    # BOR 수치 형식
    for row_idx in range(3, 3 + len(df)):
        ws.cell(row=row_idx, column=7).number_format = "0.00"

    ws.freeze_panes = "A3"


def write_model_sheet(ws, metrics: dict, X_test, y_test, y_pred):
    """모델 성능 시트 작성"""
    ws.title = "모델성능"

    ws.merge_cells("A1:D1")
    cell = ws["A1"]
    cell.value = "선형회귀 모델 성능 평가"
    cell.font = Font(bold=True, size=14, color=COLOR_HEADER_BLUE)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # 성능 지표
    ws["A3"] = "평가 지표"
    ws["B3"] = "수치"
    ws["C3"] = "설명"
    for col in ["A", "B", "C"]:
        _header_style(ws[f"{col}3"])
    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 35

    descriptions = {
        "R²": "결정계수 (1에 가까울수록 우수)",
        "MAE": "평균절대오차 (%p)",
        "RMSE": "평균제곱근오차 (%p)",
    }
    for i, (key, val) in enumerate(metrics.items(), start=4):
        ws.cell(row=i, column=1, value=key)
        ws.cell(row=i, column=2, value=val)
        ws.cell(row=i, column=3, value=descriptions[key])
        for col in range(1, 4):
            _data_style(ws.cell(row=i, column=col))
        ws.cell(row=i, column=2).number_format = "0.0000"

    # 실제 vs 예측 비교표
    ws["A8"] = "실제 BOR vs 예측 BOR (테스트셋)"
    ws["A8"].font = Font(bold=True, size=12, color=COLOR_HEADER_BLUE)
    ws.row_dimensions[8].height = 22

    ws["A9"] = "No."
    ws["B9"] = "실제 BOR (%)"
    ws["C9"] = "예측 BOR (%)"
    ws["D9"] = "오차 (%p)"
    for col in ["A", "B", "C", "D"]:
        _header_style(ws[f"{col}9"])

    for i, (actual, pred) in enumerate(zip(y_test, y_pred), start=1):
        row = 9 + i
        ws.cell(row=row, column=1, value=i)
        ws.cell(row=row, column=2, value=round(float(actual), 2))
        ws.cell(row=row, column=3, value=round(float(pred), 2))
        ws.cell(row=row, column=4, value=round(float(pred - actual), 2))
        for col in range(1, 5):
            _data_style(ws.cell(row=row, column=col))
        for col in [2, 3, 4]:
            ws.cell(row=row, column=col).number_format = "0.00"

    ws.column_dimensions["D"].width = 14

    # 비교 차트
    n_test = len(y_test)
    chart = LineChart()
    chart.title = "실제 BOR vs 예측 BOR"
    chart.style = 10
    chart.y_axis.title = "BOR (%)"
    chart.x_axis.title = "테스트 데이터 No."
    chart.height = 12
    chart.width = 20

    actual_ref = Reference(ws, min_col=2, min_row=10, max_row=9 + n_test)
    pred_ref = Reference(ws, min_col=3, min_row=10, max_row=9 + n_test)

    from openpyxl.chart import Series
    s1 = Series(actual_ref, title="실제 BOR")
    s2 = Series(pred_ref, title="예측 BOR")
    chart.series.append(s1)
    chart.series.append(s2)

    ws.add_chart(chart, "F9")


def write_prediction_sheet(ws, df: pd.DataFrame, future_df: pd.DataFrame):
    """예측 결과 시트 작성"""
    ws.title = "BOR예측"

    ws.merge_cells("A1:E1")
    cell = ws["A1"]
    cell.value = "향후 12개월 BOR 예측"
    cell.font = Font(bold=True, size=14, color=COLOR_HEADER_BLUE)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    headers = ["연월", "병상수", "예측 BOR (%)", "BOR 등급", "비고"]
    col_widths = [14, 10, 15, 12, 20]

    for col_idx, (header, width) in enumerate(zip(headers, col_widths), start=1):
        cell = ws.cell(row=2, column=col_idx, value=header)
        _header_style(cell)
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[2].height = 22

    def bor_grade(bor: float) -> tuple[str, str]:
        if bor >= 90:
            return "과밀", COLOR_BAD
        if bor >= 75:
            return "적정", COLOR_GOOD
        return "여유", COLOR_WARN

    for row_idx, (_, row) in enumerate(future_df.iterrows(), start=3):
        bor = row["예측BOR"]
        grade, bg = bor_grade(bor)
        note = "병상 과부하 위험" if bor >= 90 else ("정상 운영 범위" if bor >= 75 else "병상 활용 제고 필요")

        values = [
            row["연월"].strftime("%Y-%m"),
            int(row["병상수"]),
            bor,
            grade,
            note,
        ]
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            _data_style(cell, bg if col_idx in [3, 4] else "FFFFFF")

        ws.cell(row=row_idx, column=3).number_format = "0.00"

    ws.freeze_panes = "A3"

    # 최근 실적 + 예측 결합 차트
    recent_n = min(12, len(df))
    recent = df.tail(recent_n)[["연월", "BOR"]].copy()
    recent["연월_str"] = recent["연월"].dt.strftime("%Y-%m")
    future_chart_df = future_df[["연월", "예측BOR"]].copy()
    future_chart_df["연월_str"] = future_chart_df["연월"].dt.strftime("%Y-%m")

    # 차트용 보조 시트 데이터 작성 (같은 시트에 숨긴 영역)
    chart_start_row = 3 + len(future_df) + 2

    ws.cell(row=chart_start_row, column=1, value="연월")
    ws.cell(row=chart_start_row, column=2, value="실적BOR")
    ws.cell(row=chart_start_row, column=3, value="예측BOR")
    _header_style(ws.cell(row=chart_start_row, column=1))
    _header_style(ws.cell(row=chart_start_row, column=2))
    _header_style(ws.cell(row=chart_start_row, column=3))

    for i, (_, r) in enumerate(recent.iterrows(), start=1):
        ws.cell(row=chart_start_row + i, column=1, value=r["연월_str"])
        ws.cell(row=chart_start_row + i, column=2, value=r["BOR"])

    offset = len(recent) + 1
    for i, (_, r) in enumerate(future_chart_df.iterrows(), start=1):
        ws.cell(row=chart_start_row + offset + i - 1, column=1, value=r["연월_str"])
        ws.cell(row=chart_start_row + offset + i - 1, column=3, value=r["예측BOR"])

    total_rows = len(recent) + len(future_chart_df)

    chart = LineChart()
    chart.title = "BOR 실적 및 예측 추이"
    chart.style = 10
    chart.y_axis.title = "BOR (%)"
    chart.x_axis.title = "연월"
    chart.height = 14
    chart.width = 24

    actual_data = Reference(ws, min_col=2, min_row=chart_start_row, max_row=chart_start_row + len(recent))
    pred_data = Reference(ws, min_col=3, min_row=chart_start_row, max_row=chart_start_row + total_rows)

    chart.add_data(actual_data, titles_from_data=True)
    chart.add_data(pred_data, titles_from_data=True)

    ws.add_chart(chart, "G2")


def write_summary_sheet(ws, df: pd.DataFrame, future_df: pd.DataFrame, metrics: dict):
    """요약 대시보드 시트 작성"""
    ws.title = "요약대시보드"

    ws.merge_cells("A1:F1")
    cell = ws["A1"]
    cell.value = "🏥 BOR 예측 모델 — 요약 대시보드"
    cell.font = Font(bold=True, size=16, color=COLOR_HEADER_BLUE)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    # KPI 카드
    kpis = [
        ("최근 BOR", f"{df['BOR'].iloc[-1]:.2f}%", "최근 1개월 실적"),
        ("연평균 BOR", f"{df['BOR'].mean():.2f}%", "전체 기간 평균"),
        ("최고 BOR", f"{df['BOR'].max():.2f}%", f"{df.loc[df['BOR'].idxmax(), '연월'].strftime('%Y-%m')}"),
        ("최저 BOR", f"{df['BOR'].min():.2f}%", f"{df.loc[df['BOR'].idxmin(), '연월'].strftime('%Y-%m')}"),
        ("예측 평균 BOR", f"{future_df['예측BOR'].mean():.2f}%", "향후 12개월 예측"),
        ("모델 R²", f"{metrics['R²']:.4f}", "모델 적합도"),
    ]

    kpi_colors = [
        COLOR_HEADER_LIGHT, "FFF2CC", "C6EFCE", "FFEB9C", "DAE8FC", "E2EFDA"
    ]

    ws.row_dimensions[3].height = 50
    ws.row_dimensions[4].height = 25
    ws.row_dimensions[5].height = 20

    for col_idx, ((title, value, subtitle), bg) in enumerate(zip(kpis, kpi_colors), start=1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = 18

        ws.merge_cells(f"{col_letter}3:{col_letter}3")
        title_cell = ws.cell(row=3, column=col_idx, value=title)
        title_cell.font = Font(bold=True, size=11, color="595959")
        title_cell.alignment = Alignment(horizontal="center", vertical="bottom")
        title_cell.fill = PatternFill("solid", fgColor=bg)

        value_cell = ws.cell(row=4, column=col_idx, value=value)
        value_cell.font = Font(bold=True, size=16, color=COLOR_HEADER_BLUE)
        value_cell.alignment = Alignment(horizontal="center", vertical="center")
        value_cell.fill = PatternFill("solid", fgColor=bg)

        sub_cell = ws.cell(row=5, column=col_idx, value=subtitle)
        sub_cell.font = Font(size=9, color="767676")
        sub_cell.alignment = Alignment(horizontal="center", vertical="top")
        sub_cell.fill = PatternFill("solid", fgColor=bg)

    # 범례
    ws["A7"] = "BOR 등급 기준"
    ws["A7"].font = Font(bold=True, size=12, color=COLOR_HEADER_BLUE)

    legend = [
        ("90% 이상", "과밀", "병상 과부하 위험 — 입원 제한 검토 필요", COLOR_BAD),
        ("75% ~ 90%", "적정", "정상 운영 범위", COLOR_GOOD),
        ("75% 미만", "여유", "병상 활용률 제고 방안 검토 필요", COLOR_WARN),
    ]
    for i, (rng, grade, desc, bg) in enumerate(legend, start=8):
        for col, val in enumerate([rng, grade, desc], start=1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BORDER
            if col == 2:
                cell.font = Font(bold=True)

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 45


def build_excel(
    df: pd.DataFrame,
    future_df: pd.DataFrame,
    metrics: dict,
    X_test,
    y_test,
    y_pred,
    output_path: str,
):
    """모든 시트를 포함한 Excel 파일을 생성합니다."""
    wb = openpyxl.Workbook()

    # 기본 시트 제거 후 순서대로 추가
    default_sheet = wb.active
    wb.remove(default_sheet)

    ws_summary = wb.create_sheet("요약대시보드")
    ws_data = wb.create_sheet("실적데이터")
    ws_model = wb.create_sheet("모델성능")
    ws_pred = wb.create_sheet("BOR예측")

    write_summary_sheet(ws_summary, df, future_df, metrics)
    write_data_sheet(ws_data, df)
    write_model_sheet(ws_model, metrics, X_test, y_test, y_pred)
    write_prediction_sheet(ws_pred, df, future_df)

    wb.save(output_path)
    print(f"✅ Excel 파일 저장 완료: {output_path}")


# ─────────────────────────────────────────────
# 4. 메인 실행
# ─────────────────────────────────────────────

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    data_path = os.path.join(base_dir, "data", "hospital_data.csv")
    output_path = os.path.join(base_dir, "BOR_예측모델.xlsx")

    print("📂 데이터 로딩 중...")
    df = load_data(data_path)
    df = add_features(df)

    print(f"   총 {len(df)}개월 데이터 로드 완료")
    print(f"   BOR 범위: {df['BOR'].min():.2f}% ~ {df['BOR'].max():.2f}%")
    print(f"   평균 BOR: {df['BOR'].mean():.2f}%")

    print("\n🤖 모델 학습 중...")
    model, X_train, X_test, y_train, y_test, y_pred, metrics = train_model(df)
    print(f"   R²  : {metrics['R²']}")
    print(f"   MAE : {metrics['MAE']} %p")
    print(f"   RMSE: {metrics['RMSE']} %p")

    print("\n🔮 향후 12개월 BOR 예측 중...")
    future_df = predict_future(model, df, months=12)
    for _, row in future_df.iterrows():
        print(f"   {row['연월'].strftime('%Y-%m')}: {row['예측BOR']:.2f}%")

    print("\n📊 Excel 파일 생성 중...")
    build_excel(df, future_df, metrics, X_test, y_test, y_pred, output_path)


if __name__ == "__main__":
    main()
