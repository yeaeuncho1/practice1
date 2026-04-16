# app.py
import io
from datetime import date
from pathlib import Path

import altair as alt
import openpyxl
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill


st.set_page_config(
    page_title="AI 행사 데이터 자동 취합",
    page_icon="📊",
    layout="wide",
)

st.markdown("""
<style>
.block-container {padding-top: 1.6rem; padding-bottom: 3rem;}
.metric-card {
    padding: 1rem 1.2rem;
    border: 1px solid #e5e7eb;
    border-radius: 16px;
    background: #ffffff;
    box-shadow: 0 4px 14px rgba(0,0,0,0.04);
}
.section-card {
    padding: 1.1rem 1.3rem;
    border: 1px solid #e5e7eb;
    border-radius: 18px;
    background: #fafafa;
    margin-bottom: 1rem;
}
.small-muted {color: #6b7280; font-size: 0.92rem;}
</style>
""", unsafe_allow_html=True)


TEMPLATE_KEYS = {
    "brand": "브랜드프로모션_AI최적화_최종양식.xlsx",
    "product": "특가상품_상품행사_AI최적화_최종양식.xlsx",
}


def find_template_path(file_name: str) -> Path | None:
    search_roots = [Path.cwd(), Path("/mnt/data"), Path.cwd() / "templates", Path("/mnt/data/templates")]
    for root in search_roots:
        candidate = root / file_name
        if candidate.exists():
            return candidate

    for root in search_roots:
        if root.exists():
            for path in root.glob("*.xlsx"):
                if file_name.split("_")[0] in path.name:
                    return path
    return None


def is_selected_store(value) -> bool:
    if value is None:
        return False
    text = str(value).strip()
    return text in {"○", "O", "o", "Y", "y", "1", "TRUE", "True", "true"}


def normalize_percent(value):
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        if value <= 1:
            return round(value * 100, 1)
        return round(float(value), 1)

    text = str(value).strip().replace("%", "")
    try:
        number = float(text)
        if number <= 1:
            return round(number * 100, 1)
        return round(number, 1)
    except Exception:
        return None


def parse_brand_workbook(file_bytes: bytes, source_name: str) -> pd.DataFrame:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
    rows = []

    for sheet_name in ["30%", "20%", "10%", "기타"]:
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]
        headers = [ws.cell(row=4, column=col).value for col in range(1, ws.max_column + 1)]
        store_headers = headers[4:]

        for row_idx in range(5, ws.max_row + 1):
            team = ws.cell(row=row_idx, column=1).value
            category = ws.cell(row=row_idx, column=2).value
            brand = ws.cell(row=row_idx, column=3).value
            note = ws.cell(row=row_idx, column=4).value

            selected_stores = []
            for offset, store_name in enumerate(store_headers, start=5):
                cell_value = ws.cell(row=row_idx, column=offset).value
                if is_selected_store(cell_value):
                    selected_stores.append(store_name)

            if not any([team, category, brand, note, selected_stores]):
                continue

            rows.append({
                "유형": "브랜드프로모션",
                "할인구간": sheet_name,
                "팀": team,
                "상품군": category,
                "브랜드명": brand,
                "상품명": None,
                "최초판매가": None,
                "할인판매가": None,
                "할인율(%)": None,
                "비고": note,
                "진행점포수": len(selected_stores),
                "진행점포": ", ".join([str(x) for x in selected_stores]),
                "원본파일": source_name,
                "원본시트": sheet_name,
            })

    return pd.DataFrame(rows)


def parse_product_workbook(file_bytes: bytes, source_name: str) -> pd.DataFrame:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
    if "상품행사" not in wb.sheetnames:
        return pd.DataFrame()

    ws = wb["상품행사"]
    headers = [ws.cell(row=4, column=col).value for col in range(1, ws.max_column + 1)]
    store_headers = headers[8:]

    rows = []
    for row_idx in range(5, ws.max_row + 1):
        team = ws.cell(row=row_idx, column=1).value
        category = ws.cell(row=row_idx, column=2).value
        brand = ws.cell(row=row_idx, column=3).value
        product = ws.cell(row=row_idx, column=4).value
        origin_price = ws.cell(row=row_idx, column=5).value
        sale_price = ws.cell(row=row_idx, column=6).value
        discount_rate = ws.cell(row=row_idx, column=7).value
        note = ws.cell(row=row_idx, column=8).value

        selected_stores = []
        for offset, store_name in enumerate(store_headers, start=9):
            cell_value = ws.cell(row=row_idx, column=offset).value
            if is_selected_store(cell_value):
                selected_stores.append(store_name)

        if not any([team, category, brand, product, origin_price, sale_price, note, selected_stores]):
            continue

        rows.append({
            "유형": "상품행사",
            "할인구간": None,
            "팀": team,
            "상품군": category,
            "브랜드명": brand,
            "상품명": product,
            "최초판매가": origin_price,
            "할인판매가": sale_price,
            "할인율(%)": normalize_percent(discount_rate),
            "비고": note,
            "진행점포수": len(selected_stores),
            "진행점포": ", ".join([str(x) for x in selected_stores]),
            "원본파일": source_name,
            "원본시트": "상품행사",
        })

    return pd.DataFrame(rows)


def parse_uploaded_file(uploaded_file) -> pd.DataFrame:
    file_bytes = uploaded_file.getvalue()

    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
        sheetnames = wb.sheetnames
    except Exception:
        return pd.DataFrame()

    if "상품행사" in sheetnames:
        return parse_product_workbook(file_bytes, uploaded_file.name)

    if any(name in sheetnames for name in ["30%", "20%", "10%", "기타"]):
        return parse_brand_workbook(file_bytes, uploaded_file.name)

    return pd.DataFrame()


def add_event_metadata(df: pd.DataFrame, event_info: dict, collection_round: str) -> pd.DataFrame:
    if df.empty:
        return df

    result = df.copy()
    result.insert(0, "행사제목", event_info["title"])
    result.insert(1, "행사시작일", event_info["start_date"])
    result.insert(2, "행사종료일", event_info["end_date"])
    result.insert(3, "취합라운드", collection_round)
    return result


def build_deadline_df(deadline_labels, deadline_dates) -> pd.DataFrame:
    rows = []
    for label, due in zip(deadline_labels, deadline_dates):
        rows.append({"구분": label, "데드라인": due})
    return pd.DataFrame(rows)


def dataframe_to_excel_bytes(event_info, deadline_df, integrated_df, upload_log_df) -> bytes:
    wb = Workbook()
    ws_info = wb.active
    ws_info.title = "행사정보"

    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    info_rows = [
        ["항목", "값"],
        ["행사제목", event_info["title"]],
        ["행사시작일", event_info["start_date"]],
        ["행사종료일", event_info["end_date"]],
        ["행사기간(일)", event_info["duration_days"]],
        ["설명", event_info["description"]],
    ]

    for row in info_rows:
        ws_info.append(row)

    for cell in ws_info[1]:
        cell.fill = header_fill
        cell.font = header_font

    ws_deadline = wb.create_sheet("취합일정")
    for row in [deadline_df.columns.tolist()] + deadline_df.astype(str).values.tolist():
        ws_deadline.append(row)
    for cell in ws_deadline[1]:
        cell.fill = header_fill
        cell.font = header_font

    ws_data = wb.create_sheet("통합마스터")
    if not integrated_df.empty:
        export_df = integrated_df.copy()
        export_df["행사시작일"] = export_df["행사시작일"].astype(str)
        export_df["행사종료일"] = export_df["행사종료일"].astype(str)
        for row in [export_df.columns.tolist()] + export_df.fillna("").values.tolist():
            ws_data.append(row)
        for cell in ws_data[1]:
            cell.fill = header_fill
            cell.font = header_font
    else:
        ws_data.append(["데이터 없음"])

    ws_log = wb.create_sheet("업로드로그")
    if not upload_log_df.empty:
        for row in [upload_log_df.columns.tolist()] + upload_log_df.astype(str).values.tolist():
            ws_log.append(row)
        for cell in ws_log[1]:
            cell.fill = header_fill
            cell.font = header_font
    else:
        ws_log.append(["업로드 이력 없음"])

    for ws in wb.worksheets:
        for col_cells in ws.columns:
            max_length = 0
            col_letter = col_cells[0].column_letter
            for cell in col_cells:
                value = "" if cell.value is None else str(cell.value)
                max_length = max(max_length, len(value))
            ws.column_dimensions[col_letter].width = min(max(max_length + 2, 12), 40)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


st.title("AI 행사 데이터 자동 취합 시스템")
st.caption("기본 양식 다운로드, 행사 정보 입력, 파일 업로드, 통합 마스터 생성까지 한 화면에서 처리합니다.")

with st.sidebar:
    st.header("행사 기본 정보")
    event_title = st.text_input("행사 제목", value="2025 봄 시즌 대형 행사")
    event_period = st.columns(2)
    event_start = event_period[0].date_input("행사 시작일", value=date.today())
    event_end = event_period[1].date_input("행사 종료일", value=date.today())

    st.markdown("### 취합 데드라인")
    deadline_count = st.number_input("취합 라운드 수", min_value=1, max_value=5, value=2, step=1)

    deadline_labels = []
    deadline_dates = []
    for idx in range(int(deadline_count)):
        label_default = f"{idx + 1}차 취합일"
        cols = st.columns([1.2, 1])
        label = cols[0].text_input(f"라운드명 {idx + 1}", value=label_default, key=f"deadline_label_{idx}")
        due_date = cols[1].date_input(f"날짜 {idx + 1}", value=date.today(), key=f"deadline_date_{idx}")
        deadline_labels.append(label)
        deadline_dates.append(due_date)

    collection_round = st.selectbox("현재 업로드 라운드", deadline_labels, index=0)
    event_description = st.text_area(
        "행사 설명",
        value="표준 엑셀 양식을 기반으로 여러 부서 데이터를 자동 취합하는 행사입니다.",
        height=100,
    )

event_info = {
    "title": event_title,
    "start_date": event_start,
    "end_date": event_end,
    "duration_days": (event_end - event_start).days + 1,
    "description": event_description,
}

deadline_df = build_deadline_df(deadline_labels, deadline_dates)

col1, col2, col3 = st.columns(3)
col1.metric("행사 제목", event_title if event_title else "-")
col2.metric("행사 기간", f"{event_start} ~ {event_end}")
col3.metric("등록된 데드라인", f"{len(deadline_df)}건")

st.markdown("---")

left, right = st.columns([1, 1])

with left:
    st.subheader("1. 기본 양식 다운로드")
    brand_template_path = find_template_path(TEMPLATE_KEYS["brand"])
    product_template_path = find_template_path(TEMPLATE_KEYS["product"])

    st.markdown('<div class="section-card">부서 배포용 기본 양식을 바로 다운로드할 수 있습니다.</div>', unsafe_allow_html=True)

    if brand_template_path and brand_template_path.exists():
        with open(brand_template_path, "rb") as f:
            st.download_button(
                label="브랜드 프로모션 양식 다운로드",
                data=f.read(),
                file_name=brand_template_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    else:
        st.warning("브랜드 프로모션 양식 파일을 찾지 못했습니다.")

    if product_template_path and product_template_path.exists():
        with open(product_template_path, "rb") as f:
            st.download_button(
                label="특가상품 / 상품행사 양식 다운로드",
                data=f.read(),
                file_name=product_template_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    else:
        st.warning("특가상품 / 상품행사 양식 파일을 찾지 못했습니다.")

with right:
    st.subheader("2. 행사 정보 및 데드라인")
    st.dataframe(deadline_df, use_container_width=True, hide_index=True)
    st.markdown(f"""
<div class="section-card">
<b>행사 제목</b><br>{event_info["title"]}<br><br>
<b>행사 기간</b><br>{event_info["start_date"]} ~ {event_info["end_date"]}<br><br>
<b>행사 설명</b><br>{event_info["description"]}
</div>
""", unsafe_allow_html=True)

st.markdown("---")

st.subheader("3. 양식 업로드")
uploaded_files = st.file_uploader(
    "작성 완료된 양식을 여러 개 업로드하세요.",
    type=["xlsx"],
    accept_multiple_files=True,
    help="브랜드 프로모션 양식, 특가상품/상품행사 양식을 함께 올릴 수 있습니다.",
)

integrated_frames = []
upload_log_rows = []

if uploaded_files:
    for uploaded_file in uploaded_files:
        parsed_df = parse_uploaded_file(uploaded_file)
        file_type = "인식 실패"
        row_count = 0

        if not parsed_df.empty:
            parsed_df = add_event_metadata(parsed_df, event_info, collection_round)
            integrated_frames.append(parsed_df)
            file_type = ", ".join(parsed_df["유형"].dropna().astype(str).unique().tolist())
            row_count = len(parsed_df)

        upload_log_rows.append({
            "파일명": uploaded_file.name,
            "파일유형": file_type,
            "반영행수": row_count,
            "취합라운드": collection_round,
        })

upload_log_df = pd.DataFrame(upload_log_rows)

integrated_df = pd.concat(integrated_frames, ignore_index=True) if integrated_frames else pd.DataFrame()

if integrated_df.empty:
    st.info("업로드된 파일이 아직 없거나, 양식 구조를 인식하지 못했습니다.")
else:
    st.success(f"총 {len(uploaded_files)}개 파일에서 {len(integrated_df):,}건의 데이터를 취합했습니다.")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("업로드 파일 수", len(uploaded_files))
    m2.metric("통합 데이터 건수", f"{len(integrated_df):,}")
    m3.metric("브랜드 프로모션 건수", f"{(integrated_df['유형'] == '브랜드프로모션').sum():,}")
    m4.metric("상품행사 건수", f"{(integrated_df['유형'] == '상품행사').sum():,}")

    st.markdown("---")

    summary_by_type = (
        integrated_df.groupby("유형", dropna=False)
        .size()
        .reset_index(name="건수")
    )

    st.subheader("4. 취합 결과 요약")
    chart = (
        alt.Chart(summary_by_type)
        .mark_bar()
        .encode(
            x=alt.X("유형:N", title=None),
            y=alt.Y("건수:Q", title="건수"),
            tooltip=["유형", "건수"],
        )
        .properties(height=280)
    )
    st.altair_chart(chart, use_container_width=True)

    tab1, tab2, tab3 = st.tabs(["통합 마스터", "업로드 로그", "행사 정보"])

    with tab1:
        st.dataframe(integrated_df, use_container_width=True, hide_index=True)

    with tab2:
        st.dataframe(upload_log_df, use_container_width=True, hide_index=True)

    with tab3:
        event_summary_df = pd.DataFrame([
            {"항목": "행사 제목", "값": event_info["title"]},
            {"항목": "행사 시작일", "값": event_info["start_date"]},
            {"항목": "행사 종료일", "값": event_info["end_date"]},
            {"항목": "행사 기간(일)", "값": event_info["duration_days"]},
            {"항목": "현재 업로드 라운드", "값": collection_round},
            {"항목": "행사 설명", "값": event_info["description"]},
        ])
        st.dataframe(event_summary_df, use_container_width=True, hide_index=True)
        st.dataframe(deadline_df, use_container_width=True, hide_index=True)

    excel_bytes = dataframe_to_excel_bytes(event_info, deadline_df, integrated_df, upload_log_df)
    st.download_button(
        label="통합 결과 엑셀 다운로드",
        data=excel_bytes,
        file_name=f"{event_title}_통합마스터.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.markdown("---")
st.markdown(
    "<div class='small-muted'>실행 방법: <code>streamlit run app.py</code></div>",
    unsafe_allow_html=True,
)
