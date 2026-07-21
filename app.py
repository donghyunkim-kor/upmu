import os
import re
import pandas as pd
import streamlit as st
from datetime import datetime
from streamlit_calendar import calendar

# 페이지 기본 설정
st.set_page_config(page_title="건설 업무 지원 시스템", layout="wide")

# 폴더 설정
BASE_DATA_DIR = "uploaded_data"
BID_DIR = os.path.join(BASE_DATA_DIR, "bids")
ENG_DIR = os.path.join(BASE_DATA_DIR, "engineer")
PERF_DIR = os.path.join(BASE_DATA_DIR, "performance")

for d in [BID_DIR, ENG_DIR, PERF_DIR]:
    os.makedirs(d, exist_ok=True)


def save_uploaded_file(uploaded_file, target_dir):
    save_path = os.path.join(target_dir, uploaded_file.name)
    with open(save_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return save_path


def get_saved_files(target_dir):
    return [f for f in os.listdir(target_dir) if f.endswith(".xlsx")]


def delete_saved_file(file_path):
    if os.path.exists(file_path):
        os.remove(file_path)
        return True
    return False


# ---------------------------------------------------------
# OZ Report 입찰명 정밀 컬럼 파서
# ---------------------------------------------------------
def parse_oz_report_4schedules(file_path):
    raw_df = pd.read_excel(file_path, header=None)

    current_year = datetime.now().year
    for cell in raw_df.iloc[:5].values.flatten():
        cell_str = str(cell)
        match = re.search(r"202\d", cell_str)
        if match:
            current_year = int(match.group())
            break

    header_idx = -1
    title_col_idx = -1
    col_map = {}

    for idx, row in raw_df.iterrows():
        row_cells = [str(c).replace(" ", "").replace("\n", "") for c in row]
        row_str = "".join(row_cells)

        if "공사명" in row_str or "입찰일정" in row_str:
            header_idx = idx
            for c_idx, val in enumerate(row):
                v_clean = str(val).replace(" ", "").replace("\n", "")
                if "공사명" in v_clean:
                    title_col_idx = c_idx
                    break
            break

    if header_idx == -1:
        header_idx = 0
    if title_col_idx == -1:
        title_col_idx = 1

    data_df = raw_df.iloc[header_idx:].reset_index(drop=True)

    for r_idx in range(min(5, len(data_df))):
        row = data_df.iloc[r_idx]
        for c_idx, val in enumerate(row):
            if pd.isna(val):
                continue
            v_str = str(val).replace(" ", "").replace("\n", "")
            if "PQ" in v_str or "실적" in v_str:
                col_map["PQ"] = c_idx
            elif "협정" in v_str:
                col_map["협정"] = c_idx
            elif "등록" in v_str:
                col_map["등록"] = c_idx
            elif "입찰" in v_str and "마감" in v_str:
                col_map["입찰"] = c_idx

    parsed_events = []

    for idx in range(len(data_df)):
        row = data_df.iloc[idx]

        if len(row) <= title_col_idx or pd.isna(row.iloc[title_col_idx]):
            continue

        raw_title = str(row.iloc[title_col_idx]).strip()

        if (
            "공사명" in raw_title
            or "입찰일정" in raw_title
            or "년도" in raw_title
            or "페이지" in raw_title
            or len(raw_title) < 2
        ):
            continue

        title_lines = [
            line.strip()
            for line in raw_title.split("\n")
            if line.strip() and not line.strip().isdigit()
        ]
        if not title_lines:
            continue
        clean_title = title_lines[0]

        categories = [
            ("PQ", "PQ", "#E1BEE7"),
            ("협정", "협정", "#C8E6C9"),
            ("등록", "등록", "#FFE0B2"),
            ("입찰", "입찰", "#BBDEFB"),
        ]

        for cat_key, cat_label, color in categories:
            target_col = col_map.get(cat_key)

            if target_col is None:
                default_cols = {
                    "PQ": [4, 5],
                    "협정": [6, 7],
                    "등록": [7, 8],
                    "입찰": [8, 9],
                }
                search_cols = default_cols[cat_key]
            else:
                search_cols = [target_col]

            date_str = None
            for c_i in search_cols:
                if len(row) > c_i and pd.notna(row.iloc[c_i]):
                    val1 = str(row.iloc[c_i])
                    m = re.search(r"(\d{1,2})[/.-](\d{1,2})", val1)
                    if m:
                        date_str = f"{current_year}-{int(m.group(1)):02d}-{int(m.group(2)):02d}"
                        break

            if date_str:
                event_title = f"{cat_label}_ {clean_title}"
                parsed_events.append(
                    {
                        "title": event_title,
                        "start": date_str,
                        "end": date_str,
                        "backgroundColor": color,
                        "borderColor": color,
                        "textColor": "#1A1A1A",
                    }
                )

    return parsed_events


st.title("🏗️ 건설 입찰 및 경력/실적 통합 관리 시스템")

tab1, tab2, tab3 = st.tabs(
    ["📅 입찰 달력", "👷 경력기술자 조건 검색", "🏢 준공실적 검색"]
)

# ---------------------------------------------------------
# [TAB 1] 입찰 달력
# ---------------------------------------------------------
with tab1:
    st.header("📅 입찰 일정 달력")

    col_up, col_sel = st.columns([1, 1])

    with col_up:
        uploaded_bid = st.file_uploader(
            "OZ Report 입찰 일정 엑셀(.xlsx) 업로드",
            type=["xlsx"],
            key="bid_uploader",
        )
        if uploaded_bid is not None:
            save_uploaded_file(uploaded_bid, BID_DIR)
            st.success(f"✅ '{uploaded_bid.name}' 업로드 완료!")
            # 상태 및 캐시 완전 초기화 후 리런
            st.cache_data.clear()
            for key in ["bid_select", "bid_calendar"]:
                if key in st.session_state:
                    del st.session_state[key]
            st.rerun()

    saved_bid_files = get_saved_files(BID_DIR)

    with col_sel:
        if saved_bid_files:
            # 안전하게 인덱스 처리
            selected_bid_file = st.selectbox(
                "📁 불러올 입찰 파일 선택", saved_bid_files, key="bid_select"
            )
            if st.button("🗑️ 선택한 입찰 파일 삭제", key="del_bid"):
                file_path_to_del = os.path.join(BID_DIR, selected_bid_file)
                if delete_saved_file(file_path_to_del):
                    st.cache_data.clear()
                    for key in ["bid_select", "bid_uploader", "bid_calendar"]:
                        if key in st.session_state:
                            del st.session_state[key]
                    st.success(f"🗑️ '{selected_bid_file}' 파일이 삭제되었습니다.")
                    st.rerun()
        else:
            selected_bid_file = None
            st.info("💡 왼쪽에 OZ Report 엑셀 파일을 업로드해 주세요.")

    st.markdown("---")

    bid_events = []

    if selected_bid_file and selected_bid_file in saved_bid_files:
        file_path = os.path.join(BID_DIR, selected_bid_file)
        if os.path.exists(file_path):
            bid_events = parse_oz_report_4schedules(file_path)
            st.caption(
                f"💡 **입찰명 정밀 파싱 완료** (총 {len(bid_events)}건 일정 생성)"
            )

    st.markdown(
        """
        <div style="display: flex; gap: 20px; align-items: center; margin-bottom: 12px; font-weight: bold; font-size: 13px;">
            <span>🟣 <span style="background-color:#E1BEE7; padding:2px 8px; border-radius:4px; color:#1A1A1A;">PQ/실적</span></span>
            <span>🟢 <span style="background-color:#C8E6C9; padding:2px 8px; border-radius:4px; color:#1A1A1A;">협정마감</span></span>
            <span>🟠 <span style="background-color:#FFE0B2; padding:2px 8px; border-radius:4px; color:#1A1A1A;">등록마감</span></span>
            <span>🔵 <span style="background-color:#BBDEFB; padding:2px 8px; border-radius:4px; color:#1A1A1A;">입찰마감</span></span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    calendar_options = {
        "headerToolbar": {
            "left": "prev,next today",
            "center": "title",
            "right": "dayGridMonth,dayGridWeek",
        },
        "initialView": "dayGridMonth",
        "locale": "ko",
        "firstDay": 0,
        "height": 650,
        "selectable": True,
        "editable": False,
        "dayMaxEvents": 3,
    }

    custom_css = """
        .fc-header-toolbar { margin-bottom: 8px !important; }
        .fc-toolbar-title { font-size: 1.1rem !important; font-weight: bold; }
        .fc-col-header-cell { background-color: #0c2340 !important; color: white !important; padding: 4px 0 !important; font-size: 13px; }
        .fc-col-header-cell-cushion { color: white !important; text-decoration: none !important; }
        .fc-day-sun .fc-col-header-cell-cushion, .fc-day-sun .fc-daygrid-day-number { color: #ff4d4f !important; }
        .fc-day-sat .fc-col-header-cell-cushion, .fc-day-sat .fc-daygrid-day-number { color: #1890ff !important; }
        .fc-daygrid-day-number { font-weight: bold; font-size: 12px; text-decoration: none !important; padding: 2px 5px !important; }
        .fc-event { font-size: 11px !important; padding: 1px 3px !important; cursor: pointer; border: none !important; font-weight: 500; }
        .fc-event-main { color: #1A1A1A !important; }
        .fc-daygrid-day-frame { min-height: 85px !important; }
    """

    state = calendar(
        events=bid_events,
        options=calendar_options,
        custom_css=custom_css,
        key="bid_calendar",
    )

# ---------------------------------------------------------
# [TAB 2] 경력기술자 검색
# ---------------------------------------------------------
with tab2:
    st.header("👷 경력기술자 공종 및 경력일수 조건 검색")
    col_up, col_sel = st.columns([1, 1])
    with col_up:
        uploaded_eng = st.file_uploader(
            "새 경력 엑셀 파일(.xlsx) 업로드", type=["xlsx"], key="eng_uploader"
        )
        if uploaded_eng is not None:
            save_uploaded_file(uploaded_eng, ENG_DIR)
            st.success(f"✅ '{uploaded_eng.name}' 경력 파일 저장 완료!")
            st.cache_data.clear()
            if "eng_select" in st.session_state:
                del st.session_state["eng_select"]
            st.rerun()

    saved_eng_files = get_saved_files(ENG_DIR)
    with col_sel:
        if saved_eng_files:
            selected_eng_file = st.selectbox(
                "📁 경력 파일 목록 선택", saved_eng_files, key="eng_select"
            )
            if st.button("🗑️ 선택한 경력 파일 삭제", key="del_eng"):
                file_path_to_del = os.path.join(ENG_DIR, selected_eng_file)
                if delete_saved_file(file_path_to_del):
                    st.cache_data.clear()
                    for key in ["eng_select", "eng_uploader"]:
                        if key in st.session_state:
                            del st.session_state[key]
                    st.success(f"🗑️ '{selected_eng_file}' 파일이 삭제되었습니다.")
                    st.rerun()
        else:
            selected_eng_file = None
            st.info("💡 왼쪽에 경력 엑셀 파일을 업로드해 주세요.")

    st.markdown("---")

    if selected_eng_file and selected_eng_file in saved_eng_files:
        file_path = os.path.join(ENG_DIR, selected_eng_file)
        if os.path.exists(file_path):
            df_engineer = pd.read_excel(file_path)

            st.subheader(
                f"2. 조건 설정 (대상 파일: {selected_eng_file} / 총 {len(df_engineer)}건 레코드)"
            )
            c1, c2, c3, c4 = st.columns([1.5, 2, 1.5, 1.5])

            with c1:
                name_search = st.text_input(
                    "이름 검색 (부분 검색 가능)", "", key="eng_name"
                )
            with c2:
                type_search = st.text_input(
                    "🎯 공사종류 검색 (예: 고속도로, 준설, 교량)",
                    value="",
                    key="eng_type_search",
                )
            with c3:
                min_days = st.number_input(
                    "⏳ 최소 인정일수 (일 기준)",
                    min_value=0,
                    value=0,
                    step=30,
                    key="eng_days",
                )
            with c4:
                duty_search = st.text_input(
                    "담당업무 (예: 시공, 대리 등)", "", key="eng_duty"
                )

            filtered_df = df_engineer.copy()

            if "인정일수" in filtered_df.columns:
                filtered_df["인정일수"] = pd.to_numeric(
                    filtered_df["인정일수"], errors="coerce"
                ).fillna(0)
            else:
                filtered_df["인정일수"] = 0

            if type_search.strip() and "공사종류" in filtered_df.columns:
                filtered_df = filtered_df[
                    filtered_df["공사종류"]
                    .astype(str)
                    .str.contains(type_search.strip(), na=False)
                ]

            if duty_search.strip() and "담당업무" in filtered_df.columns:
                filtered_df = filtered_df[
                    filtered_df["담당업무"]
                    .astype(str)
                    .str.contains(duty_search.strip(), na=False)
                ]

            if name_search.strip() and "이름" in filtered_df.columns:
                filtered_df = filtered_df[
                    filtered_df["이름"]
                    .astype(str)
                    .str.contains(name_search.strip(), na=False)
                ]

            if min_days > 0 and "이름" in filtered_df.columns:
                person_days = (
                    filtered_df.groupby("이름")["인정일수"].sum().reset_index()
                )
                target_persons = person_days[person_days["인정일수"] >= min_days]["이름"]
                filtered_df = filtered_df[filtered_df["이름"].isin(target_persons)]

            st.markdown("---")

            if not filtered_df.empty and "이름" in filtered_df.columns:
                summary_df = (
                    filtered_df.groupby("이름")
                    .agg(
                        건수=("사업명", "count"),
                        총인정일수=("인정일수", "sum"),
                    )
                    .reset_index()
                )
                summary_df["추정경력년수"] = summary_df["총인정일수"].apply(
                    lambda d: f"{int(d // 365)}년 {int((d % 365) // 30)}개월 ({int(d)}일)"
                )

                st.subheader(
                    f"🎯 조건 충족 적합 인원: 총 {len(summary_df)}명 (상세 이력 {len(filtered_df)}건)"
                )
                st.markdown("##### 📌 적합 기술자 요약 목록")
                st.dataframe(summary_df, use_container_width=True)
                st.markdown("##### 📄 적합 기술자의 상세 경력 내역")
                st.dataframe(filtered_df, use_container_width=True)
            else:
                st.warning("⚠️ 설정한 조건에 맞는 기술자 경력이 없습니다.")

# ---------------------------------------------------------
# [TAB 3] 준공실적 검색
# ---------------------------------------------------------
with tab3:
    st.header("🏢 준공실적 조건 검색 (PQ / 적격심사용)")
    col_up2, col_sel2 = st.columns([1, 1])

    with col_up2:
        uploaded_perf = st.file_uploader(
            "새 실적 엑셀 파일(.xlsx) 업로드", type=["xlsx"], key="perf_uploader"
        )
        if uploaded_perf is not None:
            save_uploaded_file(uploaded_perf, PERF_DIR)
            st.success(f"✅ '{uploaded_perf.name}' 실적 파일 저장 완료!")
            st.cache_data.clear()
            if "perf_select" in st.session_state:
                del st.session_state["perf_select"]
            st.rerun()

    saved_perf_files = get_saved_files(PERF_DIR)

    with col_sel2:
        if saved_perf_files:
            selected_perf_file = st.selectbox(
                "📁 실적 파일 목록 선택", saved_perf_files, key="perf_select"
            )
            if st.button("🗑️ 선택한 실적 파일 삭제", key="del_perf"):
                file_path_to_del = os.path.join(PERF_DIR, selected_perf_file)
                if delete_saved_file(file_path_to_del):
                    st.cache_data.clear()
                    for key in ["perf_select", "perf_uploader"]:
                        if key in st.session_state:
                            del st.session_state[key]
                    st.success(f"🗑️ '{selected_perf_file}' 파일이 삭제되었습니다.")
                    st.rerun()
        else:
            selected_perf_file = None
            st.info("💡 왼쪽에 실적 엑셀 파일을 업로드해 주세요.")

    st.markdown("---")

    if selected_perf_file and selected_perf_file in saved_perf_files:
        file_path = os.path.join(PERF_DIR, selected_perf_file)
        if os.path.exists(file_path):
            df_perf = pd.read_excel(file_path)

            st.subheader(
                f"2. 실적 필터링 (선택된 파일: {selected_perf_file} / 총 {len(df_perf)}건)"
            )
            col1, col2 = st.columns(2)
            with col1:
                min_amount = st.number_input(
                    "최저 실적금액 필터 (단위: 백만원/원 기준)",
                    value=0,
                    step=100,
                )
            with col2:
                years_limit = st.slider(
                    "최근 N년 이내 실적", min_value=1, max_value=15, value=10
                )

            filtered_perf = df_perf.copy()
            if "금액" in filtered_perf.columns:
                filtered_perf = filtered_perf[
                    pd.to_numeric(filtered_perf["금액"], errors="coerce").fillna(0)
                    >= min_amount
                ]

            st.write(f"🔍 **검색 결과: 총 {len(filtered_perf)}건**")
            st.dataframe(filtered_perf, use_container_width=True)
