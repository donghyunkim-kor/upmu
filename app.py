import os
import re
import pandas as pd
import streamlit as st
from streamlit_calendar import calendar

# ---------------------------------------------------------
# [기본 페이지 설정 및 폴더 생성]
# ---------------------------------------------------------
st.set_page_config(
    page_title="통합 건설 사업 관리 시스템",
    page_icon="🏗️",
    layout="wide",
)

BID_DIR = "./uploaded_bid_files"
ENG_DIR = "./uploaded_eng_files"
PERF_DIR = "./uploaded_perf_files"

os.makedirs(BID_DIR, exist_ok=True)
os.makedirs(ENG_DIR, exist_ok=True)
os.makedirs(PERF_DIR, exist_ok=True)


# ---------------------------------------------------------
# [유틸리티 함수 정의]
# ---------------------------------------------------------
def save_uploaded_file(uploaded_file, target_dir):
    """업로드된 파일을 지정된 폴더에 저장합니다."""
    file_path = os.path.join(target_dir, uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return file_path


def get_saved_files(target_dir):
    """지정된 폴더 내의 엑셀 파일 목록을 최신 수정순으로 가져옵니다."""
    if not os.path.exists(target_dir):
        return []
    files = [f for f in os.listdir(target_dir) if f.endswith(".xlsx") or f.endswith(".xls")]
    files.sort(key=lambda x: os.path.getmtime(os.path.join(target_dir, x)), reverse=True)
    return files


def delete_saved_file(file_path):
    """지정된 파일을 삭제합니다."""
    if os.path.exists(file_path):
        os.remove(file_path)
        return True
    return False


def parse_flexible_date(val):
    """어떤 형태의 날짜/시간 텍스트든 감지하여 추출하는 유틸리티"""
    if pd.isna(val) or not str(val).strip():
        return None
    val_str = str(val).strip()
    match = re.search(r"(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})", val_str)
    if match:
        date_part = match.group(1).replace(".", "-").replace("/", "-")
        time_match = re.search(r"(\d{1,2}:\d{2})", val_str)
        if time_match:
            return f"{date_part} {time_match.group(1)}"
        return date_part
    return None


# ---------------------------------------------------------
# [TAB 1: 범용 100% 매칭 입찰 일정 파서 (막강 복구판)]
# ---------------------------------------------------------
def parse_bid_excel(file_path):
    """
    행 전체를 탐색하여 날짜 데이터가 포함된 경우 무조건 이벤트를 추출합니다.
    """
    df_raw = pd.read_excel(file_path, header=None)
    events = []
    
    for idx, row in df_raw.iterrows():
        row_values = row.dropna().tolist()
        if not row_values:
            continue
            
        row_str = " ".join([str(v) for v in row_values])
        
        # 안내 문구나 헤더, 합계 행 스킵
        if any(w in row_str for w in ["합계", "소계", "공사명/발주처", "페이지", "Printed"]):
            continue

        # 행 내에서 날짜가 포함된 셀들을 모두 찾기
        date_entries = []
        title = None
        client = "발주처 미상"

        # 첫 번째나 두 번째 텍스트가 긴 것을 공사명(제목)으로 추정
        text_candidates = [str(v).strip() for v in row_values if isinstance(v, str) and len(str(v).strip()) > 3]
        if text_candidates:
            title = text_candidates[0]
            if len(text_candidates) > 1 and "공사" not in text_candidates[1] and "입찰" not in text_candidates[1]:
                client = text_candidates[1]

        # 행 전체 셀을 돌며 날짜 찾기
        for col_idx, val in enumerate(row):
            parsed_date = parse_flexible_date(val)
            if parsed_date:
                # 컬럼 위치나 헤더명에 따라 마감 종류 유추
                label = f"마감일정_{col_idx}"
                color = "#1A535C"
                
                # 열 번호나 주변 텍스트 기반으로 이름표 부착
                if col_idx in [4, 5]:
                    label, color = "PQ마감", "#FF6B6B"
                elif col_idx in [6, 7]:
                    label, color = "협정마감", "#ECDC4"
                elif col_idx in [8, 9]:
                    label, color = "등록마감", "#FFE66D"
                elif col_idx >= 10:
                    label, color = "입찰마감", "#1A535C"

                if title and len(title) > 1:
                    events.append({
                        "title": f"[{label}] {title} ({client})",
                        "start": parsed_date,
                        "color": color,
                        "extendedProps": {"type": label, "client": client, "full_title": title}
                    })

    # 중복 이벤트 제거 (동일 날짜, 동일 제목)
    unique_events = []
    seen = set()
    for ev in events:
        identifier = (ev["title"], ev["start"])
        if identifier not in seen:
            seen.add(identifier)
            unique_events.append(ev)

    return unique_events, df_raw


# ---------------------------------------------------------
# [TAB 2: 경력기술자 파싱 및 데이터 정제]
# ---------------------------------------------------------
def load_and_clean_engineer_data(file_path):
    df_raw = pd.read_excel(file_path, header=None)
    header_idx = None
    for idx, row in df_raw.iterrows():
        row_str = " ".join(row.dropna().astype(str))
        if any(k in row_str for k in ["성명", "이름", "공사종류", "인정일수", "담당업무", "사업명"]):
            header_idx = idx
            break
            
    df = pd.read_excel(file_path, header=header_idx) if header_idx is not None else pd.read_excel(file_path)
    df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed")]
    
    col_mapping = {}
    for col in df.columns:
        c_str = str(col).replace(" ", "").strip()
        if any(x in c_str for x in ["성명", "이름", "기술자명"]): col_mapping[col] = "이름"
        elif any(x in c_str for x in ["공사종류", "공종", "전문분야"]): col_mapping[col] = "공사종류"
        elif any(x in c_str for x in ["인정일수", "경력일수", "참여일수"]): col_mapping[col] = "인정일수"
        elif any(x in c_str for x in ["담당업무", "직무", "직책"]): col_mapping[col] = "담당업무"
        elif any(x in c_str for x in ["사업명", "공사명", "프로젝트명"]): col_mapping[col] = "사업명"

    df = df.rename(columns=col_mapping)
    if "이름" not in df.columns: df["이름"] = "미상"
    if "공사종류" not in df.columns: df["공사종류"] = "-"
    if "담당업무" not in df.columns: df["담당업무"] = "-"
    if "사업명" not in df.columns: df["사업명"] = "사업명 미상"
    df["인정일수"] = pd.to_numeric(df["인정일수"], errors="coerce").fillna(0) if "인정일수" in df.columns else 0
    return df.dropna(subset=["이름"])


# ---------------------------------------------------------
# [TAB 3: 준공실적증명 파싱]
# ---------------------------------------------------------
def load_and_clean_perf_data(file_path):
    df_raw = pd.read_excel(file_path, header=None)
    header_idx = None
    for idx, row in df_raw.iterrows():
        row_str = " ".join(row.dropna().astype(str))
        if any(k in row_str for k in ["공사명", "발주처", "계약금액", "준공일자"]):
            header_idx = idx
            break
    df = pd.read_excel(file_path, header=header_idx) if header_idx is not None else pd.read_excel(file_path)
    return df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed")]


# ---------------------------------------------------------
# [메인 UI - 3개 탭]
# ---------------------------------------------------------
st.title("🏗️ 통합 건설 사업 관리 시스템")

tab1, tab2, tab3 = st.tabs(["📅 입찰 일정 캘린더", "👷 경력기술자 검색", "📜 준공실적증명 관리"])

# =========================================================
# [TAB 1] 입찰 일정 캘린더
# =========================================================
with tab1:
    st.header("📅 입찰 일정 및 마감일 캘린더")
    col_up1, col_sel1 = st.columns([1, 1])

    with col_up1:
        uploaded_bid = st.file_uploader("새 입찰 일정 엑셀 파일(.xlsx) 업로드", type=["xlsx", "xls"], key="bid_uploader")
        if uploaded_bid is not None:
            save_uploaded_file(uploaded_bid, BID_DIR)
            st.success(f"✅ '{uploaded_bid.name}' 업로드 완료!")
            st.cache_data.clear()
            st.session_state["bid_select"] = uploaded_bid.name
            st.rerun()

    saved_bid_files = get_saved_files(BID_DIR)

    with col_sel1:
        if saved_bid_files:
            if "bid_select" not in st.session_state or st.session_state["bid_select"] not in saved_bid_files:
                st.session_state["bid_select"] = saved_bid_files[0]

            selected_bid_file = st.selectbox("📁 입찰 파일 선택", saved_bid_files, key="bid_select")

            if st.button("🗑️ 파일 삭제", key="del_bid"):
                if delete_saved_file(os.path.join(BID_DIR, selected_bid_file)):
                    st.cache_data.clear()
                    if "bid_select" in st.session_state: del st.session_state["bid_select"]
                    st.rerun()
        else:
            selected_bid_file = None
            st.info("💡 입찰 엑셀 파일을 업로드해 주세요.")

    st.markdown("---")

    if selected_bid_file and selected_bid_file in saved_bid_files:
        file_path = os.path.join(BID_DIR, selected_bid_file)
        if os.path.exists(file_path):
            events, df_raw = parse_bid_excel(file_path)
            
            calendar_options = {
                "editable": False,
                "selectable": True,
                "headerToolbar": {"left": "prev,next today", "center": "title", "right": "dayGridMonth,timeGridWeek,listMonth"},
                "initialView": "dayGridMonth",
            }
            st.subheader(f"📌 {selected_bid_file} 일정 달력 (총 {len(events)}건 검출)")
            if events:
                calendar(events=events, options=calendar_options, key=f"cal_{selected_bid_file}")
            else:
                st.warning("⚠️ 엑셀 파일 내에서 날짜 데이터를 추출하지 못했습니다. 파일 구조를 확인해 주세요.")
                with st.expander("원본 데이터 미리보기"):
                    st.dataframe(df_raw.head(20))


# =========================================================
# [TAB 2] 경력기술자 검색
# =========================================================
with tab2:
    st.header("👷 경력기술자 공종 및 경력일수 조건 검색")
    col_up2, col_sel2 = st.columns([1, 1])

    with col_up2:
        uploaded_eng = st.file_uploader("새 경력 엑셀 파일 업로드", type=["xlsx", "xls"], key="eng_uploader")
        if uploaded_eng is not None:
            save_uploaded_file(uploaded_eng, ENG_DIR)
            st.success(f"✅ '{uploaded_eng.name}' 업로드 완료!")
            st.cache_data.clear()
            st.session_state["eng_select"] = uploaded_eng.name
            st.rerun()

    saved_eng_files = get_saved_files(ENG_DIR)
    with col_sel2:
        if saved_eng_files:
            if "eng_select" not in st.session_state or st.session_state["eng_select"] not in saved_eng_files:
                st.session_state["eng_select"] = saved_eng_files[0]
            selected_eng_file = st.selectbox("📁 경력 파일 선택", saved_eng_files, key="eng_select")
            if st.button("🗑️ 파일 삭제", key="del_eng"):
                if delete_saved_file(os.path.join(ENG_DIR, selected_eng_file)):
                    st.cache_data.clear()
                    if "eng_select" in st.session_state: del st.session_state["eng_select"]
                    st.rerun()
        else:
            selected_eng_file = None
            st.info("💡 경력 엑셀 파일을 업로드해 주세요.")

    st.markdown("---")
    if selected_eng_file and selected_eng_file in saved_eng_files:
        df_engineer = load_and_clean_engineer_data(os.path.join(ENG_DIR, selected_eng_file))
        c1, c2, c3, c4 = st.columns([1.5, 2, 1.5, 1.5])
        with c1: name_search = st.text_input("이름 검색", "", key="eng_name")
        with c2: type_search = st.text_input("공사종류 검색", "", key="eng_type_search")
        with c3: min_days = st.number_input("최소 인정일수", min_value=0, value=0, step=30, key="eng_days")
        with c4: duty_search = st.text_input("담당업무 검색", "", key="eng_duty")

        filtered_df = df_engineer.copy()
        if type_search.strip(): filtered_df = filtered_df[filtered_df["공사종류"].astype(str).str.contains(type_search.strip(), na=False)]
        if duty_search.strip(): filtered_df = filtered_df[filtered_df["담당업무"].astype(str).str.contains(duty_search.strip(), na=False)]
        if name_search.strip(): filtered_df = filtered_df[filtered_df["이름"].astype(str).str.contains(name_search.strip(), na=False)]
        if min_days > 0:
            person_days = filtered_df.groupby("이름")["인정일수"].sum().reset_index()
            filtered_df = filtered_df[filtered_df["이름"].isin(person_days[person_days["인정일수"] >= min_days]["이름"])]

        if not filtered_df.empty:
            summary_df = filtered_df.groupby("이름").agg(건수=("사업명", "count"), 총인정일수=("인정일수", "sum")).reset_index()
            summary_df["추정경력년수"] = summary_df["총인정일수"].apply(lambda d: f"{int(d // 365)}년 {int((d % 365) // 30)}개월")
            st.dataframe(summary_df, use_container_width=True)
            st.dataframe(filtered_df, use_container_width=True)
        else:
            st.warning("조건에 맞는 결과가 없습니다.")


# =========================================================
# [TAB 3] 준공실적증명 관리
# =========================================================
with tab3:
    st.header("📜 준공실적증명 목록 및 조회")
    col_up3, col_sel3 = st.columns([1, 1])

    with col_up3:
        uploaded_perf = st.file_uploader("새 실적 엑셀 파일 업로드", type=["xlsx", "xls"], key="perf_uploader")
        if uploaded_perf is not None:
            save_uploaded_file(uploaded_perf, PERF_DIR)
            st.success(f"✅ '{uploaded_perf.name}' 업로드 완료!")
            st.cache_data.clear()
            st.session_state["perf_select"] = uploaded_perf.name
            st.rerun()

    saved_perf_files = get_saved_files(PERF_DIR)
    with col_sel3:
        if saved_perf_files:
            if "perf_select" not in st.session_state or st.session_state["perf_select"] not in saved_perf_files:
                st.session_state["perf_select"] = saved_perf_files[0]
            selected_perf_file = st.selectbox("📁 실적 파일 선택", saved_perf_files, key="perf_select")
            if st.button("🗑️ 파일 삭제", key="del_perf"):
                if delete_saved_file(os.path.join(PERF_DIR, selected_perf_file)):
                    st.cache_data.clear()
                    if "perf_select" in st.session_state: del st.session_state["perf_select"]
                    st.rerun()
        else:
            selected_perf_file = None
            st.info("💡 준공실적 엑셀 파일을 업로드해 주세요.")

    st.markdown("---")
    if selected_perf_file and selected_perf_file in saved_perf_files:
        df_perf = load_and_clean_perf_data(os.path.join(PERF_DIR, selected_perf_file))
        perf_search = st.text_input("🔍 실적 검색", "", key="perf_search_kw")
        if perf_search.strip():
            df_perf = df_perf[df_perf.astype(str).apply(lambda row: row.str.contains(perf_search.strip(), na=False).any(), axis=1)]
        st.dataframe(df_perf, use_container_width=True)
