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


def parse_oz_date(val):
    """OZ Report 날짜 및 시간 텍스트 파싱 유틸리티"""
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
# [TAB 1: OZ Report 입찰 일정 파서 - 원본 완벽 복원]
# ---------------------------------------------------------
def parse_bid_excel(file_path):
    """
    OZ Report 엑셀 파싱 (Col 1: 공사명, Col 2: 발주처, Col 4: PQ, Col 7: 협정, Col 9: 등록, Col 10: 입찰)
    """
    df_raw = pd.read_excel(file_path, header=None)
    events = []
    
    # 1. '공사명' 또는 '입찰공고' 헤더 시작 위치 탐색
    start_row = 0
    for idx, row in df_raw.iterrows():
        row_str = " ".join(row.dropna().astype(str))
        if "공사명" in row_str or "입찰공고" in row_str:
            start_row = idx + 1
            break
            
    # 2. 데이터 행 파싱
    for idx in range(start_row, len(df_raw)):
        row = df_raw.iloc[idx]
        
        # Col 1 (공사명) 확인
        title = row.iloc[1] if len(row) > 1 else None
        
        if pd.isna(title) or not str(title).strip() or "합계" in str(title):
            continue
            
        title = str(title).strip()
        client = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else "발주처 미상"
        
        # OZ Report 지정 열 정확히 조준
        pq_date = parse_oz_date(row.iloc[4]) if len(row) > 4 else None
        joint_date = parse_oz_date(row.iloc[7]) if len(row) > 7 else None
        reg_date = parse_oz_date(row.iloc[9]) if len(row) > 9 else None
        bid_date = parse_oz_date(row.iloc[10]) if len(row) > 10 else None

        # 이벤트 등록
        if pq_date:
            events.append({
                "title": f"[PQ마감] {title} ({client})",
                "start": pq_date,
                "color": "#FF6B6B",
                "extendedProps": {"type": "PQ마감", "client": client, "full_title": title}
            })
        if joint_date:
            events.append({
                "title": f"[협정마감] {title} ({client})",
                "start": joint_date,
                "color": "#4ECDC4",
                "extendedProps": {"type": "협정마감", "client": client, "full_title": title}
            })
        if reg_date:
            events.append({
                "title": f"[등록마감] {title} ({client})",
                "start": reg_date,
                "color": "#FFE66D",
                "extendedProps": {"type": "등록마감", "client": client, "full_title": title}
            })
        if bid_date:
            events.append({
                "title": f"[입찰마감] {title} ({client})",
                "start": bid_date,
                "color": "#1A535C",
                "extendedProps": {"type": "입찰마감", "client": client, "full_title": title}
            })
            
    return events, df_raw


# ---------------------------------------------------------
# [TAB 2: 경력기술자 파싱 및 데이터 정제]
# ---------------------------------------------------------
def load_and_clean_engineer_data(file_path):
    """경력기술자 엑셀 파싱 및 헤더/컬럼 매핑"""
    df_raw = pd.read_excel(file_path, header=None)
    
    header_idx = None
    for idx, row in df_raw.iterrows():
        row_str = " ".join(row.dropna().astype(str))
        if any(k in row_str for k in ["성명", "이름", "공사종류", "인정일수", "담당업무", "사업명", "공사명"]):
            header_idx = idx
            break
            
    if header_idx is not None:
        df = pd.read_excel(file_path, header=header_idx)
    else:
        df = pd.read_excel(file_path)

    df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed")]
    
    col_mapping = {}
    for col in df.columns:
        c_str = str(col).replace(" ", "").strip()
        if any(x in c_str for x in ["성명", "이름", "기술자명"]):
            col_mapping[col] = "이름"
        elif any(x in c_str for x in ["공사종류", "공종", "전문분야", "사업분야"]):
            col_mapping[col] = "공사종류"
        elif any(x in c_str for x in ["인정일수", "경력일수", "참여일수", "일수"]):
            col_mapping[col] = "인정일수"
        elif any(x in c_str for x in ["담당업무", "직무", "직책", "수행업무"]):
            col_mapping[col] = "담당업무"
        elif any(x in c_str for x in ["사업명", "공사명", "프로젝트명"]):
            col_mapping[col] = "사업명"

    df = df.rename(columns=col_mapping)
    
    if "이름" not in df.columns: df["이름"] = "미상"
    if "공사종류" not in df.columns: df["공사종류"] = "-"
    if "담당업무" not in df.columns: df["담당업무"] = "-"
    if "사업명" not in df.columns: df["사업명"] = "사업명 미상"
        
    if "인정일수" in df.columns:
        df["인정일수"] = pd.to_numeric(df["인정일수"], errors="coerce").fillna(0)
    else:
        df["인정일수"] = 0

    df = df.dropna(subset=["이름"])
    df = df[df["이름"].astype(str).str.strip() != ""]
    return df


# ---------------------------------------------------------
# [TAB 3: 준공실적증명 파싱]
# ---------------------------------------------------------
def load_and_clean_perf_data(file_path):
    """준공실적증명 엑셀 파싱"""
    df_raw = pd.read_excel(file_path, header=None)
    
    header_idx = None
    for idx, row in df_raw.iterrows():
        row_str = " ".join(row.dropna().astype(str))
        if any(k in row_str for k in ["공사명", "발주처", "계약금액", "준공일자", "공사종류"]):
            header_idx = idx
            break
            
    if header_idx is not None:
        df = pd.read_excel(file_path, header=header_idx)
    else:
        df = pd.read_excel(file_path)

    df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed")]
    return df


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
        uploaded_bid = st.file_uploader(
            "새 입찰 일정 엑셀 파일(.xlsx) 업로드", type=["xlsx", "xls"], key="bid_uploader"
        )
        if uploaded_bid is not None:
            save_uploaded_file(uploaded_bid, BID_DIR)
            st.success(f"✅ '{uploaded_bid.name}' 입찰 파일 저장 완료!")
            st.cache_data.clear()

            # 업로드 직후 즉시 세션 반영 및 재실행
            st.session_state["bid_select"] = uploaded_bid.name
            st.rerun()

    saved_bid_files = get_saved_files(BID_DIR)

    with col_sel1:
        if saved_bid_files:
            if (
                "bid_select" not in st.session_state
                or st.session_state["bid_select"] not in saved_bid_files
            ):
                st.session_state["bid_select"] = saved_bid_files[0]

            selected_bid_file = st.selectbox(
                "📁 입찰 파일 목록 선택",
                saved_bid_files,
                key="bid_select",
            )

            if st.button("🗑️ 선택한 입찰 파일 삭제", key="del_bid"):
                file_path_to_del = os.path.join(BID_DIR, selected_bid_file)
                if delete_saved_file(file_path_to_del):
                    st.cache_data.clear()
                    if "bid_select" in st.session_state:
                        del st.session_state["bid_select"]
                    st.success(f"🗑️ '{selected_bid_file}' 파일이 삭제되었습니다.")
                    st.rerun()
        else:
            selected_bid_file = None
            st.info("💡 왼쪽에 입찰 일정 엑셀 파일을 업로드해 주세요.")

    st.markdown("---")

    if selected_bid_file and selected_bid_file in saved_bid_files:
        file_path = os.path.join(BID_DIR, selected_bid_file)
        if os.path.exists(file_path):
            events, df_raw = parse_bid_excel(file_path)

            calendar_options = {
                "editable": False,
                "selectable": True,
                "headerToolbar": {
                    "left": "prev,next today",
                    "center": "title",
                    "right": "dayGridMonth,timeGridWeek,listMonth",
                },
                "initialView": "dayGridMonth",
            }

            st.subheader(f"📌 {selected_bid_file} 일정 달력 (총 {len(events)}건 일정 표출)")
            calendar(events=events, options=calendar_options, key=f"cal_{selected_bid_file}")


# =========================================================
# [TAB 2] 경력기술자 검색
# =========================================================
with tab2:
    st.header("👷 경력기술자 공종 및 경력일수 조건 검색")
    col_up2, col_sel2 = st.columns([1, 1])

    with col_up2:
        uploaded_eng = st.file_uploader(
            "새 경력 엑셀 파일(.xlsx) 업로드", type=["xlsx", "xls"], key="eng_uploader"
        )
        if uploaded_eng is not None:
            save_uploaded_file(uploaded_eng, ENG_DIR)
            st.success(f"✅ '{uploaded_eng.name}' 경력 파일 저장 완료!")
            st.cache_data.clear()

            # 업로드 직후 즉시 세션 반영하여 새로고침 없이 사람 바로 표출
            st.session_state["eng_select"] = uploaded_eng.name
            st.rerun()

    saved_eng_files = get_saved_files(ENG_DIR)

    with col_sel2:
        if saved_eng_files:
            if (
                "eng_select" not in st.session_state
                or st.session_state["eng_select"] not in saved_eng_files
            ):
                st.session_state["eng_select"] = saved_eng_files[0]

            selected_eng_file = st.selectbox(
                "📁 경력 파일 목록 선택",
                saved_eng_files,
                key="eng_select",
            )

            if st.button("🗑️ 선택한 경력 파일 삭제", key="del_eng"):
                file_path_to_del = os.path.join(ENG_DIR, selected_eng_file)
                if delete_saved_file(file_path_to_del):
                    st.cache_data.clear()
                    if "eng_select" in st.session_state:
                        del st.session_state["eng_select"]
                    st.success(f"🗑️ '{selected_eng_file}' 파일이 삭제되었습니다.")
                    st.rerun()
        else:
            selected_eng_file = None
            st.info("💡 왼쪽에 경력 엑셀 파일을 업로드해 주세요.")

    st.markdown("---")

    if selected_eng_file and selected_eng_file in saved_eng_files:
        file_path = os.path.join(ENG_DIR, selected_eng_file)
        if os.path.exists(file_path):
            df_engineer = load_and_clean_engineer_data(file_path)

            st.subheader(
                f"🎯 조건 설정 (대상 파일: {selected_eng_file} / 총 {len(df_engineer)}건 레코드)"
            )
            c1, c2, c3, c4 = st.columns([1.5, 2, 1.5, 1.5])

            with c1:
                name_search = st.text_input("이름 검색 (부분 검색 가능)", "", key="eng_name")
            with c2:
                type_search = st.text_input("🎯 공사종류 검색 (예: 고속도로, 준설, 교량)", value="", key="eng_type_search")
            with c3:
                min_days = st.number_input("⏳ 최소 인정일수 (일 기준)", min_value=0, value=0, step=30, key="eng_days")
            with c4:
                duty_search = st.text_input("담당업무 (예: 시공, 대리 등)", "", key="eng_duty")

            filtered_df = df_engineer.copy()

            if type_search.strip():
                filtered_df = filtered_df[filtered_df["공사종류"].astype(str).str.contains(type_search.strip(), na=False)]
            if duty_search.strip():
                filtered_df = filtered_df[filtered_df["담당업무"].astype(str).str.contains(duty_search.strip(), na=False)]
            if name_search.strip():
                filtered_df = filtered_df[filtered_df["이름"].astype(str).str.contains(name_search.strip(), na=False)]
            if min_days > 0:
                person_days = filtered_df.groupby("이름")["인정일수"].sum().reset_index()
                target_persons = person_days[person_days["인정일수"] >= min_days]["이름"]
                filtered_df = filtered_df[filtered_df["이름"].isin(target_persons)]

            st.markdown("---")

            if not filtered_df.empty:
                project_col = "사업명" if "사업명" in filtered_df.columns else filtered_df.columns[0]
                
                summary_df = (
                    filtered_df.groupby("이름")
                    .agg(
                        건수=(project_col, "count"),
                        총인정일수=("인정일수", "sum"),
                    )
                    .reset_index()
                )
                summary_df["추정경력년수"] = summary_df["총인정일수"].apply(
                    lambda d: f"{int(d // 365)}년 {int((d % 365) // 30)}개월 ({int(d)}일)"
                )

                st.subheader(f"🎯 조건 충족 적합 인원: 총 {len(summary_df)}명 (상세 이력 {len(filtered_df)}건)")
                st.markdown("##### 📌 적합 기술자 요약 목록")
                st.dataframe(summary_df, use_container_width=True)
                
                st.markdown("##### 📄 적합 기술자의 상세 경력 내역")
                st.dataframe(filtered_df, use_container_width=True)
            else:
                st.warning("⚠️ 설정한 조건에 맞는 기술자 경력이 없습니다.")


# =========================================================
# [TAB 3] 준공실적증명 관리
# =========================================================
with tab3:
    st.header("📜 준공실적증명 목록 및 조회")
    col_up3, col_sel3 = st.columns([1, 1])

    with col_up3:
        uploaded_perf = st.file_uploader(
            "새 준공실적 엑셀 파일(.xlsx) 업로드", type=["xlsx", "xls"], key="perf_uploader"
        )
        if uploaded_perf is not None:
            save_uploaded_file(uploaded_perf, PERF_DIR)
            st.success(f"✅ '{uploaded_perf.name}' 실적 파일 저장 완료!")
            st.cache_data.clear()

            # 업로드 직후 즉시 세션 반영 및 재실행
            st.session_state["perf_select"] = uploaded_perf.name
            st.rerun()

    saved_perf_files = get_saved_files(PERF_DIR)

    with col_sel3:
        if saved_perf_files:
            if (
                "perf_select" not in st.session_state
                or st.session_state["perf_select"] not in saved_perf_files
            ):
                st.session_state["perf_select"] = saved_perf_files[0]

            selected_perf_file = st.selectbox(
                "📁 준공실적 파일 목록 선택",
                saved_perf_files,
                key="perf_select",
            )

            if st.button("🗑️ 선택한 실적 파일 삭제", key="del_perf"):
                file_path_to_del = os.path.join(PERF_DIR, selected_perf_file)
                if delete_saved_file(file_path_to_del):
                    st.cache_data.clear()
                    if "perf_select" in st.session_state:
                        del st.session_state["perf_select"]
                    st.success(f"🗑️ '{selected_perf_file}' 파일이 삭제되었습니다.")
                    st.rerun()
        else:
            selected_perf_file = None
            st.info("💡 왼쪽에 준공실적 엑셀 파일을 업로드해 주세요.")

    st.markdown("---")

    if selected_perf_file and selected_perf_file in saved_perf_files:
        file_path = os.path.join(PERF_DIR, selected_perf_file)
        if os.path.exists(file_path):
            df_perf = load_and_clean_perf_data(file_path)

            st.subheader(f"📋 준공실적 상세 목록 ({selected_perf_file} / 총 {len(df_perf)}건)")
            perf_search = st.text_input("🔍 실적 검색 (공사명, 발주처 등)", "", key="perf_search_kw")

            filtered_perf = df_perf.copy()
            if perf_search.strip():
                mask = filtered_perf.astype(str).apply(
                    lambda row: row.str.contains(perf_search.strip(), na=False).any(), axis=1
                )
                filtered_perf = filtered_perf[mask]

            st.dataframe(filtered_perf, use_container_width=True)
