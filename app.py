from datetime import datetime
import os
import re
import pandas as pd
import streamlit as st
from streamlit_calendar import calendar

# 페이지 기본 설정
st.set_page_config(
    page_title="건설 입찰 및 경력/실적/계약 통합 관리 시스템",
    layout="wide"
)

# 폴더 설정
BASE_DATA_DIR = "uploaded_data"
BID_DIR = os.path.join(BASE_DATA_DIR, "bids")
ENG_DIR = os.path.join(BASE_DATA_DIR, "engineer")
PERF_DIR = os.path.join(BASE_DATA_DIR, "performance")
CONTRACT_DIR = os.path.join(BASE_DATA_DIR, "contract")

for d in [BID_DIR, ENG_DIR, PERF_DIR, CONTRACT_DIR]:
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


# 🌟 테이블 출력 전용 데이터 포맷팅 함수 (날짜 시간 제거, 숫자 쉼표, None 빈칸 처리)
def clean_display_dataframe(df):
    df_formatted = df.copy()
    for col in df_formatted.columns:
        col_str = str(col)
        # 1. 날짜 관련 컬럼 (시간 제거)
        if "일" in col_str or "일자" in col_str or pd.api.types.is_datetime64_any_dtype(df_formatted[col]):
            df_formatted[col] = pd.to_datetime(df_formatted[col], errors="coerce").dt.strftime("%Y-%m-%d")
            df_formatted[col] = df_formatted[col].fillna("")
        else:
            # 2. 숫자형 데이터 컬럼 (1,000단위 쉼표 추가)
            numeric_series = pd.to_numeric(df_formatted[col], errors="coerce")
            if numeric_series.notna().sum() > 0 and not any(k in col_str for k in ["연도", "년도", "차수"]):
                df_formatted[col] = numeric_series.apply(
                    lambda x: f"{x:,.0f}" if pd.notna(x) and x == int(x) else (f"{x:,.2f}" if pd.notna(x) else "")
                )
            else:
                # 3. 텍스트 컬럼의 None, NaN 등 -> 완벽한 빈칸("")으로 치환
                df_formatted[col] = df_formatted[col].astype(str).str.replace(r"(?i)^(none|nan|nat)$", "", regex=True).str.strip()
                df_formatted[col] = df_formatted[col].replace({"nan": "", "None": "", "NONE": "", "": ""})
    return df_formatted.fillna("")


# ---------------------------------------------------------
# OZ Report 입찰명 완벽 일치 파서
# ---------------------------------------------------------
def parse_oz_report_4schedules(file_path):
    raw_df = pd.read_excel(file_path, header=None)

    pq_col = 4
    agreement_col = 7
    reg_col = 9
    bid_col = 10
    year_col = 0
    title_col = 1

    for idx, row in raw_df.iterrows():
        row_cells = [str(c).replace(" ", "").replace("\n", "").upper() for c in row]
        for c_i, cell in enumerate(row_cells):
            if "PQ" in cell or "실적" in cell:
                pq_col = c_i
            elif "협정" in cell:
                agreement_col = c_i
            elif "등록" in cell and "입찰" not in cell:
                reg_col = c_i
            elif "입찰" in cell and ("마감" in cell or "일" in cell):
                bid_col = c_i
            elif "공사명" in cell:
                title_col = c_i
            elif "년도" in cell:
                year_col = c_i

    parsed_events = []

    for idx in range(len(raw_df)):
        row = raw_df.iloc[idx]
        if len(row) <= title_col or pd.isna(row.iloc[title_col]):
            continue

        raw_title = str(row.iloc[title_col]).strip()
        if "공사명" in raw_title or "입찰일정" in raw_title or "년도" in raw_title or "페이지" in raw_title or len(raw_title) < 2:
            continue

        title_lines = [line.strip() for line in raw_title.split("\n") if line.strip() and not line.strip().isdigit()]
        if not title_lines:
            continue
        clean_title = title_lines[0]

        row_year = datetime.now().year
        if len(row) > year_col and pd.notna(row.iloc[year_col]):
            y_str = str(row.iloc[year_col]).strip()
            m_year = re.search(r"202\d", y_str)
            if m_year:
                row_year = int(m_year.group())

        schedules = [
            ("PQ", pq_col, "#E1BEE7"),
            ("협정", agreement_col, "#C8E6C9"),
            ("등록", reg_col, "#FFE0B2"),
            ("입찰", bid_col, "#BBDEFB"),
        ]

        for cat_label, c_idx, color in schedules:
            if len(row) > c_idx and pd.notna(row.iloc[c_idx]):
                val_str = str(row.iloc[c_idx]).strip()
                m_date = re.search(r"(\d{1,2})[/.-](\d{1,2})", val_str)
                if m_date:
                    month = int(m_date.group(1))
                    day = int(m_date.group(2))
                    if 1 <= month <= 12 and 1 <= day <= 31:
                        date_str = f"{row_year}-{month:02d}-{day:02d}"
                        parsed_events.append({
                            "title": f"{cat_label}_ {clean_title}",
                            "start": date_str,
                            "end": date_str,
                            "backgroundColor": color,
                            "borderColor": color,
                            "textColor": "#1A1A1A",
                        })

    return parsed_events


st.title("🏗️ 건설 입찰 및 경력/실적/계약 통합 관리 시스템")

tab1, tab2, tab3, tab4 = st.tabs(["📅 입찰 달력", "👷 경력기술자 조건 검색", "🏢 준공실적 검색", "🏗️ 계약 관리"])

# ---------------------------------------------------------
# [TAB 1] 입찰 달력
# ---------------------------------------------------------
with tab1:
    st.header("📅 입찰 일정 달력")
    col_up, col_sel = st.columns([1, 1])

    with col_up:
        uploaded_bid = st.file_uploader("OZ Report 입찰 일정 엑셀(.xlsx) 업로드", type=["xlsx"], key="bid_uploader")
        if uploaded_bid is not None:
            save_uploaded_file(uploaded_bid, BID_DIR)
            st.success(f"✅ '{uploaded_bid.name}' 업로드 완료!")
            st.cache_data.clear()
            for key in ["bid_select", "bid_calendar"]:
                if key in st.session_state:
                    del st.session_state[key]
            st.rerun()

    saved_bid_files = get_saved_files(BID_DIR)

    with col_sel:
        if saved_bid_files:
            selected_bid_file = st.selectbox("📁 불러올 입찰 파일 선택", saved_bid_files, key="bid_select")
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

    calendar_options = {
        "headerToolbar": {"left": "prev,next today", "center": "title", "right": "dayGridMonth,dayGridWeek"},
        "initialView": "dayGridMonth",
        "locale": "ko",
        "firstDay": 0,
        "height": 650,
        "selectable": True,
        "editable": False,
        "dayMaxEvents": 3,
    }
    calendar(events=bid_events, options=calendar_options, key="bid_calendar")

# ---------------------------------------------------------
# [TAB 2] 경력기술자 검색
# ---------------------------------------------------------
with tab2:
    st.header("👷 경력기술자 공종 및 경력일수 조건 검색")
    col_up, col_sel = st.columns([1, 1])
    with col_up:
        uploaded_eng = st.file_uploader("새 경력 엑셀 파일(.xlsx) 업로드", type=["xlsx"], key="eng_uploader")
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
            selected_eng_file = st.selectbox("📁 경력 파일 목록 선택", saved_eng_files, key="eng_select")
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
            c1, c2, c3, c4 = st.columns([1.5, 2, 1.5, 1.5])
            with c1:
                name_search = st.text_input("이름 검색", "", key="eng_name")
            with c2:
                type_search = st.text_input("공사종류 검색", "", key="eng_type_search")
            with c3:
                min_days = st.number_input("최소 인정일수", min_value=0, value=0, key="eng_days")
            with c4:
                duty_search = st.text_input("담당업무", "", key="eng_duty")

            filtered_df = df_engineer.copy()
            if "인정일수" in filtered_df.columns:
                filtered_df["인정일수"] = pd.to_numeric(filtered_df["인정일수"], errors="coerce").fillna(0)
            else:
                filtered_df["인정일수"] = 0

            if type_search.strip() and "공사종류" in filtered_df.columns:
                filtered_df = filtered_df[filtered_df["공사종류"].astype(str).str.contains(type_search.strip(), na=False)]
            if duty_search.strip() and "담당업무" in filtered_df.columns:
                filtered_df = filtered_df[filtered_df["담당업무"].astype(str).str.contains(duty_search.strip(), na=False)]
            if name_search.strip() and "이름" in filtered_df.columns:
                filtered_df = filtered_df[filtered_df["이름"].astype(str).str.contains(name_search.strip(), na=False)]
            if min_days > 0 and "이름" in filtered_df.columns:
                person_days = filtered_df.groupby("이름")["인정일수"].sum().reset_index()
                target_persons = person_days[person_days["인정일수"] >= min_days]["이름"]
                filtered_df = filtered_df[filtered_df["이름"].isin(target_persons)]

            st.markdown("---")
            if not filtered_df.empty and "이름" in filtered_df.columns:
                summary_df = filtered_df.groupby("이름").agg(
                    건수=("사업명", "count"),
                    총인정일수=("인정일수", "sum")
                ).reset_index()
                summary_df["추정경력년수"] = summary_df["총인정일수"].apply(
                    lambda d: f"{int(d // 365)}년 {int((d % 365) // 30)}개월 ({int(d)}일)"
                )
                st.dataframe(summary_df, use_container_width=True)
                st.dataframe(filtered_df, use_container_width=True)

# ---------------------------------------------------------
# [TAB 3] 준공실적 검색
# ---------------------------------------------------------
with tab3:
    st.header("🏢 준공실적 조건 검색")
    col_up2, col_sel2 = st.columns([1, 1])
    with col_up2:
        uploaded_perf = st.file_uploader("새 실적 엑셀 파일(.xlsx) 업로드", type=["xlsx"], key="perf_uploader")
        if uploaded_perf is not None:
            save_uploaded_file(uploaded_perf, PERF_DIR)
            st.rerun()

    saved_perf_files = get_saved_files(PERF_DIR)
    with col_sel2:
        if saved_perf_files:
            selected_perf_file = st.selectbox("실적 파일 선택", saved_perf_files, key="perf_select")
        else:
            selected_perf_file = None

    if selected_perf_file and selected_perf_file in saved_perf_files:
        file_path = os.path.join(PERF_DIR, selected_perf_file)
        if os.path.exists(file_path):
            df_perf = pd.read_excel(file_path)
            min_amount = st.number_input("최저 실적금액 필터", value=0, step=100)
            filtered_perf = df_perf.copy()
            if "금액" in filtered_perf.columns:
                filtered_perf = filtered_perf[pd.to_numeric(filtered_perf["금액"], errors="coerce").fillna(0) >= min_amount]
            st.dataframe(filtered_perf, use_container_width=True)

# ---------------------------------------------------------
# [TAB 4] 계약 관리 (낙찰, 총괄만 합산 및 차수 잔여계약금액 계산 로직 적용)
# ---------------------------------------------------------
with tab4:
    st.header("🏗️ 현장 계약 및 변경 이력 관리 (계약입력 시트 연동)")

    col_up_c, col_sel_c = st.columns([1, 1])
    with col_up_c:
        uploaded_contract = st.file_uploader("계약관리 엑셀 파일(.xlsx) 업로드", type=["xlsx"], key="contract_uploader")
        if uploaded_contract is not None:
            save_uploaded_file(uploaded_contract, CONTRACT_DIR)
            st.success(f"✅ '{uploaded_contract.name}' 계약 파일 저장 완료!")
            st.cache_data.clear()
            if "contract_select" in st.session_state:
                del st.session_state["contract_select"]
            st.rerun()

    saved_contract_files = get_saved_files(CONTRACT_DIR)
    with col_sel_c:
        if saved_contract_files:
            selected_contract_file = st.selectbox("📁 계약 파일 목록 선택", saved_contract_files, key="contract_select")
            if st.button("🗑️ 선택한 계약 파일 삭제", key="del_contract"):
                file_path_to_del = os.path.join(CONTRACT_DIR, selected_contract_file)
                if delete_saved_file(file_path_to_del):
                    st.cache_data.clear()
                    for key in ["contract_select", "contract_uploader"]:
                        if key in st.session_state:
                            del st.session_state[key]
                    st.rerun()
        else:
            selected_contract_file = None
            st.info("💡 왼쪽에 계약관리 엑셀 파일을 업로드해 주세요.")

    st.markdown("---")

    if selected_contract_file and selected_contract_file in saved_contract_files:
        file_path = os.path.join(CONTRACT_DIR, selected_contract_file)
        if os.path.exists(file_path):
            try:
                df_contract = pd.read_excel(file_path, sheet_name="계약입력")
            except Exception as e:
                st.error(f"시트 읽기 오류: {e}")
                df_contract = None

            if df_contract is not None:
                df_contract.columns = df_contract.columns.str.strip()

                if "계약명" in df_contract.columns:
                    contract_list = df_contract["계약명"].dropna().unique().tolist()
                    selected_contract = st.selectbox("🎯 조회할 계약(공사)을 선택하세요:", contract_list, key="selected_contract_box")

                    if selected_contract:
                        sub_df = df_contract[df_contract["계약명"] == selected_contract].copy()

                        first_row = sub_df.iloc[0]
                        client = first_row["발주처"] if "발주처" in df_contract.columns and pd.notna(first_row["발주처"]) else "-"

                        share_ratio = 1.0
                        if "지분율" in df_contract.columns and pd.notna(first_row["지분율"]):
                            try:
                                val = str(first_row["지분율"]).replace("%", "").strip()
                                share_ratio = float(val)
                                if share_ratio > 1:
                                    share_ratio = share_ratio / 100.0
                            except:
                                share_ratio = 1.0

                        # 🌟 '종류'가 '낙찰' 또는 '총괄'인 행들만 골라서 총공사 계약금액 및 당사 계약금액 합산
                        valid_sum_mask = sub_df["종류"].astype(str).str.contains("낙찰|총괄", na=False)
                        filtered_sum_df = sub_df[valid_sum_mask]

                        total_contract_amount = 0
                        my_contract_amount = 0

                        if "낙찰(계약)금액" in filtered_sum_df.columns:
                            total_contract_amount = pd.to_numeric(filtered_sum_df["낙찰(계약)금액"], errors="coerce").sum()

                        my_col = None
                        for col in sub_df.columns:
                            if "당사 낙찰금액" in col or "당사낙찰금액" in col:
                                my_col = col
                                break

                        if my_col and my_col in filtered_sum_df.columns:
                            my_contract_amount = pd.to_numeric(filtered_sum_df[my_col], errors="coerce").sum()

                        # 상단 요약 카드 출력
                        st.markdown(f"### 📌 [{client}] {selected_contract}")

                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            st.metric("총 공사 계약금액 (낙찰/총괄)", f"{total_contract_amount:,.0f} 원")
                        with col2:
                            st.metric("당사 지분율", f"{share_ratio * 100:.1f}%")
                        with col3:
                            st.metric("당사 총 계약금액 (낙찰/총괄)", f"{my_contract_amount:,.0f} 원")
                        with col4:
                            st.metric("변경/증감 이력", f"{len(sub_df) - 1} 건")

                        st.markdown("---")

                        # 🌟 차수 계약의 잔여계약금액(전체계약금액 - 누적기계약금액) 계산 컬럼 추가 로직
                        sub_df_calc = sub_df.copy()
                        if "낙찰(계약)금액" in sub_df_calc.columns:
                            remaining_amounts = []
                            accumulated_sum = 0
                            for idx, row in sub_df_calc.iterrows():
                                kind = str(row.get("종류", ""))
                                amt = pd.to_numeric(row.get("낙찰(계약)금액", 0), errors="coerce")
                                if pd.isna(amt):
                                    amt = 0
                                
                                if "차수" in kind and "변경" not in kind:
                                    accumulated_sum += amt
                                    rem = total_contract_amount - accumulated_sum
                                    remaining_amounts.append(rem)
                                else:
                                    remaining_amounts.append(None)
                            
                            sub_df_calc["잔여계약금액"] = remaining_amounts

                        # 포맷팅 적용
                        clean_sub_df = clean_display_dataframe(sub_df_calc)

                        st.markdown("#### 📋 상세 계약 및 변경 내역 (타임라인)")
                        display_cols = [
                            "종류",
                            "계약일(낙찰)",
                            "내용",
                            my_col if my_col else "당사 낙찰금액\n(부가세포함)",
                            "보증금액(당사)",
                            "국민주택채권(당사)",
                            "잔여계약금액",
                            "특이사항",
                            "낙찰(계약)금액",
                            "예정(기초)가격",
                            "낙찰율",
                            "착공일",
                            "준공일",
                        ]
                        actual_cols = [c for c in display_cols if c in clean_sub_df.columns]
                        st.dataframe(clean_sub_df[actual_cols], use_container_width=True)

                        # 차수별 요약 브리핑
                        st.markdown("#### 🔍 차수 및 증감 계약 요약 브리핑")
                        summary_view_cols = [
                            c for c in ["종류", "내용", "계약일(낙찰)", my_col, "잔여계약금액", "특이사항"]
                            if c in clean_sub_df.columns
                        ]
                        if summary_view_cols:
                            st.table(clean_sub_df[summary_view_cols])
                else:
                    st.error("엑셀 파일에 '계약명' 컬럼이 없습니다.")
