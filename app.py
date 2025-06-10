import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")

tabs = st.tabs([
    "사업장개요",
    "근골격계 부담작업 체크리스트",
    "유해요인조사표",
    "작업조건조사"
])

with tabs[0]:
    st.title("사업장 개요")
    사업장명 = st.text_input("사업장명")
    소재지 = st.text_input("소재지")
    업종 = st.text_input("업종")
    col1, col2 = st.columns(2)
    with col1:
        예비조사 = st.date_input("예비조사일")
        수행기관 = st.text_input("수행기관")
    with col2:
        본조사 = st.date_input("본조사일")
        성명 = st.text_input("성명")

with tabs[1]:
    st.subheader("근골격계 부담작업 체크리스트")
    columns = [
        "작업명", "단위작업명"
    ] + [f"{i}호" for i in range(1, 12)]
    data = pd.DataFrame(
        columns=columns,
        data=[["", ""] + ["X(미해당)"]*11 for _ in range(5)]
    )

    ho_options = [
        "O(해당)",
        "△(잠재위험)",
        "X(미해당)"
    ]
    column_config = {
        f"{i}호": st.column_config.SelectboxColumn(
            f"{i}호", options=ho_options, required=True
        ) for i in range(1, 12)
    }
    column_config["작업명"] = st.column_config.TextColumn("작업명")
    column_config["단위작업명"] = st.column_config.TextColumn("단위작업명")

        edited_df = st.data_editor(
        data,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config=column_config
    )
    # 반드시 session_state에 저장!
    st.session_state["checklist_df"] = edited_df

with tabs[2]:
    st.title("유해요인조사표")
    st.markdown("#### 가. 조사개요")
    col1, col2 = st.columns(2)
    with col1:
        조사일시 = st.text_input("조사일시")
        부서명 = st.text_input("부서명")
    with col2:
        조사자 = st.text_input("조사자")
        작업공정명 = st.text_input("작업공정명")
    작업명 = st.text_input("작업명", key="tab2_작업명")

    st.markdown("#### 나. 작업장 상황조사")

    def 상황조사행(항목명):
        cols = st.columns([2, 5, 3])
        with cols[0]:
            st.markdown(f"<div style='text-align:center; font-weight:bold; padding-top:0.7em;'>{항목명}</div>", unsafe_allow_html=True)
        with cols[1]:
            상태 = st.radio(
                label="",
                options=["변화없음", "감소", "증가", "기타"],
                key=f"{항목명}_상태",
                horizontal=True,
                label_visibility="collapsed"
            )
        with cols[2]:
            if 상태 == "감소":
                st.text_input("감소 - 언제부터", key=f"{항목명}_감소_시작", placeholder="언제부터", label_visibility="collapsed")
            elif 상태 == "증가":
                st.text_input("증가 - 언제부터", key=f"{항목명}_증가_시작", placeholder="언제부터", label_visibility="collapsed")
            elif 상태 == "기타":
                st.text_input("기타 - 내용", key=f"{항목명}_기타_내용", placeholder="내용", label_visibility="collapsed")
            else:
                st.markdown("&nbsp;", unsafe_allow_html=True)

    for 항목 in ["작업설비", "작업량", "작업속도", "업무변화"]:
        상황조사행(항목)
        st.markdown("<hr style='margin:0.5em 0;'>", unsafe_allow_html=True)

with tabs[3]:
    st.title("작업조건조사 (인간공학적 측면)")

    # 1단계: 작업별 주요 작업내용
    st.markdown("#### 1단계 : 작업별 주요 작업내용")
    if "checklist_df" in st.session_state:
        checklist_df = st.session_state["checklist_df"]
        # 작업명/단위작업명 리스트 추출
        작업명_list = checklist_df["작업명"].dropna().unique().tolist()
        단위작업명_list = checklist_df["단위작업명"].dropna().unique().tolist()
        st.write("작업명 목록:", 작업명_list)
        st.write("단위작업명 목록:", 단위작업명_list)
        # 표로 보여주기
        st.dataframe(checklist_df[["작업명", "단위작업명"]], use_container_width=True, hide_index=True)
    else:
        st.info("체크리스트 탭에서 작업명을 먼저 입력하세요.")

    # 2단계: 작업별 작업부하 및 작업빈도
    st.markdown("#### 2단계 : 작업별 작업부하 및 작업빈도")
    if "checklist_df" in st.session_state:
        checklist_df = st.session_state["checklist_df"]
        # 단위작업명별로 표 생성
        부하옵션 = [
            "매우쉬움(1)", "쉬움(2)", "약간 힘듦(3)", "힘듦(4)", "매우 힘듦(5)"
        ]
        빈도옵션 = [
            "3개월마다(년2-3회)((1)", "가끔(하루 또는 주2-3일에 1회)(2)", "자주(1일 4시간)(3)", "계속(1일 4시간 이상)(4)", "초과근무(1일 8시간 이상)(5)"
        ]
        rows = []
        for idx, row in checklist_df.iterrows():
            단위작업명 = row["단위작업명"]
            if not 단위작업명: continue
            col1, col2, col3, col4, col5 = st.columns([3, 3, 2, 2, 2])
            with col1:
                st.write(단위작업명)
            with col2:
                # 부담작업(호): O(해당)만 콤마로
                부담호 = [f"{i+1}" for i, v in enumerate([row[f"{i}호"] for i in range(1, 12)]) if v.startswith("O")]
                st.write(", ".join(부담호))
            with col3:
                a = st.selectbox(
                    "작업부하", 부하옵션, key=f"{단위작업명}_부하"
                )
                a_val = int(a.split("(")[-1].replace(")", ""))
            with col4:
                b = st.selectbox(
                    "작업빈도", 빈도옵션, key=f"{단위작업명}_빈도"
                )
                b_val = int(b.split("(")[-1].replace(")", ""))
            with col5:
                st.write(f"{a_val * b_val}")
    else:
        st.info("체크리스트 탭에서 작업명을 먼저 입력하세요.")
