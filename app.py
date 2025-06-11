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
    
    # 작업명 입력은 이미 체크리스트에서 입력받기 때문에 중복 제거(혹은 별도 관리 가능)
    # 작업명 = st.text_input("작업명", key="tab2_작업명")
    
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
    
    st.markdown("#### 1단계 : 작업별 주요 작업내용")
    # 작업명 및 작업내용(단위작업명)은 이미 체크리스트 탭에서 입력받은 값으로 관리하므로 별도 입력필드를 제거했습니다.
    # 예시로 체크리스트에서 첫 번째 값을 표시할 수 있습니다.
    if st.session_state.get("checklist_df") is not None:
        df = st.session_state["checklist_df"]
        if not df.empty:
            st.write("**작업명:**", df["작업명"].iloc[0])
            st.write("**작업내용(단위작업명):**", df["단위작업명"].iloc[0])
    
    st.markdown("#### 2단계 : 작업별 작업부하 및 작업빈도")
    st.write("2단계 시작됨")
    st.markdown("""
    - **작업부하**: 매우쉬움(1), 쉬움(2), 약간 힘듦(3), 힘듦(4), 매우 힘듦(5)
    - **작업빈도**: 3개월마다(1), 가끔(2), 자주(3), 계속(4), 초과근무(5)
    """)

    checklist_df = st.session_state.get("checklist_df")
    if checklist_df is not None:
        filtered = checklist_df[["단위작업명"] + [col for col in checklist_df.columns if "호" in col]]
        단위작업명_list = filtered["단위작업명"].tolist()
        부담작업호_list = []
        for idx, row in filtered.iterrows():
            해당호 = [col for col in filtered.columns if "호" in col and row[col] == "O(해당)"]
            부담작업호_list.append(", ".join(해당호) if 해당호 else "")
    else:
        단위작업명_list = ["" for _ in range(7)]
        부담작업호_list = ["" for _ in range(7)]

    row_count = max(len(단위작업명_list), 7)
    부하옵션 = ["", "매우쉬움(1)", "쉬움(2)", "약간 힘듦(3)", "힘듦(4)", "매우 힘듦(5)"]
    빈도옵션 = ["", "3개월마다(년 2-3회)(1)", "가끔(하루 또는 주2-3일에 1회)(2)", "자주(1일 4시간)(3)", "계속(1일 4시간이상)(4)", "초과근무(1일 8시간이상)(5)"]

    data = pd.DataFrame({
        "단위작업명": 단위작업명_list + [""] * (row_count - len(단위작업명_list)),
        "부담작업(호)": 부담작업호_list + [""] * (row_count - len(부담작업호_list)),
        "작업부하(A)": ["" for _ in range(row_count)],
        "작업빈도(B)": ["" for _ in range(row_count)],
        "총점": ["" for _ in range(row_count)],
    })

    column_config = {
        "단위작업명": st.column_config.TextColumn("단위작업명", width="medium", disabled=True),
        "부담작업(호)": st.column_config.TextColumn("부담작업(호)", width="medium", disabled=True),
        "작업부하(A)": st.column_config.SelectboxColumn("작업부하(A)", options=부하옵션, width="medium"),
        "작업빈도(B)": st.column_config.SelectboxColumn("작업빈도(B)", options=빈도옵션, width="medium"),
        "총점": st.column_config.TextColumn("총점", width="medium", disabled=True),
    }

    edited_df = st.data_editor(
        data,
        column_config=column_config,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="작업조건조사표"
    )

    total_sum = 0
    for i in range(len(edited_df)):
        a = edited_df.loc[i, "작업부하(A)"]
        b = edited_df.loc[i, "작업빈도(B)"]
        try:
            a_val = int(a.split("(")[-1].replace(")", "")) if "(" in str(a) else 0
            b_val = int(b.split("(")[-1].replace(")", "")) if "(" in str(b) else 0
            score = a_val * b_val if a_val and b_val else ""
            edited_df.loc[i, "총점"] = str(score) if score else ""
            if score:
                total_sum += score
        except Exception:
            edited_df.loc[i, "총점"] = ""

    st.markdown("**총합**")
    st.text_input("총합", value=str(total_sum), disabled=True, key="총합_자동계산")
