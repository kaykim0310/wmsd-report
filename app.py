import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")

# 세션 상태 초기화
if "checklist_df" not in st.session_state:
    st.session_state["checklist_df"] = pd.DataFrame()

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
    st.title("작업조건조사")
    
    # 1단계: 유해요인 기본조사
    st.subheader("1단계: 유해요인 기본조사")
    col1, col2 = st.columns(2)
    with col1:
        작업공정 = st.text_input("작업공정", key="1단계_작업공정")
    with col2:
        작업내용 = st.text_input("작업내용", key="1단계_작업내용")
    
    st.markdown("---")
    
    # 2단계: 작업별 작업부하 및 작업빈도
    st.subheader("2단계: 작업별 작업부하 및 작업빈도")
    
    # 체크리스트에서 데이터 가져오기
    if not st.session_state["checklist_df"].empty:
        checklist_data = []
        for idx, row in st.session_state["checklist_df"].iterrows():
            if row["작업명"] and row["단위작업명"]:  # 작업명과 단위작업명이 있는 경우만
                부담작업호 = []
                for i in range(1, 12):
                    if row[f"{i}호"] == "O(해당)":
                        부담작업호.append(f"{i}호")
                    elif row[f"{i}호"] == "△(잠재위험)":
                        부담작업호.append(f"{i}호(잠재)")
                
                if 부담작업호:  # 해당하는 호가 있는 경우만 추가
                    checklist_data.append({
                        "단위작업명": row["단위작업명"],
                        "부담작업(호)": ", ".join(부담작업호),
                        "작업부하(A)": "",
                        "작업빈도(B)": "",
                        "총점": 0
                    })
        
        if checklist_data:
            data = pd.DataFrame(checklist_data)
        else:
            data = pd.DataFrame({
                "단위작업명": ["" for _ in range(5)],
                "부담작업(호)": ["" for _ in range(5)],
                "작업부하(A)": ["" for _ in range(5)],
                "작업빈도(B)": ["" for _ in range(5)],
                "총점": [0 for _ in range(5)],
            })
    else:
        data = pd.DataFrame({
            "단위작업명": ["" for _ in range(5)],
            "부담작업(호)": ["" for _ in range(5)],
            "작업부하(A)": ["" for _ in range(5)],
            "작업빈도(B)": ["" for _ in range(5)],
            "총점": [0 for _ in range(5)],
        })

    부하옵션 = [
        "",
        "매우쉬움(1)", 
        "쉬움(2)", 
        "약간 힘듦(3)", 
        "힘듦(4)", 
        "매우 힘듦(5)"
    ]
    빈도옵션 = [
        "",
        "3개월마다(1)", 
        "가끔(2)", 
        "자주(3)", 
        "계속(4)", 
        "초과근무(5)"
    ]

    column_config = {
        "작업부하(A)": st.column_config.SelectboxColumn("작업부하(A)", options=부하옵션, required=False),
        "작업빈도(B)": st.column_config.SelectboxColumn("작업빈도(B)", options=빈도옵션, required=False),
        "단위작업명": st.column_config.TextColumn("단위작업명"),
        "부담작업(호)": st.column_config.TextColumn("부담작업(호)"),
        "총점": st.column_config.TextColumn("총점(자동계산)", disabled=True),
    }

    # 작업부하와 작업빈도에서 숫자 추출하는 함수
    def extract_number(value):
        if value and "(" in value and ")" in value:
            return int(value.split("(")[1].split(")")[0])
        return 0

    # 데이터 편집
    edited_df = st.data_editor(
        data,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config=column_config,
        key="작업조건_data_editor"
    )
    
    # 총점 자동 계산
    if not edited_df.empty:
        for idx in range(len(edited_df)):
            부하값 = extract_number(edited_df.at[idx, "작업부하(A)"])
            빈도값 = extract_number(edited_df.at[idx, "작업빈도(B)"])
            edited_df.at[idx, "총점"] = 부하값 * 빈도값
    
    # 엑셀 다운로드 버튼 추가
    st.markdown("---")
    if st.button("엑셀 파일로 다운로드"):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # 사업장 개요 정보
            overview_data = {
                "항목": ["사업장명", "소재지", "업종", "예비조사일", "본조사일", "수행기관", "성명"],
                "내용": [사업장명, 소재지, 업종, str(예비조사), str(본조사), 수행기관, 성명]
            }
            overview_df = pd.DataFrame(overview_data)
            overview_df.to_excel(writer, sheet_name='사업장개요', index=False)
            
            # 체크리스트
            if not st.session_state["checklist_df"].empty:
                st.session_state["checklist_df"].to_excel(writer, sheet_name='체크리스트', index=False)
            
            # 작업조건조사
            edited_df.to_excel(writer, sheet_name='작업조건조사', index=False)
            
        output.seek(0)
        st.download_button(
            label="📥 엑셀 다운로드",
            data=output,
            file_name="근골격계_유해요인조사.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
