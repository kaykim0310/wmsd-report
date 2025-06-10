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

st.subheader("2단계: 작업별 작업부하 및 작업빈도")

행_수 = 5  # 원하는 행 개수로 조정
작업부하_옵션 = ["", "1", "2", "3", "4", "5"]  # 예시 옵션, 실제 값으로 수정 가능
작업빈도_옵션 = ["", "1", "2", "3", "4", "5"]

rows = []
for i in range(행_수):
    cols = st.columns([2, 2, 2, 2, 2, 2])
    with cols[0]:
        단위작업명 = st.text_input("단위작업명", key=f"단위작업명_{i}")
    with cols[1]:
        부담작업호 = st.text_input("부담작업(호)", key=f"부담작업호_{i}")
    with cols[2]:
        작업부하 = st.selectbox("작업부하(A)", 작업부하_옵션, key=f"작업부하_{i}")
    with cols[3]:
        작업빈도 = st.selectbox("작업빈도(B)", 작업빈도_옵션, key=f"작업빈도_{i}")
    with cols[4]:
        총점 = st.text_input("총점", key=f"총점_{i}")
    rows.append({
        "단위작업명": 단위작업명,
        "부담작업(호)": 부담작업호,
        "작업부하(A)": 작업부하,
        "작업빈도(B)": 작업빈도,
        "총점": 총점
    })


with tabs[1]:
    st.subheader("근골격계 부담작업 체크리스트")
    columns = [
        "작업명", "단위작업명"
    ] + [f"{i}호" for i in range(1, 12)]
    data = pd.DataFrame(columns=columns, data=[["", ""] + ["X(미해당)"]*11 for _ in range(5)])

    ho_tooltips = {
        "1호": "노출시간: 하루 4시간 이상\n노출빈도: 반복작업\n신체부위: 손, 손가락\n작업자세 및 내용: 공구, 키보드 등 사용\n무게: -",
        "2호": "노출시간: 하루 4시간 이상\n노출빈도: 반복작업\n신체부위: 어깨, 팔\n작업자세 및 내용: 팔을 머리 위로 들어올림\n무게: -",
        "3호": "노출시간: 하루 4시간 이상\n노출빈도: 반복작업\n신체부위: 어깨, 팔\n작업자세 및 내용: 팔을 머리 위로 들어올림\n무게: -",
        "4호": "노출시간: 하루 2시간 이상\n노출빈도: 반복작업\n신체부위: 목, 허리\n작업자세 및 내용: 구부리거나 비트는 자세\n무게: 1kg이상",
        "5호": "노출시간: 하루 2시간 이상\n노출빈도: 반복작업\n신체부위: 다리, 무릎\n작업자세 및 내용: 무릎 꿇기, 쪼그리기\n무게: 4.5kg이상",
        "6호": "노출시간: 하루 2시간 이상\n노출빈도: 반복작업\n신체부위: 손가락\n작업자세 및 내용: 반복적 손작업\n무게: 25kg이상",
        "7호": "노출시간: 하루 2시간 이상\n노출빈도: 반복작업\n신체부위: 손\n작업자세 및 내용: 반복적 손작업\n무게: 10kg이상",
        "8호": "노출시간: 10회 이상\n노출빈도: 반복작업\n신체부위: 허리\n작업자세 및 내용: 중량물 들기\n무게: 4.5kg이상",
        "9호": "노출시간: 2회 이상\n노출빈도: 반복작업\n신체부위: 허리\n작업자세 및 내용: 중량물 들기\n무게: -",
        "10호": "노출시간: 하루 2시간 이상\n노출빈도: 반복작업\n신체부위: 허리\n작업자세 및 내용: 중량물 들기\n무게: -",
        "11호": "노출시간: 하루 2시간 이상\n노출빈도: 반복작업\n신체부위: 팔, 몸통\n작업자세 및 내용: 팔을 머리 위로 들어올림\n무게: -"
    }

    # 드롭다운 옵션
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

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='체크리스트')
    return output.getvalue()

st.download_button(
    label="엑셀로 저장",
    data=to_excel(edited_df),
    file_name="근골격계_부담작업_체크리스트.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

    # 체크리스트 데이터 session_state에 저장
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

    st.markdown("#### 1단계 : 작업별 주요 작업내용")
     # 여러 작업명 중 첫 번째만 예시로 사용 (여러 작업명 확장 가능)
    작업명_list = st.session_state["checklist_df"]["작업명"].dropna().unique().tolist()
    if 작업명_list:
        작업명 = st.selectbox("작업명", 작업명_list, key="작업조건조사_작업명")
    else:
        작업명 = st.text_input("작업명", key="작업조건조사_작업명")

    st.markdown("#### 2단계 : 작업별 작업부하 및 작업빈도")

    # 2단계 바로 아래 표
    st.markdown("#### 단위작업명별 입력")
    # 체크리스트에서 선택된 작업명에 해당하는 단위작업명만 추출
    if 작업명:
        filtered = st.session_state["checklist_df"][st.session_state["checklist_df"]["작업명"] == 작업명]
        table_data = []
        for idx, row in filtered.iterrows():
            단위작업명 = row["단위작업명"]
            # 부담작업(호): O(해당)만 콤마로
            부담호 = [f"{i+1}" for i, v in enumerate([row[f"{i}호"] for i in range(1, 12)]) if v.startswith("O")]
            # 드롭다운(작업부하/작업빈도)
            col1, col2, col3, col4, col5 = st.columns([3, 3, 2, 2, 2])
            with col1:
                st.write(단위작업명)
            with col2:
                st.write(", ".join(부담호))
            with col3:
                a = st.selectbox(
                    "작업부하", ["매우쉬움 (1)", "쉬움 (2)", "약간 힘듦 (3)", "힘듦 (4)", "매우 힘듦 (5)"],
                    key=f"{작업명}_{단위작업명}_부하"
                )
                a_val = int(a.split("(")[-1].replace(")", ""))
            with col4:
                b = st.selectbox(
                    "작업빈도", ["3개월마다(1)", "가끔(2)", "자주(3)", "계속(4)", "초과근무(5)"],
                    key=f"{작업명}_{단위작업명}_빈도"
                )
                b_val = int(b.split("(")[-1].replace(")", ""))
            with col5:
                st.write(f"{a_val * b_val}")


    st.dataframe(edited_df, use_container_width=True, hide_index=True)
