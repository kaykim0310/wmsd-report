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


# 예시: 체크리스트에서 작업명 리스트 추출
작업명_list = ["A", "B"]  # 실제로는 session_state["checklist_df"]["작업명"].dropna().unique().tolist()

# 현재 작업명 인덱스 (session_state로 관리)
if "작업명_idx" not in st.session_state:
    st.session_state["작업명_idx"] = 0

# 현재 작업명
if 작업명_list:
    작업명 = 작업명_list[st.session_state["작업명_idx"]]
else:
    st.warning("작업명이 없습니다.")
    st.stop()

# 1단계 표 (HTML로 구현)
st.markdown(f"""
<table style="border-collapse:collapse; width:60%;">
    <tr>
        <th style="border:1px solid #000; background:#eee; width:20%;">작업명</th>
        <td style="border:1px solid #000;">{작업명}</td>
    </tr>
    <tr>
        <th style="border:1px solid #000; background:#eee;">작업내용<br>(단위작업명)</th>
        <td style="border:1px solid #000;">
            {" , ".join(['A1', 'A2']) if 작업명=='A' else "B1"}  <!-- 실제로는 체크리스트에서 추출 -->
        </td>
    </tr>
</table>
""", unsafe_allow_html=True)

# 다음 작업명으로 이동 버튼
if st.button("다음 작업명으로"):
    if st.session_state["작업명_idx"] < len(작업명_list) - 1:
        st.session_state["작업명_idx"] += 1
        st.experimental_rerun()
    else:
        st.success("모든 작업명 입력 완료!")
    
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

    # 입력값 확인용
    st.write("입력된 체크리스트 데이터:")
    st.write(st.session_state["checklist_df"])

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

    import streamlit as st

# 1단계 표 (작업명/작업내용)
st.markdown("""
<style>
.tb1, .tb1 th, .tb1 td {
    border: 1px solid #000;
    border-collapse: collapse;
    padding: 6px 10px;
    font-size: 16px;
}
.tb1 th {
    background: #f3f3f3;
    width: 140px;
    text-align: left;
}
.tb1 {
    width: 100%;
    margin-bottom: 16px;
}
.tb2, .tb2 th, .tb2 td {
    border: 1px solid #000;
    border-collapse: collapse;
    padding: 4px 8px;
    font-size: 15px;
    text-align: center;
}
.tb2 th {
    background: #f3f3f3;
}
.tb2 {
    width: 100%;
    margin-bottom: 16px;
}
.tb3, .tb3 th, .tb3 td {
    border: 1px solid #000;
    border-collapse: collapse;
    padding: 4px 8px;
    font-size: 15px;
    text-align: center;
}
.tb3 th {
    background: #f3f3f3;
}
.tb3 {
    width: 100%;
    margin-bottom: 16px;
}
</style>
""", unsafe_allow_html=True)

st.markdown("#### 1단계 : 작업별 주요 작업내용")
st.markdown("""
<table class="tb1">
    <tr>
        <th>작업명</th>
        <td>{작업명}</td>
    </tr>
    <tr>
        <th>작업내용<br>(단위작업명)</th>
        <td>{작업내용}</td>
    </tr>
</table>
""".format(
    작업명=st.text_input("작업명", key="작업명1", label_visibility="collapsed"),
    작업내용=st.text_area("작업내용(단위작업명)", key="작업내용1", label_visibility="collapsed")
), unsafe_allow_html=True)

# 2단계 기준표
st.markdown("#### 2단계 : 작업별 작업부하 및 작업빈도")
st.markdown("""
<table class="tb2">
    <tr>
        <th rowspan="5">작업부하</th>
        <td>매우쉬움</td><td>1</td>
        <th rowspan="5">작업빈도</th>
        <td>3개월마다(년 2~3회)</td><td>1</td>
    </tr>
    <tr>
        <td>쉬움</td><td>2</td>
        <td>가끔(하루 또는 주2~3일에 1회)</td><td>2</td>
    </tr>
    <tr>
        <td>약간 힘듦</td><td>3</td>
        <td>자주(1일 4시간)</td><td>3</td>
    </tr>
    <tr>
        <td>힘듦</td><td>4</td>
        <td>계속(1일 4시간 이상)</td><td>4</td>
    </tr>
    <tr>
        <td>매우 힘듦</td><td>5</td>
        <td>초과근무 시간(1일 8시간 이상)</td><td>5</td>
    </tr>
</table>
""", unsafe_allow_html=True)

# 2단계 입력표
st.markdown("""
<table class="tb3">
    <tr>
        <th>단위작업명</th>
        <th>부담작업(호)</th>
        <th>작업부하(A)</th>
        <th>작업빈도(B)</th>
        <th>총점</th>
    </tr>
""", unsafe_allow_html=True)

# 표 입력 행 수
row_count = 7
부하옵션 = ["", "매우쉬움(1)", "쉬움(2)", "약간 힘듦(3)", "힘듦(4)", "매우 힘듦(5)"]
빈도옵션 = ["", "3개월마다(1)", "가끔(2)", "자주(3)", "계속(4)", "초과근무(5)"]

for i in range(row_count):
    cols = st.columns([2, 2, 2, 2, 1], gap="small")
    with cols[0]:
        단위작업명 = st.text_input("", key=f"unit_{i}", label_visibility="collapsed")
    with cols[1]:
        부담호 = st.text_input("", key=f"ho_{i}", label_visibility="collapsed")
    with cols[2]:
        부하 = st.selectbox("", 부하옵션, key=f"부하_{i}", label_visibility="collapsed")
        a_val = int(부하.split("(")[-1].replace(")", "")) if "(" in 부하 else 0
    with cols[3]:
        빈도 = st.selectbox("", 빈도옵션, key=f"빈도_{i}", label_visibility="collapsed")
        b_val = int(빈도.split("(")[-1].replace(")", "")) if "(" in 빈도 else 0
    with cols[4]:
        st.write(a_val * b_val if a_val and b_val else "")

st.markdown("</table>", unsafe_allow_html=True)
