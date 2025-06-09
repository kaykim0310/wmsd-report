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
    작업명 = st.text_input("작업명")

    st.markdown("#### 나. 작업장 상황조사")

    # 표 헤더
    st.markdown("""
    <table border="1" style="width:100%; border-collapse:collapse; text-align:center;">
      <tr style="background-color:#f2f2f2;">
        <th>구분</th>
        <th>변화없음</th>
        <th>변화있음<br>(언제부터)</th>
        <th>줄음<br>(언제부터)</th>
        <th>늘어남<br>(언제부터)</th>
        <th>기타</th>
      </tr>
    </table>
    """, unsafe_allow_html=True)

    def 상황조사행(항목명):
        cols = st.columns([1, 1, 2, 2, 2, 2])
        with cols[0]:
            변화없음 = st.checkbox("", key=f"{항목명}_변화없음")
        with cols[1]:
            변화있음 = st.checkbox("", key=f"{항목명}_변화있음")
        with cols[2]:
            변화있음_시작 = st.text_input("", key=f"{항목명}_변화있음_시작") if 변화있음 else ""
        with cols[3]:
            줄음 = st.checkbox("", key=f"{항목명}_줄음")
            줄음_시작 = st.text_input("", key=f"{항목명}_줄음_시작") if 줄음 else ""
        with cols[4]:
            늘어남 = st.checkbox("", key=f"{항목명}_늘어남")
            늘어남_시작 = st.text_input("", key=f"{항목명}_늘어남_시작") if 늘어남 else ""
        with cols[5]:
            기타 = st.checkbox("", key=f"{항목명}_기타")
            기타_내용 = st.text_input("", key=f"{항목명}_기타_내용") if 기타 else ""
        # 구분(항목명) 라벨은 표 왼쪽에 별도로 출력
        st.markdown(
            f"<div style='position:relative; top:-60px; left:10px; width:100px; font-weight:bold;'>{항목명}</div>",
            unsafe_allow_html=True
        )

    # 각 행(항목) 반복
    for 항목 in ["작업설비", "작업량", "작업속도", "업무변화"]:
        상황조사행(항목)
        st.markdown("<hr style='margin:0.5em 0;'>", unsafe_allow_html=True)
