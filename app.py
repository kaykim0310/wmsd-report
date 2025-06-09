import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")

tabs = st.tabs([
    "사업장개요",
    "근골격계 부담작업 체크리스트",
    "유해요인조사표"
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
        "부", "팀", "작업명", "단위작업명", "일일 해당작업 시간", "중량(kg)",
        "1호", "2호", "3호", "4호", "5호", "6호", "7호", "8호", "9호", "10호", "11호"
    ]
    data = pd.DataFrame(columns=columns, data=[[""]*len(columns) for _ in range(5)])

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

    column_config = {}
    for col in columns:
        if col in ho_tooltips:
            column_config[col] = st.column_config.TextColumn(
                col,
                help=ho_tooltips[col]
            )
        else:
            column_config[col] = st.column_config.TextColumn(col)

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

with tabs[2]:
    st.title("유해요인조사표")

    st.markdown("#### 가. 조사개요")
    # 조사개요 입력
    col1, col2 = st.columns(2)
    with col1:
        조사일시 = st.text_input("조사일시", key="조사일시")
        부서명 = st.text_input("부서명", key="부서명")
    with col2:
        조사자 = st.text_input("조사자", key="조사자")
        작업공정명 = st.text_input("작업공정명", key="작업공정명")
    작업명 = st.text_input("작업명", key="작업명")

    # 조사개요 표로 출력
    st.markdown(f"""
    <table border="1" style="width:500px; text-align:center; border-collapse:collapse; margin-bottom:30px;">
      <tr>
        <td style="width:120px;">조사일시</td>
        <td style="width:180px;">{조사일시}</td>
        <td style="width:80px;">조사자</td>
        <td style="width:120px;">{조사자}</td>
      </tr>
      <tr>
        <td>부서명</td>
        <td colspan="3">{부서명}</td>
      </tr>
      <tr>
        <td>작업공정명</td>
        <td colspan="3">{작업공정명}</td>
      </tr>
      <tr>
        <td>작업명</td>
        <td colspan="3">{작업명}</td>
      </tr>
    </table>
    """, unsafe_allow_html=True)

    st.markdown("#### 나. 작업장 상황조사")

    def 상황조사행(항목명):
        st.markdown(f"<b>{항목명}</b>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns([1, 2, 3])
        with col1:
            변화없음 = st.checkbox("변화없음", key=f"{항목명}_변화없음")
        with col2:
            변화있음 = st.checkbox("변화있음", key=f"{항목명}_변화있음")
            변화시작 = st.text_input("변화있음(언제부터)", key=f"{항목명}_변화시작") if 변화있음 else ""
        with col3:
            if 항목명 != "작업설비":
                줄음 = st.checkbox("줄음", key=f"{항목명}_줄음")
                줄음_시작 = st.text_input("줄음(언제부터)", key=f"{항목명}_줄음_시작") if 줄음 else ""
                늘어남 = st.checkbox("늘어남", key=f"{항목명}_늘어남")
                늘어남_시작 = st.text_input("늘어남(언제부터)", key=f"{항목명}_늘어남_시작") if 늘어남 else ""
                기타 = st.checkbox("기타", key=f"{항목명}_기타")
                기타_내용 = st.text_input("기타(내용)", key=f"{항목명}_기타_내용") if 기타 else ""

    상황조사행("작업설비")
    st.markdown("---")
    상황조사행("작업량")
    st.markdown("---")
    상황조사행("작업속도")
    st.markdown("---")
    상황조사행("업무변화")
