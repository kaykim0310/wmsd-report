import streamlit as st
import pandas as pd
from io import BytesIO

# 전체 페이지 넓게
st.set_page_config(layout="wide")

# 탭 구성
tabs = st.tabs([
    "사업장개요",
    "근골격계 부담작업 체크리스트"
])

# 1. 사업장개요 탭
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

# 2. 체크리스트 탭
with tabs[1]:
    st.subheader("근골격계 부담작업 체크리스트")
    columns = [
        "부", "팀", "작업명", "단위작업명", "일일 해당작업 시간", "중량(kg)",
        "1호", "2호", "3호", "4호", "5호", "6호", "7호", "8호", "9호", "10호", "11호", "12호"
    ]
    data = pd.DataFrame(columns=columns, data=[[""]*len(columns) for _ in range(5)])  # 5행 예시

    edited_df = st.data_editor(
        data,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True
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
