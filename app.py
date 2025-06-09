import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")

tabs = st.tabs([
    "사업장개요",
    "근골격계 부담작업 체크리스트"
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

    # 1~11호 설명을 한 줄로 표기
    ho_desc = [
        "하루 4시간 이상, 반복작업, 손/손가락, 공구·키보드 등 사용, -",
        "하루 4시간 이상, 반복작업, 어깨/팔, 팔 머리 위로, -",
        "하루 4시간 이상, 반복작업, 어깨/팔, 팔 머리 위로, -",
        "하루 2시간 이상, 반복작업, 목/허리, 구부리기·비틀기, 1kg이상",
        "하루 2시간 이상, 반복작업, 다리/무릎, 무릎꿇기·쪼그리기, 4.5kg이상",
        "하루 2시간 이상, 반복작업, 손가락, 반복적 손작업, 25kg이상",
        "하루 2시간 이상, 반복작업, 손, 반복적 손작업, 10kg이상",
        "10회 이상, 반복작업, 허리, 중량물 들기, 4.5kg이상",
        "2회 이상, 반복작업, 허리, 중량물 들기, -",
        "하루 2시간 이상, 반복작업, 허리, 중량물 들기, -",
        "하루 2시간 이상, 반복작업, 팔/몸통, 팔 머리 위로, -"
    ]
    # 설명 행을 표로 출력
    st.markdown(
        "<div style='overflow-x:auto;'>"
        "<table style='width:100%; text-align:center; font-size:12px; border-collapse:collapse;'>"
        "<tr>"
        "<th style='border:1px solid #ccc;'></th>"*6 +
        "".join([f"<th style='border:1px solid #ccc;'>{desc}</th>" for desc in ho_desc]) +
        "</tr>"
        "</table></div>",
        unsafe_allow_html=True
    )

    # 실제 입력 테이블
    columns = [
        "부", "팀", "작업명", "단위작업명", "일일 해당작업 시간", "중량(kg)",
        "1호", "2호", "3호", "4호", "5호", "6호", "7호", "8호", "9호", "10호", "11호"
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
