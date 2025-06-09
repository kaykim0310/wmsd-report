import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")

tabs = st.tabs([
    "사업장개요",
    "근골격계 부담작업 체크리스트"
])

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
