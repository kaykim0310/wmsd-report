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

    # 1~11호 이미지 파일명 리스트 (예시)
    ho_images = [
        "1.png", "2.png", "3.png", "4.png", "5.png", "6.png",
        "7.png", "8.png", "9.png", "10.png", "11.png"
    ]
    # 이미지 파일을 static 폴더 등에 넣고, 아래처럼 표시
    img_html = "<div style='display:flex; gap:8px; justify-content:left;'>"
    for img in ho_images:
        img_html += f"<div><img src='https://your_image_url/{img}' width='40'><br><span style='font-size:12px;'>호</span></div>"
    img_html += "</div>"
    st.markdown(img_html, unsafe_allow_html=True)

    # 실제 입력 테이블
    columns = [
        "부", "팀", "작업명", "단위작업명", "일일 해당작업 시간", "중량(kg)",
        "1호", "2호", "3호", "4호", "5호", "6호", "7호", "8호", "9호", "10호", "11호"
    ]
    data = pd.DataFrame(columns=columns, data=[[""]*len(columns) for _ in range(5)])

    edited_df = st.data_editor(
        data,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True
    )
