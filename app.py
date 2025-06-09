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

    # 설명용 표 (가이드라인)
    st.markdown("""
    <table border="1" style="width:100%; text-align:center; font-size:14px;">
      <tr>
        <th>구분</th>
        <th>시간</th>
        <th>노출시간</th>
        <th>노출빈도</th>
        <th>신체부위</th>
        <th>작업자세 및 내용</th>
      </tr>
      <tr>
        <td>1호</td>
        <td>하루 4시간 이상</td>
        <td>반복작업</td>
        <td>손, 손가락</td>
        <td>공구, 키보드 등 사용</td>
        <td>반복적 손작업</td>
      </tr>
      <tr>
        <td>2호</td>
        <td>하루 4시간 이상</td>
        <td>반복작업</td>
        <td>어깨, 팔</td>
        <td>팔을 머리 위로 들어올림</td>
        <td>반복적 팔작업</td>
      </tr>
      <!-- 필요시 3~12호도 추가 -->
    </table>
    """, unsafe_allow_html=True)

    # 실제 입력 테이블
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
