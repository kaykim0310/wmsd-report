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

    # 설명용 표 (1~11호, 각 항목별 설명)
    st.markdown("""
    <div style="overflow-x:auto;">
    <table border="1" style="width:100%; text-align:center; font-size:13px; border-collapse:collapse;">
      <tr style="background-color:#f2f2f2;">
        <th>구분</th>
        <th>1호</th>
        <th>2호</th>
        <th>3호</th>
        <th>4호</th>
        <th>5호</th>
        <th>6호</th>
        <th>7호</th>
        <th>8호</th>
        <th>9호</th>
        <th>10호</th>
        <th>11호</th>
      </tr>
      <tr>
        <td>노출시간</td>
        <td>하루 4시간 이상</td>
        <td>하루 4시간 이상</td>
        <td>하루 4시간 이상</td>
        <td>하루 2시간 이상</td>
        <td>하루 2시간 이상</td>
        <td>하루 2시간 이상</td>
        <td>하루 2시간 이상</td>
        <td>10회 이상</td>
        <td>2회 이상</td>
        <td>하루 2시간 이상</td>
        <td>하루 2시간 이상</td>
      </tr>
      <tr>
        <td>노출빈도</td>
        <td>반복작업</td>
        <td>반복작업</td>
        <td>반복작업</td>
        <td>반복작업</td>
        <td>반복작업</td>
        <td>반복작업</td>
        <td>반복작업</td>
        <td>반복작업</td>
        <td>반복작업</td>
        <td>반복작업</td>
        <td>반복작업</td>
      </tr>
      <tr>
        <td>신체부위</td>
        <td>손, 손가락</td>
        <td>어깨, 팔</td>
        <td>어깨, 팔</td>
        <td>목, 허리</td>
        <td>다리, 무릎</td>
        <td>손가락</td>
        <td>손</td>
        <td>허리</td>
        <td>허리</td>
        <td>허리</td>
        <td>팔, 몸통</td>
      </tr>
      <tr>
        <td>작업자세 및 내용</td>
        <td>공구, 키보드 등 사용</td>
        <td>팔을 머리 위로 들어올림</td>
        <td>팔을 머리 위로 들어올림</td>
        <td>구부리거나 비트는 자세</td>
        <td>무릎 꿇기, 쪼그리기</td>
        <td>반복적 손작업</td>
        <td>반복적 손작업</td>
        <td>중량물 들기</td>
        <td>중량물 들기</td>
        <td>중량물 들기</td>
        <td>팔을 머리 위로 들어올림</td>
      </tr>
      <tr>
        <td>무게</td>
        <td>-</td>
        <td>-</td>
        <td>-</td>
        <td>1kg이상인 물건</td>
        <td>4.5kg이상인 물건</td>
        <td>25kg이상</td>
        <td>10kg이상</td>
        <td>4.5kg이상</td>
        <td>-</td>
        <td>-</td>
        <td>-</td>
      </tr>
    </table>
    </div>
    """, unsafe_allow_html=True)

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
