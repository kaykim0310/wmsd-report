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
        <th style="width:10%;">구분</th>
        <th style="width:12%;">변화없음</th>
        <th style="width:18%;">변화있음<br>(언제부터)</th>
        <th style="width:18%;">줄음<br>(언제부터)</th>
        <th style="width:18%;">늘어남<br>(언제부터)</th>
        <th style="width:18%;">기타</th>
      </tr>
    </table>
    """, unsafe_allow_html=True)

    def 상황조사행(항목명):
        cols = st.columns([1, 1, 2, 2, 2, 2])
        with cols[0]:
            st.markdown(f"<div style='text-align:center; font-weight:bold; padding-top:0.7em;'>{항목명}</div>", unsafe_allow_html=True)
        with cols[1]:
            st.markdown("<div style='text-align:center;'>", unsafe_allow_html=True)
            변화없음 = st.checkbox("", key=f"{항목명}_변화없음")
            st.markdown("</div>", unsafe_allow_html=True)
        with cols[2]:
            st.markdown("<div style='display:flex; flex-direction:column; align-items:center;'>", unsafe_allow_html=True)
            변화있음 = st.checkbox("", key=f"{항목명}_변화있음")
            변화있음_시작 = st.text_input("", key=f"{항목명}_변화있음_시작", placeholder="언제부터", label_visibility="collapsed") if 변화있음 else ""
            st.markdown("</div>", unsafe_allow_html=True)
        with cols[3]:
            st.markdown("<div style='display:flex; flex-direction:column; align-items:center;'>", unsafe_allow_html=True)
            줄음 = st.checkbox("", key=f"{항목명}_줄음")
            줄음_시작 = st.text_input("", key=f"{항목명}_줄음_시작", placeholder="언제부터", label_visibility="collapsed") if 줄음 else ""
            st.markdown("</div>", unsafe_allow_html=True)
        with cols[4]:
            st.markdown("<div style='display:flex; flex-direction:column; align-items:center;'>", unsafe_allow_html=True)
            늘어남 = st.checkbox("", key=f"{항목명}_늘어남")
            늘어남_시작 = st.text_input("", key=f"{항목명}_늘어남_시작", placeholder="언제부터", label_visibility="collapsed") if 늘어남 else ""
            st.markdown("</div>", unsafe_allow_html=True)
        with cols[5]:
            st.markdown("<div style='display:flex; flex-direction:column; align-items:center;'>", unsafe_allow_html=True)
            기타 = st.checkbox("", key=f"{항목명}_기타")
            기타_내용 = st.text_input("", key=f"{항목명}_기타_내용", placeholder="내용", label_visibility="collapsed") if 기타 else ""
            st.markdown("</div>", unsafe_allow_html=True)

    for 항목 in ["작업설비", "작업량", "작업속도", "업무변화"]:
        상황조사행(항목)
        st.markdown("<hr style='margin:0.5em 0;'>", unsafe_allow_html=True)
