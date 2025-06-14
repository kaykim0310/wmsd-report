import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")

# 세션 상태 초기화
if "checklist_df" not in st.session_state:
    st.session_state["checklist_df"] = pd.DataFrame()

tabs = st.tabs([
    "사업장개요",
    "근골격계 부담작업 체크리스트",
    "유해요인조사표",
    "작업조건조사",
    "정밀조사"
])

with tabs[0]:
    st.title("사업장 개요")
    사업장명 = st.text_input("사업장명", key="사업장명")
    소재지 = st.text_input("소재지", key="소재지")
    업종 = st.text_input("업종", key="업종")
    col1, col2 = st.columns(2)
    with col1:
        예비조사 = st.date_input("예비조사일", key="예비조사")
        수행기관 = st.text_input("수행기관", key="수행기관")
    with col2:
        본조사 = st.date_input("본조사일", key="본조사")
        성명 = st.text_input("성명", key="성명")

with tabs[1]:
    st.subheader("근골격계 부담작업 체크리스트")
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
    st.title("작업조건조사")
    
    # 체크리스트에서 작업명 목록 가져오기
    작업명_목록 = []
    if not st.session_state["checklist_df"].empty:
        작업명_목록 = st.session_state["checklist_df"]["작업명"].dropna().unique().tolist()
    
    if not 작업명_목록:
        st.warning("⚠️ 먼저 '근골격계 부담작업 체크리스트' 탭에서 작업명을 입력해주세요.")
    else:
        # 작업명 선택
        selected_작업명 = st.selectbox(
            "작업명 선택",
            작업명_목록,
            key="작업명_선택"
        )
        
        st.info(f"📋 총 {len(작업명_목록)}개의 작업이 있습니다. 각 작업별로 1,2,3단계를 작성하세요.")
        
        # 선택된 작업에 대한 1,2,3단계
        with st.container():
            # 1단계: 유해요인 기본조사
            st.subheader(f"1단계: 유해요인 기본조사 - [{selected_작업명}]")
            col1, col2 = st.columns(2)
            with col1:
                작업공정 = st.text_input("작업공정", value=selected_작업명, key=f"1단계_작업공정_{selected_작업명}")
            with col2:
                작업내용 = st.text_input("작업내용", key=f"1단계_작업내용_{selected_작업명}")
            
            st.markdown("---")
            
            # 2단계: 작업별 작업부하 및 작업빈도
            st.subheader(f"2단계: 작업별 작업부하 및 작업빈도 - [{selected_작업명}]")
            
            # 선택된 작업명에 해당하는 체크리스트 데이터 가져오기
            checklist_data = []
            if not st.session_state["checklist_df"].empty:
                # 디버깅을 위한 정보 표시
                작업_체크리스트 = st.session_state["checklist_df"][
                    st.session_state["checklist_df"]["작업명"] == selected_작업명
                ]
                
                for idx, row in 작업_체크리스트.iterrows():
                    if row["단위작업명"]:  # 단위작업명이 있는 경우만
                        부담작업호 = []
                        for i in range(1, 12):
                            if row[f"{i}호"] == "O(해당)":
                                부담작업호.append(f"{i}호")
                            elif row[f"{i}호"] == "△(잠재위험)":
                                부담작업호.append(f"{i}호(잠재)")
                        
                        # 부담작업이 있든 없든 단위작업명이 있으면 추가
                        checklist_data.append({
                            "단위작업명": row["단위작업명"],
                            "부담작업(호)": ", ".join(부담작업호) if 부담작업호 else "미해당",
                            "작업부하(A)": "",
                            "작업빈도(B)": "",
                            "총점": 0
                        })
            
            # 데이터프레임 생성
            if checklist_data:
                data = pd.DataFrame(checklist_data)
            else:
                # 기본 빈 데이터프레임
                data = pd.DataFrame({
                    "단위작업명": ["" for _ in range(3)],
                    "부담작업(호)": ["" for _ in range(3)],
                    "작업부하(A)": ["" for _ in range(3)],
                    "작업빈도(B)": ["" for _ in range(3)],
                    "총점": [0 for _ in range(3)],
                })

            부하옵션 = [
                "",
                "매우쉬움(1)", 
                "쉬움(2)", 
                "약간 힘듦(3)", 
                "힘듦(4)", 
                "매우 힘듦(5)"
            ]
            빈도옵션 = [
                "",
                "3개월마다(1)", 
                "가끔(2)", 
                "자주(3)", 
                "계속(4)", 
                "초과근무(5)"
            ]

            column_config = {
                "작업부하(A)": st.column_config.SelectboxColumn("작업부하(A)", options=부하옵션, required=False),
                "작업빈도(B)": st.column_config.SelectboxColumn("작업빈도(B)", options=빈도옵션, required=False),
                "단위작업명": st.column_config.TextColumn("단위작업명"),
                "부담작업(호)": st.column_config.TextColumn("부담작업(호)"),
                "총점": st.column_config.TextColumn("총점(자동계산)", disabled=True),
            }

            # 작업부하와 작업빈도에서 숫자 추출하는 함수
            def extract_number(value):
                if value and "(" in value and ")" in value:
                    return int(value.split("(")[1].split(")")[0])
                return 0

            # 총점 계산 함수
            def calculate_total_score(row):
                부하값 = extract_number(row["작업부하(A)"])
                빈도값 = extract_number(row["작업빈도(B)"])
                return 부하값 * 빈도값

            # 데이터 편집
            edited_df = st.data_editor(
                data,
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                column_config=column_config,
                key=f"작업조건_data_editor_{selected_작업명}"
            )
            
            # 편집된 데이터를 세션 상태에 저장
            st.session_state[f"작업조건_data_{selected_작업명}"] = edited_df
            
            # 총점 자동 계산 후 다시 표시
            if not edited_df.empty:
                # 총점 계산
                display_df = edited_df.copy()
                for idx in range(len(display_df)):
                    display_df.at[idx, "총점"] = calculate_total_score(display_df.iloc[idx])
                
                # 계산된 데이터를 다시 표시
                st.markdown("##### 계산 결과")
                st.dataframe(
                    display_df,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "단위작업명": st.column_config.TextColumn("단위작업명"),
                        "부담작업(호)": st.column_config.TextColumn("부담작업(호)"),
                        "작업부하(A)": st.column_config.TextColumn("작업부하(A)"),
                        "작업빈도(B)": st.column_config.TextColumn("작업빈도(B)"),
                        "총점": st.column_config.NumberColumn("총점(자동계산)", format="%d"),
                    }
                )
                
                st.info("💡 총점은 작업부하(A) × 작업빈도(B)로 자동 계산됩니다.")
            
            # 3단계: 유해요인평가
            st.markdown("---")
            st.subheader(f"3단계: 유해요인평가 - [{selected_작업명}]")
            
            # 작업명과 근로자수 입력
            col1, col2 = st.columns(2)
            with col1:
                평가_작업명 = st.text_input("작업명", value=selected_작업명, key=f"3단계_작업명_{selected_작업명}")
            with col2:
                평가_근로자수 = st.text_input("근로자수", key=f"3단계_근로자수_{selected_작업명}")
            
            # 사진 업로드 및 설명 입력
            st.markdown("#### 작업 사진 및 설명")
            
            # 사진 개수 선택
            num_photos = st.number_input("사진 개수", min_value=1, max_value=10, value=3, key=f"사진개수_{selected_작업명}")
            
            # 각 사진별로 업로드와 설명 입력
            for i in range(num_photos):
                st.markdown(f"##### 사진 {i+1}")
                col1, col2 = st.columns([1, 2])
                
                with col1:
                    uploaded_file = st.file_uploader(
                        f"사진 {i+1} 업로드",
                        type=['png', 'jpg', 'jpeg'],
                        key=f"사진_{i+1}_업로드_{selected_작업명}"
                    )
                    if uploaded_file:
                        st.image(uploaded_file, caption=f"사진 {i+1}", use_column_width=True)
                
                with col2:
                    photo_description = st.text_area(
                        f"사진 {i+1} 설명",
                        height=150,
                        key=f"사진_{i+1}_설명_{selected_작업명}",
                        placeholder="이 사진에 대한 설명을 입력하세요..."
                    )
                
                st.markdown("---")
            
            # 작업별로 관련된 유해요인에 대한 원인분석
            st.markdown("---")
            st.subheader(f"작업별로 관련된 유해요인에 대한 원인분석 - [{selected_작업명}]")
            
            # 2단계에서 입력한 데이터 가져오기
            부담작업_정보 = []
            if 'display_df' in locals() and not display_df.empty:
                for idx, row in display_df.iterrows():
                    if row["단위작업명"] and row["부담작업(호)"]:
                        부담작업_정보.append({
                            "단위작업명": row["단위작업명"],
                            "부담작업호": row["부담작업(호)"]
                        })
            
            # 7개 행의 데이터 준비
            원인분석_data = []
            for i in range(7):
                if i < len(부담작업_정보):
                    부담작업_텍스트 = f"부담작업({부담작업_정보[i]['부담작업호']})"
                    단위작업명_기본값 = 부담작업_정보[i]['단위작업명']
                else:
                    부담작업_텍스트 = "부담작업(해당사항없음)"
                    단위작업명_기본값 = ""
                    
                원인분석_data.append({
                    "번호": str(i+1),
                    "단위작업명": 단위작업명_기본값,
                    "유해요인": "",
                    "부담작업": 부담작업_텍스트,
                    "발생원인": "",
                    "비고": ""
                })
            
            원인분석_df = pd.DataFrame(원인분석_data)
            
            # 컬럼 설정
            원인분석_column_config = {
                "번호": st.column_config.TextColumn("", disabled=True, width=50),
                "단위작업명": st.column_config.TextColumn("단위작업명", width=200),
                "유해요인": st.column_config.TextColumn("유해요인", width=200),
                "부담작업": st.column_config.TextColumn("", disabled=True, width=180),
                "발생원인": st.column_config.TextColumn("발생원인", width=300),
                "비고": st.column_config.TextColumn("비고", width=150)
            }
            
            # 데이터 편집기
            원인분석_edited_df = st.data_editor(
                원인분석_df,
                use_container_width=True,
                hide_index=True,
                column_config=원인분석_column_config,
                key=f"원인분석_data_editor_{selected_작업명}",
                disabled=["번호", "부담작업"]
            )
            
            # 원인분석 데이터도 세션 상태에 저장
            st.session_state[f"원인분석_data_{selected_작업명}"] = 원인분석_edited_df
    
    # 엑셀 다운로드 버튼 추가
    st.markdown("---")
    if st.button("엑셀 파일로 다운로드"):
        try:
            output = BytesIO()
            
            # 작업명 목록 다시 가져오기
            작업명_목록_다운로드 = []
            if "checklist_df" in st.session_state and not st.session_state["checklist_df"].empty:
                작업명_목록_다운로드 = st.session_state["checklist_df"]["작업명"].dropna().unique().tolist()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # 사업장 개요 정보
                overview_data = {
                    "항목": ["사업장명", "소재지", "업종", "예비조사일", "본조사일", "수행기관", "성명"],
                    "내용": [
                        st.session_state.get("사업장명", ""),
                        st.session_state.get("소재지", ""),
                        st.session_state.get("업종", ""),
                        str(st.session_state.get("예비조사", "")),
                        str(st.session_state.get("본조사", "")),
                        st.session_state.get("수행기관", ""),
                        st.session_state.get("성명", "")
                    ]
                }
                overview_df = pd.DataFrame(overview_data)
                overview_df.to_excel(writer, sheet_name='사업장개요', index=False)
                
                # 체크리스트
                if "checklist_df" in st.session_state and not st.session_state["checklist_df"].empty:
                    st.session_state["checklist_df"].to_excel(writer, sheet_name='체크리스트', index=False)
                
                # 각 작업별 데이터 저장
                for 작업명 in 작업명_목록_다운로드:
                    # 작업조건조사 데이터 저장
                    data_key = f"작업조건_data_{작업명}"
                    if data_key in st.session_state:
                        작업_df = st.session_state[data_key]
                        if isinstance(작업_df, pd.DataFrame) and not 작업_df.empty:
                            export_df = 작업_df.copy()
                            
                            # 총점 계산
                            def calc_score(row):
                                def extract_num(value):
                                    if pd.isna(value) or value == "":
                                        return 0
                                    if isinstance(value, str) and "(" in value and ")" in value:
                                        try:
                                            return int(value.split("(")[1].split(")")[0])
                                        except:
                                            return 0
                                    return 0
                                
                                부하값 = extract_num(row.get("작업부하(A)", ""))
                                빈도값 = extract_num(row.get("작업빈도(B)", ""))
                                return 부하값 * 빈도값
                            
                            for idx in range(len(export_df)):
                                export_df.at[idx, "총점"] = calc_score(export_df.iloc[idx])
                            
                            # 시트 이름 정리 (특수문자 제거)
                            sheet_name = f'작업조건_{작업명}'.replace('/', '_').replace('\\', '_')[:31]
                            export_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # 3단계 유해요인평가 데이터 저장
                    평가_작업명 = st.session_state.get(f"3단계_작업명_{작업명}", 작업명)
                    평가_근로자수 = st.session_state.get(f"3단계_근로자수_{작업명}", "")
                    
                    평가_data = {
                        "작업명": [평가_작업명],
                        "근로자수": [평가_근로자수]
                    }
                    
                    # 사진 설명 추가
                    사진개수 = st.session_state.get(f"사진개수_{작업명}", 3)
                    for i in range(사진개수):
                        설명 = st.session_state.get(f"사진_{i+1}_설명_{작업명}", "")
                        평가_data[f"사진{i+1}_설명"] = [설명]
                    
                    if 평가_작업명 or 평가_근로자수:  # 데이터가 있을 때만 저장
                        평가_df = pd.DataFrame(평가_data)
                        sheet_name = f'유해요인평가_{작업명}'.replace('/', '_').replace('\\', '_')[:31]
                        평가_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # 원인분석 데이터 저장
                    원인분석_key = f"원인분석_data_{작업명}"
                    if 원인분석_key in st.session_state:
                        원인분석_df = st.session_state[원인분석_key]
                        if isinstance(원인분석_df, pd.DataFrame) and not 원인분석_df.empty:
                            sheet_name = f'원인분석_{작업명}'.replace('/', '_').replace('\\', '_')[:31]
                            원인분석_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # 정밀조사 데이터 저장 (있는 경우만)
                if st.session_state.get("정밀_작업공정명", "") or st.session_state.get("정밀_작업명", ""):
                    정밀_data_rows = []
                    
                    # 기본 정보
                    정밀_data_rows.append(["작업공정명", st.session_state.get("정밀_작업공정명", "")])
                    정밀_data_rows.append(["작업명", st.session_state.get("정밀_작업명", "")])
                    정밀_data_rows.append([])  # 빈 행
                    정밀_data_rows.append(["작업별로 관련된 유해요인에 대한 원인분석"])
                    정밀_data_rows.append(["작업분석 및 평가도구", "분석결과", "만점"])
                    
                    # 원인분석 데이터
                    if "정밀_원인분석_data" in st.session_state:
                        원인분석_df = st.session_state["정밀_원인분석_data"]
                        for _, row in 원인분석_df.iterrows():
                            if row.get("작업분석 및 평가도구", "") or row.get("분석결과", "") or row.get("만점", ""):
                                정밀_data_rows.append([
                                    row.get("작업분석 및 평가도구", ""),
                                    row.get("분석결과", ""),
                                    row.get("만점", "")
                                ])
                    
                    if len(정밀_data_rows) > 5:  # 헤더 이후에 데이터가 있는 경우만
                        정밀_sheet_df = pd.DataFrame(정밀_data_rows)
                        정밀_sheet_df.to_excel(writer, sheet_name="정밀조사", index=False, header=False)
                
            output.seek(0)
            st.download_button(
                label="📥 엑셀 다운로드",
                data=output,
                file_name="근골격계_유해요인조사.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"엑셀 파일 생성 중 오류가 발생했습니다: {str(e)}")
            st.info("데이터를 입력한 후 다시 시도해주세요.")
