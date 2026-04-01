import streamlit as st
import pandas as pd
import plotly.express as px

# 1. 웹 화면 설정
st.set_page_config(page_title="사출 생산성 통합 분석기", layout="wide")
st.title("📊 사출 생산성 통합 분석 리포트")

# 2. 파일 업로드
uploaded_files = st.file_uploader("분석할 파일들을 선택하세요 (여러 개 선택 가능)", type=['xlsx', 'csv'], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    
    for file in uploaded_files:
        try:
            # 🚨 [근본 해결] 복잡한 꼼수 삭제! 파일의 첫 줄을 정직하게 제목으로 바로 읽습니다.
            if file.name.endswith('.csv'):
                temp_df = pd.read_csv(file)
            else:
                temp_df = pd.read_excel(file)
            
            # 컬럼명에 섞인 엔터키나 공백을 깔끔하게 제거
            temp_df.columns = [str(c).replace('\n', '').replace('\r', '').strip() for c in temp_df.columns]
            
            # 엑셀마다 조금씩 다른 이름을 표준 이름으로 강제 변경
            name_map = {
                '작업장 [설비]': '설비명', '작업장[설비]': '설비명',
                '품목명': '품명',
                '합계': '합계수량', '합게수량': '합계수량', 
                '종합 효율': '종합효율', '목표 효율': '목표효율'
            }
            temp_df = temp_df.rename(columns=name_map)
            
            # OPEN ISSUE 처리 (Unnamed로 표기된 마지막 비고란을 찾아 이름 변경)
            for col in temp_df.columns:
                if 'Unnamed' in col or 'ISSUE' in col.upper():
                    temp_df = temp_df.rename(columns={col: 'OPEN ISSUE'})
                    break
            
            # 중복 컬럼 원천 차단 (만약 동일한 이름이 있으면 첫 번째 1개만 남김)
            temp_df = temp_df.loc[:, ~temp_df.columns.duplicated(keep='first')]
            
            if '품명' not in temp_df.columns:
                temp_df['품명'] = ""
            
            # 생산일 추출 (파일명에서 글자 제거)
            clean_date = file.name.split('.')[0].replace('일일 생산성_', '').replace('사출생산팀 일일 생산성자료_', '')
            temp_df['생산일'] = clean_date
            
            # 불필요한 합계 행 제거
            if '설비명' in temp_df.columns:
                temp_df = temp_df[~temp_df['설비명'].astype(str).str.contains('TOTAL|합계|GRAND', na=False, case=False)]
            
            all_data.append(temp_df)
        except Exception as e:
            st.error(f"파일 읽기 오류 ({file.name}): {e}")

    if all_data:
        df = pd.concat(all_data, ignore_index=True)
        
        # 병합 후 혹시 모를 중복 컬럼 한 번 더 제거
        df = df.loc[:, ~df.columns.duplicated(keep='first')]
        
        # 숫자 데이터 변환 (계산을 위해)
        num_cols = ['양품수량', '불량수량', '합계수량', '투입시간', '가동시간', '비가동시간', '정미시간', '종합효율', '목표효율', '양품율', '성능가동율', '시간가동율']
        for col in num_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # ---------------------------------------------------------
        # [사이드바 정밀 필터]
        # ---------------------------------------------------------
        st.sidebar.header("🎯 정밀 필터링")
        if '설비명' in df.columns:
            all_machines = sorted(df['설비명'].astype(str).unique())
            selected_machines = st.sidebar.multiselect("설비 선택", all_machines, default=all_machines)
        else:
            selected_machines = []
        
        # 안전한 품목 필터링 (에러 방지용 문자열 변환)
        df['품명_필터'] = df['품명'].fillna("").astype(str).str.strip()
        df['품명_필터'] = df['품명_필터'].replace(['0', '0.0', 'nan', 'NaN'], "")
        actual_prods = sorted([p for p in df['품명_필터'].unique() if p != ""])
        selected_prod = st.sidebar.selectbox("품목 선택", ["전체 품목"] + actual_prods)

        f_df = df[df['설비명'].isin(selected_machines)] if '설비명' in df.columns else df.copy()
        if selected_prod != "전체 품목":
            f_df = f_df[f_df['품명_필터'] == selected_prod]

        # ---------------------------------------------------------
        # [그래프 분석 탭]
        # ---------------------------------------------------------
        tab1, tab2 = st.tabs(["📈 일별 추이 분석", "🔍 상세 데이터 분석"])

        with tab1:
            c1, c2 = st.columns(2)
            with c1:
                st.subheader("일별 평균 종합효율(OEE)")
                if '종합효율' in f_df.columns:
                    active_oee = f_df[f_df['종합효율'] > 0]
                    if not active_oee.empty:
                        daily_oee = active_oee.groupby('생산일')['종합효율'].mean().reset_index()
                        fig1 = px.line(daily_oee, x='생산일', y='종합효율', markers=True, text=daily_oee['종합효율'].apply(lambda x: f'{x:.1%}'))
                        fig1.update_layout(yaxis_tickformat='.0%', yaxis_range=[0, 1.1])
                        st.plotly_chart(fig1, use_container_width=True)
            with c2:
                st.subheader("일별 총 비가동 시간(분)")
                if '비가동시간' in f_df.columns:
                    daily_stop = f_df.groupby('생산일')['비가동시간'].sum().reset_index()
                    fig2 = px.bar(daily_stop, x='생산일', y='비가동시간', text_auto='.1f', color_discrete_sequence=['#FF4B4B'])
                    st.plotly_chart(fig2, use_container_width=True)

        with tab2:
            st.subheader("설비별 가동 효율 변화")
            if '종합효율' in f_df.columns and not f_df[f_df['종합효율'] > 0].empty:
                fig3 = px.line(f_df[f_df['종합효율'] > 0], x='생산일', y='종합효율', color='설비명', markers=True)
                fig3.update_layout(yaxis_tickformat='.0%', yaxis_range=[0, 1.1])
                st.plotly_chart(fig3, use_container_width=True)

        # ---------------------------------------------------------
        # [통합 원본 데이터 표]
        # ---------------------------------------------------------
        st.write("---")
        st.subheader("📂 통합 원본 데이터 (검토용)")
        
        if '설비명' in f_df.columns:
            display_df = f_df.sort_values(by=['생산일', '설비명']).copy()
        else:
            display_df = f_df.copy()
        
        # 관리자님 요청 순서
        target_order = [
            '생산일', '설비명', '품명', '종합효율', '성능가동율', '시간가동율', '양품율', '목표효율',
            '투입시간', '가동시간', '비가동시간', '정미시간', '양품수량', '불량수량', '합계수량', 'OPEN ISSUE'
        ]
        display_df = display_df[[c for c in target_order if c in display_df.columns]]

        # 비가동 설비 빈칸 처리 및 서식(소수점/콤마) 지정
        def finalize_row(row):
            val = str(row.get('품명', '')).strip()
            is_idle = (val == '') or (val in ['0', '0.0', 'nan', 'NaN'])
            
            res = row.copy()
            for col in res.index:
                if is_idle and col not in ['생산일', '설비명']:
                    res[col] = ""
                else:
                    if col in ['종합효율', '성능가동율', '시간가동율', '양품율', '목표효율']:
                        try: res[col] = f"{float(res[col]):.1%}" if res[col] != "" else ""
                        except: pass
                    elif col in ['양품수량', '불량수량', '합계수량']:
                        try: res[col] = f"{int(float(res[col])):,.0f}" if res[col] != "" else ""
                        except: pass
                    elif col in ['투입시간', '가동시간', '비가동시간', '정미시간']:
                        try: res[col] = f"{float(res[col]):.1f}" if res[col] != "" else ""
                        except: pass
            if is_idle: 
                res['품명'] = ""
            return res

        st.dataframe(display_df.apply(finalize_row, axis=1), use_container_width=True)
else:
    st.info("왼쪽 상단에서 생산성 파일을 업로드해 주세요.")
