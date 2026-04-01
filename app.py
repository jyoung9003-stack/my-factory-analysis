import streamlit as st
import pandas as pd
import plotly.express as px
import re
import io

# 1. 웹 화면 설정
st.set_page_config(page_title="사출 생산성 통합 분석기", layout="wide")
st.title("📊 사출 생산성 통합 분석 리포트")

# 2. 파일 업로드
uploaded_files = st.file_uploader("분석할 파일들을 선택하세요", type=['xlsx', 'csv'], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    
    for file in uploaded_files:
        try:
            # 파일 읽기
            file_bytes = file.read()
            if file.name.endswith('.csv'):
                raw_df = pd.read_csv(io.BytesIO(file_bytes))
            else:
                raw_df = pd.read_excel(io.BytesIO(file_bytes))
            
            # [핵심] 제목 줄(Header) 자동 찾기 로직
            # '설비' 또는 '호기'라는 글자가 있는 행을 찾습니다.
            header_row = 0
            for i, row in raw_df.iterrows():
                row_str = " ".join(row.astype(str))
                if '설비' in row_str or '호기' in row_str:
                    header_row = i + 1
                    break
            
            # 찾은 위치를 바탕으로 데이터를 다시 정규화하여 읽기
            if file.name.endswith('.csv'):
                temp_df = pd.read_csv(io.BytesIO(file_bytes), skiprows=header_row-1 if header_row > 0 else 0)
            else:
                temp_df = pd.read_excel(io.BytesIO(file_bytes), skiprows=header_row-1 if header_row > 0 else 0)
            
            # 항목명 정리
            temp_df.columns = [" ".join(str(c).split()) for c in temp_df.columns]
            
            # 항목명 표준화 (여러 양식 대응)
            name_map = {
                '작업장 [설비]': '설비명', '품목명': '품명', '합계': '합게수량', 
                '종합 효율': '종합효율', '목표 효율': '목표효율', 'Unnamed: 14': 'OPEN ISSUE'
            }
            temp_df = temp_df.rename(columns=name_map)
            
            # [수정 1] 생산일 추출 ("일일 생산성_" 제거)
            file_date = file.name.split('.')[0].replace('일일 생산성_', '').replace('사출생산팀 일일 생산성자료_', '')
            temp_df['생산일'] = file_date
            
            # 불필요한 합계(TOTAL) 행 제거
            if '설비명' in temp_df.columns:
                temp_df = temp_df[~temp_df['설비명'].astype(str).str.contains('TOTAL|합계|GRAND', na=False, case=False)]
            
            all_data.append(temp_df)
        except Exception as e:
            st.error(f"파일 읽기 오류 ({file.name}): {e}")

    if all_data:
        df = pd.concat(all_data, ignore_index=True)
        
        # 숫자 데이터 변환 및 결측치 처리
        num_cols = ['양품수량', '불량수량', '합게수량', '투입시간', '가동시간', '비가동시간', '정미시간', '종합효율', '목표효율']
        for col in num_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # ---------------------------------------------------------
        # [사이드바 필터링]
        # ---------------------------------------------------------
        st.sidebar.header("🎯 데이터 필터링")
        all_machines = sorted(df['설비명'].unique())
        selected_machines = st.sidebar.multiselect("분석할 설비 선택", all_machines, default=all_machines)
        
        df['품명_필터'] = df['품명'].fillna("비가동").replace(0, "비가동").astype(str)
        actual_prods = sorted([p for p in df['품명_필터'].unique() if p != "비가동" and p != "nan"])
        selected_prod = st.sidebar.selectbox("특정 품목 분석", ["전체 품목"] + actual_prods)

        filtered_df = df[df['설비명'].isin(selected_machines)]
        if selected_prod != "전체 품목":
            filtered_df = filtered_df[filtered_df['품명_필터'] == selected_prod]

        # ---------------------------------------------------------
        # 3. 분석 결과 그래프
        # ---------------------------------------------------------
        tab1, tab2 = st.tabs(["📈 기간별 추이", "🔍 설비/품목별 상세"])

        with tab1:
            c1, c2 = st.columns(2)
            with c1:
                st.subheader("일별 평균 종합효율(OEE)")
                active_df = filtered_df[filtered_df['종합효율'] > 0]
                if not active_df.empty:
                    daily_oee = active_df.groupby('생산일')['종합효율'].mean().reset_index()
                    fig1 = px.line(daily_oee, x='생산일', y='종합효율', markers=True, text=daily_oee['종합효율'].apply(lambda x: f'{x:.1%}'))
                    fig1.update_layout(yaxis_tickformat='.0%', yaxis_range=[0, 1.1])
                    st.plotly_chart(fig1, use_container_width=True)
            with c2:
                st.subheader("일별 총 비가동 시간(분)")
                daily_stop = filtered_df.groupby('생산일')['비가동시간'].sum().reset_index()
                fig2 = px.bar(daily_stop, x='생산일', y='비가동시간', text_auto='.1f', color_discrete_sequence=['#FF4B4B'])
                st.plotly_chart(fig2, use_container_width=True)

        with tab2:
            st.subheader("선택 조건별 효율 변화")
            if not filtered_df[filtered_df['종합효율'] > 0].empty:
                fig3 = px.line(filtered_df[filtered_df['종합효율'] > 0], x='생산일', y='종합효율', color='설비명', markers=True)
                fig3.update_layout(yaxis_tickformat='.0%', yaxis_range=[0, 1.1])
                st.plotly_chart(fig3, use_container_width=True)

        # ---------------------------------------------------------
        # 4. 통합 원본 데이터 (요청하신 순서 및 빈칸 처리)
        # ---------------------------------------------------------
        st.write("---")
        st.subheader("📂 통합 원본 데이터 (검토용)")
        
        display_df = filtered_df.sort_values(by=['생산일', '설비명']).copy()
        
        # [수정 2] 요청하신 항목 순서
        target_order = [
            '생산일', '설비명', '품명', '종합효율', '성능가동율', '시간가동율', '양품율', '목표효율',
            '투입시간', '가동시간', '비가동시간', '정미시간', '양품수량', '불량수량', '합게수량', 'OPEN ISSUE'
        ]
        display_df = display_df[[c for c in target_order if c in display_df.columns]]

        # [수정 3] 비가동 설비 빈칸 처리 및 숫자 포맷팅
        def format_final_table(row):
            # 품명이 없거나 0인 경우 비가동으로 판단
            is_idle = pd.isna(row['품명']) or str(row['품명']).strip() in ['0', '0.0', '', 'nan', 'NaN']
            res = row.copy()
            
            for col in res.index:
                if is_idle and col not in ['생산일', '설비명']:
                    res[col] = ""
                else:
                    # 백분율 항목
                    if col in ['종합효율', '성능가동율', '시간가동율', '양품율', '목표효율']:
                        try: res[col] = f"{float(res[col]):.1%}" if res[col] != "" else ""
                        except: pass
                    # 수량 항목 (콤마)
                    elif col in ['양품수량', '불량수량', '합게수량']:
                        try: res[col] = f"{int(float(res[col])):,.0f}" if res[col] != "" else ""
                        except: pass
                    # 시간 항목 (소수점 1자리) [수정 4]
                    elif col in ['투입시간', '가동시간', '비가동시간', '정미시간']:
                        try: res[col] = f"{float(res[col]):.1f}" if res[col] != "" else ""
                        except: pass
            
            if is_idle: res['품명'] = ""
            return res

        final_table = display_df.apply(format_final_table, axis=1)
        st.dataframe(final_table, use_container_width=True)
else:
    st.info("파일을 업로드해 주세요.")
