import streamlit as st
import pandas as pd
import plotly.express as px

# 1. 웹 화면 설정
st.set_page_config(page_title="사출 생산성 통합 분석기", layout="wide")
st.title("📊 사출 생산성 통합 분석 리포트")

# 2. 파일 업로드
uploaded_files = st.file_uploader("분석할 파일들을 선택하세요", type=['xlsx', 'csv'], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    
    for file in uploaded_files:
        try:
            if file.name.endswith('.csv'):
                temp_df = pd.read_csv(file)
            else:
                temp_df = pd.read_excel(file)
            
            # 항목명 정리 및 OPEN ISSUE 변경
            temp_df.columns = [str(c).strip() for c in temp_df.columns]
            temp_df = temp_df.rename(columns={'Unnamed: 14': 'OPEN ISSUE'})
            
            # 생산일 추출 (날짜만 남기기)
            file_date = file.name.split('.')[0].replace('일일 생산성_', '')
            temp_df['생산일'] = file_date
            
            all_data.append(temp_df)
        except Exception as e:
            st.error(f"파일 읽기 오류 ({file.name}): {e}")

    if all_data:
        df = pd.concat(all_data, ignore_index=True)
        
        # 데이터 숫자 변환
        cols_to_fix = ['양품수량', '불량수량', '합게수량', '투입시간', '가동시간', '비가동시간', '정미시간', '종합효율', '목표효율']
        for col in cols_to_fix:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # 항목 순서 재배치
        desired_order = [
            '생산일', '설비명', '품명', '종합효율', '성능가동율', '시간가동율', '양품율', '목표효율',
            '투입시간', '가동시간', '비가동시간', '정미시간', '양품수량', '불량수량', '합게수량', 'OPEN ISSUE'
        ]
        df = df[[c for c in desired_order if c in df.columns]]
        df = df.sort_values(by=['생산일', '설비명']).reset_index(drop=True)

        # ---------------------------------------------------------
        # [추가] 3. 설비별/품목별 정밀 필터 (사이드바)
        # ---------------------------------------------------------
        st.sidebar.header("🎯 데이터 필터링")
        
        # 설비 선택 (멀티 선택 가능)
        all_machines = sorted(df['설비명'].unique())
        selected_machines = st.sidebar.multiselect("분석할 설비를 선택하세요", all_machines, default=all_machines)
        
        # 품목 선택 (멀티 선택 가능)
        # 0이나 NaN인 값은 '비가동'으로 표시하여 필터링 가능하게 함
        df['품명_필터'] = df['품명'].fillna("비가동").replace(0, "비가동")
        all_products = sorted([p for p in df['품명_필터'].unique() if p != "비가동"])
        all_products = ["전체 품목"] + all_products
        
        selected_prod = st.sidebar.selectbox("특정 품목만 보기", all_products)

        # 필터 적용
        filtered_df = df[df['설비명'].isin(selected_machines)]
        if selected_prod != "전체 품목":
            filtered_df = filtered_df[filtered_df['품명_필터'] == selected_prod]

        # ---------------------------------------------------------
        # 4. 분석 결과 탭 (필터링된 데이터 기반)
        # ---------------------------------------------------------
        tab1, tab2 = st.tabs(["📈 기간별 추이 분석", "🔍 설비별 정밀 분석"])

        with tab1:
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("🗓️ 일별 평균 종합효율(OEE)")
                active_df = filtered_df[filtered_df['종합효율'] > 0]
                if not active_df.empty:
                    daily_oee = active_df.groupby('생산일')['종합효율'].mean().reset_index()
                    fig1 = px.line(daily_oee, x='생산일', y='종합효율', markers=True, 
                                   text=daily_oee['종합효율'].apply(lambda x: f'{x:.1%}'))
                    fig1.update_layout(yaxis_tickformat='.0%', yaxis_range=[0, 1.1])
                    st.plotly_chart(fig1, use_container_width=True)
                else:
                    st.warning("선택한 조건에 해당하는 가동 데이터가 없습니다.")
            
            with col2:
                st.subheader("⏱️ 일별 총 비가동 시간(분)")
                daily_stop = filtered_df.groupby('생산일')['비가동시간'].sum().reset_index()
                fig2 = px.bar(daily_stop, x='생산일', y='비가동시간', text_auto='.1f', color_discrete_sequence=['#FF4B4B'])
                st.plotly_chart(fig2, use_container_width=True)

        with tab2:
            st.subheader("🔍 선택 설비별 변화")
            if len(selected_machines) > 0:
                fig3 = px.line(filtered_df[filtered_df['종합효율'] > 0], x='생산일', y='종합효율', color='설비명', markers=True)
                fig3.update_layout(yaxis_tickformat='.0%', yaxis_range=[0, 1.1])
                st.plotly_chart(fig3, use_container_width=True)

        # ---------------------------------------------------------
        # 5. 통합 원본 데이터 (검토용) - [수정] 0을 빈칸으로
        # ---------------------------------------------------------
        st.write("---")
        st.subheader("📂 통합 원본 데이터 (검토용)")

        # 출력을 위한 복사본 생성
        display_df = filtered_df.copy()
        
        # 비가동 설비(품명이 0이나 빈값인 경우)는 수치 데이터를 빈칸으로 처리
        numeric_cols = ['종합효율', '성능가동율', '시간가동율', '양품율', '목표효율', '투입시간', '가동시간', '비가동시간', '정미시간', '양품수량', '불량수량', '합게수량']
        
        # 품명이 없거나 0인 행 찾기
        mask = (display_df['품명'].isna()) | (display_df['품명'] == 0) | (display_df['품명'] == "")
        
        # 스타일 지정을 위한 포맷 함수
        def format_table(styler):
            # 먼저 기본 포맷 적용
            styler
