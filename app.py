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
            
            # [수정 1] "일일 생산성_" 단어 제거하고 날짜만 추출
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

        # ---------------------------------------------------------
        # [수정 2] 사이드바 정밀 필터 추가
        # ---------------------------------------------------------
        st.sidebar.header("🎯 데이터 필터링")
        
        # 설비 필터
        all_machines = sorted(df['설비명'].unique())
        selected_machines = st.sidebar.multiselect("분석할 설비를 선택하세요", all_machines, default=all_machines)
        
        # 품목 필터 (비가동 제외한 실제 품목 리스트)
        df['품명_필터'] = df['품명'].fillna("비가동").replace(0, "비가동")
        actual_products = sorted([p for p in df['품명_필터'].unique() if p != "비가동"])
        selected_prod = st.sidebar.selectbox("특정 품목만 보기", ["전체 품목"] + actual_products)

        # 필터 적용
        filtered_df = df[df['설비명'].isin(selected_machines)]
        if selected_prod != "전체 품목":
            filtered_df = filtered_df[filtered_df['품명_필터'] == selected_prod]

        # ---------------------------------------------------------
        # 3. 분석 결과 탭
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
            
            with col2:
                st.subheader("⏱️ 일별 총 비가동 시간(분)")
                daily_stop = filtered_df.groupby('생산일')['비가동시간'].sum().reset_index()
                fig2 = px.bar(daily_stop, x='생산일', y='비가동시간', text_auto='.1f', color_discrete_sequence=['#FF4B4B'])
                st.plotly_chart(fig2, use_container_width=True)

        with tab2:
            st.subheader("🔍 선택 설비별 변화 추이")
            if not filtered_df[filtered_df['종합효율'] > 0].empty:
                fig3 = px.line(filtered_df[filtered_df['종합효율'] > 0], x='생산일', y='종합효율', color='설비명', markers=True)
                fig3.update_layout(yaxis_tickformat='.0%', yaxis_range=[0, 1.1])
                st.plotly_chart(fig3, use_container_width=True)

        # ---------------------------------------------------------
        # 4. 통합 원본 데이터 (검토용) - [수정] 0을 빈칸으로 & 순서 정렬
        # ---------------------------------------------------------
        st.write("---")
        st.subheader("📂 통합 원본 데이터 (검토용)")

        # 표시용 데이터 생성
        display_df = filtered_df.sort_values(by=['생산일', '설비명']).copy()
        
        # 항목 순서 재배치
        target_order = [
            '생산일', '설비명', '품명', '종합효율', '성능가동율', '시간가동율', '양품율', '목표효율',
            '투입시간', '가동시간', '비가동시간', '정미시간', '양품수량', '불량수량', '합게수량', 'OPEN ISSUE'
        ]
        display_df = display_df[[c for c in target_order if c in display_df.columns]]

        # [핵심] 비가동 설비(품명이 0이나 NaN인 경우) 빈칸 처리 및 서식 적용
        def make_clean_table(row):
            # 품명이 없으면 수치 데이터를 빈칸으로
            is_idle = pd.isna(row['품명']) or row['품명'] == 0 or str(row['품명']).strip() == ""
            
            res = row.copy()
            pct_cols = ['종합효율', '성능가동율', '시간가동율', '양품율', '목표효율']
            num_cols = ['양품수량', '불량수량', '합게수량']
            time_cols = ['투입시간', '가동시간', '비가동시간', '정미시간']
            
            for col in pct_cols:
                if col in res: res[col] = "" if is_idle else f"{float(res[col]):.1%}"
            for col in num_cols:
                if col in res: res[col] = "" if is_idle else f"{int(res[col]):,}"
            for col in time_cols:
                if col in res: res[col] = "" if is_idle else f"{float(res[col]):.1f}"
            
            # 품명이 0이면 빈칸으로 표시
            if is_idle: res['품명'] = ""
            return res

        final_display = display_df.apply(make_clean_table, axis=1)
        st.dataframe(final_display)
else:
    st.info("파일을 업로드해 주세요.")
