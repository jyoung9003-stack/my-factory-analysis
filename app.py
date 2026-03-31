import streamlit as st
import pandas as pd
import plotly.express as px
import re

# 1. 웹 화면 설정
st.set_page_config(page_title="사출 생산성 통합 분석기", layout="wide")
st.title("📈 사출 생산성 및 비가동 통합 분석 (추이 리포트)")

# 2. 여러 파일 업로드
uploaded_files = st.file_uploader("일별 엑셀/CSV 파일들을 모두 선택하세요 (Shift/Ctrl 키 활용)", type=['xlsx', 'csv'], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    
    for file in uploaded_files:
        try:
            # 파일 읽기
            if file.name.endswith('.csv'):
                temp_df = pd.read_csv(file)
            else:
                temp_df = pd.read_excel(file)
            
            # 제목 정리 (줄바꿈 제거 및 공백 정리)
            temp_df.columns = [" ".join(str(c).split()) for c in temp_df.columns]
            
            # [핵심] 파일명에서 날짜 추출 (예: 03.02, 0311 등 숫자와 점만 추출)
            # 파일명에서 숫자와 점(.)이 섞인 부분을 찾아 '분석날짜' 컬럼으로 강제 지정
            date_info = re.findall(r'\d{2}\.\d{2}|\d{4}', file.name)
            temp_df['분석날짜'] = date_info[0] if date_info else file.name
            
            all_data.append(temp_df)
        except Exception as e:
            st.error(f"{file.name} 파일을 읽는 중 오류가 발생했습니다: {e}")

    if all_data:
        # 모든 파일 합치기
        df = pd.concat(all_data, ignore_index=True)
        # 날짜순으로 정렬
        df = df.sort_values(by='분석날짜')

        # 3. 사이드바 자동/수동 설정
        st.sidebar.header("📊 분석 항목 설정")
        
        # 이름 기반으로 칸 자동 찾기
        def find_best_col(keywords, default_idx):
            options = [c for c in df.columns if any(k in c for k in keywords)]
            return options[0] if options else df.columns[default_idx]

        m_col = st.sidebar.selectbox("설비명(호기) 칸", df.columns, index=list(df.columns).index(find_best_col(['설비', '호기'], 0)))
        o_col = st.sidebar.selectbox("종합효율 칸", df.columns, index=list(df.columns).index(find_best_col(['종합', '효율'], 1)))
        s_col = st.sidebar.selectbox("비가동 시간 칸", df.columns, index=list(df.columns).index(find_best_col(['비가동', '정지'], 2)))

        # 데이터 숫자 변환 (계산이 가능하도록)
        df[o_col] = pd.to_numeric(df[o_col], errors='coerce').fillna(0)
        df[s_col] = pd.to_numeric(df[s_col], errors='coerce').fillna(0)

        # 4. 분석 결과 탭
        tab1, tab2 = st.tabs(["📊 전체 추이 분석", "🔍 설비별 상세 분석"])

        with tab1:
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("🗓️ 날짜별 평균 종합효율(OEE)")
                daily_oee = df.groupby('분석날짜')[o_col].mean().reset_index()
                fig1 = px.line(daily_oee, x='분석날짜', y=o_col, markers=True, text=daily_oee[o_col].apply(lambda x: f'{x:.1%}'))
                st.plotly_chart(fig1, use_container_width=True)
            with col2:
                st.subheader("⏱️ 날짜별 총 비가동 시간(분)")
                daily_stop = df.groupby('분석날짜')[s_col].sum().reset_index()
                fig2 = px.bar(daily_stop, x='분석날짜', y=s_col, text_auto=True, color_discrete_sequence=['#EF553B'])
                st.plotly_chart(fig2, use_container_width=True)

        with tab2:
            st.subheader("🔍 특정 설비 추이 보기")
            selected_machine = st.selectbox("분석할 설비를 선택하세요", sorted(df[m_col].unique()))
            machine_df = df[df[m_col] == selected_machine]
            
            fig3 = px.area(machine_df, x='분석날짜', y=o_col, title=f"{selected_machine} 효율 변화")
            st.plotly_chart(fig3, use_container_width=True)

        st.write("---")
        st.subheader("📂 통합 데이터 확인")
        st.dataframe(df)
