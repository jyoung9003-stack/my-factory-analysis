import streamlit as st
import pandas as pd
import plotly.express as px
import re

# 1. 웹 화면 설정 및 스타일
st.set_page_config(page_title="사출 생산성 통합 분석기", layout="wide")
st.title("📈 사출 공정 일별 생산성 추이 리포트")
st.markdown("9개 이상의 파일을 통합하여 설비별 효율 변화를 분석합니다.")

# 2. 파일 업로드
uploaded_files = st.file_uploader("분석할 엑셀/CSV 파일들을 모두 선택하세요", type=['xlsx', 'csv'], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    
    for file in uploaded_files:
        try:
            # 파일 읽기 (제목이 두 줄인 경우를 대비해 유연하게 로드)
            if file.name.endswith('.csv'):
                temp_df = pd.read_csv(file)
            else:
                temp_df = pd.read_excel(file)
            
            # [헤더 정리] 줄바꿈 제거, 앞뒤 공백 제거
            temp_df.columns = [" ".join(str(c).split()) for c in temp_df.columns]
            
            # [날짜 추출] 파일명에서 '03.02' 같은 날짜 형식을 찾아 '분석날짜'로 통일
            date_match = re.search(r'\d{2}\.\d{2}|\d{4}', file.name)
            temp_df['분석날짜'] = date_match.group() if date_match else file.name
            
            all_data.append(temp_df)
        except Exception as e:
            st.error(f"파일 읽기 오류 ({file.name}): {e}")

    if all_data:
        # 모든 데이터 통합
        df = pd.concat(all_data, ignore_index=True)
        
        # [중요] 날짜순으로 데이터 정렬 (여기서 '날짜'가 아닌 '분석날짜' 사용)
        df = df.sort_values(by='분석날짜').reset_index(drop=True)

        # 3. 사이드바 분석 설정 (오류 방지 로직)
        st.sidebar.header("📊 분석 항목 설정")
        
        # 필터링할 컬럼 찾기 함수
        def get_col(keywords, default_idx):
            options = [c for c in df.columns if any(k in c for k in keywords)]
            return options[0] if options else df.columns[default_idx] if len(df.columns) > default_idx else df.columns[0]

        # 관리자님이 사이드바에서 선택할 항목들
        m_col = st.sidebar.selectbox("1. 설비명(호기) 칸", df.columns, index=list(df.columns).index(get_col(['설비', '호기'], 0)))
        o_col = st.sidebar.selectbox("2. 종합효율 칸", df.columns, index=list(df.columns).index(get_col(['종합', '효율'], 12)))
        s_col = st.sidebar.selectbox("3. 비가동 시간 칸", df.columns, index=list(df.columns).index(get_col(['비가동', '정지'], 7)))

        # 데이터 숫자화 (계산 가능하도록 변환)
        df[o_col] = pd.to_numeric(df[o_col], errors='coerce').fillna(0)
        df[s_col] = pd.to_numeric(df[s_col], errors='coerce').fillna(0)

        # 4. 분석 결과 표시
        tab1, tab2 = st.tabs(["📊 기간별 추이", "🔍 설비별 상세"])

        with tab1:
            c1, c2 = st.columns(2)
            with c1:
                st.subheader("🗓️ 일별 평균 종합효율(OEE)")
                daily_oee = df.groupby('분석날짜')[o_col].mean().reset_index()
                fig1 = px.line(daily_oee, x='분석날짜', y=o_col, markers=True, 
                               text=daily_oee[o_col].apply(lambda x: f'{x:.1%}'))
                fig1.update_traces(textposition="top center")
                st.plotly_chart(fig1, use_container_width=True)
            
            with c2:
                st.subheader("⏱️ 일별 총 비가동 시간(분)")
                daily_stop = df.groupby('분석날짜')[s_col].sum().reset_index()
                fig2 = px.bar(daily_stop, x='분석날짜', y=s_col, text_auto=True, color_discrete_sequence=['#FF4B4B'])
                st.plotly_chart(fig2, use_container_width=True)

        with tab2:
            st.subheader("🔍 설비별 가동 성적표")
            unique_machines = sorted(df[m_col].unique())
            target_machine = st.selectbox("분석할 호기를 선택하세요", unique_machines)
            
            m_df = df[df[m_col] == target_machine]
            fig3 = px.area(m_df, x='분석날짜', y=o_col, title=f"{target_machine} 효율 변화")
            st.plotly_chart(fig3, use_container_width=True)

        st.write("---")
