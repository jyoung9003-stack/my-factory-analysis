import streamlit as st
import pandas as pd
import plotly.express as px
import re

# 1. 웹 화면 설정
st.set_page_config(page_title="사출 공정 통합 분석기", layout="wide")
st.title("📊 사출 생산성 통합 분석 리포트")
st.markdown("9개 이상의 파일을 통합하여 일별 추이 및 설비 성적을 분석합니다.")

# 2. 파일 업로드
uploaded_files = st.file_uploader("분석할 파일들을 모두 선택하세요 (Shift/Ctrl 키 활용)", type=['xlsx', 'csv'], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    
    for file in uploaded_files:
        try:
            # 파일 읽기
            if file.name.endswith('.csv'):
                temp_df = pd.read_csv(file)
            else:
                temp_df = pd.read_excel(file)
            
            # [항목명 정리] 공백 제거 및 특수문자 정리
            temp_df.columns = [" ".join(str(c).split()) for c in temp_df.columns]
            
            # [날짜 추출] 파일명에서 20260302 또는 03.02 같은 패턴 추출
            date_match = re.search(r'\d{4}\d{2}\d{2}|\d{2}\.\d{2}', file.name)
            temp_df['분석날짜'] = date_match.group() if date_match else file.name
            
            # [#N/A 및 빈칸 처리] 생산 안한 설비는 모두 0으로 채움
            temp_df = temp_df.fillna(0)
            
            all_data.append(temp_df)
        except Exception as e:
            st.error(f"파일 읽기 오류 ({file.name}): {e}")

    if all_data:
        # 모든 데이터 통합 및 정렬
        df = pd.concat(all_data, ignore_index=True)
        df = df.sort_values(by='분석날짜').reset_index(drop=True)

        # 3. 사이드바 항목 설정 (관리자님이 정리한 이름으로 자동 매칭)
        st.sidebar.header("📊 분석 항목 확인")
        
        # 키워드로 정확한 열 찾기
        def find_col(keywords, default_idx):
            options = [c for c in df.columns if any(k in c for k in keywords)]
            return options[0] if options else df.columns[default_idx]

        m_col = st.sidebar.selectbox("1. 설비명 칸", df.columns, index=list(df.columns).index(find_col(['설비'], 0)))
        o_col = st.sidebar.selectbox("2. 종합효율 칸", df.columns, index=list(df.columns).index(find_col(['종합효율'], 12)))
        s_col = st.sidebar.selectbox("3. 비가동 시간 칸", df.columns, index=list(df.columns).index(find_col(['비가동'], 7)))

        # 데이터 숫자 변환 (백분율 계산을 위해 소수점 유지)
        df[o_col] = pd.to_numeric(df[o_col], errors='coerce').fillna(0)
        df[s_col] = pd.to_numeric(df[s_col], errors='coerce').fillna(0)

        # 4. 분석 결과 탭
        tab1, tab2 = st.tabs(["📈 기간별 추이 분석", "🔍 설비별 정밀 분석"])

        with tab1:
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("🗓️ 일별 평균 종합효율(OEE)")
                daily_oee = df.groupby('분석날짜')[o_col].mean().reset_index()
                # y축을 백분율(%)로 표시
                fig1 = px.line(daily_oee, x='분석날짜', y=o_col, markers=True, text=daily_oee[o_col].apply(lambda x: f'{x:.1%}'))
                fig1.update_layout(yaxis_tickformat='.0%') # y축 서식
                fig1.update_traces(textposition="top center")
                st.plotly_chart(fig1, use_container_width=True)
            
            with col2:
                st.subheader("⏱️ 일별 총 비가동 시간(분)")
                daily_stop = df.groupby('분석날짜')[s_col].sum().reset_index()
                fig2 = px.bar(daily_stop, x='분석날짜', y=s_col, text_auto='.1f', color_discrete_sequence=['#FF4B4B'])
                st.plotly_chart(fig2, use_container_width=True)

        with tab2:
            st.subheader("🔍 호기별 가동 현황")
            target_machine = st.selectbox("분석할 설비를 선택하세요", sorted(df[m_col].unique()))
            m_df = df[df[m_col] == target_machine]
            
            fig3 = px.area(m_df, x='분석날짜', y=o_col, title=f"{target_machine} 효율 변화")
            fig3.update_layout(yaxis_tickformat='.0%') # y축 서식
            st.plotly_chart(fig3, use_container_width=True)

        st.write("---")
        st.subheader("📂 통합 데이터 검토")
        # 테이블에서도 백분율로 보이게 서식 지정
        st.dataframe(df.style.format({o_col: '{:.1%}', '양품율': '{:.1%}', '성능가동율': '{:.1%}', '시간가동율': '{:.1%}'}))
else:
    st.info("왼쪽 상단에서 일별 생산성 파일들을 업로드해 주세요.")
