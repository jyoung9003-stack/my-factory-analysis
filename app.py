import streamlit as st
import pandas as pd
import plotly.express as px

# 1. 웹 화면 설정
st.set_page_config(page_title="사출 생산성 정밀 분석기", layout="wide")
st.title("📊 사출 생산성 통합 분석 리포트")

# 2. 파일 업로드
uploaded_files = st.file_uploader("분석할 파일들을 선택하세요", type=['xlsx', 'csv'], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    
    for file in uploaded_files:
        try:
            # 파일 읽기
            if file.name.endswith('.csv'):
                temp_df = pd.read_csv(file)
            else:
                temp_df = pd.read_excel(file)
            
            # [항목명 정리]
            temp_df.columns = [str(c).strip() for c in temp_df.columns]
            
            # [날짜 지정] 사용자의 요청대로 파일명을 그대로 사용 (확장자 제외)
            file_date = file.name.split('.')[0]
            temp_df['분석날짜'] = file_date
            
            # [#N/A 및 빈칸 처리] 생산 안 한 곳은 0으로 채움
            temp_df = temp_df.fillna(0)
            
            all_data.append(temp_df)
        except Exception as e:
            st.error(f"파일 읽기 오류 ({file.name}): {e}")

    if all_data:
        # 모든 데이터 통합 및 정렬
        df = pd.concat(all_data, ignore_index=True)
        df = df.sort_values(by='분석날짜').reset_index(drop=True)

        # 3. 데이터 숫자 변환 (계산 가능하도록)
        # 콤마나 문자가 섞여있을 경우를 대비해 숫자로 강제 변환
        cols_to_fix = ['양품수량', '불량수량', '합게수량', '투입시간', '가동시간', '비가동시간', '종합효율']
        for col in cols_to_fix:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # 4. 분석 결과 탭
        tab1, tab2 = st.tabs(["📈 기간별 추이 분석", "🔍 설비별 정밀 분석"])

        with tab1:
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("🗓️ 일별 평균 종합효율(OEE)")
                daily_oee = df.groupby('분석날짜')['종합효율'].mean().reset_index()
                # 그래프 소수점을 백분율(%)로 표기
                fig1 = px.line(daily_oee, x='분석날짜', y='종합효율', markers=True, 
                               text=daily_oee['종합효율'].apply(lambda x: f'{x:.1%}'))
                fig1.update_layout(yaxis_tickformat='.0%')
                fig1.update_traces(textposition="top center")
                st.plotly_chart(fig1, use_container_width=True)
            
            with col2:
                st.subheader("⏱️ 일별 총 비가동 시간(분)")
                daily_stop = df.groupby('분석날짜')['비가동시간'].sum().reset_index()
                fig2 = px.bar(daily_stop, x='분석날짜', y='비가동시간', text_auto='.1f')
                st.plotly_chart(fig2, use_container_width=True)

        with tab2:
            st.subheader("🔍 설비별 성적 확인")
            target_machine = st.selectbox("분석할 설비를 선택하세요", sorted(df['설비명'].unique()))
            m_df = df[df['설비명'] == target_machine]
            
            fig3 = px.area(m_df, x='분석날짜', y='종합효율', title=f"{target_machine} 효율 변화")
            fig3.update_layout(yaxis_tickformat='.0%')
            st.plotly_chart(fig3, use_container_width=True)

        st.write("---")
        st.subheader("📂 통합 원본 데이터 (검토용)")
        
        # [천 단위 콤마 및 백분율 서식 적용]
        formatted_df = df.style.format({
            '양품수량': '{:,.0f}',
            '불량수량': '{:,.0f}',
            '합게수량': '{:,.0f}',
            '종합효율': '{:.1%}',
            '양품율': '{:.1%}',
            '성능가동율': '{:.1%}',
            '시간가동율': '{:.1%}'
        })
        st.dataframe(formatted_df)
else:
    st.info("왼쪽 상단에서 파일을 업로드해 주세요.")
