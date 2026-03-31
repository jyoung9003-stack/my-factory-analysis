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
            
            # [항목명 정리] 양쪽 공백 제거
            temp_df.columns = [str(c).strip() for c in temp_df.columns]
            
            # [수정 3] Unnamed: 14 항목을 OPEN ISSUE로 변경
            temp_df = temp_df.rename(columns={'Unnamed: 14': 'OPEN ISSUE'})
            
            # [수정 1] '분석날짜'를 '생산일'로 변경 (파일명 기준)
            file_date = file.name.split('.')[0]
            temp_df['생산일'] = file_date
            
            # 결측치 처리
            temp_df = temp_df.fillna(0)
            
            all_data.append(temp_df)
        except Exception as e:
            st.error(f"파일 읽기 오류 ({file.name}): {e}")

    if all_data:
        df = pd.concat(all_data, ignore_index=True)
        
        # [데이터 숫자 변환]
        cols_to_fix = ['양품수량', '불량수량', '합게수량', '투입시간', '가동시간', '비가동시간', '종합효율', '목표효율']
        for col in cols_to_fix:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # [수정 2] 항목 순서 재배치 (관리자님 요청 순서 + 추가 항목)
        desired_order = [
            '생산일', '설비명', '품명', '종합효율', '성능가동율', '시간가동율', '양품율', '목표효율',
            '투입시간', '가동시간', '비가동시간', '정미시간', '양품수량', '불량수량', '합게수량'
        ]
        # 만약 OPEN ISSUE가 있으면 리스트 끝에 추가
        if 'OPEN ISSUE' in df.columns:
            desired_order.append('OPEN ISSUE')
        
        # 실제 존재하는 컬럼만 선별하여 재배치
        existing_cols = [c for c in desired_order if c in df.columns]
        df = df[existing_cols]

        # 데이터 정렬 (생산일 -> 설비명 순)
        df = df.sort_values(by=['생산일', '설비명']).reset_index(drop=True)

        tab1, tab2 = st.tabs(["📈 기간별 추이 분석", "🔍 설비별 정밀 분석"])

        with tab1:
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("🗓️ 일별 평균 종합효율(OEE)")
                active_df = df[df['종합효율'] > 0]
                daily_oee = active_df.groupby('생산일')['종합효율'].mean().reset_index()
                
                fig1 = px.line(daily_oee, x='생산일', y='종합효율', markers=True, 
                               text=daily_oee['종합효율'].apply(lambda x: f'{x:.1%}'))
                fig1.update_layout(yaxis_tickformat='.0%', yaxis_range=[0, 1.1])
                fig1.update_traces(textposition="top center")
                st.plotly_chart(fig1, use_container_width=True)
            
            with col2:
                st.subheader("⏱️ 일별 총 비가동 시간(분)")
                daily_stop = df.groupby('생산일')['비가동시간'].sum().reset_index()
                fig2 = px.bar(daily_stop, x='생산일', y='비가동시간', text_auto='.1f')
                st.plotly_chart(fig2, use_container_width=True)

        with tab2:
            st.subheader("🔍 설비별 성적 확인")
            target_machine = st.selectbox("분석할 설비를 선택하세요", sorted(df['설비명'].unique()))
            m_df = df[df['설비명'] == target_machine]
            
            fig3 = px.area(m_df, x='생산일', y='종합효율', title=f"{target_machine} 효율 변화")
            fig3.update_layout(yaxis_tickformat='.0%', yaxis_range=[0, 1.1])
            st.plotly_chart(fig3, use_container_width=True)

        st.write("---")
        st.subheader("📂 통합 원본 데이터 (검토용)")
        
        # [수정 4] 목표효율 백분율 서식 추가 및 천 단위 콤마
        formatted_df = df.style.format({
            '양품수량': '{:,.0f}',
            '불량수량': '{:,.0f}',
            '합게수량': '{:,.0f}',
            '종합효율': '{:.1%}',
            '목표효율': '{:.1%}',
            '양품율': '{:.1%}',
            '성능가동율': '{:.1%}',
            '시간가동율': '{:.1%}'
        })
        st.dataframe(formatted_df)
else:
    st.info("파일을 업로드해 주세요.")
