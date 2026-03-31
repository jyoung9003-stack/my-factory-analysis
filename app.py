import streamlit as st
import pandas as pd
import plotly.express as px
import re

# 1. 웹 화면 설정
st.set_page_config(page_title="사출 생산성 정밀 분석기", layout="wide")
st.title("🏭 사출 공정 통합 추이 리포트 (데이터 정밀 매칭)")

# 2. 파일 업로드
uploaded_files = st.file_uploader("분석할 9개 이상의 파일을 선택하세요", type=['xlsx', 'csv'], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    
    for file in uploaded_files:
        try:
            # [수정] 헤더가 두 줄인 양식을 고려하여 첫 2행을 읽음
            if file.name.endswith('.csv'):
                # CSV의 경우 첫 2행을 읽어 제목을 합침
                raw_df = pd.read_csv(file, header=[0, 1])
            else:
                raw_df = pd.read_excel(file, header=[0, 1])
            
            # 두 줄의 제목을 하나로 합치기 (예: '운영 시간' + '비가동' = '운영 시간 비가동')
            new_cols = []
            for col in raw_df.columns:
                c1 = "" if "Unnamed" in str(col[0]) else str(col[0])
                c2 = "" if "Unnamed" in str(col[1]) else str(col[1])
                combined = (c1 + " " + c2).strip()
                new_cols.append(combined)
            
            raw_df.columns = new_cols
            temp_df = raw_df.copy()
            
            # 파일명에서 날짜 추출 (03.02 등)
            date_match = re.search(r'\d{2}\.\d{2}|\d{4}', file.name)
            temp_df['분석날짜'] = date_match.group() if date_match else file.name
            
            all_data.append(temp_df)
        except Exception as e:
            st.error(f"파일 읽기 오류 ({file.name}): {e}")

    if all_data:
        # 데이터 통합 및 정렬
        df = pd.concat(all_data, ignore_index=True)
        df = df.sort_values(by='분석날짜').reset_index(drop=True)

        # 3. 사이드바 항목 설정 (정밀 매칭)
        st.sidebar.header("📊 데이터 칸 맞추기")
        st.sidebar.info("값의 숫자가 다르다면 여기서 정확한 칸을 다시 골라주세요.")

        # 키워드로 자동 찾기 (가장 정확한 칸 우선)
        m_col = st.sidebar.selectbox("1. 설비명 칸", df.columns, index=list(df.columns).index([c for c in df.columns if '설비' in c][0]) if [c for c in df.columns if '설비' in c] else 0)
        
        # '종합 효율'이라는 글자가 포함된 모든 칸 중 첫 번째 선택
        o_options = [c for c in df.columns if '종합' in c and '효율' in c]
        o_col = st.sidebar.selectbox("2. 종합효율(OEE) 칸", df.columns, index=list(df.columns).index(o_options[0]) if o_options else 0)
        
        # '비가동'이라는 글자가 포함된 모든 칸 중 첫 번째 선택
        s_options = [c for c in df.columns if '비가동' in c]
        s_col = st.sidebar.selectbox("3. 비가동 시간 칸", df.columns, index=list(df.columns).index(s_options[0]) if s_options else 0)

        # 4. 데이터 숫자 변환 및 계산
        df[o_col] = pd.to_numeric(df[o_col], errors='coerce').fillna(0)
        df[s_col] = pd.to_numeric(df[s_col], errors='coerce').fillna(0)

        # 5. 결과 탭 표시
        tab1, tab2 = st.tabs(["📊 기간별 추이", "🔍 설비별 상세"])

        with tab1:
            c1, c2 = st.columns(2)
            with c1:
                st.subheader("일별 평균 종합효율(OEE)")
                daily_oee = df.groupby('분석날짜')[o_col].mean().reset_index()
                fig1 = px.line(daily_oee, x='분석날짜', y=o_col, markers=True, text=daily_oee[o_col].apply(lambda x: f'{x:.1%}'))
                st.plotly_chart(fig1, use_container_width=True)
            
            with c2:
                st.subheader("일별 총 비가동 시간(분/시간)")
                daily_stop = df.groupby('분석날짜')[s_col].sum().reset_index()
                fig2 = px.bar(daily_stop, x='분석날짜', y=s_col, text_auto='.2f')
                st.plotly_chart(fig2, use_container_width=True)

        with tab2:
            machine = st.selectbox("분석할 호기를 선택하세요", sorted(df[m_col].unique()))
            m_df = df[df[m_col] == machine]
            fig3 = px.area(m_df, x='분석날짜', y=o_col, title=f"{machine} 효율 변화")
            st.plotly_chart(fig3, use_container_width=True)

        st.write("---")
        st.subheader("📂 데이터 검증 표")
        st.write("현재 선택된 칸의 값들입니다. 엑셀과 대조해 보세요.")
        st.dataframe(df[['분석날짜', m_col, o_col, s_col]])
