import streamlit as st
import pandas as pd
import plotly.express as px

# 1. 웹 화면 설정
st.set_page_config(page_title="사출 생산성 추이 분석기", layout="wide")
st.title("📈 일별 생산성 및 비가동 추이 분석")

# 2. 여러 파일 업로드 기능 (accept_multiple_files=True 추가)
uploaded_files = st.file_uploader("일별 엑셀 파일들을 모두 선택하여 올리세요", type=['xlsx', 'csv'], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    
    for file in uploaded_files:
        # 파일 읽기
        if file.name.endswith('.csv'):
            temp_df = pd.read_csv(file)
        else:
            temp_df = pd.read_excel(file)
        
        # 제목 정리 (공백 제거 등)
        temp_df.columns = [" ".join(str(c).split()) for c in temp_df.columns]
        
        # 파일명에서 날짜 추출하거나 데이터 내 날짜 컬럼 활용 (없으면 파일명 사용)
        if '날짜' not in temp_df.columns:
            temp_df['날짜'] = file.name.split('_')[-1].replace('.xlsx', '').replace('.csv', '')
            
        all_data.append(temp_df)
    
    # 모든 데이터를 하나로 합치기
    df = pd.concat(all_data, ignore_index=True)
    df['날짜'] = pd.to_numeric(df['날짜'], errors='coerce').fillna(df['날짜']) # 날짜 형식 정리
    df = df.sort_values(by='날짜') # 날짜순 정렬

    # 3. 데이터 분석 (종합 효율, 비가동 시간 찾기)
    oee_col = [c for c in df.columns if '종합 효율' in c or '종합효율' in c][0]
    stop_col = [c for c in df.columns if '비가동' in c][0]
    machine_col = [c for c in df.columns if '설비' in c][0]

    # 숫자 데이터 변환
    df[oee_col] = pd.to_numeric(df[oee_col], errors='coerce').fillna(0)
    df[stop_col] = pd.to_numeric(df[stop_col], errors='coerce').fillna(0)

    # 4. 추이 그래프 그리기
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("🗓️ 날짜별 종합효율(OEE) 추이")
        # 일별 평균 OEE 계산
        daily_oee = df.groupby('날짜')[oee_col].mean().reset_index()
        fig1 = px.line(daily_oee, x='날짜', y=oee_col, markers=True, title="전체 설비 평균 OEE 추이")
        st.plotly_chart(fig1, use_container_width=True)

    with col2:
        st.subheader("⏱️ 날짜별 비가동 시간 추이")
        # 일별 총 비가동 시간 계산
        daily_stop = df.groupby('날짜')[stop_col].sum().reset_index()
        fig2 = px.bar(daily_stop, x='날짜', y=stop_col, title="일별 총 비가동 시간(분)")
        st.plotly_chart(fig2, use_container_width=True)

    # 5. 설비별 상세 분석
    st.subheader("🔍 설비별 성적표 (선택한 기간 합계)")
    selected_machine = st.selectbox("분석할 설비를 선택하세요", df[machine_col].unique())
    machine_df = df[df[machine_col] == selected_machine]
    
    fig3 = px.area(machine_df, x='날짜', y=oee_col, title=f"{selected_machine}의 효율 변화")
    st.plotly_chart(fig3, use_container_width=True)

    st.write("---")
    st.write("📂 전체 통합 데이터")
    st.dataframe(df)
