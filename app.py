import streamlit as st
import pandas as pd
import plotly.express as px

# 1. 웹 화면 설정
st.set_page_config(page_title="사출 생산성 추이 분석기", layout="wide")
st.title("📈 일별 생산성 및 비가동 추이 분석")

# 2. 여러 파일 업로드 기능
uploaded_files = st.file_uploader("일별 엑셀 파일들을 모두 선택하여 올리세요", type=['xlsx', 'csv'], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    
    for file in uploaded_files:
        if file.name.endswith('.csv'):
            temp_df = pd.read_csv(file)
        else:
            # 엑셀의 헤더 위치를 유연하게 잡기 위해 header=[0,1] 등을 시도할 수 있으나, 
            # 여기서는 파일 제목 정리에 집중합니다.
            temp_df = pd.read_excel(file)
        
        # 제목 정리 (공백 및 줄바꿈 제거)
        temp_df.columns = ["".join(str(c).split()) for c in temp_df.columns]
        
        # 날짜 정보 추가
        if '날짜' not in temp_df.columns:
            # 파일명에서 숫자 4자리(날짜) 추출 시도
            import re
            date_match = re.search(r'\d{4}', file.name)
            temp_df['날짜'] = date_match.group() if date_match else file.name
            
        all_data.append(temp_df)
    
    df = pd.concat(all_data, ignore_index=True)
    df = df.sort_values(by='날짜')

    # 3. 항목 찾기 (더 강력한 검색 로직)
    # '종합'과 '효율'이 포함된 열 찾기
    oee_cols = [c for c in df.columns if '종합' in c and '효율' in c]
    # '비가동'이 포함된 열 찾기
    stop_cols = [c for c in df.columns if '비가동' in c]
    # '설비'가 포함된 열 찾기
    machine_cols = [c for c in df.columns if '설비' in c]

    # 항목을 하나라도 못 찾으면 안내 메시지 출력
    if not oee_cols or not stop_cols or not machine_cols:
        st.error("⚠️ 엑셀 파일에서 필요한 항목(종합효율, 비가동, 설비명)을 찾을 수 없습니다.")
        st.info(f"현재 파일에서 찾은 항목들: {list(df.columns)}")
        st.stop()

    oee_col = oee_cols[0]
    stop_col = stop_cols[0]
    machine_col = machine_cols[0]

    # 숫자 데이터 변환
    df[oee_col] = pd.to_numeric(df[oee_col], errors='coerce').fillna(0)
    df[stop_col] = pd.to_numeric(df[stop_col], errors='coerce').fillna(0)

    # 4. 추이 그래프 그리기
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("🗓️ 날짜별 종합효율(OEE) 추이")
        daily_oee = df.groupby('날짜')[oee_col].mean().reset_index()
        fig1 = px.line(daily_oee, x='날짜', y=oee_col, markers=True)
        st.plotly_chart(fig1, use_container_width=True)

    with col2:
        st.subheader("⏱️ 날짜별 비가동 시간 추이")
        daily_stop = df.groupby('날짜')[stop_col].sum().reset_index()
        fig2 = px.bar(daily_stop, x='날짜', y=stop_col)
        st.plotly_chart(fig2, use_container_width=True)

    # 5. 설비별 상세 분석
    st.subheader("🔍 설비별 성적표")
    selected_machine = st.selectbox("분석할 설비를 선택하세요", df[machine_col].unique())
    machine_df = df[df[machine_col] == selected_machine]
    
    fig3 = px.area(machine_df, x='날짜', y=oee_col, title=f"{selected_machine} 효율 변화")
    st.plotly_chart(fig3, use_container_width=True)

    st.write("---")
    st.dataframe(df)
