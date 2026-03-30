import streamlit as st
import pandas as pd
import plotly.express as px

# 1. 웹 화면 설정
st.set_page_config(page_title="사출 생산성 분석기", layout="wide")
st.title("🏭 일일 생산성 및 오픈이슈 분석 리포트")

# 2. 파일 업로드 기능
uploaded_file = st.file_uploader("정리하신 엑셀 파일을 업로드하세요 (XLSX)", type=['xlsx'])

if uploaded_file:
    # 엑셀 읽기 (관리자님의 03.26 양식 기준)
    df = pd.read_excel(uploaded_file, header=1) # 두 번째 줄이 실제 제목인 경우
    
    # 데이터 정리 (공백 제거 등)
    df.columns = [c.replace('\n', ' ') for c in df.columns]
    
    # 3. 핵심 지표 계산 (OEE 요약)
    avg_oee = df['종합 효율'].mean()
    st.metric("오늘의 평균 종합효율(OEE)", f"{avg_oee*100:.1f}%")

    # 4. 오픈 이슈 하이라이트
    st.subheader("⚠️ 집중 관리 필요 이슈 (OPEN ISSUE)")
    issue_df = df[df['OPEN ISSUE'].notna()]
    if not issue_df.empty:
        for _, row in issue_df.iterrows():
            st.error(f"**[{row['설비명']}]** : {row['OPEN ISSUE']}")
    else:
        st.success("오늘 발생한 특이사항이 없습니다.")

    # 5. 시각화 그래프
    st.subheader("📈 설비별 가동 현황")
    fig = px.bar(df, x='설비명', y='종합 효율', color='종합 효율',
                 color_continuous_scale='RdYlGn', range_y=[0, 1])
    st.plotly_chart(fig, use_container_width=True)

    st.write("---")
    st.dataframe(df) # 상세 표 보여주기