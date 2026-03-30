import streamlit as st
import pandas as pd
import plotly.express as px

# 1. 웹 화면 설정
st.set_page_config(page_title="사출 생산성 분석기", layout="wide")
st.title("🏭 일일 생산성 및 오픈이슈 분석 리포트")

# 2. 파일 업로드 기능 (XLSX와 CSV 모두 지원)
uploaded_file = st.file_uploader("정리하신 엑셀 파일을 업로드하세요", type=['xlsx', 'csv'])

if uploaded_file:
    # 파일 확장자에 따라 읽는 방식 자동 선택
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        # 엑셀의 첫 번째 줄부터 읽어옵니다
        df = pd.read_excel(uploaded_file)
    
    # [핵심] 제목(컬럼명) 정리하기: 줄바꿈 제거, 앞뒤 공백 제거
    df.columns = [" ".join(str(c).split()) for c in df.columns]
    
    # '종합 효율'이라는 글자가 포함된 칸을 자동으로 찾습니다
    target_col = [c for c in df.columns if '종합 효율' in c or '종합효율' in c]
    
    if target_col:
        col_name = target_col[0]
        # 숫자 데이터로 변환 (오류 방지)
        df[col_name] = pd.to_numeric(df[col_name], errors='coerce').fillna(0)
        
        # 3. 핵심 지표 계산
        avg_oee = df[col_name].mean()
        st.metric("오늘의 평균 종합효율(OEE)", f"{avg_oee*100:.1f}%")

        # 4. 오픈 이슈 하이라이트 (이름에 'OPEN'이 들어간 칸 찾기)
        issue_col = [c for c in df.columns if 'OPEN' in c.upper() or '이슈' in c]
        if issue_col:
            st.subheader("⚠️ 집중 관리 필요 이슈 (OPEN ISSUE)")
            issue_df = df[df[issue_col[0]].notna() & (df[issue_col[0]] != "")]
            if not issue_df.empty:
                for _, row in issue_df.iterrows():
                    # 설비명 칸 찾기
                    machine_col = [c for c in df.columns if '설비' in c][0]
                    st.error(f"**[{row[machine_col]}]** : {row[issue_col[0]]}")
            else:
                st.success("오늘 발생한 특이사항이 없습니다.")

        # 5. 시각화 그래프
        st.subheader("📈 설비별 가동 현황")
        machine_col = [c for c in df.columns if '설비' in c][0]
        fig = px.bar(df, x=machine_col, y=col_name, color=col_name,
                     color_continuous_scale='RdYlGn', range_y=[0, 1])
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.error("엑셀 파일에서 '종합 효율' 칸을 찾을 수 없습니다. 양식을 확인해주세요.")

    st.write("---")
    st.write("📂 업로드된 데이터 상세 보기")
    st.dataframe(df)
