import streamlit as st
import pandas as pd
import plotly.express as px
import re
import io
import unicodedata

# 1. 웹 화면 설정
st.set_page_config(page_title="사출 생산성 통합 분석기", layout="wide")
st.title("📊 사출 생산성 통합 분석 리포트")

# 2. 파일 업로드
uploaded_files = st.file_uploader("분석할 파일들을 선택하세요 (여러 개 선택 가능)", type=['xlsx', 'csv'], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    
    for file in uploaded_files:
        try:
            file_bytes = file.read()
            if file.name.endswith('.csv'):
                raw_df = pd.read_csv(io.BytesIO(file_bytes))
            else:
                raw_df = pd.read_excel(io.BytesIO(file_bytes))
            
            # [헤더 자동 찾기]
            header_row = 0
            for i, row in raw_df.iterrows():
                row_str = " ".join(row.astype(str))
                if '설비' in row_str:
                    header_row = i + 1
                    break
            
            # 헤더 위치 반영 읽기
            if file.name.endswith('.csv'):
                temp_df = pd.read_csv(io.BytesIO(file_bytes), skiprows=header_row-1 if header_row > 0 else 0)
            else:
                temp_df = pd.read_excel(io.BytesIO(file_bytes), skiprows=header_row-1 if header_row > 0 else 0)
            
            # 🚨 [가장 강력한 방어막 1] 
            # 항목명 정화 및 Mac/Windows 글자 깨짐(유니코드) 완벽 통일
            temp_df.columns = [unicodedata.normalize('NFC', " ".join(str(c).split())) for c in temp_df.columns]
            
            name_map = {
                '작업장 [설비]': '설비명', '품목명': '품명', '합계': '합계수량', '합게수량': '합계수량', 
                '종합 효율': '종합효율', '목표 효율': '목표효율', 'Unnamed: 14': 'OPEN ISSUE'
            }
            temp_df = temp_df.rename(columns=name_map)
            
            # 🚨 [가장 강력한 방어막 2] 
            # 숨겨진 칸이나 중복된 이름이 발견되면, 무조건 첫 번째 열만 가져오도록 강제 단일화
            clean_temp = []
            for col in temp_df.columns.unique():
                col_data = temp_df[col]
                if isinstance(col_data, pd.DataFrame):
                    clean_temp.append(col_data.iloc[:, 0].rename(col))
                else:
                    clean_temp.append(col_data)
            temp_df = pd.concat(clean_temp, axis=1)
            
            # 품명 누락 대비 안전장치
            if '품명' not in temp_df.columns:
                temp_df['품명'] = ""
            
            # 생산일 추출
            clean_date = file.name.split('.')[0].replace('일일 생산성_', '').replace('사출생산팀 일일 생산성자료_', '')
            temp_df['생산일'] = clean_date
            
            # 불필요한 합계 행 제거
            if '설비명' in temp_df.columns:
                temp_df = temp_df[~temp_df['설비명'].astype(str).str.contains('TOTAL|합계|GRAND', na=False, case=False)]
            
            all_data.append(temp_df)
        except Exception as e:
            st.error(f"파일 읽기 오류 ({file.name}): {e}")

    if all_data:
        df = pd.concat(all_data, ignore_index=True)
        
        # 🚨 [가장 강력한 방어막 3] 파일 병합 후에도 중복 열 재확인
        clean_cols = []
        for col in df.columns.unique():
            col_data = df[col]
            if isinstance(col_data, pd.DataFrame):
                clean_cols.append(col_data.iloc[:, 0].rename(col))
            else:
                clean_cols.append(col_data)
        df = pd.concat(clean_cols, axis=1)
        
        # 숫자 데이터 변환
        num_cols = ['양품수량', '불량수량', '합계수량', '투입시간', '가동시간', '비가동시간', '정미시간', '종합효율', '목표효율', '양품율', '성능가동율', '시간가동율']
        for col in num_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # ---------------------------------------------------------
        # [사이드바 필터]
        # ---------------------------------------------------------
        st.sidebar.header("🎯 정밀 필터링")
        selected_machines = st.sidebar.multiselect("설비 선택", sorted(df['설비명'].unique()), default=sorted(df['설비명'].unique()))
        
        df['품명_필터'] = df['품명'].fillna("").replace(0, "").astype(str)
        actual_prods = sorted([p for p in df['품명_필터'].unique() if p.strip() != "" and p != "nan"])
        selected_prod = st.sidebar.selectbox("품목 선택", ["전체 품목"] + actual_prods)

        f_df = df[df['설비명'].isin(selected_machines)]
        if selected_prod != "전체 품목":
            f_df = f_df[f_df['품명_필터'] == selected_prod]

        # ---------------------------------------------------------
        # [그래프 분석]
        # ---------------------------------------------------------
        tab1, tab2 = st.tabs(["📈 일별 추이 분석", "🔍 상세 데이터 분석"])

        with tab1:
            c1, c2 = st.columns(2)
            with c1:
                st.subheader("일별 평균 종합효율(OEE)")
                active_oee = f_df[f_df['종합효율'] > 0]
                if not active_oee.empty:
                    daily_oee = active_oee.groupby('생산일')['종합효율'].mean().reset_index()
                    fig1 = px.line(daily_oee, x='생산일', y='종합효율', markers=True, text=daily_oee['종합효율'].apply(lambda x: f'{x:.1%}'))
                    fig1.update_layout(yaxis_tickformat='.0%', yaxis_range=[0, 1.1])
                    st.plotly_chart(fig1, use_container_width=True)
            with c2:
                st.subheader("일별 총 비가동 시간(분)")
                daily_stop = f_df.groupby('생산일')['비가동시간'].sum().reset_index()
                fig2 = px.bar(daily_stop, x='생산일', y='비가동시간', text_auto='.1f', color_discrete_sequence=['#FF4B4B'])
                st.plotly_chart(fig2, use_container_width=True)

        with tab2:
            st.subheader("설비별 가동 효율 변화")
            if not f_df[f_df['종합효율'] > 0].empty:
                fig3 = px.line(f_df[f_df['종합효율'] > 0], x='생산일', y='종합효율', color='설비명', markers=True)
                fig3.update_layout(yaxis_tickformat='.0%', yaxis_range=[0, 1.1])
                st.plotly_chart(fig3, use_container_width=True)

        # ---------------------------------------------------------
        # [데이터 표] - 비가동 빈칸 처리 및 요청하신 순서 정렬
        # ---------------------------------------------------------
        st.write("---")
        st.subheader("📂 통합 원본 데이터 (검토용)")
        
        display_df = f_df.sort_values(by=['생산일', '설비명']).copy()
        
        # 항목 순서 재배치
        target_order = [
            '생산일', '설비명', '품명', '종합효율', '성능가동율', '시간가동율', '양품율', '목표효율',
            '투입시간', '가동시간', '비가동시간', '정미시간', '양품수량', '불량수량', '합계수량', 'OPEN ISSUE'
        ]
        display_df = display_df[[c for c in target_order if c in display_df.columns]]

        # 비가동 설비 빈칸 처리 및 서식 적용
        def finalize_row(row):
            is_idle = pd.isna(row['품명']) or str(row['품명']).strip() in ['0', '0.0', '', 'nan', 'NaN']
            res = row.copy()
            for col in res.index:
                if is_idle and col not in ['생산일', '설비명']:
                    res[col] = ""
                else:
                    if col in ['종합효율', '성능가동율', '시간가동율', '양품율', '목표효율']:
                        try: res[col] = f"{float(res[col]):.1%}" if res[col] != "" else ""
                        except: pass
                    elif col in ['양품수량', '불량수량', '합계수량']:
                        try: res[col] = f"{int(float(res[col])):,.0f}" if res[col] != "" else ""
                        except: pass
                    elif col in ['투입시간', '가동시간', '비가동시간', '정미시간']:
                        try: res[col] = f"{float(res[col]):.1f}" if res[col] != "" else ""
                        except: pass
            if is_idle: res['품명'] = ""
            return res

        st.dataframe(display_df.apply(finalize_row, axis=1), use_container_width=True)
else:
    st.info("왼쪽 상단에서 생산성 파일을 업로드해 주세요.")
