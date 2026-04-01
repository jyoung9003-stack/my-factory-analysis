import streamlit as st
import pandas as pd
import plotly.express as px
import re

# 1. 웹 화면 설정
st.set_page_config(page_title="사출 생산성 통합 분석기", layout="wide")
st.title("📊 사출 생산성 통합 분석 리포트")

# 2. 파일 업로드
uploaded_files = st.file_uploader("분석할 파일들을 선택하세요 (여러 개 선택 가능)", type=['xlsx', 'csv'], accept_multiple_files=True)

# 🚨 반드시 추출해야 할 16개 항목
target_cols = ['생산일', '설비명', '품명', '양품수량', '불량수량', '합계수량', '투입시간', '가동시간', '비가동시간', '정미시간', '양품율', '성능가동율', '시간가동율', '종합효율', '목표효율', 'OPEN ISSUE']

if uploaded_files:
    all_records = []
    
    for file in uploaded_files:
        try:
            if file.name.endswith('.csv'):
                temp_df = pd.read_csv(file)
            else:
                temp_df = pd.read_excel(file)
            
            # 컬럼명 공백/줄바꿈 제거
            temp_df.columns = [str(c).replace('\n', '').replace('\r', '').strip() for c in temp_df.columns]
            
            # 항목명 강제 통일
            name_map = {
                '작업장 [설비]': '설비명', '작업장[설비]': '설비명', '품목명': '품명',
                '합계': '합계수량', '합게수량': '합계수량', '종합 효율': '종합효율', '목표 효율': '목표효율'
            }
            temp_df = temp_df.rename(columns=name_map)
            
            for col in temp_df.columns:
                if 'Unnamed' in col or 'ISSUE' in col.upper():
                    temp_df = temp_df.rename(columns={col: 'OPEN ISSUE'})
                    break
            
            # 🚨 [수정 1] 생산일 정밀 추출: 파일명 꼬리(.csv, - Sheet 등)를 다 자르고 '_' 뒤의 날짜만 뽑기
            name_only = re.sub(r'\.xlsx?.*$', '', file.name) # .xlsx 뒤에 붙은 찌꺼기 제거
            name_only = re.sub(r'\.csv$', '', name_only)     # .csv 제거
            
            if '_' in name_only:
                clean_date = name_only.split('_')[-1].strip()
            else:
                clean_date = name_only.strip()
            
            # 핀셋 추출 로직 (에러 방지)
            for _, row in temp_df.iterrows():
                machine_val = str(row.get('설비명', ''))
                if isinstance(machine_val, pd.Series): 
                    machine_val = str(machine_val.iloc[0])
                if 'TOTAL' in machine_val.upper() or '합계' in machine_val or 'GRAND' in machine_val.upper():
                    continue
                    
                record = {'생산일': clean_date}
                
                for col in target_cols:
                    if col == '생산일': continue
                    if col in temp_df.columns:
                        val = row[col]
                        if isinstance(val, pd.Series):
                            val = val.iloc[0]
                        record[col] = val
                    else:
                        record[col] = None
                        
                all_records.append(record)
                
        except Exception as e:
            st.error(f"파일 읽기 오류 ({file.name}): {e}")

    if all_records:
        df = pd.DataFrame(all_records)
        
        num_cols = ['양품수량', '불량수량', '합계수량', '투입시간', '가동시간', '비가동시간', '정미시간', '종합효율', '목표효율', '양품율', '성능가동율', '시간가동율']
        for col in num_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # ---------------------------------------------------------
        # [사이드바 필터]
        # ---------------------------------------------------------
        st.sidebar.header("🎯 정밀 필터링")
        
        df['설비명'] = df['설비명'].fillna("").astype(str)
        all_machines = sorted([m for m in df['설비명'].unique() if m.strip() != ""])
        
        # 🚨 [수정 2] 설비 선택: 다 지우는 방식에서 -> 빈칸에서 클릭해 추가하는 방식으로 변경
        selected_machines = st.sidebar.multiselect(
            "설비 선택 (클릭하여 추가하세요)", 
            all_machines, 
            default=[], 
            placeholder="여기를 클릭하여 설비를 고르세요"
        )
        
        df['품명_필터'] = df['품명'].fillna("").astype(str).str.strip()
        df['품명_필터'] = df['품명_필터'].replace(['0', '0.0', 'nan', 'NaN', 'None'], "")
        actual_prods = sorted([p for p in df['품명_필터'].unique() if p != ""])
        selected_prod = st.sidebar.selectbox("품목 선택", ["전체 품목"] + actual_prods)

        # 미선택 시 전체 데이터 보여주기
        if len(selected_machines) == 0:
            f_df = df.copy()
        else:
            f_df = df[df['설비명'].isin(selected_machines)].copy()
            
        if selected_prod != "전체 품목":
            f_df = f_df[f_df['품명_필터'] == selected_prod]

        # ---------------------------------------------------------
        # [그래프 분석 탭] - 세로로 넓게 배치
        # ---------------------------------------------------------
        tab1, tab2 = st.tabs(["📈 일별 추이 분석", "🔍 상세 데이터 분석"])

        with tab1:
            # 🚨 [수정 3] 좌우 분할(columns)을 없애고 위아래로 큼직하게 배치
            st.subheader("🗓️ 일별 평균 종합효율(OEE)")
            active_oee = f_df[f_df['종합효율'] > 0]
            if not active_oee.empty:
                daily_oee = active_oee.groupby('생산일')['종합효율'].mean().reset_index().sort_values(by='생산일')
                fig1 = px.line(daily_oee, x='생산일', y='종합효율', markers=True, text=daily_oee['종합효율'].apply(lambda x: f'{x:.1%}'))
                fig1.update_layout(yaxis_tickformat='.0%', yaxis_range=[0, 1.1], height=400) # 높이도 넉넉하게
                st.plotly_chart(fig1, use_container_width=True)
            else:
                st.info("가동된 설비의 효율 데이터가 없습니다.")
            
            st.write("---") # 위아래 그래프 구분선
            
            st.subheader("⏱️ 일별 총 비가동 시간(분)")
            daily_stop = f_df.groupby('생산일')['비가동시간'].sum().reset_index().sort_values(by='생산일')
            fig2 = px.bar(daily_stop, x='생산일', y='비가동시간', text_auto='.1f', color_discrete_sequence=['#FF4B4B'])
            fig2.update_layout(height=400)
            st.plotly_chart(fig2, use_container_width=True)

        with tab2:
            st.subheader("설비별 가동 효율 변화")
            if not f_df[f_df['종합효율'] > 0].empty:
                plot_df = f_df[f_df['종합효율'] > 0].sort_values(by='생산일')
                fig3 = px.line(plot_df, x='생산일', y='종합효율', color='설비명', markers=True)
                fig3.update_layout(yaxis_tickformat='.0%', yaxis_range=[0, 1.1], height=500)
                st.plotly_chart(fig3, use_container_width=True)
            else:
                st.info("가동된 설비의 효율 데이터가 없습니다.")

        # ---------------------------------------------------------
        # [통합 원본 데이터 표]
        # ---------------------------------------------------------
        st.write("---")
        st.subheader("📂 통합 원본 데이터 (검토용)")
        
        display_df = f_df.sort_values(by=['생산일', '설비명']).copy()
        display_df = display_df[target_cols] 

        def finalize_row(row):
            val = str(row.get('품명', '')).strip()
            is_idle = (val == '') or (val in ['0', '0.0', 'nan', 'NaN', 'None'])
            
            res = row.copy()
            for col in res.index:
                if is_idle and col not in ['생산일', '설비명']:
                    res[col] = ""
                else:
                    if col in ['종합효율', '성능가동율', '시간가동율', '양품율', '목표효율']:
                        try: res[col] = f"{float(res[col]):.1%}" if pd.notnull(res[col]) and res[col] != "" else ""
                        except: pass
                    elif col in ['양품수량', '불량수량', '합계수량']:
                        try: res[col] = f"{int(float(res[col])):,.0f}" if pd.notnull(res[col]) and res[col] != "" else ""
                        except: pass
                    elif col in ['투입시간', '가동시간', '비가동시간', '정미시간']:
                        try: res[col] = f"{float(res[col]):.1f}" if pd.notnull(res[col]) and res[col] != "" else ""
                        except: pass
            if is_idle: 
                res['품명'] = ""
            for k, v in res.items():
                if pd.isna(v) or v == 'None': res[k] = ""
            return res

        final_table = display_df.apply(finalize_row, axis=1)
        st.dataframe(final_table, use_container_width=True)
else:
    st.info("왼쪽 상단에서 생산성 파일을 업로드해 주세요.")
