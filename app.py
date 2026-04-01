import streamlit as st
import pandas as pd
import plotly.express as px
import re

# 1. 웹 화면 설정
st.set_page_config(page_title="사출 생산성 통합 분석기", layout="wide")
st.title("📊 사출 생산성 통합 분석 리포트")

# 2. 파일 업로드
uploaded_files = st.file_uploader("분석할 파일들을 선택하세요 (여러 개 선택 가능)", type=['xlsx', 'csv'], accept_multiple_files=True)

# 🚨 [수정 5] 합계수량 -> 총 생산수량 명칭 변경
target_cols = ['생산일', '설비명', '품명', '양품수량', '불량수량', '총 생산수량', '투입시간', '가동시간', '비가동시간', '정미시간', '양품율', '성능가동율', '시간가동율', '종합효율', '목표효율', 'OPEN ISSUE']

if uploaded_files:
    all_records = []
    daily_totals_data = {} 
    
    for file in uploaded_files:
        try:
            if file.name.endswith('.csv'):
                temp_df = pd.read_csv(file)
            else:
                temp_df = pd.read_excel(file)
            
            temp_df.columns = [str(c).replace('\n', '').replace('\r', '').strip() for c in temp_df.columns]
            
            # 항목명 강제 통일
            name_map = {
                '작업장 [설비]': '설비명', '작업장[설비]': '설비명', '품목명': '품명',
                '합계': '총 생산수량', '합게수량': '총 생산수량', '종합 효율': '종합효율', '목표 효율': '목표효율'
            }
            temp_df = temp_df.rename(columns=name_map)
            
            for col in temp_df.columns:
                if 'Unnamed' in col or 'ISSUE' in col.upper():
                    temp_df = temp_df.rename(columns={col: 'OPEN ISSUE'})
                    break
            
            # 생산일 추출 (예: 260331)
            date_match = re.search(r'\d{8}', file.name)
            if date_match:
                clean_date = date_match.group()[2:]
            else:
                clean_date = file.name.split('.')[0]
            
            # M47 셀(합계) 종합효율 원본 추출
            daily_total_oee = None
            for _, row in temp_df.iterrows():
                machine_val = str(row.get('설비명', ''))
                if isinstance(machine_val, pd.Series): machine_val = str(machine_val.iloc[0])
                
                if 'TOTAL' in machine_val.upper() or '합계' in machine_val or 'GRAND' in machine_val.upper():
                    val = row.get('종합효율', None)
                    if isinstance(val, pd.Series): val = val.iloc[0]
                    try: daily_total_oee = float(val)
                    except: pass
                    break
            
            if daily_total_oee is None:
                try:
                    val = temp_df['종합효율'].iloc[45]
                    if isinstance(val, pd.Series): val = val.iloc[0]
                    daily_total_oee = float(val)
                except:
                    daily_total_oee = 0
            
            daily_totals_data[clean_date] = daily_total_oee

            # 핀셋 추출 로직
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
                        if isinstance(val, pd.Series): val = val.iloc[0]
                        record[col] = val
                    else:
                        record[col] = None
                        
                all_records.append(record)
                
        except Exception as e:
            st.error(f"파일 읽기 오류 ({file.name}): {e}")

    if all_records:
        df = pd.DataFrame(all_records)
        
        daily_df = pd.DataFrame(list(daily_totals_data.items()), columns=['생산일', '공장종합효율']).sort_values(by='생산일')
        
        num_cols = ['양품수량', '불량수량', '총 생산수량', '투입시간', '가동시간', '비가동시간', '정미시간', '종합효율', '목표효율', '양품율', '성능가동율', '시간가동율']
        for col in num_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # ---------------------------------------------------------
        # [사이드바 필터]
        # ---------------------------------------------------------
        st.sidebar.header("🎯 정밀 필터링")
        
        df['설비명'] = df['설비명'].fillna("").astype(str)
        all_machines = sorted([m for m in df['설비명'].unique() if m.strip() != ""])
        
        selected_machines = st.sidebar.multiselect(
            "설비 선택 (클릭하여 추가하세요)", 
            all_machines, 
            default=[], 
            placeholder="여기를 클릭하여 설비를 고르세요"
        )
        
        # 🚨 [수정 6] 선택한 설비에서 생산된 제품만 목록에 표시되도록 연동!
        if len(selected_machines) == 0:
            pool_df = df.copy()
        else:
            pool_df = df[df['설비명'].isin(selected_machines)].copy()
            
        pool_df['품명_필터'] = pool_df['품명'].fillna("").astype(str).str.strip()
        pool_df['품명_필터'] = pool_df['품명_필터'].replace(['0', '0.0', 'nan', 'NaN', 'None'], "")
        actual_prods = sorted([p for p in pool_df['품명_필터'].unique() if p != ""])
        
        selected_prod = st.sidebar.selectbox("품목 선택 (해당 설비 생산품)", ["전체 품목"] + actual_prods)

        f_df = pool_df.copy()
        if selected_prod != "전체 품목":
            f_df = f_df[f_df['품명_필터'] == selected_prod]

        # ---------------------------------------------------------
        # [그래프 분석 탭]
        # ---------------------------------------------------------
        tab1, tab2 = st.tabs(["📈 일별 추이 분석", "📝 일별 OPEN ISSUE 분석"])

        with tab1:
            # 🚨 [수정 1 & 수정 6 대응] 필터 적용 여부에 따라 그래프가 자동으로 바뀝니다.
            if len(selected_machines) == 0 and selected_prod == "전체 품목":
                st.subheader("🗓️ 사출생산팀 일별 종합 효율(%)")
                plot_df = daily_df.copy()
                y_val = '공장종합효율'
            else:
                st.subheader("🗓️ 선택 조건(설비/품목) 일별 평균 종합 효율(%)")
                active_oee = f_df[f_df['종합효율'] > 0]
                plot_df = active_oee.groupby('생산일')['종합효율'].mean().reset_index().sort_values(by='생산일')
                y_val = '종합효율'

            if not plot_df.empty:
                # 🚨 [수정 2] 가로축(X축) 날짜 아래에 효율(%) 텍스트를 겹치지 않게 두 줄로 배치
                plot_df['x_label'] = plot_df.apply(lambda row: f"{row['생산일']}<br><span style='font-size:12px;color:gray;'>({row[y_val]:.1%})</span>", axis=1)
                
                fig1 = px.line(plot_df, x='x_label', y=y_val, markers=True, text=plot_df[y_val].apply(lambda x: f'{x:.1%}'))
                
                # 🚨 [수정 2] 그래프 선 위에 있는 숫자도 천장에 닿지 않도록 공간 확보 및 위치 조정
                fig1.update_traces(textposition="top center", textfont=dict(size=14, color="#1f77b4")) 
                fig1.update_xaxes(type='category', title="") 
                fig1.update_yaxes(title="종합효율", tickformat='.0%', range=[0, 1.2]) # 천장을 120%까지 높임
                fig1.update_layout(height=450, margin=dict(t=50))
                st.plotly_chart(fig1, use_container_width=True)
            else:
                st.info("종합효율 데이터가 없습니다.")
            
            st.write("---")
            
            st.subheader("⏱️ 일별 총 비가동 시간(분)")
            daily_stop = f_df.groupby('생산일')['비가동시간'].sum().reset_index().sort_values(by='생산일')
            fig2 = px.bar(daily_stop, x='생산일', y='비가동시간', text_auto='.1f', color_discrete_sequence=['#FF4B4B'])
            fig2.update_traces(textposition="outside") # 막대그래프 숫자도 밖으로 빼냄
            fig2.update_xaxes(type='category', title="")
            fig2.update_yaxes(title="비가동시간 (분)")
            fig2.update_layout(height=400, margin=dict(t=50))
            st.plotly_chart(fig2, use_container_width=True)

        with tab2:
            # 🚨 [수정 3] 어지러운 상세 그래프를 없애고 '일별 OPEN ISSUE 현황' 표로 대체!
            st.subheader("📝 일별 특이사항(OPEN ISSUE) 현황")
            st.markdown("선택하신 설비/품목에서 발생한 주요 이슈와 당일의 효율을 나란히 확인하세요.")
            
            def has_issue(val):
                val_str = str(val).strip()
                return val_str != "" and val_str.lower() not in ['0', '0.0', 'nan', 'none']
                
            issue_df = f_df[f_df['OPEN ISSUE'].apply(has_issue)].copy()
            
            if not issue_df.empty:
                issue_display = issue_df[['생산일', '설비명', '품명', '종합효율', 'OPEN ISSUE']].sort_values(by=['생산일', '설비명'])
                issue_display['종합효율'] = issue_display['종합효율'].apply(lambda x: f"{float(x):.1%}" if pd.notnull(x) and str(x).strip()!="" else "")
                
                st.dataframe(issue_display, hide_index=True, use_container_width=True)
            else:
                st.info("선택된 조건에 해당하는 특이사항(OPEN ISSUE)이 없습니다.")

        # ---------------------------------------------------------
        # [통합 원본 데이터 표]
        # ---------------------------------------------------------
        st.write("---")
        st.subheader("📂 사출생산팀 일일 생산성 자료")
        
        display_df = f_df.sort_values(by=['생산일', '설비명']).copy()
        
        # 🚨 [수정 4] 관리자님이 요청하신 11개 항목 순서 재배치
        target_order = [
            '생산일', '설비명', '품명', '종합효율', '양품율', '성능가동율', '시간가동율', 
            '총 생산수량', '양품수량', '불량수량', 'OPEN ISSUE'
        ]
        display_df = display_df[[c for c in target_order if c in display_df.columns]]

        def finalize_row(row):
            val = str(row.get('품명', '')).strip()
            is_idle = (val == '') or (val in ['0', '0.0', 'nan', 'NaN', 'None'])
            
            res = row.copy()
            for col in res.index:
                if is_idle and col not in ['생산일', '설비명']:
                    res[col] = ""
                else:
                    if col in ['종합효율', '성능가동율', '시간가동율', '양품율']:
                        try: res[col] = f"{float(res[col]):.1%}" if pd.notnull(res[col]) and res[col] != "" else ""
                        except: pass
                    elif col in ['양품수량', '불량수량', '총 생산수량']:
                        try: res[col] = f"{int(float(res[col])):,.0f}" if pd.notnull(res[col]) and res[col] != "" else ""
                        except: pass
            if is_idle: 
                res['품명'] = ""
            for k, v in res.items():
                if pd.isna(v) or v == 'None': res[k] = ""
            return res

        final_table = display_df.apply(finalize_row, axis=1)
        
        # 🚨 [수정 2] 표 상단 헤더를 그룹화(MultiIndex)하여 구분 / 생산성 / 생산실적 분류!
        multi_cols = [
            ('구분', '생산일'), ('구분', '설비명'), ('구분', '품명'),
            ('생산성', '종합효율'), ('생산성', '양품율'), ('생산성', '성능가동율'), ('생산성', '시간가동율'),
            ('생산실적', '총 생산수량'), ('생산실적', '양품수량'), ('생산실적', '불량수량'),
            ('OPEN ISSUE', 'OPEN ISSUE')
        ]
        final_table.columns = pd.MultiIndex.from_tuples(multi_cols)
        
        st.dataframe(final_table, hide_index=True, use_container_width=True)
else:
    st.info("왼쪽 상단에서 생산성 파일을 업로드해 주세요.")
