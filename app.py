import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re

# 1. 웹 화면 및 폰트 강제 설정
st.set_page_config(page_title="사출생산팀 일일 생산성 정밀 분석", layout="wide")

st.markdown("""
<style>
    /* 전체 폰트를 세련된 'Pretendard' 또는 '맑은 고딕'으로 변경 */
    html, body, [class*="css"] {
        font-family: 'Pretendard', 'Noto Sans KR', 'Malgun Gothic', sans-serif !important;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 사출생산팀 일일 생산성 정밀 분석")

# 2. 파일 업로드
uploaded_files = st.file_uploader("분석할 파일들을 선택하세요 (여러 개 선택 가능)", type=['xlsx', 'csv'], accept_multiple_files=True)

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
            
            name_map = {
                '작업장 [설비]': '설비명', '작업장[설비]': '설비명', '품목명': '품명',
                '합계': '총 생산수량', '합게수량': '총 생산수량', '종합 효율': '종합효율', '목표 효율': '목표효율'
            }
            temp_df = temp_df.rename(columns=name_map)
            
            for col in temp_df.columns:
                if 'Unnamed' in col or 'ISSUE' in col.upper():
                    temp_df = temp_df.rename(columns={col: 'OPEN ISSUE'})
                    break
            
            date_match = re.search(r'\d{8}', file.name)
            if date_match:
                clean_date = date_match.group()[2:]
            else:
                clean_date = file.name.split('.')[0]
            
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

        # 줄바꿈 정밀 적용
        def format_issue(text):
            val = str(text).strip()
            if val in ['', '0', '0.0', 'nan', 'NaN', 'None']:
                return ""
            val = val.replace('\r\n', '\n')
            val = re.sub(r'(?<!\n)\*', '\n*', val)
            val = re.sub(r'(?<!\n)-\.', '\n-.', val)
            val = re.sub(r'(?<!\n)→', '\n→', val)
            return val.strip()

        df['OPEN ISSUE'] = df['OPEN ISSUE'].apply(format_issue)

        # ---------------------------------------------------------
        # [사이드바 필터]
        # ---------------------------------------------------------
        st.sidebar.header("🎯 정밀 필터링")
        
        df['설비명'] = df['설비명'].fillna("").astype(str)
        all_machines = sorted([m for m in df['설비명'].unique() if m.strip() != ""])
        
        selected_machines = st.sidebar.multiselect("설비 선택 (클릭하여 추가)", all_machines, default=[])
        
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
        tab1, tab2 = st.tabs(["📈 일별 추이 분석", "📝 일별 OPEN ISSUE 현황"])

        with tab1:
            is_factory_view = (len(selected_machines) == 0 and selected_prod == "전체 품목")
            
            if is_factory_view:
                st.subheader("🗓️ 사출생산팀 일별 종합 효율(%)")
                plot_df = daily_df.copy()
                y_val = '공장종합효율'
            else:
                st.subheader("🗓️ 선택 조건(설비/품목) 일별 평균 종합 효율(%)")
                active_oee = f_df[f_df['종합효율'] > 0]
                plot_df = active_oee.groupby('생산일')['종합효율'].mean().reset_index().sort_values(by='생산일')
                y_val = '종합효율'

            if not plot_df.empty:
                plot_df['x_label'] = plot_df.apply(lambda row: f"{row['생산일']}<br><span style='font-size:11px;color:gray;'>({row[y_val]:.1%})</span>", axis=1)
                
                fig1 = px.line(plot_df, x='x_label', y=y_val, text=plot_df[y_val].apply(lambda x: f'{x:.1%}'))
                
                # 🚨 [수정 1] 메인 라인 디자인 (잘림 방지 cliponaxis=False 적용)
                fig1.update_traces(
                    mode='lines+markers+text',
                    line=dict(shape='spline', width=3, color='#1f77b4'), 
                    marker=dict(size=10, color='white', line=dict(width=2, color='#1f77b4')),
                    textposition="top center", 
                    textfont=dict(size=14, color="#1f77b4", weight="bold"),
                    fill='tozeroy', fillcolor='rgba(31, 119, 180, 0.1)',
                    cliponaxis=False
                )
                
                # 🚨 [수정 1] 목표 효율(86%)을 범례(Legend)로 생성하여 겹침 원천 차단
                if is_factory_view:
                    fig1.add_trace(go.Scatter(
                        x=plot_df['x_label'], 
                        y=[0.86] * len(plot_df), 
                        mode='lines', 
                        name='목표 효율 (86%)', 
                        line=dict(color='#FF4B4B', dash='dash', width=2),
                        hoverinfo='skip'
                    ))
                
                # 좌우 여백을 넓혀 끝부분 글자 잘림 방지
                fig1.update_xaxes(type='category', title="", showgrid=False, range=[-0.5, len(plot_df)-0.5]) 
                fig1.update_yaxes(title="종합효율", tickformat='.0%', range=[0, 1.2], showgrid=True, gridcolor='rgba(200,200,200,0.2)') 
                fig1.update_layout(height=450, margin=dict(l=40, r=40, t=40, b=40), plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', legend=dict(yanchor="top", y=1.1, xanchor="right", x=1))
                st.plotly_chart(fig1, use_container_width=True)
            else:
                st.info("종합효율 데이터가 없습니다.")
            
            st.write("---")
            
            st.subheader("⏱️ 일별 총 비가동 시간(분)")
            daily_stop = f_df.groupby('생산일')['비가동시간'].sum().reset_index().sort_values(by='생산일')
            fig2 = px.bar(daily_stop, x='생산일', y='비가동시간', text_auto='.1f')
            
            fig2.update_traces(
                marker=dict(color='#FF4B4B', opacity=0.8), 
                textposition="outside",
                textfont=dict(size=13, weight="bold"),
                cliponaxis=False
            )
            fig2.update_xaxes(type='category', title="", showgrid=False, range=[-0.5, len(daily_stop)-0.5])
            fig2.update_yaxes(title="비가동시간 (분)", showgrid=True, gridcolor='rgba(200,200,200,0.2)')
            fig2.update_layout(height=400, margin=dict(l=40, r=40, t=40, b=40), plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig2, use_container_width=True)

        # ---------------------------------------------------------
        # [HTML 렌더링 헬퍼 함수] (스트림릿 표 한계 극복)
        # ---------------------------------------------------------
        def render_custom_html_table(styler_obj, is_multiindex=False):
            html = styler_obj.hide(axis="index").to_html()
            
            if is_multiindex:
                # 🚨 [수정 2 & 4] OPEN ISSUE 2칸 병합 처리 (정규식 활용 HTML 수정)
                html = re.sub(r'<th class="col_heading level0 col10".*?>OPEN ISSUE</th>', r'<th class="col_heading level0 col10" rowspan="2" style="vertical-align: middle;">OPEN ISSUE</th>', html)
                html = re.sub(r'<th class="col_heading level1 col10".*?></th>', '', html)
            
            # 🚨 [수정 3] 중앙 정렬 및 줄바꿈(pre-wrap) 강제 적용 CSS
            wrapped_html = f"""
            <div style="width: 100%; max-height: 600px; overflow: auto; border: 1px solid #e6e9ef; border-radius: 5px; box-shadow: 0 2px 5px rgba(0,0,0,0.05);">
                <style>
                    .my-perfect-table {{
                        width: 100%;
                        border-collapse: collapse;
                        background-color: white;
                        font-size: 13.5px;
                        color: #31333F;
                    }}
                    .my-perfect-table th {{
                        background-color: #f8f9fb;
                        border: 1px solid #e6e9ef;
                        padding: 10px;
                        text-align: center !important;
                        vertical-align: middle !important;
                        font-weight: 700;
                        position: sticky;
                        top: 0;
                        z-index: 1;
                    }}
                    .my-perfect-table thead tr:nth-child(2) th {{
                        top: 40px;
                    }}
                    .my-perfect-table td {{
                        border: 1px solid #e6e9ef;
                        padding: 8px 10px;
                        text-align: center !important;
                        vertical-align: middle !important;
                    }}
                    .my-perfect-table td:last-child {{
                        white-space: pre-wrap !important; 
                        text-align: left !important; 
                        min-width: 350px;
                        line-height: 1.5;
                    }}
                </style>
                {html.replace('<table', '<table class="my-perfect-table"')}
            </div>
            """
            st.markdown(wrapped_html, unsafe_allow_html=True)

        # ---------------------------------------------------------
        # [표 출력 영역]
        # ---------------------------------------------------------
        with tab2:
            st.subheader("📝 일별 특이사항(OPEN ISSUE) 현황")
            st.markdown("발생한 주요 이슈와 당일의 효율을 확인하세요. (클릭 없이 바로 줄바꿈되어 보입니다)")
            
            issue_df = f_df[f_df['OPEN ISSUE'] != ""].copy()
            if not issue_df.empty:
                issue_display = issue_df[['생산일', '설비명', '품명', '종합효율', 'OPEN ISSUE']].sort_values(by=['생산일', '설비명'])
                
                style_issue = pd.DataFrame('', index=issue_display.index, columns=issue_display.columns)
                for i, idx in enumerate(issue_display.index):
                    try:
                        oee = float(issue_df.loc[idx, '종합효율'])
                        target = float(issue_df.loc[idx, '목표효율'])
                        if oee < target:
                            style_issue.iloc[i, style_issue.columns.get_loc('종합효율')] = 'color: #FF4B4B; font-weight: bold;'
                    except: pass
                
                issue_display['종합효율'] = issue_display['종합효율'].apply(lambda x: f"{float(x):.1%}" if pd.notnull(x) and str(x).strip()!="" else "")
                
                styler = issue_display.style.apply(lambda _: style_issue, axis=None)
                render_custom_html_table(styler, is_multiindex=False)
            else:
                st.info("선택된 조건에 해당하는 특이사항(OPEN ISSUE)이 없습니다.")

        st.write("---")
        st.subheader("📂 사출생산팀 일일 생산성 자료")
        
        display_df = f_df.sort_values(by=['생산일', '설비명']).copy()
        
        target_order = [
            '생산일', '설비명', '품명', '종합효율', '양품율', '성능가동율', '시간가동율', 
            '총 생산수량', '양품수량', '불량수량', 'OPEN ISSUE'
        ]
        
        style_main = pd.DataFrame('', index=display_df.index, columns=target_order)
        for i in range(len(display_df)):
            try:
                oee = float(display_df.iloc[i]['종합효율'])
                target = float(display_df.iloc[i]['목표효율'])
                if oee < target:
                    style_main.iloc[i, style_main.columns.get_loc('종합효율')] = 'color: #FF4B4B; font-weight: bold;'
            except: pass

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
        
        multi_cols = [
            ('구분', '생산일'), ('구분', '설비명'), ('구분', '품명'),
            ('생산성', '종합효율'), ('생산성', '양품율'), ('생산성', '성능가동율'), ('생산성', '시간가동율'),
            ('생산실적', '총 생산수량'), ('생산실적', '양품수량'), ('생산실적', '불량수량'),
            ('OPEN ISSUE', '') # 🚨 아랫단을 빈칸으로 처리하고 HTML에서 병합합니다.
        ]
        final_table.columns = pd.MultiIndex.from_tuples(multi_cols)
        style_main.columns = pd.MultiIndex.from_tuples(multi_cols)
        
        final_styler = final_table.style.apply(lambda _: style_main, axis=None)
        
        # HTML 커스텀 렌더링 호출
        render_custom_html_table(final_styler, is_multiindex=True)

else:
    st.info("왼쪽 상단에서 생산성 파일을 업로드해 주세요.")
