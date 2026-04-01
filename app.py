import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
import os
import html
from collections import Counter

# 1. 웹 화면 및 폰트 설정
st.set_page_config(page_title="사출생산팀 일일 생산성 정밀 분석", layout="wide")

st.markdown("""
<style>
    html, body, [class*="css"] {
        font-family: 'Pretendard', 'Noto Sans KR', 'Malgun Gothic', sans-serif !important;
        background-color: #F8F9FA;
    }
    .metric-card {
        background-color: white; border: 1px solid #E9ECEF; border-radius: 8px;
        padding: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); text-align: center; margin-bottom: 20px;
    }
    .metric-title { font-size: 14px; color: #6C757D; font-weight: bold; margin-bottom: 5px; }
    .metric-value.best { font-size: 20px; color: #1F77B4; font-weight: 900; }
    .metric-value.worst { font-size: 20px; color: #FF4B4B; font-weight: 900; }
    .metric-date { font-size: 12px; color: #ADB5BD; }
</style>
""", unsafe_allow_html=True)

# 2. 로고 및 타이틀
col1, col2 = st.columns([1, 10])
with col1:
    logo_path = "듀링로고_가로형_빨강_JPG.jpg"
    if os.path.exists(logo_path): st.image(logo_path, width=100)
    else: st.markdown("<h2 style='color: #FF2A2A; font-weight: 900; margin-top: 10px;'>DÜRING</h2>", unsafe_allow_html=True)
with col2:
    st.markdown("<h1 style='margin-top: 0px; color: #212529;'>사출생산팀 일일 생산성 정밀 분석</h1>", unsafe_allow_html=True)

# 3. 데이터 수집 및 정제
uploaded_files = st.file_uploader("분석할 파일들을 선택하세요 (여러 개 선택 가능)", type=['xlsx', 'csv'], accept_multiple_files=True)

target_cols = ['생산일', '설비명', '품명', '양품수량', '불량수량', '총 생산수량', '투입시간', '가동시간', '비가동시간', '정미시간', '양품율', '성능가동율', '시간가동율', '종합효율', '목표효율', 'OPEN ISSUE']

if uploaded_files:
    all_records = []
    daily_totals_data = {} 
    
    for file in uploaded_files:
        try:
            if file.name.endswith('.csv'): temp_df = pd.read_csv(file)
            else: temp_df = pd.read_excel(file)
            
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
            clean_date = date_match.group()[2:] if date_match else file.name.split('.')[0]
            
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
                try: daily_total_oee = float(temp_df['종합효율'].iloc[45]) if not isinstance(temp_df['종합효율'].iloc[45], pd.Series) else float(temp_df['종합효율'].iloc[45].iloc[0])
                except: daily_total_oee = 0
            
            daily_totals_data[clean_date] = daily_total_oee

            for _, row in temp_df.iterrows():
                machine_val = str(row.get('설비명', ''))
                if isinstance(machine_val, pd.Series): machine_val = str(machine_val.iloc[0])
                if 'TOTAL' in machine_val.upper() or '합계' in machine_val or 'GRAND' in machine_val.upper(): continue
                    
                record = {'생산일': clean_date}
                for col in target_cols:
                    if col == '생산일': continue
                    if col in temp_df.columns:
                        val = row[col]
                        if isinstance(val, pd.Series): val = val.iloc[0]
                        record[col] = val
                    else: record[col] = None
                all_records.append(record)
                
        except Exception as e:
            st.error(f"파일 읽기 오류 ({file.name}): {e}")

    if all_records:
        df = pd.DataFrame(all_records)
        daily_df = pd.DataFrame(list(daily_totals_data.items()), columns=['생산일', '공장종합효율']).sort_values(by='생산일')
        
        num_cols = ['양품수량', '불량수량', '총 생산수량', '투입시간', '가동시간', '비가동시간', '정미시간', '종합효율', '목표효율', '양품율', '성능가동율', '시간가동율']
        for col in num_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # 줄바꿈 정제 함수
        def format_issue(text):
            val = str(text).strip()
            if val in ['', '0', '0.0', 'nan', 'NaN', 'None']: return ""
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
        selected_machines = st.sidebar.multiselect("설비 선택", all_machines, default=[], placeholder="전체 설비")
        
        if len(selected_machines) == 0: pool_df = df.copy()
        else: pool_df = df[df['설비명'].isin(selected_machines)].copy()
            
        pool_df['품명_필터'] = pool_df['품명'].fillna("").astype(str).str.strip()
        pool_df['품명_필터'] = pool_df['품명_필터'].replace(['0', '0.0', 'nan', 'NaN', 'None'], "")
        actual_prods = sorted([p for p in pool_df['품명_필터'].unique() if p != ""])
        selected_prod = st.sidebar.selectbox("품목 선택 (해당 설비 생산품)", ["전체 품목"] + actual_prods)

        f_df = pool_df.copy()
        if selected_prod != "전체 품목":
            f_df = f_df[f_df['품명_필터'] == selected_prod]

        # ---------------------------------------------------------
        # [HTML 테이블 렌더링 헬퍼 함수] - 에러 방지 및 디자인
        # ---------------------------------------------------------
        def render_styler_to_html(styler, is_multi=False):
            # 🚨 특수문자 충돌 에러 완전 방지를 위해 Pandas Styler 내장 기능 사용
            html_str = styler.to_html(escape=True) 
            
            # 🚨 강제 줄바꿈 및 정중앙 정렬 CSS 주입
            wrapped_html = f"""
            <div style="width: 100%; max-height: 500px; overflow: auto; border: 1px solid #DEE2E6; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); margin-bottom: 30px;">
                <style>
                    .custom-table {{ width: 100%; border-collapse: collapse; font-size: 13px; color: #333; background-color: white; }}
                    .custom-table th {{ background-color: #F8F9FA; border: 1px solid #DEE2E6; padding: 10px; text-align: center !important; vertical-align: middle !important; font-weight: bold; position: sticky; top: 0; z-index: 2; }}
                    .custom-table thead tr:nth-child(2) th {{ top: 38px; }}
                    .custom-table td {{ border: 1px solid #DEE2E6; padding: 8px 10px; text-align: center !important; vertical-align: middle !important; }}
                    .custom-table tbody tr:hover {{ background-color: #F1F3F5; }}
                    /* OPEN ISSUE 열(가장 마지막 열) 왼쪽 정렬 및 줄바꿈 강제 적용 */
                    .custom-table td:last-child {{ text-align: left !important; white-space: pre-wrap !important; min-width: 300px; line-height: 1.5; }}
                </style>
                {html_str.replace('<table', '<table class="custom-table"')}
            </div>
            """
            # 멀티인덱스(OPEN ISSUE 두칸) 강제 병합 처리
            if is_multi:
                wrapped_html = re.sub(r'<th class="col_heading level0 col10".*?>OPEN ISSUE</th>', r'<th class="col_heading level0 col10" rowspan="2" style="vertical-align: middle;">OPEN ISSUE</th>', wrapped_html)
                wrapped_html = re.sub(r'<th class="col_heading level1 col10".*?>OPEN ISSUE</th>', '', wrapped_html)
            
            st.markdown(wrapped_html, unsafe_allow_html=True)

        # ---------------------------------------------------------
        # 탭 구성: [Tab 1. 생산 추이 및 요약] / [Tab 2. OPEN ISSUE 정밀 분석]
        # ---------------------------------------------------------
        tab1, tab2 = st.tabs(["📈 생산 추이 및 요약", "📝 OPEN ISSUE 정밀 분석"])

        # =========================================================
        # TAB 1: 생산 추이 및 요약
        # =========================================================
        with tab1:
            is_factory_view = (len(selected_machines) == 0 and selected_prod == "전체 품목")
            
            if is_factory_view:
                st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 사출생산팀 일별 종합 효율(%)</h3>", unsafe_allow_html=True)
                plot_df = daily_df.copy()
                plot_df['목표효율'] = 0.86
                y_val = '공장종합효율'
            else:
                st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 선택 조건(설비/품목) 일별 평균 종합 효율(%)</h3>", unsafe_allow_html=True)
                active_oee = f_df[f_df['종합효율'] > 0]
                plot_df = active_oee.groupby('생산일')[['종합효율', '목표효율']].mean().reset_index().sort_values(by='생산일')
                y_val = '종합효율'

            if not plot_df.empty:
                colors = ['#FF4B4B' if row[y_val] < row['목표효율'] else '#1F77B4' for _, row in plot_df.iterrows()]
                plot_df['x_label'] = plot_df.apply(lambda row: f"{row['생산일']}<br><span style='font-size:11px;color:gray;'>({row[y_val]:.1%})</span>", axis=1)
                
                fig1 = go.Figure()
                fig1.add_trace(go.Scatter(
                    x=plot_df['x_label'], y=plot_df[y_val], mode='lines',
                    line=dict(shape='spline', width=3, color='#1F77B4'),
                    fill='tozeroy', fillcolor='rgba(31, 119, 180, 0.05)', hoverinfo='skip', showlegend=False
                ))
                fig1.add_trace(go.Scatter(
                    x=plot_df['x_label'], y=plot_df[y_val], mode='markers+text',
                    text=plot_df[y_val].apply(lambda x: f'{x:.1%}'), textposition="top center",
                    marker=dict(size=10, color='white', line=dict(width=2.5, color=colors)),
                    textfont=dict(size=14, color=colors, weight="bold"), showlegend=False, cliponaxis=False
                ))
                
                target_val = 0.86 if is_factory_view else plot_df['목표효율'].mean()
                fig1.add_trace(go.Scatter(
                    x=plot_df['x_label'], y=[target_val] * len(plot_df), mode='lines', 
                    name=f'목표 효율 ({target_val:.1%})', line=dict(color='#ADB5BD', dash='dash', width=2)
                ))
                
                fig1.update_xaxes(type='category', title="", showgrid=False, range=[-0.5, len(plot_df)-0.5]) 
                fig1.update_yaxes(title="종합효율", tickformat='.0%', range=[0, 1.2], showgrid=True, gridcolor='rgba(230,230,230,0.5)') 
                fig1.update_layout(height=450, margin=dict(l=40, r=40, t=40, b=40), plot_bgcolor='white', paper_bgcolor='white', legend=dict(yanchor="top", y=1.1, xanchor="right", x=1))
                st.plotly_chart(fig1, use_container_width=True)
                
                sorted_df = plot_df.sort_values(by=y_val, ascending=False)
                best_3 = sorted_df.head(3)
                worst_3 = sorted_df.tail(3).sort_values(by=y_val, ascending=True)
                
                st.markdown("#### 🏆 종합효율 BEST 3 & 🚨 WORST 3")
                b_cols = st.columns(3)
                w_cols = st.columns(3)
                
                for i, (_, r) in enumerate(best_3.iterrows()):
                    with b_cols[i]: st.markdown(f"<div class='metric-card'><div class='metric-title'>BEST {i+1}</div><div class='metric-value best'>{r[y_val]:.1%}</div><div class='metric-date'>{r['생산일']}</div></div>", unsafe_allow_html=True)
                for i, (_, r) in enumerate(worst_3.iterrows()):
                    with w_cols[i]: st.markdown(f"<div class='metric-card'><div class='metric-title'>WORST {i+1}</div><div class='metric-value worst'>{r[y_val]:.1%}</div><div class='metric-date'>{r['생산일']}</div></div>", unsafe_allow_html=True)
            else:
                st.info("종합효율 데이터가 없습니다.")
            
            st.write("---")
            st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 일별 총 비가동 시간(분)</h3>", unsafe_allow_html=True)
            daily_stop = f_df.groupby('생산일')['비가동시간'].sum().reset_index().sort_values(by='생산일')
            fig2 = px.bar(daily_stop, x='생산일', y='비가동시간', text_auto='.1f')
            
            fig2.update_traces(marker=dict(color='#E07A5F', opacity=0.9), textposition="outside", textfont=dict(size=13, weight="bold", color="#E07A5F"), cliponaxis=False)
            fig2.update_xaxes(type='category', title="", showgrid=False, range=[-0.5, len(daily_stop)-0.5])
            fig2.update_yaxes(title="비가동시간 (분)", showgrid=True, gridcolor='rgba(230,230,230,0.5)')
            fig2.update_layout(height=350, margin=dict(l=40, r=40, t=40, b=40), plot_bgcolor='white', paper_bgcolor='white')
            st.plotly_chart(fig2, use_container_width=True)

            # 🚨 [수정 3] Tab 1 하단에 '사출생산팀 일일 생산성 자료' 복구 및 배치
            st.write("---")
            st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 사출생산팀 일일 생산성 자료</h3>", unsafe_allow_html=True)
            
            display_df = f_df.sort_values(by=['생산일', '설비명']).copy()
            target_order = ['생산일', '설비명', '품명', '종합효율', '양품율', '성능가동율', '시간가동율', '총 생산수량', '양품수량', '불량수량', 'OPEN ISSUE']
            
            # 목표 미달 빨간색 강조용 마스크 생성
            style_main = pd.DataFrame('', index=display_df.index, columns=target_order)
            for i in range(len(display_df)):
                try:
                    if float(display_df.iloc[i]['종합효율']) < float(display_df.iloc[i]['목표효율']):
                        style_main.iloc[i, style_main.columns.get_loc('종합효율')] = 'color: #FF4B4B; font-weight: bold;'
                except: pass

            display_df = display_df[[c for c in target_order if c in display_df.columns]]

            def finalize_row(row):
                val = str(row.get('품명', '')).strip()
                is_idle = (val == '') or (val in ['0', '0.0', 'nan', 'NaN', 'None'])
                res = row.copy()
                for col in res.index:
                    if is_idle and col not in ['생산일', '설비명']: res[col] = ""
                    else:
                        if col in ['종합효율', '성능가동율', '시간가동율', '양품율']:
                            try: res[col] = f"{float(res[col]):.1%}" if pd.notnull(res[col]) and res[col] != "" else ""
                            except: pass
                        elif col in ['양품수량', '불량수량', '총 생산수량']:
                            try: res[col] = f"{int(float(res[col])):,.0f}" if pd.notnull(res[col]) and res[col] != "" else ""
                            except: pass
                if is_idle: res['품명'] = ""
                for k, v in res.items():
                    if pd.isna(v) or v == 'None': res[k] = ""
                return res

            final_table = display_df.apply(finalize_row, axis=1)
            
            # 멀티인덱스(카테고리) 설정
            multi_cols = [
                ('구분', '생산일'), ('구분', '설비명'), ('구분', '품명'),
                ('생산성', '종합효율'), ('생산성', '양품율'), ('생산성', '성능가동율'), ('생산성', '시간가동율'),
                ('생산실적', '총 생산수량'), ('생산실적', '양품수량'), ('생산실적', '불량수량'),
                ('OPEN ISSUE', 'OPEN ISSUE') 
            ]
            final_table.columns = pd.MultiIndex.from_tuples(multi_cols)
            style_main.columns = pd.MultiIndex.from_tuples(multi_cols)
            
            # 스타일러 적용 및 HTML 렌더링 호출
            final_styler = final_table.style.apply(lambda _: style_main, axis=None).hide(axis="index")
            render_styler_to_html(final_styler, is_multi=True)

        # =========================================================
        # TAB 2: OPEN ISSUE 정밀 분석
        # =========================================================
        with tab2:
            st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 빈출 이슈 (OPEN ISSUE 키워드 분석)</h3>", unsafe_allow_html=True)
            st.markdown("선택된 기간/조건 동안 가장 많이 발생한 특이사항의 핵심 단어입니다. (효율 개선 집중 포인트)")
            
            issue_df = f_df[f_df['OPEN ISSUE'] != ""].copy()
            
            if not issue_df.empty:
                all_text = " ".join(issue_df['OPEN ISSUE'].astype(str))
                words = re.findall(r'[가-힣]+', all_text)
                stopwords = {'주간', '야간', '확인', '점검', '가동', '조치', '이상', '완료', '발생', '후', '및', '설비', '생산', '연속', '특이사항', '동작'}
                filtered_words = [w for w in words if w not in stopwords and len(w) > 1]
                
                if filtered_words:
                    word_counts = Counter(filtered_words).most_common(5)
                    wc_cols = st.columns(len(word_counts))
                    for i, (word, count) in enumerate(word_counts):
                        with wc_cols[i]:
                            st.markdown(f"<div style='background-color:white; padding:15px; border-radius:8px; text-align:center; border:1px solid #E9ECEF; box-shadow: 0 2px 4px rgba(0,0,0,0.05);'><div style='font-size:18px; font-weight:900; color:#1F77B4;'>{word}</div><div style='font-size:13px; color:#6C757D; margin-top:5px;'>{count}건 발생</div></div>", unsafe_allow_html=True)
                else:
                    st.write("반복되는 유의미한 키워드가 감지되지 않았습니다.")
                
                st.write("---")
                
                # 🚨 [수정 1] Tab 2 하단에 '일별, 설비별 OPEN ISSUE' 복구
                st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 일별, 설비별 OPEN ISSUE 상세</h3>", unsafe_allow_html=True)
                
                issue_display = issue_df[['생산일', '설비명', '품명', '종합효율', 'OPEN ISSUE']].sort_values(by=['생산일', '설비명'])
                
                # 목표 미달 빨간색 강조용 마스크 생성
                style_issue = pd.DataFrame('', index=issue_display.index, columns=issue_display.columns)
                for i, idx in enumerate(issue_display.index):
                    try:
                        if float(issue_df.loc[idx, '종합효율']) < float(issue_df.loc[idx, '목표효율']):
                            style_issue.iloc[i, style_issue.columns.get_loc('종합효율')] = 'color: #FF4B4B; font-weight: bold;'
                    except: pass
                
                issue_display['종합효율'] = issue_display['종합효율'].apply(lambda x: f"{float(x):.1%}" if pd.notnull(x) and str(x).strip()!="" else "")
                
                # 스타일러 적용 및 HTML 렌더링 호출
                issue_styler = issue_display.style.apply(lambda _: style_issue, axis=None).hide(axis="index")
                render_styler_to_html(issue_styler, is_multi=False)
            else:
                st.info("선택된 조건에 해당하는 특이사항(OPEN ISSUE)이 없습니다.")

else:
    st.info("왼쪽 상단에서 생산성 파일을 업로드해 주세요.")
