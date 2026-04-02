import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
import os
import numpy as np
from collections import Counter
from datetime import datetime

# 1. 웹 화면 및 폰트/스타일, 자동번역 차단 설정 🚨
st.set_page_config(page_title="사출생산팀 일일 생산성 정밀 분석", layout="wide")

st.markdown("""
<meta name="google" content="notranslate">
<style>
    html, body, [class*="css"] {
        font-family: 'Pretendard', 'Noto Sans KR', 'Malgun Gothic', sans-serif !important;
        background-color: #F8F9FA;
        translate: no; 
    }
    .metric-card {
        background-color: white; border: 1px solid #E9ECEF; border-radius: 8px;
        padding: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); text-align: center; margin-bottom: 20px;
        min-height: 160px;
    }
    .metric-title { font-size: 14px; color: #6C757D; font-weight: bold; margin-bottom: 5px; }
    .metric-value.best { font-size: 20px; color: #1F77B4; font-weight: 900; }
    .metric-value.worst { font-size: 20px; color: #FF4B4B; font-weight: 900; }
</style>
""", unsafe_allow_html=True)

# 2. 로고 및 타이틀
col1, col2 = st.columns([1, 10])
with col1:
    logo_path = "듀링로고_가로형_빨강_JPG.jpg"
    if os.path.exists(logo_path): st.image(logo_path, width=100)
    else: st.markdown("<h2 style='color: #FF2A2A; font-weight: 900; margin-top: 10px;' class='notranslate'>DÜRING</h2>", unsafe_allow_html=True)
with col2:
    st.markdown("<h1 style='margin-top: 0px; color: #212529;' class='notranslate'>사출생산팀 일일 생산성 정밀 분석</h1>", unsafe_allow_html=True)

# 에러 방지 함수
def safe_float(val):
    try:
        if isinstance(val, pd.Series): return float(val.iloc[0])
        if pd.isna(val) or val == '' or val is None: return 0.0
        return float(val)
    except: return 0.0

target_cols = ['생산일', '설비명', '품명', '양품수량', '불량수량', '총 생산수량', '투입시간', '가동시간', '비가동시간', '정미시간', '양품율', '성능가동율', '시간가동율', '종합효율', '목표효율', 'OPEN ISSUE']

# 3. 데이터 수집 로직
data_to_process = []
DATA_DIR = "data"
if os.path.exists(DATA_DIR):
    for file_name in os.listdir(DATA_DIR):
        if file_name.startswith("~$"): continue 
        if file_name.endswith('.xlsx') or file_name.endswith('.csv'):
            file_path = os.path.join(DATA_DIR, file_name)
            try:
                if file_name.endswith('.csv'): df = pd.read_csv(file_path)
                else: df = pd.read_excel(file_path)
                data_to_process.append((file_name, df))
            except Exception as e:
                st.error(f"고정 데이터 읽기 오류 ({file_name}): {e}")

uploaded_files = st.file_uploader("📂 새로운 일일 생산성 파일이 있다면 추가로 업로드하세요 (선택사항)", type=['xlsx', 'csv'], accept_multiple_files=True)
if uploaded_files:
    for file in uploaded_files:
        try:
            if file.name.endswith('.csv'): df = pd.read_csv(file)
            else: df = pd.read_excel(file)
            data_to_process.append((file.name, df))
        except Exception as e:
            st.error(f"업로드 파일 읽기 오류 ({file.name}): {e}")

if data_to_process:
    all_records = []
    daily_totals_data = {} 
    
    for file_name, temp_df in data_to_process:
        temp_df = temp_df.loc[:, ~temp_df.columns.duplicated(keep='first')]
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
        
        date_match = re.search(r'\d{8}', file_name)
        if date_match:
            raw_date = date_match.group()[2:]
            try:
                dt = datetime.strptime(raw_date, '%y%m%d')
                weekdays = ['월', '화', '수', '목', '금', '토', '일']
                clean_date = f"{dt.strftime('%y')}년 {dt.month}월 {dt.day}일 ({weekdays[dt.weekday()]})"
                sort_key = raw_date 
            except:
                clean_date = raw_date
                sort_key = raw_date
        else:
            clean_date = file_name.split('.')[0]
            sort_key = clean_date
        
        daily_total_oee = None
        for _, row in temp_df.iterrows():
            machine_val = str(row.get('설비명', ''))
            if isinstance(machine_val, pd.Series): machine_val = str(machine_val.iloc[0])
            if 'TOTAL' in machine_val.upper() or '합계' in machine_val or 'GRAND' in machine_val.upper():
                val = row.get('종합효율', None)
                daily_total_oee = safe_float(val)
                break
        
        if daily_total_oee is None or daily_total_oee == 0.0:
            try: daily_total_oee = safe_float(temp_df['종합효율'].iloc[45])
            except: daily_total_oee = 0.0
        
        if sort_key not in daily_totals_data:
            daily_totals_data[sort_key] = {'생산일': clean_date, '공장종합효율': daily_total_oee}
        else:
            if daily_total_oee > 0:
                daily_totals_data[sort_key]['공장종합효율'] = daily_total_oee

        for _, row in temp_df.iterrows():
            machine_val = str(row.get('설비명', ''))
            if isinstance(machine_val, pd.Series): machine_val = str(machine_val.iloc[0])
            if 'TOTAL' in machine_val.upper() or '합계' in machine_val or 'GRAND' in machine_val.upper(): continue
                
            record = {'sort_key': sort_key, '생산일': clean_date}
            for col in target_cols:
                if col == '생산일': continue
                if col in temp_df.columns:
                    val = row[col]
                    if isinstance(val, pd.Series): val = val.iloc[0]
                    record[col] = val
                else: record[col] = None
            all_records.append(record)

    if all_records:
        df = pd.DataFrame(all_records).sort_values(by='sort_key').reset_index(drop=True)
        df = df.drop(columns=['sort_key'])
        
        daily_list = [{'sort_key': k, **v} for k, v in daily_totals_data.items()]
        daily_df = pd.DataFrame(daily_list).sort_values(by='sort_key').reset_index(drop=True)
        daily_df = daily_df.drop(columns=['sort_key'])
        
        num_cols = ['양품수량', '불량수량', '총 생산수량', '투입시간', '가동시간', '비가동시간', '정미시간', '종합효율', '목표효율', '양품율', '성능가동율', '시간가동율']
        for col in num_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

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
        # [사이드바 필터] (모든 탭 공통 적용)
        # ---------------------------------------------------------
        st.sidebar.header("🎯 정밀 필터링")
        
        df['설비명'] = df['설비명'].fillna("").astype(str)
        all_dates = [d for d in df['생산일'].unique() if str(d).strip() != ""]
        selected_dates = st.sidebar.multiselect("📅 생산일 선택", all_dates, default=[], placeholder="전체 생산일")
        
        if len(selected_dates) == 0: 
            date_filtered_df = df.copy()
            daily_df_filtered = daily_df.copy()
        else: 
            date_filtered_df = df[df['생산일'].isin(selected_dates)].copy()
            daily_df_filtered = daily_df[daily_df['생산일'].isin(selected_dates)].copy()

        all_machines = sorted([m for m in date_filtered_df['설비명'].unique() if m.strip() != ""])
        selected_machines = st.sidebar.multiselect("⚙️ 설비 선택", all_machines, default=[], placeholder="전체 설비")
        
        if len(selected_machines) == 0: pool_df = date_filtered_df.copy()
        else: pool_df = date_filtered_df[date_filtered_df['설비명'].isin(selected_machines)].copy()
            
        pool_df['품명_필터'] = pool_df['품명'].fillna("").astype(str).str.strip()
        pool_df['품명_필터'] = pool_df['품명_필터'].replace(['0', '0.0', 'nan', 'NaN', 'None'], "")
        actual_prods = sorted([p for p in pool_df['품명_필터'].unique() if p != ""])
        selected_prod = st.sidebar.selectbox("📦 품목 선택 (해당 설비 생산품)", ["전체 품목"] + actual_prods)

        f_df = pool_df[pool_df['품명_필터'] == selected_prod].copy() if selected_prod != "전체 품목" else pool_df.copy()

        # HTML 테이블 렌더링 헬퍼 함수
        def render_styler_to_html(styler, is_multi=False):
            html_str = styler.to_html(escape=True) 
            wrapped_html = f"""
            <div style="width: 100%; max-height: 500px; overflow: auto; border: 1px solid #DEE2E6; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); margin-bottom: 30px;">
                <style>
                    .custom-table {{ width: 100%; border-collapse: collapse; font-size: 13px; color: #333; background-color: white; }}
                    .custom-table th {{ background-color: #F8F9FA; border: 1px solid #DEE2E6; padding: 10px; text-align: center !important; vertical-align: middle !important; font-weight: bold; position: sticky; top: 0; z-index: 2; }}
                    .custom-table thead tr:nth-child(2) th {{ top: 38px; }}
                    .custom-table td {{ border: 1px solid #DEE2E6; padding: 8px 10px; text-align: center !important; vertical-align: middle !important; }}
                    .custom-table tbody tr:hover {{ background-color: #F1F3F5; }}
                    .custom-table td:last-child {{ text-align: left !important; white-space: pre-wrap !important; min-width: 300px; line-height: 1.5; }}
                </style>
                {html_str.replace('<table', '<table class="custom-table notranslate"')}
            </div>
            """
            if is_multi:
                wrapped_html = re.sub(r'<th class="col_heading level0 col10".*?>OPEN ISSUE</th>', r'<th class="col_heading level0 col10" rowspan="2" style="vertical-align: middle;">OPEN ISSUE</th>', wrapped_html)
                wrapped_html = re.sub(r'<th class="col_heading level1 col10".*?>OPEN ISSUE</th>', '', wrapped_html)
            st.markdown(wrapped_html, unsafe_allow_html=True)

        # ---------------------------------------------------------
        # 🚨 5개 탭 구성
        # ---------------------------------------------------------
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📈 종합 효율 및 비가동 추이 분석", 
            "📝 OPEN ISSUE 현황", 
            "📅 일별 생산성 및 비가동 현황", 
            "🏆/🚨 종합효율 BEST & WORST 분석 현황",
            "🛑 비가동시간 BEST & WORST 분석 현황"
        ])

        # =========================================================
        # TAB 1: 종합 효율 및 비가동 추이 분석
        # =========================================================
        with tab1:
            is_factory_view = (len(selected_machines) == 0 and selected_prod == "전체 품목")
            
            if is_factory_view:
                st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 사출생산팀 일별 종합 효율(%)</h3>", unsafe_allow_html=True)
                plot_df = daily_df_filtered.copy()
                plot_df['목표효율'] = 0.86
                y_val = '공장종합효율'
            else:
                st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 선택 조건(설비/품목) 일별 평균 종합 효율(%)</h3>", unsafe_allow_html=True)
                active_oee = f_df[f_df['종합효율'] > 0]
                plot_df = active_oee.groupby('생산일', sort=False)[['종합효율', '목표효율']].mean().reset_index()
                y_val = '종합효율'

            if not plot_df.empty:
                colors = []
                for _, row in plot_df.iterrows():
                    if safe_float(row[y_val]) < safe_float(row['목표효율']): colors.append('#FF4B4B')
                    else: colors.append('#1F77B4')
                
                plot_df['x_label'] = plot_df.apply(lambda row: f"{row['생산일']}<br><span style='font-size:11px;color:gray;'>({safe_float(row[y_val]):.1%})</span>", axis=1)
                
                fig1 = go.Figure()
                fig1.add_trace(go.Scatter(x=plot_df['x_label'], y=plot_df[y_val], mode='lines', line=dict(shape='spline', width=3, color='#1F77B4'), fill='tozeroy', fillcolor='rgba(31, 119, 180, 0.05)', hoverinfo='skip', showlegend=False))
                fig1.add_trace(go.Scatter(x=plot_df['x_label'], y=plot_df[y_val], mode='markers+text', text=plot_df[y_val].apply(lambda x: f'{safe_float(x):.1%}'), textposition="top center", marker=dict(size=10, color='white', line=dict(width=2.5, color=colors)), textfont=dict(size=14, color=colors, weight="bold"), showlegend=False, cliponaxis=False))
                target_val = 0.86 if is_factory_view else plot_df['목표효율'].mean()
                fig1.add_trace(go.Scatter(x=plot_df['x_label'], y=[target_val] * len(plot_df), mode='lines', name=f'목표 효율 ({target_val:.1%})', line=dict(color='#ADB5BD', dash='dash', width=2)))
                
                fig1.update_xaxes(type='category', title="", showgrid=False, range=[-0.5, len(plot_df)-0.5]) 
                fig1.update_yaxes(title="종합효율", tickformat='.0%', range=[0, 1.2], showgrid=True, gridcolor='rgba(230,230,230,0.5)') 
                fig1.update_layout(height=450, margin=dict(l=40, r=40, t=40, b=40), plot_bgcolor='white', paper_bgcolor='white', legend=dict(yanchor="top", y=1.1, xanchor="right", x=1))
                st.plotly_chart(fig1, use_container_width=True)
            else: st.info("종합효율 데이터가 없습니다.")
            
            st.write("---")
            st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 일별 총 비가동 시간(분)</h3>", unsafe_allow_html=True)
            
            daily_stop = f_df.groupby('생산일', sort=False)['비가동시간'].sum().reset_index()
            fig2 = px.bar(daily_stop, x='생산일', y='비가동시간', text_auto='.1f')
            fig2.update_traces(marker=dict(color='#E07A5F', opacity=0.9), textposition="outside", textfont=dict(size=13, weight="bold", color="#E07A5F"), cliponaxis=False)
            fig2.update_xaxes(type='category', title="", showgrid=False, range=[-0.5, len(daily_stop)-0.5])
            fig2.update_yaxes(title="비가동시간 (분)", showgrid=True, gridcolor='rgba(230,230,230,0.5)')
            fig2.update_layout(height=350, margin=dict(l=40, r=40, t=40, b=40), plot_bgcolor='white', paper_bgcolor='white')
            st.plotly_chart(fig2, use_container_width=True)

        # =========================================================
        # TAB 2: OPEN ISSUE 현황
        # =========================================================
        with tab2:
            st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> OPEN ISSUE 키워드 분석</h3>", unsafe_allow_html=True)
            st.markdown("선택된 조건 동안 반복적으로 발생한 주요 불량 및 이슈 문구입니다.")
            
            issue_df = f_df[f_df['OPEN ISSUE'] != ""].copy()
            if not issue_df.empty:
                all_text = " ".join(issue_df['OPEN ISSUE'].astype(str))
                all_text = re.sub(r'(주간|야간|주,|야,|주야간)\s*', '', all_text)
                words = re.findall(r'[가-힣A-Za-z0-9]+', all_text)
                stopwords = {'확인', '점검', '가동', '조치', '완료', '발생', '설비', '생산', '연속', '특이사항', '대기', '진행', '시간', '정도', '이후'}
                filtered_words = [w for w in words if w not in stopwords and len(w) > 1]
                bigrams = [f"{filtered_words[i]} {filtered_words[i+1]}" for i in range(len(filtered_words) - 1)]
                if not bigrams and filtered_words: bigrams = filtered_words

                if bigrams:
                    word_counts = Counter(bigrams).most_common(5)
                    wc_cols = st.columns(len(word_counts))
                    for i, (word, count) in enumerate(word_counts):
                        with wc_cols[i]: st.markdown(f"<div style='background-color:white; padding:15px; border-radius:8px; text-align:center; border:1px solid #E9ECEF; box-shadow: 0 2px 4px rgba(0,0,0,0.05);'><div style='font-size:16px; font-weight:900; color:#1F77B4;'>{word}</div><div style='font-size:13px; color:#6C757D; margin-top:5px;'>{count}건 감지</div></div>", unsafe_allow_html=True)
                else: st.write("반복되는 유의미한 키워드가 감지되지 않았습니다.")
                
                st.write("---")
                st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 일별, 설비별 OPEN ISSUE 상세</h3>", unsafe_allow_html=True)
                
                issue_display = issue_df[['생산일', '설비명', '품명', '종합효율', 'OPEN ISSUE']].reset_index(drop=True)
                base_issue_df = issue_df.copy().reset_index(drop=True) # 🚨 에러 방지용 베이스 데이터
                
                issue_display['종합효율'] = issue_display['종합효율'].apply(lambda x: f"{safe_float(x):.1%}")
                
                # 🚨 에러 없는 안전한 Row-by-Row 렌더링 로직 (MultiIndex 버그 회피)
                def style_issue_row(row):
                    styles = [''] * len(row)
                    idx = row.name
                    try:
                        oee = safe_float(base_issue_df.loc[idx, '종합효율'])
                        tgt = safe_float(base_issue_df.loc[idx, '목표효율'])
                        if 0 < oee < tgt:
                            pos = row.index.get_loc('종합효율')
                            if isinstance(pos, np.ndarray): pos = np.where(pos)[0][0]
                            styles[pos] = 'color: #FF4B4B; font-weight: bold;'
                    except: pass
                    return styles
                
                issue_styler = issue_display.style.apply(style_issue_row, axis=1).hide(axis="index")
                render_styler_to_html(issue_styler, is_multi=False)
            else: st.info("선택된 조건에 해당하는 특이사항(OPEN ISSUE)이 없습니다.")

        # =========================================================
        # TAB 3: 일별 생산성 및 비가동 현황
        # =========================================================
        with tab3:
            st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 일일 생산성 자료 (전체 상세 데이터)</h3>", unsafe_allow_html=True)
            
            display_df = f_df.reset_index(drop=True)
            base_df = display_df.copy() # 🚨 비교용 안전 베이스
            
            target_order = ['생산일', '설비명', '품명', '종합효율', '양품율', '성능가동율', '시간가동율', '총 생산수량', '양품수량', '불량수량', 'OPEN ISSUE']
            for c in target_order:
                if c not in display_df.columns: display_df[c] = ""
            display_df = display_df[target_order]

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
            multi_cols = [('구분', '생산일'), ('구분', '설비명'), ('구분', '품명'), ('생산성', '종합효율'), ('생산성', '양품율'), ('생산성', '성능가동율'), ('생산성', '시간가동율'), ('생산실적', '총 생산수량'), ('생산실적', '양품수량'), ('생산실적', '불량수량'), ('OPEN ISSUE', 'OPEN ISSUE')]
            final_table.columns = pd.MultiIndex.from_tuples(multi_cols)
            
            # 🚨 에러 없는 안전한 Row-by-Row 색상 적용 로직
            def style_main_row(row):
                styles = [''] * len(row)
                idx = row.name
                try:
                    oee = safe_float(base_df.loc[idx, '종합효율'])
                    tgt = safe_float(base_df.loc[idx, '목표효율'])
                    if 0 < oee < tgt:
                        pos = row.index.get_loc(('생산성', '종합효율'))
                        if isinstance(pos, np.ndarray): pos = np.where(pos)[0][0]
                        styles[pos] = 'color: #FF4B4B; font-weight: bold;'
                except: pass
                return styles

            final_styler = final_table.style.apply(style_main_row, axis=1).hide(axis="index")
            render_styler_to_html(final_styler, is_multi=True)

        # =========================================================
        # TAB 4: 종합효율 BEST & WORST 분석 현황
        # =========================================================
        with tab4:
            st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 종합효율 BEST 5 & WORST 5 요인 분석</h3>", unsafe_allow_html=True)
            st.markdown("가장 효율이 높았던 5건과 가장 저조했던 5건의 상세 내역(풀네임 품목 및 OPEN ISSUE)을 비교 분석합니다.")
            
            valid_df = f_df[f_df['종합효율'] > 0].copy()
            if not valid_df.empty:
                # --- BEST 5 ---
                st.markdown("<h4 style='color: #1F77B4; margin-top: 20px; font-weight: 800;'>🏆 최고 효율 (BEST 5)</h4>", unsafe_allow_html=True)
                best5_df = valid_df.sort_values(by=['종합효율', '생산일'], ascending=[False, False]).head(5)
                best5_display = best5_df[['생산일', '설비명', '품명', '종합효율', 'OPEN ISSUE']].reset_index(drop=True)
                best5_display['종합효율'] = best5_display['종합효율'].apply(lambda x: f"{safe_float(x):.1%}")
                
                def style_best_row(row):
                    styles = [''] * len(row)
                    try:
                        pos = row.index.get_loc('종합효율')
                        if isinstance(pos, np.ndarray): pos = np.where(pos)[0][0]
                        styles[pos] = 'color: #1F77B4; font-weight: bold;'
                    except: pass
                    return styles
                
                best_styler = best5_display.style.apply(style_best_row, axis=1).hide(axis="index")
                render_styler_to_html(best_styler, is_multi=False)
                
                # --- WORST 5 ---
                st.markdown("<h4 style='color: #FF4B4B; margin-top: 40px; font-weight: 800;'>🚨 최저 효율 (WORST 5)</h4>", unsafe_allow_html=True)
                worst5_df = valid_df.sort_values(by=['종합효율', '생산일'], ascending=[True, False]).head(5)
                worst5_display = worst5_df[['생산일', '설비명', '품명', '종합효율', 'OPEN ISSUE']].reset_index(drop=True)
                worst5_display['종합효율'] = worst5_display['종합효율'].apply(lambda x: f"{safe_float(x):.1%}")
                
                def style_worst_row(row):
                    styles = [''] * len(row)
                    try:
                        pos = row.index.get_loc('종합효율')
                        if isinstance(pos, np.ndarray): pos = np.where(pos)[0][0]
                        styles[pos] = 'color: #FF4B4B; font-weight: bold;'
                    except: pass
                    return styles
                
                worst_styler = worst5_display.style.apply(style_worst_row, axis=1).hide(axis="index")
                render_styler_to_html(worst_styler, is_multi=False)
            else: st.info("분석할 가동 데이터가 없습니다.")

        # =========================================================
        # TAB 5: 비가동시간 BEST & WORST 분석 현황
        # =========================================================
        with tab5:
            st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 비가동시간 BEST 5 & WORST 5 요인 분석</h3>", unsafe_allow_html=True)
            st.markdown("비가동시간이 가장 적었던(BEST) 5건과 가장 많았던(WORST) 5건의 상세 내역과 이슈를 분석합니다.")
            
            valid_dt_df = f_df.copy()
            if not valid_dt_df.empty:
                # --- 최소 비가동 (BEST 5) ---
                st.markdown("<h4 style='color: #20C997; margin-top: 20px; font-weight: 800;'>🏆 최소 비가동 (BEST 5)</h4>", unsafe_allow_html=True)
                best5_dt = valid_dt_df.sort_values(by=['비가동시간', '종합효율'], ascending=[True, False]).head(5)
                best5_dt_display = best5_dt[['생산일', '설비명', '품명', '비가동시간', 'OPEN ISSUE']].reset_index(drop=True)
                best5_dt_display['비가동시간'] = best5_dt_display['비가동시간'].apply(lambda x: f"{safe_float(x):.1f}분")
                
                def style_best_dt_row(row):
                    styles = [''] * len(row)
                    try:
                        pos = row.index.get_loc('비가동시간')
                        if isinstance(pos, np.ndarray): pos = np.where(pos)[0][0]
                        styles[pos] = 'color: #20C997; font-weight: bold;'
                    except: pass
                    return styles
                
                best_dt_styler = best5_dt_display.style.apply(style_best_dt_row, axis=1).hide(axis="index")
                render_styler_to_html(best_dt_styler, is_multi=False)
                
                # --- 최대 비가동 (WORST 5) ---
                st.markdown("<h4 style='color: #E07A5F; margin-top: 40px; font-weight: 800;'>🚨 최대 비가동 (WORST 5)</h4>", unsafe_allow_html=True)
                worst5_dt = valid_dt_df.sort_values(by=['비가동시간', '종합효율'], ascending=[False, True]).head(5)
                worst5_dt_display = worst5_dt[['생산일', '설비명', '품명', '비가동시간', 'OPEN ISSUE']].reset_index(drop=True)
                worst5_dt_display['비가동시간'] = worst5_dt_display['비가동시간'].apply(lambda x: f"{safe_float(x):.1f}분")
                
                def style_worst_dt_row(row):
                    styles = [''] * len(row)
                    try:
                        pos = row.index.get_loc('비가동시간')
                        if isinstance(pos, np.ndarray): pos = np.where(pos)[0][0]
                        styles[pos] = 'color: #E07A5F; font-weight: bold;'
                    except: pass
                    return styles
                
                worst_dt_styler = worst5_dt_display.style.apply(style_worst_dt_row, axis=1).hide(axis="index")
                render_styler_to_html(worst_dt_styler, is_multi=False)
            else: st.info("분석할 가동 데이터가 없습니다.")

else:
    st.info("데이터 파일이 없습니다. GitHub의 'data' 폴더에 엑셀 파일을 넣거나, 아래 버튼을 통해 직접 파일을 업로드해주세요.")
