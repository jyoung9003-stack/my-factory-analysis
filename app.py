import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
import os
import numpy as np
from collections import Counter
from datetime import datetime

# 1. 웹 화면 및 폰트/스타일, 자동번역 차단 설정
st.set_page_config(page_title="사출생산팀 일일 생산성 분석 리포트", layout="wide")

st.markdown("""
<meta name="google" content="notranslate">
<meta property="og:title" content="사출생산팀 일일 생산성 분석 리포트">
<meta property="og:description" content="사출생산팀 종합효율, 비가동, 특이사항 정밀 분석 대시보드입니다.">
<meta property="og:site_name" content="DÜRING 사출생산팀">
<style>
    html, body, [class*="css"] {
        font-family: 'Pretendard', 'Noto Sans KR', 'Malgun Gothic', sans-serif !important;
        background-color: #F8F9FA;
        translate: no; 
    }
    .metric-card {
        background-color: white; border: 1px solid #E9ECEF; border-radius: 8px;
        padding: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); text-align: center; margin-bottom: 20px;
        min-height: 160px;
    }
    .metric-title { font-size: 13px; color: #6C757D; font-weight: bold; margin-bottom: 5px; }
    .metric-value.best { font-size: 18px; color: #1F77B4; font-weight: 900; }
    .metric-value.worst { font-size: 18px; color: #FF4B4B; font-weight: 900; }
    .machine-summary { background-color: #e9ecef; padding: 12px; border-radius: 8px; margin-bottom: 15px; border-left: 5px solid #1F77B4; font-size: 15px; color: #343a40;}
</style>
""", unsafe_allow_html=True)

# 2. 로고 및 타이틀
col1, col2 = st.columns([1.5, 10])
with col1:
    logo_path = "듀링로고_가로형_빨강_JPG.jpg"
    if os.path.exists(logo_path): st.image(logo_path, width=100)
    else: st.markdown("<h2 style='color: #FF2A2A; font-weight: 900; margin-top: 10px; white-space: nowrap;' class='notranslate'>DÜRING</h2>", unsafe_allow_html=True)
with col2:
    st.markdown("<h1 style='margin-top: 0px; color: #212529;' class='notranslate'>사출생산팀 일일 생산성 정밀 분석</h1>", unsafe_allow_html=True)

# 🌟 [핵심 수정] CSV 포맷 클렌징 기능을 탑재한 안전한 숫자 변환 함수
def safe_float(val):
    try:
        if isinstance(val, pd.Series): val = val.iloc[0]
        if pd.isna(val) or val is None: return 0.0
        
        # 공백, 쉼표 등 불순물 제거
        v_str = str(val).strip().replace(',', '').replace(' ', '')
        
        # 엑셀의 빈칸 처리 방식이나 에러 텍스트는 0으로 치환
        if v_str in ['', '-', '#DIV/0!', '#N/A', 'nan', 'None']: return 0.0
        
        # 퍼센트(%) 기호가 있으면 지우고 100으로 나눠서 실수로 변환
        if '%' in v_str:
            return float(v_str.replace('%', '')) / 100.0
            
        return float(v_str)
    except:
        return 0.0

target_cols = ['생산일', '설비명', '품명', '양품수량', '불량수량', '총 생산수량', '투입시간', '가동시간', '비가동시간', '정미시간', '양품율', '성능가동율', '시간가동율', '종합효율', '목표효율', 'OPEN ISSUE']
target_order = ['생산일', '설비명', '품명', '종합효율', '양품율', '성능가동율', '시간가동율', '총 생산수량', '양품수량', '불량수량', 'OPEN ISSUE']
multi_cols = [
    ('구분', '생산일'), ('구분', '설비명'), ('구분', '품명'),
    ('생산성', '종합효율'), ('생산성', '양품율'), ('생산성', '성능가동율'), ('생산성', '시간가동율'),
    ('생산실적', '총 생산수량'), ('생산실적', '양품수량'), ('생산실적', '불량수량'),
    ('OPEN ISSUE', 'OPEN ISSUE')
]

# 3. 데이터 수집 로직
data_to_process = []
DATA_DIR = "data"
if os.path.exists(DATA_DIR):
    for file_name in os.listdir(DATA_DIR):
        if file_name.startswith("~$"): continue 
        if file_name.endswith('.xlsx') or file_name.endswith('.csv'):
            file_path = os.path.join(DATA_DIR, file_name)
            try:
                if file_name.endswith('.csv'): 
                    try: df = pd.read_csv(file_path, encoding='utf-8')
                    except UnicodeDecodeError: df = pd.read_csv(file_path, encoding='cp949')
                else: df = pd.read_excel(file_path)
                data_to_process.append((file_name, df))
            except Exception as e:
                st.error(f"고정 데이터 읽기 오류 ({file_name}): {e}")

uploaded_files = st.file_uploader("📂 새로운 일일 생산성 파일이 있다면 추가로 업로드하세요 (선택사항)", type=['xlsx', 'csv'], accept_multiple_files=True)
if uploaded_files:
    for file in uploaded_files:
        try:
            if file.name.endswith('.csv'): 
                try: df = pd.read_csv(file, encoding='utf-8')
                except UnicodeDecodeError:
                    file.seek(0)
                    df = pd.read_csv(file, encoding='cp949')
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
                month_str = f"{dt.strftime('%y')}년 {dt.month}월"
                sort_key = raw_date 
            except:
                clean_date = raw_date
                month_str = "분류 안됨"
                sort_key = raw_date
        else:
            clean_date = file_name.split('.')[0]
            month_str = "분류 안됨"
            sort_key = clean_date
        
        daily_total_oee = 0.0
        daily_total_perf = 0.0
        daily_total_avail = 0.0
        daily_total_qual = 0.0
        
        for _, row in temp_df.iterrows():
            machine_val = str(row.get('설비명', ''))
            if isinstance(machine_val, pd.Series): machine_val = str(machine_val.iloc[0])
            if 'TOTAL' in machine_val.upper() or '합계' in machine_val or 'GRAND' in machine_val.upper():
                daily_total_oee = safe_float(row.get('종합효율', 0.0))
                daily_total_perf = safe_float(row.get('성능가동율', 0.0))
                daily_total_avail = safe_float(row.get('시간가동율', 0.0))
                daily_total_qual = safe_float(row.get('양품율', 0.0))
                break
        
        if daily_total_oee == 0.0:
            try: 
                daily_total_oee = safe_float(temp_df['종합효율'].iloc[45])
                daily_total_perf = safe_float(temp_df['성능가동율'].iloc[45])
                daily_total_avail = safe_float(temp_df['시간가동율'].iloc[45])
                daily_total_qual = safe_float(temp_df['양품율'].iloc[45])
            except: pass
        
        if sort_key not in daily_totals_data:
            daily_totals_data[sort_key] = {
                '생산일': clean_date, '생산월': month_str, 
                '공장종합효율': daily_total_oee,
                '공장성능가동율': daily_total_perf,
                '공장시간가동율': daily_total_avail,
                '공장양품율': daily_total_qual
            }
        else:
            if daily_total_oee > 0:
                daily_totals_data[sort_key]['공장종합효율'] = daily_total_oee
                daily_totals_data[sort_key]['공장성능가동율'] = daily_total_perf
                daily_totals_data[sort_key]['공장시간가동율'] = daily_total_avail
                daily_totals_data[sort_key]['공장양품율'] = daily_total_qual

        for _, row in temp_df.iterrows():
            machine_val = str(row.get('설비명', ''))
            if isinstance(machine_val, pd.Series): machine_val = str(machine_val.iloc[0])
            if 'TOTAL' in machine_val.upper() or '합계' in machine_val or 'GRAND' in machine_val.upper(): continue
                
            record = {'sort_key': sort_key, '생산월': month_str, '생산일': clean_date}
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
        daily_list = [{'sort_key': k, **v} for k, v in daily_totals_data.items()]
        daily_df = pd.DataFrame(daily_list).sort_values(by='sort_key').reset_index(drop=True)
        
        # 🌟 [핵심 수정] CSV에서 넘어온 문자열(%, 콤마 등)을 모두 깔끔한 숫자로 변환합니다.
        num_cols = ['양품수량', '불량수량', '총 생산수량', '투입시간', '가동시간', '비가동시간', '정미시간', '종합효율', '목표효율', '양품율', '성능가동율', '시간가동율']
        for col in num_cols:
            if col in df.columns:
                df[col] = df[col].apply(safe_float)

        def format_issue(text):
            val = str(text).strip()
            if val in ['', '0', '0.0', 'nan', 'NaN', 'None']: return ""
            val = val.replace('\r\n', '\n')
            val = re.sub(r'(?<!\n)\*', '\n*', val)
            val = re.sub(r'(?<!\n)-\.', '\n-.', val)
            val = re.sub(r'(?<!\n)→', '\n→ ', val)
            return val.strip()

        df['OPEN ISSUE'] = df['OPEN ISSUE'].apply(format_issue)

        def split_issue_to_columns(issue_text):
            lines = [line.strip() for line in str(issue_text).split('\n') if line.strip()]
            if not lines:
                return "<div style='font-size:12px; color:#ADB5BD; margin-top:6px; background-color:#F8F9FA; padding:8px; border-radius:4px; border:1px dashed #E9ECEF;'>📝 기록된 특이사항(OPEN ISSUE) 없음</div>"
                
            day_lines, night_lines, general_lines = [], [], []
            current = general_lines
            has_shift = False
            
            for line in lines:
                clean_line = line.replace(' ', '')
                if '*주간' in clean_line or line.startswith('주간'):
                    current = day_lines
                    has_shift = True
                    line = re.sub(r'^\*?\s*주간\s*', '', line).strip()
                elif '*야간' in clean_line or line.startswith('야간'):
                    current = night_lines
                    has_shift = True
                    line = re.sub(r'^\*?\s*야간\s*', '', line).strip()
                
                if line: current.append(line)
                
            if not has_shift:
                return f"<div style='font-size:13px; color:#495057; line-height:1.6;'>{'<br>'.join(lines)}</div>"
                
            if general_lines: day_lines = general_lines + day_lines
                
            day_html = '<br>'.join(day_lines) if day_lines else "<span style='color:#ADB5BD; font-size:12px;'>특이사항 없음</span>"
            night_html = '<br>'.join(night_lines) if night_lines else "<span style='color:#ADB5BD; font-size:12px;'>특이사항 없음</span>"
            
            return f"""
            <div style='display: flex; gap: 10px; margin-top: 5px; width: 100%; min-width: 400px;'>
                <div style='flex: 1; background-color: #F8F9FA; border: 1px solid #E9ECEF; border-radius: 4px; padding: 10px; border-top: 3px solid #FFC107;'>
                    <div style='font-size:11px; font-weight:bold; color:#E0A800; margin-bottom:4px;'>☀️ 주간</div>
                    <div style='font-size:13px; color:#495057; line-height:1.6;'>{day_html}</div>
                </div>
                <div style='flex: 1; background-color: #F8F9FA; border: 1px solid #E9ECEF; border-radius: 4px; padding: 10px; border-top: 3px solid #343A40;'>
                    <div style='font-size:11px; font-weight:bold; color:#495057; margin-bottom:4px;'>🌙 야간</div>
                    <div style='font-size:13px; color:#495057; line-height:1.6;'>{night_html}</div>
                </div>
            </div>
            """

        st.sidebar.header("🎯 정밀 필터링")
        df['설비명'] = df['설비명'].fillna("").astype(str)
        all_months_sidebar = [m for m in df['생산월'].unique() if str(m).strip() != ""]
        selected_months_sidebar = st.sidebar.multiselect("📅 생산월 선택", all_months_sidebar, default=[], placeholder="전체 월")
        
        if len(selected_months_sidebar) == 0: 
            month_filtered_df = df.copy()
            daily_month_filtered = daily_df.copy()
        else: 
            month_filtered_df = df[df['생산월'].isin(selected_months_sidebar)].copy()
            daily_month_filtered = daily_df[daily_df['생산월'].isin(selected_months_sidebar)].copy()

        all_dates = [d for d in month_filtered_df['생산일'].unique() if str(d).strip() != ""]
        selected_dates = st.sidebar.multiselect("📆 생산일 선택", all_dates, default=[], placeholder="전체 생산일")
        
        if len(selected_dates) == 0: 
            date_filtered_df = month_filtered_df.copy()
            daily_df_filtered = daily_month_filtered.copy()
        else: 
            date_filtered_df = month_filtered_df[month_filtered_df['생산일'].isin(selected_dates)].copy()
            daily_df_filtered = daily_month_filtered[daily_month_filtered['생산일'].isin(selected_dates)].copy()

        all_machines = sorted([m for m in date_filtered_df['설비명'].unique() if m.strip() != ""])
        selected_machines = st.sidebar.multiselect("⚙️ 설비 선택", all_machines, default=[], placeholder="전체 설비")
        
        if len(selected_machines) == 0: pool_df = date_filtered_df.copy()
        else: pool_df = date_filtered_df[date_filtered_df['설비명'].isin(selected_machines)].copy()
            
        pool_df['품명_필터'] = pool_df['품명'].fillna("").astype(str).str.strip()
        pool_df['품명_필터'] = pool_df['품명_필터'].replace(['0', '0.0', 'nan', 'NaN', 'None'], "")
        actual_prods = sorted([p for p in pool_df['품명_필터'].unique() if p != ""])
        selected_prod = st.sidebar.selectbox("📦 품목 선택 (해당 설비 생산품)", ["전체 품목"] + actual_prods)

        f_df = pool_df[pool_df['품명_필터'] == selected_prod].copy() if selected_prod != "전체 품목" else pool_df.copy()

        def render_styler_to_html(styler, is_multi=False):
            try: html_str = styler.to_html(escape=False) 
            except: html_str = styler.to_html()
            wrapped_html = f"""<div style="width: 100%; max-height: 500px; overflow: auto; border: 1px solid #DEE2E6; border-radius: 5px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); margin-bottom: 30px;">
<style>
.custom-table {{ width: 100%; border-collapse: collapse; font-size: 13px; color: #333; background-color: white; }}
.custom-table th {{ background-color: #F8F9FA; border: 1px solid #DEE2E6; padding: 10px; text-align: center !important; vertical-align: middle !important; font-weight: bold; position: sticky; top: 0; z-index: 2; }}
.custom-table thead tr:nth-child(2) th {{ top: 38px; }}
.custom-table td {{ border: 1px solid #DEE2E6; padding: 8px 10px; text-align: center !important; vertical-align: middle !important; }}
.custom-table tbody tr:hover {{ background-color: #F1F3F5; }}
.custom-table td:last-child {{ text-align: left !important; white-space: pre-wrap !important; min-width: 450px; line-height: 1.5; }}
</style>
{html_str.replace('<table', '<table class="custom-table notranslate"')}
</div>"""
            if is_multi:
                wrapped_html = re.sub(r'<th class="col_heading level0 col10".*?>OPEN ISSUE</th>', r'<th class="col_heading level0 col10" rowspan="2" style="vertical-align: middle;">OPEN ISSUE</th>', wrapped_html)
                wrapped_html = re.sub(r'<th class="col_heading level1 col10".*?>OPEN ISSUE</th>', '', wrapped_html)
            st.markdown(wrapped_html, unsafe_allow_html=True)

        tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
            "📈 종합 효율 추이 분석", 
            "📝 OPEN ISSUE 현황", 
            "📅 일별 생산성 상세 현황", 
            "🏆 종합효율 BEST & WORST",
            "🛑 비가동시간 BEST & WORST",
            "📉 효율 급변 구간 정밀 추적",
            "🤖 AI 데이터 챗봇 (Beta)" 
        ])

        # =========================================================
        # TAB 1: 종합 효율 요약 및 추이 분석
        # =========================================================
        with tab1:
            is_factory_view = (len(selected_machines) == 0 and selected_prod == "전체 품목")
            
            if is_factory_view:
                plot_df = daily_df_filtered.copy()
                plot_df['목표효율'] = 0.86
                y_val = '공장종합효율'; p_val = '공장성능가동율'; a_val = '공장시간가동율'; q_val = '공장양품율'
            else:
                active_oee = f_df[f_df['종합효율'] > 0]
                plot_df = active_oee.groupby(['sort_key', '생산월', '생산일'])[['종합효율', '목표효율', '성능가동율', '시간가동율', '양품율']].mean().reset_index().sort_values('sort_key')
                y_val = '종합효율'; p_val = '성능가동율'; a_val = '시간가동율'; q_val = '양품율'

            st.markdown("<h3 style='font-weight: 900; color: #212529; margin-top: 10px;'><span style='color: #FF4B4B;'>■</span> 최근 5일 생산성 요약</h3>", unsafe_allow_html=True)
            if not plot_df.empty:
                recent_5_df = plot_df.sort_values('sort_key').tail(5)
                r_cols = st.columns(5)
                for i, (_, r) in enumerate(recent_5_df.iterrows()):
                    oee_val = safe_float(r[y_val])
                    tgt_val = safe_float(r['목표효율'])
                    perf_val = safe_float(r[p_val])
                    avail_val = safe_float(r[a_val])
                    qual_val = safe_float(r[q_val])

                    color = "#1F77B4" if oee_val >= tgt_val else "#FF4B4B"
                    bg_color = "#F8FBFF" if oee_val >= tgt_val else "#FFF5F5"
                    
                    with r_cols[i]:
                        st.markdown(f"""
                        <div style='background-color: {bg_color}; border: 1px solid {color}40; border-top: 4px solid {color}; border-radius: 8px; padding: 12px; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.05); margin-bottom: 10px;'>
                            <div style='font-size: 14px; color: #495057; font-weight: bold; margin-bottom: 8px;'>{r['생산일']}</div>
                            <div style='font-size: 26px; font-weight: 900; color: {color}; margin-bottom: 8px;'>{oee_val:.1%}</div>
                            <div style='display: flex; justify-content: space-between; font-size: 11px; color: #6C757D; background-color: rgba(255,255,255,0.6); padding: 4px 6px; border-radius: 4px;'>
                                <span>시간 <b>{avail_val:.1%}</b></span>
                                <span>성능 <b>{perf_val:.1%}</b></span>
                                <span>양품 <b>{qual_val:.1%}</b></span>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
            else: st.info("최근 5일 데이터가 없습니다.")

            st.markdown("<hr style='border:1px solid #DEE2E6; margin-top:15px; margin-bottom:30px;'>", unsafe_allow_html=True)

            def render_horizontal_card(row, rank, is_best, y_col='종합효율'):
                rank_color = "#1F77B4" if is_best else "#FF4B4B"
                bg_color = "#F8FBFF" if is_best else "#FFF5F5"
                title_text = "BEST" if is_best else "WORST"
                val = safe_float(row[y_col])

                day_data = f_df[(f_df['생산일'] == row['생산일']) & (f_df['종합효율'] > 0)].sort_values(by='종합효율', ascending=not is_best).head(3)
                issue_html = ""
                for _, r in day_data.iterrows():
                    m_name = str(r['설비명']).split(' - ')[0]
                    oee_c = "#1F77B4" if is_best else "#FF4B4B"
                    issue_h = split_issue_to_columns(r['OPEN ISSUE'])
                    issue_html += f"<div style='margin-bottom: 12px;'><strong style='color:{oee_c}; font-size:14px;'>[{m_name}]</strong> <span style='font-size:14px; font-weight:bold; color:#212529;'>{r['품명']} ({safe_float(r['종합효율']):.1%})</span>{issue_h}</div>"

                return f"""
                <div style='background-color: white; border: 1px solid {rank_color}40; border-radius: 8px; margin-bottom: 15px; display: flex; flex-direction: row; box-shadow: 0 2px 4px rgba(0,0,0,0.02); overflow: hidden;'>
                    <div style='flex: 0 0 180px; display: flex; flex-direction: column; justify-content: center; align-items: center; border-right: 1px dashed #DEE2E6; padding: 15px; background-color: {bg_color};'>
                        <div style='font-size: 14px; color: {rank_color}; font-weight: bold; margin-bottom: 5px;'>{title_text} {rank}</div>
                        <div style='font-size: 15px; font-weight: 900; color: #212529; text-align: center; word-break: keep-all;'>{row['생산일']}</div>
                        <div style='font-size: 22px; font-weight: 900; color: {rank_color}; margin-top: 8px;'>{val:.1%}</div>
                    </div>
                    <div style='flex: 1; padding: 15px 20px; text-align: left; display: flex; flex-direction: column; justify-content: center;'>
                        {issue_html}
                    </div>
                </div>
                """

            months = list(dict.fromkeys(plot_df['생산월'].tolist())) if not plot_df.empty else []
            
            if not plot_df.empty:
                st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 월별 종합 효율 추이 및 요인 분석</h3>", unsafe_allow_html=True)
                
                selected_months_tab1 = st.multiselect("📅 조회할 월(Month)을 선택하세요 (기본값: 최근 월)", options=months, default=[months[-1]] if months else [], key='t1_m')
                
                if not selected_months_tab1: st.warning("선택된 월이 없습니다.")
                
                for m in selected_months_tab1:
                    m_plot_df = plot_df[plot_df['생산월'] == m].copy()
                    if m_plot_df.empty: continue
                    
                    colors = ['#FF4B4B' if safe_float(row[y_val]) < safe_float(row['목표효율']) else '#1F77B4' for _, row in m_plot_df.iterrows()]
                    m_plot_df['x_label'] = m_plot_df.apply(lambda row: f"{row['생산일']}<br><span style='font-size:11px;color:gray;'>({safe_float(row[y_val]):.1%})</span>", axis=1)
                    
                    fig1 = go.Figure()
                    fig1.add_trace(go.Scatter(x=m_plot_df['x_label'], y=m_plot_df[y_val], mode='lines', line=dict(shape='spline', width=3, color='#1F77B4'), fill='tozeroy', fillcolor='rgba(31, 119, 180, 0.05)', hoverinfo='skip', showlegend=False))
                    fig1.add_trace(go.Scatter(x=m_plot_df['x_label'], y=m_plot_df[y_val], mode='markers+text', text=m_plot_df[y_val].apply(lambda x: f'{safe_float(x):.1%}'), textposition="top center", marker=dict(size=10, color='white', line=dict(width=2.5, color=colors)), textfont=dict(size=14, color=colors, weight="bold"), showlegend=False, cliponaxis=False))
                    target_val = 0.86 if is_factory_view else m_plot_df['목표효율'].mean()
                    fig1.add_trace(go.Scatter(x=m_plot_df['x_label'], y=[target_val] * len(m_plot_df), mode='lines', name=f'목표 효율 ({target_val:.1%})', line=dict(color='#ADB5BD', dash='dash', width=2)))
                    
                    fig1.update_xaxes(type='category', categoryorder='array', categoryarray=m_plot_df['x_label'], title="", showgrid=False) 
                    fig1.update_yaxes(title="종합효율", tickformat='.0%', range=[0, 1.2], showgrid=True, gridcolor='rgba(230,230,230,0.5)') 
                    fig1.update_layout(height=350, margin=dict(l=40, r=40, t=40, b=40), plot_bgcolor='white', paper_bgcolor='white', legend=dict(yanchor="top", y=1.1, xanchor="right", x=1))
                    st.plotly_chart(fig1, use_container_width=True)

                    st.markdown(f"<h5 style='color: #495057; margin-top: 20px;'>⚙️ {m} 일자별 가동 설비 대수 분석</h5>", unsafe_allow_html=True)
                    m_f_df = f_df[f_df['생산월'] == m].copy()
                    machine_counts = m_f_df[m_f_df['종합효율'] > 0].groupby(['sort_key', '생산일'])['설비명'].nunique().reset_index().sort_values('sort_key')
                    machine_counts.rename(columns={'설비명': '가동설비수'}, inplace=True)
                    
                    if not machine_counts.empty:
                        avg_m = machine_counts['가동설비수'].mean()
                        max_m = machine_counts['가동설비수'].max()
                        st.markdown(f"<div class='machine-summary'>💡 해당 월 일평균 <b>{avg_m:.1f}대</b> 가동 (최대 가동: <b>{max_m}대</b>)</div>", unsafe_allow_html=True)
                        fig_m = px.bar(machine_counts, x='생산일', y='가동설비수', text_auto=True)
                        fig_m.update_traces(marker_color='#6C757D', textposition="outside", textfont=dict(weight="bold"))
                        fig_m.update_xaxes(type='category', categoryorder='array', categoryarray=machine_counts['생산일'], title="", showgrid=False)
                        fig_m.update_layout(height=250, margin=dict(l=20,r=20,t=20,b=20), plot_bgcolor='white', yaxis_title="가동 대수 (대)")
                        st.plotly_chart(fig_m, use_container_width=True)
                    
                    m_sorted_df = m_plot_df.sort_values(by=y_val, ascending=False)
                    best_5 = m_sorted_df.head(5)
                    worst_5 = m_sorted_df.tail(5).sort_values(by=y_val, ascending=True)
                    
                    st.markdown(f"<h5 style='color: #1F77B4; margin-top: 30px; margin-bottom: 15px;'>🏆 {m} 종합효율 BEST 5</h5>", unsafe_allow_html=True)
                    if not best_5.empty:
                        for i, (_, r) in enumerate(best_5.iterrows()):
                            st.markdown(render_horizontal_card(r, i+1, True, y_val), unsafe_allow_html=True)
                    
                    st.markdown(f"<h5 style='color: #FF4B4B; margin-top: 30px; margin-bottom: 15px;'>🚨 {m} 종합효율 WORST 5</h5>", unsafe_allow_html=True)
                    if not worst_5.empty:
                        for i, (_, r) in enumerate(worst_5.iterrows()):
                            st.markdown(render_horizontal_card(r, i+1, False, y_val), unsafe_allow_html=True)

        # =========================================================
        # TAB 2: OPEN ISSUE 현황
        # =========================================================
        with tab2:
            st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> OPEN ISSUE 현황 및 정밀 분석</h3>", unsafe_allow_html=True)
            months_t2 = list(dict.fromkeys(f_df['생산월'].tolist())) if not f_df.empty else []
            sel_m_t2 = st.multiselect("📅 조회할 월(Month)을 선택하세요 (기본값: 최근 월)", options=months_t2, default=[months_t2[-1]] if months_t2 else [], key='t2_m')
            tab2_df = f_df[f_df['생산월'].isin(sel_m_t2)].copy() if sel_m_t2 else f_df.copy()
            
            st.markdown("선택하신 기간 내에 반복적으로 발생한 주요 불량 및 이슈 문구입니다.")
            issue_df = tab2_df[tab2_df['OPEN ISSUE'] != ""].copy()
            
            if not issue_df.empty:
                all_text = " ".join(issue_df['OPEN ISSUE'].astype(str))
                all_text = re.sub(r'(주간|야간|주,|야,|주야간)\s*', '', all_text)
                words = re.findall(r'[가-힣A-Za-z0-9]+', all_text)
                stopwords = {'확인', '점검', '가동', '조치', '완료', '발생', '설비', '생산', '연속', '특이사항', '대기', '진행', '시간', '정도', '이후'}
                filtered_words = [w for w in words if w not in stopwords and len(w) > 1]
                bigrams = [f"{filtered_words[i]} {filtered_words[i+1]}" for i in range(len(filtered_words) - 1)]
                if bigrams:
                    word_counts = Counter(bigrams).most_common(5)
                    wc_cols = st.columns(len(word_counts))
                    for i, (word, count) in enumerate(word_counts):
                        with wc_cols[i]: st.markdown(f"<div style='background-color:white; padding:15px; border-radius:8px; text-align:center; border:1px solid #E9ECEF; box-shadow: 0 2px 4px rgba(0,0,0,0.05);'><div style='font-size:16px; font-weight:900; color:#1F77B4;'>{word}</div><div style='font-size:13px; color:#6C757D; margin-top:5px;'>{count}건 감지</div></div>", unsafe_allow_html=True)
                
                st.write("---")
                st.markdown("<h4 style='font-weight: 800; color: #212529;'>특정 생산일 OPEN ISSUE 조회</h4>", unsafe_allow_html=True)
                all_dates_rev_t2 = list(reversed([d for d in tab2_df['생산일'].unique() if str(d).strip() != ""]))
                if all_dates_rev_t2:
                    selected_date_t2 = st.selectbox("조회할 생산일을 선택하세요", all_dates_rev_t2, key='tab2_date')
                    day_issue_df = issue_df[issue_df['생산일'] == selected_date_t2].copy().reset_index(drop=True)
                    if not day_issue_df.empty:
                        base_day_issue_df = day_issue_df.copy()
                        day_issue_display = day_issue_df[['생산일', '설비명', '품명', '종합효율', 'OPEN ISSUE']].copy()
                        day_issue_display['종합효율'] = day_issue_display['종합효율'].apply(lambda x: f"{safe_float(x):.1%}")
                        day_issue_display['OPEN ISSUE'] = day_issue_display['OPEN ISSUE'].apply(split_issue_to_columns)
                        
                        def style_day_issue_row(row):
                            styles = [''] * len(row)
                            idx = row.name
                            try:
                                if 0 < safe_float(base_day_issue_df.loc[idx, '종합효율']) < safe_float(base_day_issue_df.loc[idx, '목표효율']):
                                    pos = row.index.get_loc('종합효율')
                                    if isinstance(pos, np.ndarray): pos = np.where(pos)[0][0]
                                    styles[pos] = 'color: #FF4B4B; font-weight: bold;'
                            except: pass
                            return styles
                            
                        render_styler_to_html(day_issue_display.style.apply(style_day_issue_row, axis=1).hide(axis="index"))
                    else: st.info("선택하신 일자에는 작성된 특이사항이 없습니다.")

                st.write("---")
                st.markdown("<h4 style='font-weight: 800; color: #212529;'>조회 월(Month) OPEN ISSUE 전체 상세</h4>", unsafe_allow_html=True)
                issue_display = issue_df[['생산일', '설비명', '품명', '종합효율', '목표효율', 'OPEN ISSUE']].reset_index(drop=True)
                base_issue_df = issue_display.copy()
                issue_display = issue_display.drop(columns=['목표효율'])
                issue_display['종합효율'] = issue_display['종합효율'].apply(lambda x: f"{safe_float(x):.1%}")
                issue_display['OPEN ISSUE'] = issue_display['OPEN ISSUE'].apply(split_issue_to_columns)
                
                def style_issue_row(row):
                    styles = [''] * len(row)
                    idx = row.name
                    try:
                        if 0 < safe_float(base_issue_df.loc[idx, '종합효율']) < safe_float(base_issue_df.loc[idx, '목표효율']):
                            pos = row.index.get_loc('종합효율')
                            if isinstance(pos, np.ndarray): pos = np.where(pos)[0][0]
                            styles[pos] = 'color: #FF4B4B; font-weight: bold;'
                    except: pass
                    return styles
                    
                render_styler_to_html(issue_display.style.apply(style_issue_row, axis=1).hide(axis="index"))
            else: st.info("해당 월에 기록된 특이사항(OPEN ISSUE)이 없습니다.")

        # =========================================================
        # TAB 3: 일별 생산성 및 비가동 현황
        # =========================================================
        with tab3:
            st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 일일 생산성 상세 현황</h3>", unsafe_allow_html=True)
            months_t3 = list(dict.fromkeys(f_df['생산월'].tolist())) if not f_df.empty else []
            sel_m_t3 = st.multiselect("📅 조회할 월(Month)을 선택하세요 (기본값: 최근 월)", options=months_t3, default=[months_t3[-1]] if months_t3 else [], key='t3_m')
            tab3_df = f_df[f_df['생산월'].isin(sel_m_t3)].copy() if sel_m_t3 else f_df.copy()
            
            all_dates_rev_t3 = list(reversed([d for d in tab3_df['생산일'].unique() if str(d).strip() != ""]))
            if all_dates_rev_t3:
                selected_date = st.selectbox("📆 조회할 상세 생산일을 선택하세요", all_dates_rev_t3, key='tab3_date')
                day_df = tab3_df[tab3_df['생산일'] == selected_date].copy().reset_index(drop=True)
                
                st.write("---")
                st.markdown(f"#### 📊 {selected_date} 설비별 종합효율 비교")
                active_day = day_df[day_df['종합효율'] > 0].sort_values(by='종합효율', ascending=False)
                
                if not active_day.empty:
                    bar_colors = ['#FF4B4B' if safe_float(row['종합효율']) < safe_float(row['목표효율']) else '#1F77B4' for _, row in active_day.iterrows()]
                    fig3 = px.bar(active_day, x='설비명', y='종합효율', text_auto='.1%')
                    fig3.update_traces(marker_color=bar_colors, textposition="outside", textfont=dict(size=12, weight="bold"))
                    fig3.update_xaxes(title="", tickangle=45)
                    fig3.update_yaxes(title="종합효율", tickformat='.0%', range=[0, 1.2], showgrid=True, gridcolor='rgba(230,230,230,0.5)')
                    fig3.update_layout(height=400, margin=dict(l=40, r=40, t=40, b=80), plot_bgcolor='white', paper_bgcolor='white')
                    st.plotly_chart(fig3, use_container_width=True)
                else: st.info("해당 일자에 가동된 설비 데이터가 없습니다.")
                
                st.write("---")
                st.markdown(f"#### 🛑 {selected_date} 설비별 총 비가동 시간(시간) 비교")
                downtime_day = day_df[day_df['비가동시간'] > 0].sort_values(by='비가동시간', ascending=False)
                
                if not downtime_day.empty:
                    fig4 = px.bar(downtime_day, x='설비명', y='비가동시간', text_auto='.1f')
                    fig4.update_traces(marker_color='#E07A5F', textposition="outside", textfont=dict(size=12, weight="bold"))
                    fig4.update_xaxes(title="", tickangle=45)
                    fig4.update_yaxes(title="총 비가동시간(시간)", showgrid=True, gridcolor='rgba(230,230,230,0.5)')
                    fig4.update_layout(height=400, margin=dict(l=40, r=40, t=40, b=80), plot_bgcolor='white', paper_bgcolor='white')
                    st.plotly_chart(fig4, use_container_width=True)
                else: st.info("비가동 시간이 발생한 설비가 없습니다.")
                
                st.write("---")
                st.markdown(f"#### 📂 {selected_date} 전체 상세 표")
                
                display_day = day_df.copy()
                base_day_df = display_day.copy()
                for c in target_order:
                    if c not in display_day.columns: display_day[c] = ""
                display_day = display_day[target_order]
                
                display_day['OPEN ISSUE'] = display_day['OPEN ISSUE'].apply(split_issue_to_columns)
                
                def finalize_day_row(row):
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

                final_day_table = display_day.apply(finalize_day_row, axis=1)
                final_day_table.columns = pd.MultiIndex.from_tuples(multi_cols)
                
                def style_day_row(row):
                    styles = [''] * len(row)
                    idx = row.name
                    try:
                        if 0 < safe_float(base_day_df.loc[idx, '종합효율']) < safe_float(base_day_df.loc[idx, '목표효율']):
                            pos = row.index.get_loc(('생산성', '종합효율'))
                            if isinstance(pos, np.ndarray): pos = np.where(pos)[0][0]
                            styles[pos] = 'color: #FF4B4B; font-weight: bold;'
                    except: pass
                    return styles

                render_styler_to_html(final_day_table.style.apply(style_day_row, axis=1).hide(axis="index"), is_multi=True)
            else: st.info("분석할 생산일 데이터가 없습니다.")

        # =========================================================
        # TAB 4: 종합효율 BEST & WORST 분석 현황
        # =========================================================
        with tab4:
            st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 종합효율 BEST 5 & WORST 5 요인 분석</h3>", unsafe_allow_html=True)
            months_t4 = list(dict.fromkeys(f_df['생산월'].tolist())) if not f_df.empty else []
            sel_m_t4 = st.multiselect("📅 조회할 월(Month)을 선택하세요 (기본값: 최근 월)", options=months_t4, default=[months_t4[-1]] if months_t4 else [], key='t4_m')
            tab4_df = f_df[f_df['생산월'].isin(sel_m_t4)].copy() if sel_m_t4 else f_df.copy()
            
            valid_df = tab4_df[tab4_df['종합효율'] > 0].copy()
            if not valid_df.empty:
                st.markdown("<h4 style='color: #1F77B4; margin-top: 20px; font-weight: 800;'>🏆 BEST 5</h4>", unsafe_allow_html=True)
                best5_df = valid_df.sort_values(by=['종합효율', '생산일'], ascending=[False, False]).head(5)
                best5_display = best5_df[['생산일', '설비명', '품명', '종합효율', 'OPEN ISSUE']].reset_index(drop=True)
                best5_display['종합효율'] = best5_display['종합효율'].apply(lambda x: f"{safe_float(x):.1%}")
                best5_display['OPEN ISSUE'] = best5_display['OPEN ISSUE'].apply(split_issue_to_columns)
                
                def style_best_row(row):
                    styles = [''] * len(row)
                    try:
                        pos = row.index.get_loc('종합효율')
                        if isinstance(pos, np.ndarray): pos = np.where(pos)[0][0]
                        styles[pos] = 'color: #1F77B4; font-weight: bold;'
                    except: pass
                    return styles
                
                render_styler_to_html(best5_display.style.apply(style_best_row, axis=1).hide(axis="index"), is_multi=False)
                
                st.markdown("<h4 style='color: #FF4B4B; margin-top: 40px; font-weight: 800;'>🚨 WORST 5</h4>", unsafe_allow_html=True)
                worst5_df = valid_df.sort_values(by=['종합효율', '생산일'], ascending=[True, False]).head(5)
                worst5_display = worst5_df[['생산일', '설비명', '품명', '종합효율', 'OPEN ISSUE']].reset_index(drop=True)
                worst5_display['종합효율'] = worst5_display['종합효율'].apply(lambda x: f"{safe_float(x):.1%}")
                worst5_display['OPEN ISSUE'] = worst5_display['OPEN ISSUE'].apply(split_issue_to_columns)
                
                def style_worst_row(row):
                    styles = [''] * len(row)
                    try:
                        pos = row.index.get_loc('종합효율')
                        if isinstance(pos, np.ndarray): pos = np.where(pos)[0][0]
                        styles[pos] = 'color: #FF4B4B; font-weight: bold;'
                    except: pass
                    return styles
                
                render_styler_to_html(worst5_display.style.apply(style_worst_row, axis=1).hide(axis="index"), is_multi=False)
            else: st.info("해당 월에 분석할 가동 데이터가 없습니다.")

        # =========================================================
        # TAB 5: 비가동시간 BEST & WORST 분석 현황
        # =========================================================
        with tab5:
            st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 비가동시간 요인 정밀 분석</h3>", unsafe_allow_html=True)
            months_t5 = list(dict.fromkeys(f_df['생산월'].tolist())) if not f_df.empty else []
            sel_m_t5 = st.multiselect("📅 조회할 월(Month)을 선택하세요 (기본값: 최근 월)", options=months_t5, default=[months_t5[-1]] if months_t5 else [], key='t5_m')
            tab5_df = f_df[f_df['생산월'].isin(sel_m_t5)].copy() if sel_m_t5 else f_df.copy()
            
            valid_dt_df = tab5_df.copy()
            if not valid_dt_df.empty:
                st.markdown("<h4 style='color: #20C997; margin-top: 20px; font-weight: 800;'>🏆 최소 비가동 BEST 5</h4>", unsafe_allow_html=True)
                best5_dt = valid_dt_df.sort_values(by=['비가동시간', '종합효율'], ascending=[True, False]).head(5)
                best5_dt_display = best5_dt[['생산일', '설비명', '품명', '비가동시간', 'OPEN ISSUE']].reset_index(drop=True)
                best5_dt_display['비가동시간'] = best5_dt_display['비가동시간'].apply(lambda x: f"{safe_float(x):.1f}시간")
                best5_dt_display['OPEN ISSUE'] = best5_dt_display['OPEN ISSUE'].apply(split_issue_to_columns)
                
                def style_best_dt_row(row):
                    styles = [''] * len(row)
                    try:
                        pos = row.index.get_loc('비가동시간')
                        if isinstance(pos, np.ndarray): pos = np.where(pos)[0][0]
                        styles[pos] = 'color: #20C997; font-weight: bold;'
                    except: pass
                    return styles
                
                render_styler_to_html(best5_dt_display.style.apply(style_best_dt_row, axis=1).hide(axis="index"), is_multi=False)
                
                st.markdown("<h4 style='color: #E07A5F; margin-top: 40px; font-weight: 800;'>🚨 최대 비가동 WORST 5</h4>", unsafe_allow_html=True)
                worst5_dt = valid_dt_df.sort_values(by=['비가동시간', '종합효율'], ascending=[False, True]).head(5)
                worst5_dt_display = worst5_dt[['생산일', '설비명', '품명', '비가동시간', 'OPEN ISSUE']].reset_index(drop=True)
                worst5_dt_display['비가동시간'] = worst5_dt_display['비가동시간'].apply(lambda x: f"{safe_float(x):.1f}시간")
                worst5_dt_display['OPEN ISSUE'] = worst5_dt_display['OPEN ISSUE'].apply(split_issue_to_columns)
                
                def style_worst_dt_row(row):
                    styles = [''] * len(row)
                    try:
                        pos = row.index.get_loc('비가동시간')
                        if isinstance(pos, np.ndarray): pos = np.where(pos)[0][0]
                        styles[pos] = 'color: #E07A5F; font-weight: bold;'
                    except: pass
                    return styles
                
                render_styler_to_html(worst5_dt_display.style.apply(style_worst_dt_row, axis=1).hide(axis="index"), is_multi=False)
            else: st.info("해당 월에 분석할 비가동 데이터가 없습니다.")

        # =========================================================
        # TAB 6: 효율 급변(급증/급감) 구간 분석
        # =========================================================
        with tab6:
            st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 개별 설비 및 품목 기준 효율 급변(급증/급감) 정밀 추적</h3>", unsafe_allow_html=True)
            months_t6 = list(dict.fromkeys(f_df['생산월'].tolist())) if not f_df.empty else []
            sel_m_t6 = st.multiselect("📅 조회할 월(Month)을 선택하세요 (기본값: 최근 월)", options=months_t6, default=[months_t6[-1]] if months_t6 else [], key='t6_m')
            tab6_df = f_df[f_df['생산월'].isin(sel_m_t6)].copy() if sel_m_t6 else f_df.copy()
            
            detail_trend = tab6_df[tab6_df['종합효율'] > 0].sort_values(by=['설비명', '품명', 'sort_key']).copy()
            
            if not detail_trend.empty and len(detail_trend) > 1:
                detail_trend['전일대비_변동폭'] = detail_trend.groupby(['설비명', '품명'])['종합효율'].diff()
                detail_trend['이전_생산일'] = detail_trend.groupby(['설비명', '품명'])['생산일'].shift(1)
                detail_trend['이전_종합효율'] = detail_trend.groupby(['설비명', '품명'])['종합효율'].shift(1)
                
                diff_df = detail_trend.dropna(subset=['전일대비_변동폭']).copy()
                
                if not diff_df.empty:
                    machine_drops = diff_df.sort_values(by='전일대비_변동폭', ascending=True).head(5)
                    machine_surges = diff_df.sort_values(by='전일대비_변동폭', ascending=False).head(5)
                    
                    def get_diff_styler(df_input, is_drop=True):
                        disp = df_input[['설비명', '품명', '이전_생산일', '이전_종합효율', '생산일', '종합효율', '전일대비_변동폭', 'OPEN ISSUE']].copy()
                        disp['OPEN ISSUE'] = disp['OPEN ISSUE'].apply(split_issue_to_columns)
                        disp['이전_종합효율'] = disp['이전_종합효율'].apply(lambda x: f"{safe_float(x):.1%}")
                        disp['종합효율'] = disp['종합효율'].apply(lambda x: f"{safe_float(x):.1%}")
                        sign = "▼" if is_drop else "▲"
                        color_code = "#FF4B4B" if is_drop else "#1F77B4"
                        disp['변동폭'] = disp['전일대비_변동폭'].apply(lambda x: f"{safe_float(x):+.1%} {sign}")
                        disp = disp[['설비명', '품명', '이전_생산일', '이전_종합효율', '생산일', '종합효율', '변동폭', 'OPEN ISSUE']]
                        disp.columns = ['설비명', '품명', '이전 생산일', '이전 효율', '현재 생산일', '현재 효율', '변동폭', '해당일자 OPEN ISSUE (원인)']

                        style_df = pd.DataFrame('', index=disp.index, columns=disp.columns)
                        col_idx = style_df.columns.get_loc('변동폭')
                        if isinstance(col_idx, np.ndarray): col_idx = np.where(col_idx)[0][0]
                        for i in range(len(disp)):
                            style_df.iat[i, col_idx] = f'color: {color_code}; font-weight: bold;'
                        return disp.style.apply(lambda _: style_df, axis=None).hide(axis="index")

                    st.markdown("<h4 style='color: #FF4B4B; margin-top: 20px; font-weight: 800;'>📉 효율 급락 (최악의 하락폭 TOP 5)</h4>", unsafe_allow_html=True)
                    render_styler_to_html(get_diff_styler(machine_drops, is_drop=True), is_multi=False)
                    
                    st.markdown("<h4 style='color: #1F77B4; margin-top: 20px; font-weight: 800;'>📈 효율 급등 (최고의 상승폭 TOP 5)</h4>", unsafe_allow_html=True)
                    render_styler_to_html(get_diff_styler(machine_surges, is_drop=False), is_multi=False)
                    
                else: st.info("동일한 설비/품목이 2일 이상 연속으로 생산된 데이터가 없어 변동폭을 계산할 수 없습니다.")
            else: st.info("해당 월에 비교할 가동 데이터가 부족합니다.")

        # =========================================================
        # TAB 7: AI 데이터 챗봇 (Beta)
        # =========================================================
        with tab7:
            st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #1F77B4;'>■</span> 🤖 AI 생산 데이터 챗봇 (Beta)</h3>", unsafe_allow_html=True)
            st.markdown("사이드바에서 **선택한 조건(생산월, 생산일, 설비 등)의 데이터**를 바탕으로 챗봇이 답변합니다. 무엇이든 물어보세요!")
            
            api_key = st.text_input("🔑 OpenAI API Key를 입력하세요 (현재는 테스트 모드입니다)", type="password")
            
            if "messages" not in st.session_state:
                st.session_state.messages = [{"role": "assistant", "content": "안녕하세요! 사출생산팀 AI 어시스턴트입니다. 왼쪽 필터가 적용된 현재 데이터를 바탕으로 분석해 드립니다. 무엇이 궁금하신가요?"}]
                
            for msg in st.session_state.messages:
                st.chat_message(msg["role"]).write(msg["content"])
                
            if prompt := st.chat_input("질문을 입력하세요... (예: 종합효율이 80% 이하인 설비 알려줘)"):
                st.session_state.messages.append({"role": "user", "content": prompt})
                st.chat_message("user").write(prompt)
                
                if not api_key:
                    dummy_response = f"💡 **[체험 모드]** API 키가 연결되지 않았습니다.\n\n나중에 API 키를 연결하시면, 시스템이 스스로 데이터를 분석하여 **'{prompt}'**에 대한 정확한 답변을 찾아줄 것입니다.\n\n> *예시 답변: 질문하신 기간 내 종합효율 80% 이하 설비는 총 2대(50호기, 21호기)이며, 주요 비가동 사유는 로보트 알람 및 히터 단선입니다.*"
                    st.session_state.messages.append({"role": "assistant", "content": dummy_response})
                    st.chat_message("assistant").write(dummy_response)
                else:
                    try:
                        from pandasai import SmartDataframe
                        from pandasai.llm.openai import OpenAI
                        llm = OpenAI(api_token=api_key)
                        sdf = SmartDataframe(f_df, config={"llm": llm})
                        response = sdf.chat(prompt)
                        st.session_state.messages.append({"role": "assistant", "content": str(response)})
                        st.chat_message("assistant").write(str(response))
                    except Exception as e:
                        err_msg = f"❌ 분석 중 오류가 발생했습니다. (API 키 오류 또는 'pandasai' 라이브러리 설치가 필요합니다)\n\n상세 오류: {e}"
                        st.session_state.messages.append({"role": "assistant", "content": err_msg})
                        st.chat_message("assistant").write(err_msg)

else:
    st.info("데이터 파일이 없습니다. GitHub의 'data' 폴더에 엑셀/CSV 파일을 넣거나 직접 파일을 업로드해주세요.")
