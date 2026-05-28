import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
import os
from datetime import datetime

# ==========================================
# 🌟 1. 기본 설정 및 테마 (현장 가시성 극대화)
# ==========================================
st.set_page_config(page_title="설비별 정밀 분석 대시보드", layout="wide", initial_sidebar_state="collapsed")

components.html(
    """<script>
        const parent = window.parent.document;
        parent.documentElement.lang = 'ko';
        parent.documentElement.setAttribute('translate', 'no');
        parent.body.classList.add('notranslate');
        const meta = parent.createElement('meta');
        meta.name = "google"; meta.content = "notranslate";
        parent.head.appendChild(meta);
    </script>""", width=0, height=0
)

st.markdown("""
<meta name="google" content="notranslate">
<style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.8/dist/web/static/pretendard.css');
    html, body, [class*="css"] { font-family: 'Pretendard', sans-serif !important; background-color: #F8FAFC; color: #0F172A; translate: no; }
    
    /* 섹션 배너 디자인 (더 진하게) */
    .section-banner { background-color: #ffffff; border: 1px solid #E2E8F0; border-left: 8px solid #D91B1B; padding: 18px 24px; border-radius: 12px; margin-top: 35px; margin-bottom: 25px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); }
    .section-banner h3 { margin: 0; font-weight: 900; color: #0F172A; font-size: 22px; letter-spacing: -0.5px; }
    
    /* 인포그래픽 요약 카드 */
    .analysis-report-card { background-color: #F1F5F9; border-left: 5px solid #3B82F6; border-radius: 8px; padding: 20px 25px; margin-bottom: 25px; }
    
    /* 핵심 지표 (Metric) 카드 디자인 강화 */
    .metric-card-container { background-color: #FFFFFF; border-radius: 16px; padding: 25px 20px; text-align: center; border: 1px solid #E2E8F0; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); transition: transform 0.2s; }
    .metric-card-container:hover { transform: translateY(-2px); box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1); }
    .metric-title { font-size: 16px; color: #475569; margin-bottom: 12px; font-weight: 700; }
    .metric-value-box { display: flex; align-items: center; justify-content: center; gap: 8px; }
    .metric-value { font-size: 42px; font-weight: 900; letter-spacing: -1.5px; line-height: 1; }
    .metric-icon { font-size: 24px; }
    
    /* 터치 친화적 큼직한 설비 선택 버튼 */
    div.stButton > button {
        width: 100%;
        height: 75px;
        background-color: #FFFFFF;
        border: 2px solid #CBD5E1;
        color: #1E293B;
        font-size: 20px !important;
        font-weight: 800 !important;
        border-radius: 12px;
        transition: all 0.2s ease-in-out;
        box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
    }
    div.stButton > button:hover { border-color: #3B82F6; color: #1D4ED8; background-color: #EFF6FF; transform: scale(1.02); }
    div.stButton > button:active { background-color: #2563EB !important; color: white !important; border-color: #2563EB; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 🌟 2. 함수 정의
# ==========================================
def safe_float(val):
    try:
        if isinstance(val, pd.Series): val = val.iloc[0]
        if pd.isna(val) or val is None: return 0.0
        v_str = str(val).strip().replace(',', '').replace(' ', '')
        if v_str in ['', '-', '#DIV/0!', '#N/A', 'nan', 'None']: return 0.0
        if '%' in v_str: return float(v_str.replace('%', '')) / 100.0
        return float(v_str)
    except: return 0.0

def render_section_title(text): st.markdown(f"<div class='section-banner'><h3>{text}</h3></div>", unsafe_allow_html=True)
def render_tab_insight(title, content): st.markdown(f"<div class='analysis-report-card'><h4 style='margin-top:0; color:#1E293B; font-weight:800; font-size:17px; margin-bottom:10px;'>{title}</h4><div style='line-height:1.6; font-size:15px; color:#334155; margin-bottom:0;'>{content}</div></div>", unsafe_allow_html=True)
def render_trendy_metric(title, value_str, color, icon): st.markdown(f"<div class='metric-card-container'><div class='metric-title'>{title}</div><div class='metric-value-box'><span class='metric-value' style='color: {color};'>{value_str}</span><span class='metric-icon' style='color: {color};'>{icon}</span></div></div>", unsafe_allow_html=True)

def split_issue_to_columns(issue_text):
    lines = [line.strip() for line in str(issue_text).split('\n') if line.strip()]
    if not lines or str(issue_text).strip() in ['nan', '0', '0.0', 'None']: return "<div style='font-size:13px; color:#94A3B8; padding:8px; font-weight: 500;'>✔ 특이사항 없음</div>"
    d_l, n_l, g_l = [], [], []; has_s = False; curr = g_l
    for line in lines:
        cl = line.replace(' ', '')
        if '*주간' in cl or line.startswith('주간'): curr = d_l; has_s = True; line = re.sub(r'^\*?\s*주간\s*', '', line).strip()
        elif '*야간' in cl or line.startswith('야간'): curr = n_l; has_s = True; line = re.sub(r'^\*?\s*야간\s*', '', line).strip()
        if line: curr.append(line)
    if not has_s: return f"<div style='font-size:14px; font-weight: 600; color:#334155;'>{'<br>'.join(lines)}</div>"
    d_h = '<br>'.join(g_l + d_l) if (g_l + d_l) else "-"; n_h = '<br>'.join(n_l) if n_l else "-"
    return f"<div style='display: flex; gap: 8px;'><div style='flex: 1; background-color: #FFFBEB; padding: 10px; border-radius: 6px; border-top: 3px solid #F59E0B;'><div style='font-size:12px; font-weight:900; color:#B45309; margin-bottom:4px;'>☀️ 주간</div><div style='font-size:13px; font-weight:600;'>{d_h}</div></div><div style='flex: 1; background-color: #F1F5F9; padding: 10px; border-radius: 6px; border-top: 3px solid #334155;'><div style='font-size:12px; font-weight:900; color:#1E293B; margin-bottom:4px;'>🌙 야간</div><div style='font-size:13px; font-weight:600;'>{n_h}</div></div></div>"

def format_issue(text):
    val = str(text).strip()
    if val in ['', '0', '0.0', 'nan', 'None']: return ""
    val = val.replace('\r\n', '\n'); val = re.sub(r'(?<!\n)\*', '\n*', val); val = re.sub(r'(?<!\n)-\.', '\n-.', val); val = re.sub(r'(?<!\n)→', '\n→ ', val)
    return val.strip()

def render_styler_to_html(styler):
    try: html_str = styler.to_html(escape=False)
    except: html_str = styler.to_html()
    # 인포그래픽 스타일의 깔끔하고 대비가 강한 표 디자인 적용
    html_str = html_str.replace('<table', '<table style="width: 100%; border-collapse: collapse; font-size: 14px; text-align: center; background-color: white;"')
    html_str = html_str.replace('<th', '<th style="background-color: #1E293B; color: #FFFFFF; border: 1px solid #334155; padding: 14px; font-weight: 700; font-size: 15px;"')
    html_str = html_str.replace('<td', '<td style="border: 1px solid #E2E8F0; padding: 12px; font-weight: 500; color: #334155;"')
    st.markdown(f"<div style='width:100%; overflow-x:auto; border:1px solid #CBD5E1; border-radius:10px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); margin-bottom:25px;'>{html_str}</div>", unsafe_allow_html=True)

# ==========================================
# 🌟 3. 데이터 로드 
# ==========================================
target_cols = ['생산일', '설비명', '품명', '양품수량', '불량수량', '총 생산수량', '투입시간', '가동시간', '비가동시간', '종합효율', '양품율', 'OPEN ISSUE']
data_to_process = []
DATA_DIR = "data"

if os.path.exists(DATA_DIR):
    for file_name in os.listdir(DATA_DIR):
        if file_name.startswith("~$"): continue 
        if file_name.endswith(('.xlsx', '.csv')):
            file_path = os.path.join(DATA_DIR, file_name)
            try:
                if file_name.endswith('.csv'): 
                    try: t_df = pd.read_csv(file_path, encoding='utf-8')
                    except: t_df = pd.read_csv(file_path, encoding='cp949')
                else: t_df = pd.read_excel(file_path)
                data_to_process.append((file_name, t_df))
            except: pass

if data_to_process:
    all_records = []
    daily_totals_data = {} 
    week_arr = ['월', '화', '수', '목', '금', '토', '일']
    
    for file_name, temp_df in data_to_process:
        temp_df.columns = [str(c).replace('\n', '').replace('\r', '').strip() for c in temp_df.columns]
        temp_df = temp_df.rename(columns={'작업장 [설비]': '설비명', '작업장[설비]': '설비명', '품목명': '품명', '합계': '총 생산수량', '합게수량': '총 생산수량', '종합 효율': '종합효율'})
        for col in temp_df.columns:
            if 'Unnamed' in col or 'ISSUE' in col.upper(): temp_df = temp_df.rename(columns={col: 'OPEN ISSUE'}); break
        
        date_match = re.search(r'\d{8}', file_name)
        if date_match:
            raw_date = date_match.group()[2:]
            try:
                dt = datetime.strptime(raw_date, '%y%m%d')
                clean_date = f"{dt.strftime('%y')}년 {dt.month}월 {dt.day}일 ({week_arr[dt.weekday()]})"
                month_str = f"{dt.strftime('%y')}년 {dt.month}월"; sort_key = raw_date 
            except: clean_date = raw_date; month_str = "분류 안됨"; sort_key = raw_date
        else: clean_date = file_name.split('.')[0]; month_str = "분류 안됨"; sort_key = clean_date
        
        d_total_oee, daily_val, fallback_val, last_valid_val = 0.0, 0.0, 0.0, 0.0
        for _, row in temp_df.iterrows():
            row_str = "".join([str(v).replace(' ', '').upper() for v in row.values])
            val = safe_float(row.get('종합효율', 0.0))
            if val > 0:
                if any(kw in row_str for kw in ['당월', '월간', '누계', 'MONTH']): continue
                if any(kw in row_str for kw in ['금일', '당일', 'TODAY', 'DAILY']): daily_val = val
                elif any(kw in row_str for kw in ['TOTAL', '합계', '총합', '전체', '총계']): fallback_val = val
                last_valid_val = val
        if daily_val > 0: d_total_oee = daily_val
        elif fallback_val > 0: d_total_oee = fallback_val
        else: d_total_oee = last_valid_val
            
        if sort_key not in daily_totals_data: daily_totals_data[sort_key] = {'생산일': clean_date, '생산월': month_str, '공장종합효율': d_total_oee}

        for _, row in temp_df.iterrows():
            m_val = str(row.get('설비명')).strip()
            if m_val.lower() in ['', 'nan', 'none', '#n/a'] or any(kw in m_val.upper() for kw in ['TOTAL', '합계']): continue
            record = {'sort_key': sort_key, '생산월': month_str, '생산일': clean_date}
            for col in target_cols:
                if col != '생산일': record[col] = row[col] if col in temp_df.columns else None
            all_records.append(record)

    df = pd.DataFrame(all_records).sort_values(by='sort_key').reset_index(drop=True)
    date_mapping = dict(zip(df['생산일'], df['sort_key']))
    daily_df = pd.DataFrame([{'sort_key': k, **v} for k, v in daily_totals_data.items()]).sort_values(by='sort_key').reset_index(drop=True)
    for col in ['양품수량', '불량수량', '총 생산수량', '투입시간', '가동시간', '비가동시간', '종합효율', '양품율']:
        if col in df.columns: df[col] = df[col].apply(safe_float)
    df['OPEN ISSUE'] = df['OPEN ISSUE'].apply(format_issue)

    # =========================================================
    # 🌟 4. 레이아웃 및 필터
    # =========================================================
    title_col1, title_col2 = st.columns([0.6, 9.4])
    with title_col1:
        try: st.image("logo.png", width=80) 
        except: st.markdown("<div style='font-size: 50px; margin-top: 5px;'>🏭</div>", unsafe_allow_html=True)
    with title_col2: st.markdown("<h1 style='margin-top: 15px; margin-bottom: 25px; font-weight: 900; font-size: 34px; color: #0F172A;'>사출생산팀 현장 설비 모니터링</h1>", unsafe_allow_html=True)

    f1, f2 = st.columns(2)
    all_months = [m for m in df['생산월'].unique() if str(m).strip() != ""]
    with f1: sel_m_side = st.multiselect("📅 조회할 월 선택", all_months, default=[all_months[-1]] if all_months else [])
    
    m_f_df = df[df['생산월'].isin(sel_m_side)].copy() if sel_m_side else df.copy()
    all_dates = list(m_f_df['생산일'].unique())
    all_dates.sort(key=lambda x: date_mapping.get(x, ""), reverse=True)
    with f2: sel_d_side = st.multiselect("📆 특정 일자만 집중 분석 (선택 안하면 월 전체)", all_dates, default=[])
    f_df = m_f_df[m_f_df['생산일'].isin(sel_d_side)].copy() if sel_d_side else m_f_df.copy()

    st.markdown("<div style='margin-bottom: 25px;'></div>", unsafe_allow_html=True)
    
    # 🌟 5. 메인 탭 설정 
    tab1, tab2 = st.tabs(["📈 사출생산팀 종합효율 추이", "🎯 설비별 정밀 분석 (버튼 터치)"])

    # -----------------------------------------------------
    # TAB 1: 종합 효율 추이 
    # -----------------------------------------------------
    with tab1:
        p_df = daily_df[daily_df['생산월'].isin(sel_m_side)].copy() if sel_m_side else daily_df.copy()
        if sel_d_side: p_df = p_df[p_df['생산일'].isin(sel_d_side)]
            
        render_section_title("공장 전체 종합효율(OEE) 추이")
        if not p_df.empty:
            avg_oee = p_df['공장종합효율'].mean()
            render_tab_insight("💡 현장 운영 가이드", f"조회하신 기간 동안 사출 공정의 평균 OEE는 <b><span style='color:#3B82F6; font-size:18px;'>{avg_oee:.1%}</span></b>를 기록했습니다. 86.0% 목표 달성을 위해 아래 설비별 정밀 분석 탭에서 취약 설비의 비가동 요인을 점검해 주십시오.")
            
            bar_colors = ['#3B82F6' if safe_float(row['공장종합효율']) >= 0.86 else '#EF4444' for _, row in p_df.iterrows()]
            fig_oee = go.Figure(go.Bar(x=p_df['생산일'], y=p_df['공장종합효율'], text=p_df['공장종합효율'].apply(lambda x: f"{x:.1%}"), textposition='auto', marker_color=bar_colors, textfont=dict(size=14, weight='bold', color='white')))
            fig_oee.update_layout(plot_bgcolor='rgba(0,0,0,0)', height=450, yaxis=dict(tickformat='.0%', range=[0, 1.0]), margin=dict(t=20))
            st.plotly_chart(fig_oee, use_container_width=True)
        else: st.info("조건에 해당하는 데이터가 없습니다.")

    # -----------------------------------------------------
    # TAB 2: 설비별 정밀 분석 (인포그래픽 스타일)
    # -----------------------------------------------------
    with tab2:
        render_section_title("👆 분석할 설비 번호를 터치하세요")
        
        machine_list = sorted([m for m in f_df['설비명'].unique() if m and str(m).strip() != 'nan'])
        
        if machine_list:
            if 'selected_mach' not in st.session_state:
                st.session_state.selected_mach = machine_list[0]
                
            cols = st.columns(6)
            for i, mach in enumerate(machine_list):
                short_name = mach.split(' - ')[0].strip()
                if cols[i % 6].button(short_name, key=f"btn_{mach}"):
                    st.session_state.selected_mach = mach

            tgt_mach = st.session_state.selected_mach
            st.markdown(f"<div style='background-color:#1E293B; color:white; padding:20px; border-radius:12px; margin-top:30px; margin-bottom:20px; text-align:center;'><h2 style='margin:0; font-weight:900;'>💻 {tgt_mach} 집중 분석 리포트</h2></div>", unsafe_allow_html=True)
            
            t7_df = f_df[f_df['설비명'] == tgt_mach].copy().sort_values('sort_key')
            
            if not t7_df.empty:
                avg_oee = t7_df['종합효율'].apply(safe_float).mean()
                total_down = t7_df['비가동시간'].apply(safe_float).sum()
                issue_count = t7_df['OPEN ISSUE'].apply(lambda x: 0 if str(x).strip() in ['', 'nan', '0', '0.0'] else 1).sum()
                
                c1, c2, c3 = st.columns(3)
                with c1: render_trendy_metric("기간 내 평균 OEE", f"{avg_oee:.1%}", "#2563EB" if avg_oee >= 0.86 else "#DC2626", "📈")
                with c2: render_trendy_metric("누적 비가동 손실", f"{total_down:.1f}h", "#DC2626" if total_down > 0 else "#059669", "🛑")
                with c3: render_trendy_metric("이슈 발생 일수", f"{issue_count}일", "#D97706" if issue_count > 0 else "#059669", "📝")
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                chart_col, table_col = st.columns([1, 1.3])
                
                with chart_col:
                    st.markdown("<h4 style='font-weight:800; color:#0F172A; margin-bottom:15px;'>📊 일자별 OEE 흐름도</h4>", unsafe_allow_html=True)
                    fig7 = go.Figure(go.Scatter(
                        x=t7_df['생산일'], y=t7_df['종합효율'], mode='lines+markers+text',
                        text=t7_df['종합효율'].apply(lambda x: f"{x:.1%}"), textposition="top center",
                        line=dict(color='#2563EB', width=4), marker=dict(size=12, color='#0F172A'), textfont=dict(size=14, weight='bold')
                    ))
                    fig7.update_layout(plot_bgcolor='rgba(0,0,0,0)', height=400, yaxis=dict(tickformat='.0%', range=[0, 1.0]), margin=dict(l=0, r=0, t=10, b=0))
                    st.plotly_chart(fig7, use_container_width=True)
                
                with table_col:
                    st.markdown("<h4 style='font-weight:800; color:#0F172A; margin-bottom:15px;'>📋 세부 조업 실적 및 이슈 이력</h4>", unsafe_allow_html=True)
                    disp_t7 = t7_df[['생산일', '품명', '종합효율', '비가동시간', '총 생산수량', 'OPEN ISSUE']].copy()
                    
                    for idx, row in disp_t7.iterrows():
                        disp_t7.at[idx, '종합효율'] = f"{safe_float(row['종합효율']):.1%}"
                        disp_t7.at[idx, '총 생산수량'] = f"{int(safe_float(row['총 생산수량'])):,}"
                        disp_t7.at[idx, '비가동시간'] = f"{safe_float(row['비가동시간']):.1f}h"
                        
                    disp_t7['OPEN ISSUE'] = disp_t7['OPEN ISSUE'].apply(split_issue_to_columns)
                    render_styler_to_html(disp_t7.style.hide(axis="index"))
                    
            else: st.warning("해당 설비는 선택한 기간 내 가동 데이터가 없습니다.")
        else: st.info("분석할 설비 데이터가 존재하지 않습니다.")

else: st.info("GitHub data 폴더에 CSV/Excel 파일을 넣어주세요.")
