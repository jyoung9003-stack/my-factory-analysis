import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import re
import os
import numpy as np
import textwrap
from collections import Counter
from datetime import datetime

# ==========================================
# 🌟 [와이드 스크린 에디션] 테마 설정
# ==========================================
st.set_page_config(
    page_title="사출생산팀 일일 생산성 정밀 분석", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 크롬 자동 번역 강제 중지 스크립트
components.html(
    """
    <script>
        const parent = window.parent.document;
        parent.documentElement.lang = 'ko';
        parent.documentElement.setAttribute('translate', 'no');
        parent.body.classList.add('notranslate');
        const meta = parent.createElement('meta');
        meta.name = "google";
        meta.content = "notranslate";
        parent.head.appendChild(meta);
    </script>
    """,
    width=0, height=0
)

st.markdown("""
<meta name="google" content="notranslate">
<style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.8/dist/web/static/pretendard.css');
    
    html, body, [class*="css"] {
        font-family: 'Pretendard', sans-serif !important;
        background-color: #FFFFFF; 
        color: #1E293B; 
        translate: no; 
    }

    .section-banner {
        background-color: #ffffff;
        border: 1px solid #F1F5F9;
        border-left: 6px solid #D91B1B; 
        padding: 16px 22px;
        border-radius: 10px;
        margin-top: 30px;
        margin-bottom: 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }
    .section-banner h3 {
        margin: 0;
        font-weight: 800;
        color: #1E293B;
        font-size: 19px;
    }

    .analysis-report-card {
        background-color: #FFF5F5; 
        border: 1px solid #FEE2E2; 
        border-radius: 12px;
        padding: 20px 25px;
        margin-bottom: 25px;
        box-shadow: 0 2px 4px rgba(217, 27, 27, 0.05);
    }

    .trendy-card {
        background-color: #FFFFFF;
        border-radius: 12px;
        padding: 24px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        border: 1px solid #F1F5F9;
        margin-bottom: 24px;
    }
    
    .metric-card-container {
        background-color: #FFFFFF;
        border-radius: 12px;
        padding: 20px;
        text-align: center;
        border: 1px solid #F1F5F9;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.03);
    }
    .metric-title {
        font-size: 13px;
        color: #64748B; 
        margin-bottom: 10px;
        font-weight: 500;
    }
    .metric-value-box {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 6px;
    }
    .metric-value {
        font-size: 32px;
        font-weight: 800;
        letter-spacing: -1px;
    }
    .metric-icon {
        font-size: 18px;
    }

    .dashboard-header {
        background: linear-gradient(135deg, #9F1239, #D91B1B); 
        color: #FFFFFF;
        border-radius: 12px;
        padding: 25px;
        margin-bottom: 25px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 🌟 [함수 사전 정의 구역]
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

def render_section_title(text):
    st.markdown(f"<div class='section-banner'><h3>{text}</h3></div>", unsafe_allow_html=True)

def render_tab_insight(title, content):
    st.markdown(f"""
    <div class='analysis-report-card'>
        <h4 style='margin-top:0; color:#9F1239; font-weight:800; font-size:16px; margin-bottom:12px;'>{title}</h4>
        <div style='line-height:1.7; font-size:14.5px; color:#334155; margin-bottom:0;'>
            {content}
        </div>
    </div>
    """, unsafe_allow_html=True)

def get_status_color(oee, tgt=0.86):
    color = "#3B82F6" if oee >= tgt else "#D91B1B" 
    bg = "#EBF5FF" if oee >= tgt else "#FFF5F5"
    icon = "✓" if oee >= tgt else "⚠"
    return color, bg, icon

def render_trendy_metric(title, value_str, color, icon):
    st.markdown(f"""
    <div class="metric-card-container">
        <div class="metric-title">{title}</div>
        <div class="metric-value-box">
            <span class="metric-value" style="color: {color};">{value_str}</span>
            <span class="metric-icon" style="color: {color};">{icon}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

def get_natural_issue_summary(issue_text):
    if not issue_text or str(issue_text).strip() in ['', 'nan', 'None', '0', '0.0']:
        return "별도의 특이사항 없음"
    text = str(issue_text)
    text = re.sub(r'[\*→\-\.※○●]', ' ', text)
    text = re.sub(r'\b(주간|야간|주, 야간|주야간|상기\s*\d+건|해당|발생)\b', ' ', text)
    phrases = [p.strip() for p in text.split('\n') if p.strip()]
    valid_phrases = []
    for p in phrases:
        clean_p = re.sub(r'\s+', ' ', p).strip()
        if len(clean_p) > 3: valid_phrases.append(clean_p)
    if not valid_phrases: return "명확히 기록되지 않은 원인"
    if len(valid_phrases) == 1: return f"{valid_phrases[0]}"
    elif len(valid_phrases) == 2: return f"{valid_phrases[0]} 및 {valid_phrases[1]}"
    else: return f"{valid_phrases[0]}, {valid_phrases[1]} 등 복합적 요인"

def split_issue_to_columns(issue_text):
    lines = [line.strip() for line in str(issue_text).split('\n') if line.strip()]
    if not lines: return "<div style='font-size:12px; color:#ADB5BD; background-color:#F8FAFC; padding:8px; border-radius:4px; border:1px dashed #E9ECEF;'>📝 특이사항 없음</div>"
    day_lines, night_lines, general_lines = [], [], []
    has_shift = False
    curr = general_lines
    for line in lines:
        cl = line.replace(' ', '')
        if '*주간' in cl or line.startswith('주간'): curr = day_lines; has_shift = True; line = re.sub(r'^\*?\s*주간\s*', '', line).strip()
        elif '*야간' in cl or line.startswith('야간'): curr = night_lines; has_shift = True; line = re.sub(r'^\*?\s*야간\s*', '', line).strip()
        if line: curr.append(line)
    if not has_shift: return f"<div style='font-size:13px; color:#495057;'>{'<br>'.join(lines)}</div>"
    d_h = '<br>'.join(general_lines + day_lines) if (general_lines + day_lines) else "없음"
    n_h = '<br>'.join(night_lines) if night_lines else "없음"
    return f"<div style='display: flex; gap: 8px; margin-top: 5px;'><div style='flex: 1; background-color: #F8FAFC; border: 1px solid #E9ECEF; border-radius: 6px; padding: 10px; border-top: 3px solid #FBBF24;'><div style='font-size:11px; font-weight:bold; color:#B45309;'>☀️ 주간</div><div style='font-size:13px;'>{d_h}</div></div><div style='flex: 1; background-color: #F8FAFC; border: 1px solid #E9ECEF; border-radius: 6px; padding: 10px; border-top: 3px solid #1E293B;'><div style='font-size:11px; font-weight:bold; color:#1E293B;'>🌙 야간</div><div style='font-size:13px;'>{n_h}</div></div></div>"

def format_issue(text):
    val = str(text).strip()
    if val in ['', '0', '0.0', 'nan', 'None']: return ""
    val = val.replace('\r\n', '\n')
    val = re.sub(r'(?<!\n)\*', '\n*', val)
    val = re.sub(r'(?<!\n)-\.', '\n-.', val)
    val = re.sub(r'(?<!\n)→', '\n→ ', val)
    return val.strip()

def render_styler_to_html(styler, is_multi=False):
    try: html_str = styler.to_html(escape=False)
    except: html_str = styler.to_html()
    html_str = html_str.replace('<table', '<table class="custom-table notranslate"')
    wrapped_html = f"""<div style='width: 100%; max-height: 500px; overflow: auto; border: 1px solid #E2E8F0; border-radius: 8px; box-shadow: 0 1px 3px 0 rgba(0,0,0,0.03); background-color: white; margin-bottom: 24px;'><style>.custom-table {{ width: 100%; border-collapse: collapse; font-size: 13px; color: #1E293B; background-color: white; }}.custom-table th {{ background-color: #F8FAFC; border: 1px solid #E2E8F0; padding: 12px 16px; text-align: center !important; font-weight: 600; color: #475569; position: sticky; top: 0; z-index: 2; }}.custom-table thead tr:nth-child(2) th {{ top: 40px; }}.custom-table td {{ border: 1px solid #F1F5F9; padding: 10px 16px; text-align: center !important; }}.custom-table td:last-child {{ text-align: left !important; min-width: 450px; line-height: 1.5; }} .custom-table tr:hover {{ background-color: #F8FAFC; }}</style>{html_str}</div>"""
    if is_multi:
        wrapped_html = re.sub(r'<th class=\"col_heading level0 col10\".*?>OPEN ISSUE</th>', r'<th class=\"col_heading level0 col10\" rowspan=\"2\" style=\"vertical-align: middle;\">OPEN ISSUE</th>', wrapped_html)
        wrapped_html = re.sub(r'<th class=\"col_heading level1 col10\".*?>OPEN ISSUE</th>', '', wrapped_html)
    st.markdown(wrapped_html, unsafe_allow_html=True)


# ==========================================
# 🌟 데이터 로드 구역
# ==========================================
target_cols = ['생산일', '설비명', '품명', '양품수량', '불량수량', '총 생산수량', '투입시간', '가동시간', '비가동시간', '정미시간', '양품율', '성능가동율', '시간가동율', '종합효율', '목표효율', 'OPEN ISSUE']
target_order = ['생산일', '설비명', '품명', '종합효율', '양품율', '성능가동율', '시간가동율', '총 생산수량', '양품수량', '불량수량', 'OPEN ISSUE']
multi_cols = [
    ('구분', '생산일'), ('구분', '설비명'), ('구분', '품명'),
    ('생산성', '종합효율'), ('생산성', '양품율'), ('생산성', '성능가동율'), ('생산성', '시간가동율'),
    ('생산실적', '총 생산수량'), ('생산실적', '양품수량'), ('생산실적', '불량수량'),
    ('OPEN ISSUE', 'OPEN ISSUE')
]

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
            except Exception as e: st.error(f"데이터 읽기 오류: {e}")

if data_to_process:
    all_records = []
    daily_totals_data = {} 
    
    for file_name, temp_df in data_to_process:
        temp_df.columns = [str(c).replace('\n', '').replace('\r', '').strip() for c in temp_df.columns]
        name_map = {'작업장 [설비]': '설비명', '작업장[설비]': '설비명', '품목명': '품명', '합계': '총 생산수량', '합게수량': '총 생산수량', '종합 효율': '종합효율', '목표 효율': '목표효율'}
        temp_df = temp_df.rename(columns=name_map)
        for col in temp_df.columns:
            if 'Unnamed' in col or 'ISSUE' in col.upper(): temp_df = temp_df.rename(columns={col: 'OPEN ISSUE'}); break
        
        date_match = re.search(r'\d{8}', file_name)
        if date_match:
            raw_date = date_match.group()[2:]
            try:
                dt = datetime.strptime(raw_date, '%y%m%d')
                week_arr = ['월', '화', '수', '목', '금', '토', '일']
                clean_date = f"{dt.strftime('%y')}년 {dt.month}월 {dt.day}일 ({week_arr[dt.weekday()]})"
                month_str = f"{dt.strftime('%y')}년 {dt.month}월"
                sort_key = raw_date 
            except: clean_date = raw_date; month_str = "분류 안됨"; sort_key = raw_date
        else: clean_date = file_name.split('.')[0]; month_str = "분류 안됨"; sort_key = clean_date
        
        d_total_oee = 0.0
        for _, row in temp_df.iterrows():
            row_str = "".join([str(v).replace(' ', '').upper() for v in row.values])
            if any(kw in row_str for kw in ['TOTAL', '합계', '총합', '전체', '총계', '평균']):
                val = row.get('종합효율', 0.0)
                if pd.notna(val) and str(val).strip() != '':
                    d_total_oee = safe_float(val)
                    if d_total_oee > 0: break
        
        if d_total_oee == 0.0:
            try: 
                last_val = safe_float(temp_df['종합효율'].iloc[-1])
                if last_val > 0: d_total_oee = last_val
            except: pass
            
        if d_total_oee == 0.0:
            try:
                valid_oees = [safe_float(x) for x in temp_df['종합효율'] if safe_float(x) > 0]
                if valid_oees: d_total_oee = sum(valid_oees) / len(valid_oees)
            except: pass

        if sort_key not in daily_totals_data:
            daily_totals_data[sort_key] = {'생산일': clean_date, '생산월': month_str, '공장종합효율': d_total_oee}

        for _, row in temp_df.iterrows():
            m_val = str(row.get('설비명', '')).strip()
            if m_val in ['', 'nan', '설비명'] or 'TOTAL' in m_val.upper() or '합계' in m_val: continue
            record = {'sort_key': sort_key, '생산월': month_str, '생산일': clean_date}
            for col in target_cols:
                if col != '생산일': record[col] = row[col] if col in temp_df.columns else None
            all_records.append(record)

    df = pd.DataFrame(all_records).sort_values(by='sort_key').reset_index(drop=True)
    date_mapping = dict(zip(df['생산일'], df['sort_key']))
    daily_df = pd.DataFrame([{'sort_key': k, **v} for k, v in daily_totals_data.items()]).sort_values(by='sort_key').reset_index(drop=True)
    
    for col in ['양품수량', '불량수량', '총 생산수량', '투입시간', '가동시간', '비가동시간', '정미시간', '종합효율', '목표효율', '양품율', '성능가동율', '시간가동율']:
        if col in df.columns: df[col] = df[col].apply(safe_float)

    df['OPEN ISSUE'] = df['OPEN ISSUE'].apply(format_issue)

    # =========================================================
    # 🌟 우측 상단 배치 필터 레이아웃
    # =========================================================
    header_col, filter_col = st.columns([1, 2.5])
    
    with header_col:
        # 메인 타이틀 변경
        st.markdown("<h1 style='margin-top: 15px; margin-bottom: 10px; color: #1E293B; font-weight: 900; font-size: 30px;' class='notranslate'>사출생산팀 생산성 및 OPEN ISSUE 분석 리포트</h1>", unsafe_allow_html=True)

    with filter_col:
        st.markdown("<div style='margin-top: 25px;'></div>", unsafe_allow_html=True) 
        f1, f2, f3, f4 = st.columns(4)
        
        all_months = [m for m in df['생산월'].unique() if str(m).strip() != ""]
        # 멀티셀렉트 Placeholder 통일 (전체 생산월 등)
        with f1: sel_m_side = st.multiselect("📅 생산월", all_months, default=[all_months[-1]] if all_months else [], placeholder="전체 생산월")
        m_f_df = df[df['생산월'].isin(sel_m_side)].copy() if sel_m_side else df.copy()
        
        all_dates = list(m_f_df['생산일'].unique())
        all_dates.sort(key=lambda x: date_mapping.get(x, ""), reverse=True)
        with f2: sel_d_side = st.multiselect("📆 생산일", all_dates, default=[all_dates[0]] if all_dates else [], placeholder="전체 생산일")
        d_f_df = m_f_df[m_f_df['생산일'].isin(sel_d_side)].copy() if sel_d_side else m_f_df.copy()
        
        all_machines = sorted([m for m in d_f_df['설비명'].unique() if m.strip() != ""])
        with f3: sel_mach_side = st.multiselect("⚙️ 설비", all_machines, placeholder="전체 설비")
        pool_df = d_f_df[d_f_df['설비명'].isin(sel_mach_side)].copy() if sel_mach_side else d_f_df.copy()
        
        actual_prods = sorted([p for p in pool_df['품명'].fillna("").astype(str).str.strip().unique() if p not in ["", "0", "0.0", "nan", "NaN", "None"]])
        with f4: sel_prod = st.selectbox("📦 품목", ["전체 품목"] + actual_prods)
        f_df = pool_df[pool_df['품명'].str.strip() == sel_prod].copy() if sel_prod != "전체 품목" else pool_df.copy()

    st.markdown("<div style='margin-bottom: 20px;'></div>", unsafe_allow_html=True)
    
    # 🌟 탭 명칭 완벽 적용
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📈 사출생산팀 종합효율 추이", 
        "📝 OPEN ISSUE 현황", 
        "📅 설비 가동 현황 및 생산성, 비가동 분석", 
        "🏆 종합효율 BEST&WORST", 
        "🛑 비가동 BEST", 
        "🤖 사출생산팀 생산성 AI 챗봇"
    ])

    # =========================================================
    # TAB 1: 종합 효율 추이
    # =========================================================
    with tab1:
        tab1_df = m_f_df.copy()
        if sel_mach_side: tab1_df = tab1_df[tab1_df['설비명'].isin(sel_mach_side)]
        if sel_prod != "전체 품목": tab1_df = tab1_df[tab1_df['품명'].str.strip() == sel_prod]

        is_fac = (not sel_mach_side and sel_prod == "전체 품목")
        
        if is_fac: 
            p_df = daily_df[daily_df['생산월'].isin(sel_m_side)].copy() if sel_m_side else daily_df.copy()
            y_v = '공장종합효율'
        else: 
            act_oee = tab1_df[tab1_df['종합효율'] > 0]
            p_df = act_oee.groupby(['sort_key', '생산월', '생산일'])[['종합효율']].mean(numeric_only=True).reset_index().sort_values('sort_key')
            y_v = '종합효율'
            
        render_section_title("최근 5일 종합효율 요약")
        if not p_df.empty:
            r5 = p_df.sort_values('sort_key').tail(5)
            r_cols = st.columns(5)
            for i, (_, r) in enumerate(r5.iterrows()):
                oee = safe_float(r[y_v])
                clr, _, icn = get_status_color(oee, 0.86)
                with r_cols[i]: 
                    render_trendy_metric(r['생산일'], f"{oee:.1%}", clr, icn)
        
        render_section_title("월별 종합효율 추이")
        mons = list(dict.fromkeys(p_df['생산월'].tolist()))
        sel_mons = st.multiselect("📅 조회할 월 선택", mons, default=[mons[-1]] if mons else [], key='t1_m', placeholder="전체 생산월")
        
        for m in sel_mons:
            m_p = p_df[p_df['생산월'] == m].copy()
            if m_p.empty: continue
            
            avg_oee = m_p[y_v].mean()
            max_row = m_p.loc[m_p[y_v].idxmax()]
            min_row = m_p.loc[m_p[y_v].idxmin()]
            c_text = f"<b>{m}</b> 기준 <b>평균 종합효율은 {avg_oee:.1%}</b>입니다.<br>최고 생산일은 <b>{max_row['생산일']} ({max_row[y_v]:.1%})</b>, 최저 생산일은 <b>{min_row['생산일']} ({min_row[y_v]:.1%})</b>로 확인됩니다."
            render_tab_insight(f"📊 {m} 생산성 요약", c_text)
            
            text_colors = ['#3B82F6' if safe_float(row[y_v]) >= 0.86 else '#D91B1B' for _, row in m_p.iterrows()]
            fig_oee = go.Figure()
            fig_oee.add_trace(go.Bar(x=m_p['생산일'], y=m_p[y_v], text=m_p[y_v].apply(lambda x: f"{x:.1%}"), textposition='auto', marker_color=text_colors))
            fig_oee.update_layout(title=dict(text=f"📊 {m} 추이 차트", font=dict(size=16, weight=800)), plot_bgcolor='rgba(0,0,0,0)', height=350, yaxis=dict(tickformat='.0%', range=[0, 1.0]))
            st.plotly_chart(fig_oee, use_container_width=True)

    # =========================================================
    # TAB 2: OPEN ISSUE 정밀 조회
    # =========================================================
    with tab2:
        render_section_title("OPEN ISSUE 현황")
        mons2 = list(dict.fromkeys(f_df['생산월'].tolist()))
        sel_m2 = st.multiselect("📅 조회할 월 선택", mons2, default=[mons2[-1]] if mons2 else [], key='t2_m', placeholder="전체 생산월")
        t2_df = f_df[f_df['생산월'].isin(sel_m2)].copy()
        
        all_d2 = list(t2_df['생산일'].unique())
        all_d2.sort(key=lambda x: date_mapping.get(x, ""), reverse=True)
        
        if all_d2:
            sd2 = st.selectbox("조회할 일자", all_d2, key='tab2_date')
            issue_df = t2_df[t2_df['생산일'] == sd2].copy().sort_values(by='설비명').reset_index(drop=True)
            if not issue_df.empty:
                issue_disp = issue_df[['생산일', '설비명', '품명', '종합효율', 'OPEN ISSUE']].copy()
                for idx, row in issue_disp.iterrows():
                    if str(row['품명']).strip() in ['', 'nan', '0', '0.0']: 
                        issue_disp.at[idx, '품명'] = ""
                        issue_disp.at[idx, '종합효율'] = ""
                    else: 
                        issue_disp.at[idx, '종합효율'] = f"{safe_float(row['종합효율']):.1%}"
                issue_disp['OPEN ISSUE'] = issue_disp['OPEN ISSUE'].apply(lambda x: split_issue_to_columns(x))
                render_styler_to_html(issue_disp.style.hide(axis="index"))

    # =========================================================
    # TAB 3: 일일 상세 현황
    # =========================================================
    with tab3:
        render_section_title("일일 생산성 현황")
        
        all_d3 = list(f_df['생산일'].unique())
        all_d3.sort(key=lambda x: date_mapping.get(x, ""), reverse=True)
        
        if all_d3:
            sd3 = st.selectbox("📅 조회할 일자", all_d3, key='tab3_date')
            day_df = f_df[f_df['생산일'] == sd3].copy().sort_values(by='설비명').reset_index(drop=True)
            day_df['설비_짧은명'] = day_df['설비명'].apply(lambda x: str(x).split(' - ')[0].strip())
            active_day = day_df[day_df['종합효율'] > 0].sort_values(by='종합효율', ascending=False)
            
            if not active_day.empty:
                day_total_val = daily_df[daily_df['생산일'] == sd3]['공장종합효율'].iloc[0] if not daily_df[daily_df['생산일'] == sd3].empty else active_day['종합효율'].mean(numeric_only=True)
                worst_r = active_day.iloc[-1]
                w_issue_sum = get_natural_issue_summary(worst_r['OPEN ISSUE'])
                
                c_text3 = f"<b>{sd3}</b> 전체 설비 <b>합계 종합효율은 {day_total_val:.1%}</b>입니다.<br>효율이 가장 저조한 <b>{str(worst_r['설비명']).split(' - ')[0]} ({worst_r['품명']}, {worst_r['종합효율']:.1%})</b>는 <b><span style='color:#D91B1B;'>[{w_issue_sum}]</span></b> 문제가 핵심 원인입니다."
                render_tab_insight(f"📊 {sd3} 일일 가동 총평", c_text3)

                best_html = "".join([f"<div style='margin-bottom:10px;'><b>{i+1}. {str(r['설비명']).split(' - ')[0]}</b> <span style='float:right; font-weight:bold;'>{r['종합효율']:.1%}</span><br><span style='font-size:12px; opacity:0.8;'>{r['품명']}</span></div>" for i, (_, r) in enumerate(active_day.head(5).iterrows())])
                worst_html = "".join([f"<div style='margin-bottom:10px;'><b>{i+1}. {str(r['설비명']).split(' - ')[0]}</b> <span style='float:right; font-weight:bold;'>{r['종합효율']:.1%}</span><br><span style='font-size:12px; opacity:0.8;'>{r['품명']}</span></div>" for i, (_, r) in enumerate(active_day.tail(5).sort_values(by='종합효율').iterrows())])
                
                active_count = active_day['설비명'].nunique()
                total_count = day_df['설비명'].nunique()
                
                st.markdown(f"""<div class='dashboard-header'><div style='font-size: 16px; opacity: 0.9; margin-bottom: 5px; font-weight: 500;'>💡 {sd3} 생산 요약</div><div style='font-size: 32px; font-weight: 900; margin-bottom: 25px; letter-spacing: -1.5px;'>총 <span style='color: #E2E8F0;'>{total_count}</span>대 중 <span style='color: #FBBF24;'>{active_count}</span>대 가동 중</div><div style='display: flex; gap: 20px;'><div style='flex: 1; background: rgba(255,255,255,0.1); padding: 20px; border-radius: 10px;'><div style='color: #4ADE80; font-weight: 800; font-size: 16px; margin-bottom: 15px;'>🏆 종합효율 BEST 5</div>{best_html}</div><div style='flex: 1; background: rgba(255,255,255,0.1); padding: 20px; border-radius: 10px;'><div style='color: #FFAAAA; font-weight: 800; font-size: 16px; margin-bottom: 15px;'>🚨 종합효율 WORST 5</div>{worst_html}</div></div></div>""", unsafe_allow_html=True)
                
                st.markdown(f"#### 📊 {sd3} 종합효율 비교")
                bar_clrs = ['#3B82F6' if safe_float(row['종합효율']) >= 0.86 else '#D91B1B' for _, row in active_day.iterrows()]
                fig3 = px.bar(active_day, x='설비_짧은명', y='종합효율', text_auto='.1%')
                fig3.update_traces(marker_color=bar_clrs, textfont=dict(weight="bold", color='white'))
                fig3.update_layout(plot_bgcolor='rgba(0,0,0,0)', height=400, yaxis=dict(tickformat='.0%', range=[0, 1.0]), xaxis_title="")
                st.plotly_chart(fig3, use_container_width=True)
                
                st.write("---")
                st.markdown(f"#### 🛑 {sd3} 비가동 시간 비교")
                downtime_day = active_day.sort_values(by='비가동시간', ascending=False)
                fig_dt = px.bar(downtime_day[downtime_day['비가동시간']>0], x='설비_짧은명', y='비가동시간', text_auto='.1f')
                fig_dt.update_traces(marker_color='#EF4444')
                fig_dt.update_layout(plot_bgcolor='rgba(0,0,0,0)', height=350, yaxis_title="시간(h)", xaxis_title="")
                st.plotly_chart(fig_dt, use_container_width=True)
                
                disp_day = day_df[target_order].copy()
                for idx, row in disp_day.iterrows():
                    prod = str(row['품명']).strip()
                    if prod in ['', 'nan', '0', '0.0']:
                        disp_day.at[idx, '품명'] = ""
                        for c in ['종합효율', '양품율', '성능가동율', '시간가동율', '양품수량', '불량수량', '총 생산수량']: 
                            if c in disp_day.columns: disp_day.at[idx, c] = ""
                    else:
                        for c in ['종합효율', '양품율', '성능가동율', '시간가동율']: 
                            if c in disp_day.columns: disp_day.at[idx, c] = f"{safe_float(row[c]):.1%}"
                        for c in ['양품수량', '불량수량', '총 생산수량']: 
                            if c in disp_day.columns: disp_day.at[idx, c] = f"{int(safe_float(row[c])):,}"
                
                disp_day['OPEN ISSUE'] = disp_day['OPEN ISSUE'].apply(lambda x: split_issue_to_columns(x))
                disp_day.columns = pd.MultiIndex.from_tuples(multi_cols)
                
                def style_day_row(row):
                    styles = [''] * len(row)
                    idx = row.name
                    try:
                        if 0 < safe_float(day_df.loc[idx, '종합효율']) < safe_float(day_df.loc[idx, '목표효율']):
                            pos = row.index.get_loc(('생산성', '종합효율'))
                            if isinstance(pos, np.ndarray): pos = np.where(pos)[0][0]
                            styles[pos] = 'color: #D91B1B; font-weight: 800;' 
                    except: pass
                    return styles
                
                render_styler_to_html(disp_day.style.apply(style_day_row, axis=1).hide(axis="index"), is_multi=True)

    # =========================================================
    # TAB 4: BEST & WORST
    # =========================================================
    with tab4:
        # 🌟 수식어 제거 및 타이틀 간소화
        render_section_title("종합효율 BEST 5 & WORST 5")
        mons4 = list(dict.fromkeys(f_df['생산월'].tolist()))
        sel_m4 = st.multiselect("📅 조회할 월 선택", mons4, default=[mons4[-1]] if mons4 else [], key='t4_m', placeholder="전체 생산월")
        t4_df = f_df[(f_df['생산월'].isin(sel_m4)) & (f_df['종합효율'] > 0)].copy()
        if not t4_df.empty:
            w5 = t4_df.sort_values(by='종합효율').head(5)
            w_details = "".join([f"📍 <b>{rw['생산일']}</b> - <b>{str(rw['설비명']).split(' - ')[0]}</b> ({rw['품명']}, {rw['종합효율']:.1%}) ➔ <span style='color:#D91B1B;'>{get_natural_issue_summary(rw['OPEN ISSUE'])}</span><br>" for _, rw in w5.iterrows()])
            render_tab_insight("📊 집중관리 대상", f"기간 내 하위 설비군 주요 이슈입니다:<br><div style='background-color:rgba(217,27,27,0.03); padding:15px; border-radius:8px; margin-top:10px; border-left:4px solid #D91B1B; line-height: 1.7;'>{w_details}</div>")
            
            for label, asc in [("🏆 BEST 5", False), ("🚨 WORST 5", True)]:
                st.markdown(f"<h4 style='font-weight: 800; color: #1E293B; margin-top: 15px;'>{label}</h4>", unsafe_allow_html=True)
                res = t4_df.sort_values(by='종합효율', ascending=asc).head(5)
                res_disp = res[['생산일', '설비명', '품명', '종합효율', 'OPEN ISSUE']].copy()
                for idx, row in res_disp.iterrows():
                    if str(row['품명']).strip() in ['', 'nan']: 
                        res_disp.at[idx, '품명'] = ""
                        res_disp.at[idx, '종합효율'] = ""
                    else: 
                        res_disp.at[idx, '종합효율'] = f"{safe_float(row['종합효율']):.1%}"
                res_disp['OPEN ISSUE'] = res_disp['OPEN ISSUE'].apply(lambda x: split_issue_to_columns(x))
                render_styler_to_html(res_disp.style.hide(axis="index"))

    # =========================================================
    # TAB 5: 비가동 정밀 분석
    # =========================================================
    with tab5:
        # 🌟 수식어 제거 및 타이틀 간소화
        render_section_title("비가동시간 WORST 현황")
        mons5 = list(dict.fromkeys(f_df['생산월'].tolist()))
        sel_m5 = st.multiselect("📅 조회할 월 선택", mons5, default=[mons5[-1]] if mons5 else [], key='t5_m', placeholder="전체 생산월")
        t5_df = f_df[f_df['생산월'].isin(sel_m5)].copy()
        if not t5_df.empty:
            w_dt = t5_df.sort_values(by='비가동시간', ascending=False).head(10)
            w_dt_details = "".join([f"🛑 <b>{rw['생산일']}</b> - <b>{str(rw['설비명']).split(' - ')[0]}</b> ({rw['품명']}, {rw['비가동시간']:.1f}h) ➔ <span style='color:#D91B1B;'>{get_natural_issue_summary(rw['OPEN ISSUE'])}</span><br>" for _, rw in w_dt.head(3).iterrows()])
            render_tab_insight("🛑 핵심 비가동 원인", f"최장 비가동 상위 설비 및 원인입니다:<br><div style='background-color:rgba(217,27,27,0.03); padding:15px; border-radius:8px; margin-top:10px; border-left:4px solid #D91B1B; line-height: 1.7;'>{w_dt_details}</div>")
            
            # 🌟 괄호 제거 및 간소화
            st.markdown("<h4 style='font-weight: 800; color: #1E293B;'>🚨 WORST 10</h4>", unsafe_allow_html=True)
            res_disp = w_dt[['생산일', '설비명', '품명', '비가동시간', 'OPEN ISSUE']].copy()
            for idx, row in res_disp.iterrows():
                if str(row['품명']).strip() in ['', 'nan', '0', '0.0']: res_disp.at[idx, '품명'] = ""
            res_disp['비가동시간'] = res_disp['비가동시간'].apply(lambda x: f"{safe_float(x):.1f}h")
            res_disp['OPEN ISSUE'] = res_disp['OPEN ISSUE'].apply(lambda x: split_issue_to_columns(x))
            render_styler_to_html(res_disp.style.hide(axis="index"))

    # =========================================================
    # TAB 6: 챗봇
    # =========================================================
    with tab6:
        render_section_title("🤖 AI 생산 데이터 챗봇")
        ak = st.text_input("🔑 OpenAI API Key 입력", type="password")
        if "msgs" not in st.session_state: st.session_state.msgs = [{"role": "assistant", "content": "사출생산팀 데이터 분석을 도와드리는 AI 챗봇입니다."}]
        for m in st.session_state.msgs: st.chat_message(m["role"]).write(m["content"])
        if pr := st.chat_input("데이터에 관해 질문하세요"):
            st.session_state.msgs.append({"role": "user", "content": pr}); st.chat_message("user").write(pr)
            if not ak: st.chat_message("assistant").write("💡 API 키를 입력해 주세요.")
            else: st.chat_message("assistant").write("데이터를 분석 중입니다...")

else: st.info("GitHub data 폴더에 CSV 파일을 넣어주세요.")
